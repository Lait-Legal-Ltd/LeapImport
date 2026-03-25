using ClosedXML.Excel;
using LeapMergeDoc.Models;
using MySql.Data.MySqlClient;
using System.Globalization;
using System.IO;

namespace LeapMergeDoc.Services
{
    public class JournalEntryImportService
    {
        private readonly string _connectionString;
        private readonly Action<string> _logAction;

        // Fixed transaction date: 31/12/2025
        private readonly DateTime _transactionDate = new DateTime(2025, 12, 31, 17, 49, 47);

        public JournalEntryImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _logAction = logAction;
        }

        public List<JournalExcelData> ReadExcelData(string filePath)
        {
            var data = new List<JournalExcelData>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();

                if (range == null)
                {
                    _logAction("No data found in Excel file.");
                    return data;
                }

                var rowCount = range.RowCount();
                var colCount = range.ColumnCount();

                _logAction($"Found {rowCount - 1} data rows in Excel file.");

                // Read headers
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cell(1, col).GetString().Trim());
                }

                _logAction($"Headers: {string.Join(", ", headers)}");

                // Normalize headers for matching
                var normalizedHeaders = headers.Select(h => h.ToLower().Replace("_", "").Replace(" ", "").Replace(".", "")).ToList();

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new JournalExcelData();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetString().Trim();
                        var header = normalizedHeaders[col - 1];

                        switch (header)
                        {
                            case "matter":
                                rowData.Matter = cellValue;
                                break;
                            case "client":
                                rowData.Client = cellValue;
                                break;
                            case "matterdescription":
                                rowData.MatterDescription = cellValue;
                                break;
                            case "lasttransdate":
                                rowData.LastTransDate = ParseDate(cellValue);
                                break;
                            case "amount":
                                if (decimal.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amount))
                                {
                                    rowData.Amount = amount;
                                }
                                break;
                        }
                    }

                    // Only include rows with Matter and Amount
                    if (!string.IsNullOrEmpty(rowData.Matter))
                    {
                        data.Add(rowData);
                    }
                }
            }

            return data;
        }

        private DateTime? ParseDate(string? dateString)
        {
            if (string.IsNullOrEmpty(dateString))
                return null;

            string[] formats = {
                "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd",
                "dd-MM-yyyy", "MM-dd-yyyy", "d/M/yyyy",
                "M/d/yyyy", "yyyy/MM/dd"
            };

            foreach (string format in formats)
            {
                if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                {
                    return result;
                }
            }

            if (DateTime.TryParse(dateString, out DateTime generalResult))
            {
                return generalResult;
            }

            return null;
        }

        public (List<JournalImportData> records, JournalImportSummary summary) ProcessExcelData(List<JournalExcelData> excelData)
        {
            var records = new List<JournalImportData>();
            var summary = new JournalImportSummary();

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                int lineNumber = 1;

                foreach (var row in excelData)
                {
                    var caseInfo = GetCaseIdByReference(connection, row.Matter!);

                    var importData = new JournalImportData
                    {
                        CaseId = caseInfo?.caseId,
                        LedgerCardId = caseInfo?.ledgerCardId,
                        CaseReference = row.Matter,
                        ClientName = row.Client ?? "Unknown Client",
                        Balance = row.Amount,
                        Description = $"Opening Balance - {row.Matter}",
                        AccountType = "case",
                        LineNumber = lineNumber++,
                        IsFound = caseInfo != null
                    };

                    records.Add(importData);

                    summary.TotalRecords++;
                    summary.TotalAmount += row.Amount;

                    if (caseInfo != null)
                    {
                        summary.FoundCases++;
                        summary.FoundAmount += row.Amount;
                    }
                    else
                    {
                        summary.NotFoundCases++;
                        summary.NotFoundAmount += row.Amount;
                        _logAction($"⚠️ Case not found: '{row.Matter}'");
                    }
                }
            }

            return (records, summary);
        }

        private (int caseId, int ledgerCardId)? GetCaseIdByReference(MySqlConnection connection, string caseReference)
        {
            try
            {
                string sql = @"
                    SELECT c.case_id, COALESCE(l.ledger_card_id, c.case_id) as ledger_card_id
                    FROM tbl_case_details_general c
                    LEFT JOIN tbl_acc_ledger_cards l ON l.fk_case_id = c.case_id
                    WHERE c.case_reference_auto = @caseReference 
                       OR c.case_reference_manual = @caseReference
                    LIMIT 1";

                using (var cmd = new MySqlCommand(sql, connection))
                {
                    cmd.Parameters.AddWithValue("@caseReference", caseReference);

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return (reader.GetInt32("case_id"), reader.GetInt32("ledger_card_id"));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction($"Error looking up case reference '{caseReference}': {ex.Message}");
            }

            return null;
        }

        public (int success, int errors) ImportJournalEntries(List<JournalImportData> records)
        {
            int successCount = 0;
            int errorCount = 0;

            // Filter only found cases
            var foundRecords = records.Where(r => r.IsFound).ToList();

            if (foundRecords.Count == 0)
            {
                _logAction("No valid records to import.");
                return (0, 0);
            }

            // Calculate total amount
            decimal totalAmount = foundRecords.Sum(r => r.Balance);

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // 1. Insert Transaction
                        int transactionId = InsertTransaction(connection, transaction, totalAmount);
                        _logAction($"Created transaction ID: {transactionId}");

                        // 2. Get next journal entry number
                        int journalEntryNumber = GetNextJournalEntryNumber(connection, transaction);

                        // 3. Insert Journal Entry Header
                        int journalEntryId = InsertJournalEntryHeader(connection, transaction, transactionId, journalEntryNumber, totalAmount);
                        _logAction($"Created journal entry ID: {journalEntryId}, Number: {journalEntryNumber}");

                        // 4. Insert Bank Line (Debit)
                        InsertJournalEntryLine(connection, transaction, journalEntryId, 1, "bank", 1,
                            totalAmount, 0, "Opening Balance");

                        // 5. Insert Bank Transaction
                        InsertClientBankTransaction(connection, transaction, transactionId, journalEntryNumber, totalAmount);

                        int lineNumber = 2;

                        // 6. Insert Case Lines (Credit) and Ledger Card Transactions
                        foreach (var data in foundRecords)
                        {
                            try
                            {
                                // Insert journal entry line (Credit)
                                InsertJournalEntryLine(connection, transaction, journalEntryId, lineNumber++,
                                    "ledger", data.LedgerCardId ?? data.CaseId!.Value, 0, data.Balance,
                                    $"Opening Balance - {data.CaseReference}");

                                // Insert ledger card transaction
                                InsertLedgerCardTransaction(connection, transaction, data, transactionId, journalEntryNumber);

                                successCount++;
                                _logAction($"✅ Imported: {data.CaseReference} - {data.Balance:C}");
                            }
                            catch (Exception ex)
                            {
                                errorCount++;
                                _logAction($"❌ Error importing {data.CaseReference}: {ex.Message}");
                            }
                        }

                        transaction.Commit();
                        _logAction("Transaction committed successfully.");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logAction($"Transaction rolled back: {ex.Message}");
                        throw;
                    }
                }
            }

            return (successCount, errorCount);
        }

        private int InsertTransaction(MySqlConnection connection, MySqlTransaction transaction, decimal totalAmount)
        {
            string sql = @"
                INSERT INTO tbl_acc_transactions (
                    fk_branch_id, fk_transaction_type_id, fk_transaction_sub_type_id, 
                    transaction_details, transaction_reference, transaction_amount,
                    is_cancelled, post_by, post_date_time, transaction_date, transaction_type
                ) VALUES (
                    @branchId, @typeId, @subTypeId, @details, @reference, @amount,
                    @isCancelled, @postBy, @postDateTime, @transactionDate, @transactionType
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@branchId", 1);
                cmd.Parameters.AddWithValue("@typeId", 1);
                cmd.Parameters.AddWithValue("@subTypeId", 4);
                cmd.Parameters.AddWithValue("@details", "Opening Balance");
                cmd.Parameters.AddWithValue("@reference", "JRN 1");
                cmd.Parameters.AddWithValue("@amount", Math.Abs(totalAmount));
                cmd.Parameters.AddWithValue("@isCancelled", false);
                cmd.Parameters.AddWithValue("@postBy", 1);
                cmd.Parameters.AddWithValue("@postDateTime", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@transactionDate", _transactionDate);
                cmd.Parameters.AddWithValue("@transactionType", "Add");

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private int GetNextJournalEntryNumber(MySqlConnection connection, MySqlTransaction transaction)
        {
            string sql = "SELECT COALESCE(MAX(journal_entry_number), 0) + 1 FROM tbl_acc_journal_entry";
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private int InsertJournalEntryHeader(MySqlConnection connection, MySqlTransaction transaction,
            int transactionId, int journalEntryNumber, decimal totalAmount)
        {
            string sql = @"
                INSERT INTO tbl_acc_journal_entry (
                    fk_branch_id, fk_transaction_id, journal_entry_number, reference,
                    journal_entry_description, journal_entry_date, current_date_time,
                    staff_id, total, is_canceled
                ) VALUES (
                    @branchId, @transactionId, @journalNumber, @reference,
                    @description, @date, @datetime, @staffId, @total, @isCanceled
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@branchId", 1);
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@journalNumber", journalEntryNumber);
                cmd.Parameters.AddWithValue("@reference", $"JRN-{journalEntryNumber}");
                cmd.Parameters.AddWithValue("@description", "Opening Balance");
                cmd.Parameters.AddWithValue("@date", _transactionDate);
                cmd.Parameters.AddWithValue("@datetime", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@staffId", 1);
                cmd.Parameters.AddWithValue("@total", Math.Abs(totalAmount));
                cmd.Parameters.AddWithValue("@isCanceled", false);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void InsertJournalEntryLine(MySqlConnection connection, MySqlTransaction transaction,
            int journalEntryId, int lineNumber, string accountType, int accountId,
            decimal debit, decimal credit, string description)
        {
            string sql = @"
                INSERT INTO tbl_acc_journal_entry_lines (
                    fk_journal_entry_id, fk_branch_id, line_number, account_type,
                    fk_account_id, description, debit, credit, date_time, fk_user_id, is_deleted
                ) VALUES (
                    @journalEntryId, @branchId, @lineNumber, @accountType,
                    @accountId, @description, @debit, @credit, @datetime, @userId, @isDeleted
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@journalEntryId", journalEntryId);
                cmd.Parameters.AddWithValue("@branchId", 1);
                cmd.Parameters.AddWithValue("@lineNumber", lineNumber);
                cmd.Parameters.AddWithValue("@accountType", accountType);
                cmd.Parameters.AddWithValue("@accountId", accountId);
                cmd.Parameters.AddWithValue("@description", description);
                cmd.Parameters.AddWithValue("@debit", debit);
                cmd.Parameters.AddWithValue("@credit", credit);
                cmd.Parameters.AddWithValue("@datetime", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@userId", 1);
                cmd.Parameters.AddWithValue("@isDeleted", false);

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertClientBankTransaction(MySqlConnection connection, MySqlTransaction transaction,
            int transactionId, int journalEntryNumber, decimal totalAmount)
        {
            string sql = @"
                INSERT INTO tbl_acc_client_bank_transactions (
                    fk_branch_id, fk_transaction_id, fk_client_bank_id, transaction_date_time,
                    transaction, reference, details, dr_amount, cr_amount,
                    balance_pre, balance_post, is_cancelled, is_reconciled, fk_bank_reconciliation_id
                ) VALUES (
                    @branchId, @transactionId, @clientBankId, @transactionDateTime,
                    @transaction, @reference, @details, @drAmount, @crAmount,
                    @balancePre, @balancePost, @isCancelled, @isReconciled, @bankReconciliationId
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@branchId", 1);
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@clientBankId", 1);
                cmd.Parameters.AddWithValue("@transactionDateTime", _transactionDate);
                cmd.Parameters.AddWithValue("@transaction", "Opening Balance");
                cmd.Parameters.AddWithValue("@reference", $"JRN-{journalEntryNumber}");
                cmd.Parameters.AddWithValue("@details", "Opening Balance");
                cmd.Parameters.AddWithValue("@drAmount", Math.Abs(totalAmount));
                cmd.Parameters.AddWithValue("@crAmount", 0);
                cmd.Parameters.AddWithValue("@balancePre", 0);
                cmd.Parameters.AddWithValue("@balancePost", totalAmount);
                cmd.Parameters.AddWithValue("@isCancelled", false);
                cmd.Parameters.AddWithValue("@isReconciled", false);
                cmd.Parameters.AddWithValue("@bankReconciliationId", DBNull.Value);

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertLedgerCardTransaction(MySqlConnection connection, MySqlTransaction transaction,
            JournalImportData data, int transactionId, int journalEntryNumber)
        {
            decimal clientBalance = Math.Abs(data.Balance);
            string clientBalanceType = data.Balance > 0 ? "CR" : "DR";

            string sql = @"
                INSERT INTO tbl_acc_ledger_card_transactions (
                    fk_branch_id, fk_transaction_id, fk_ledger_card_id, transaction_date_time,
                    details, office_dr, office_cr, office_bal, office_bal_type,
                    client_dr, client_cr, client_bal, client_bal_type, total,
                    ledger_reference, is_cancelled
                ) VALUES (
                    @branchId, @transactionId, @ledgerCardId, @transactionDateTime,
                    @details, @officeDr, @officeCr, @officeBal, @officeBalType,
                    @clientDr, @clientCr, @clientBal, @clientBalType, @total,
                    @ledgerReference, @isCancelled
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@branchId", 1);
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@ledgerCardId", data.LedgerCardId ?? data.CaseId);
                cmd.Parameters.AddWithValue("@transactionDateTime", _transactionDate);
                cmd.Parameters.AddWithValue("@details", "Opening Balance");
                cmd.Parameters.AddWithValue("@officeDr", 0);
                cmd.Parameters.AddWithValue("@officeCr", 0);
                cmd.Parameters.AddWithValue("@officeBal", 0);
                cmd.Parameters.AddWithValue("@officeBalType", "DR");
                cmd.Parameters.AddWithValue("@clientDr", 0);
                cmd.Parameters.AddWithValue("@clientCr", data.Balance);
                cmd.Parameters.AddWithValue("@clientBal", clientBalance);
                cmd.Parameters.AddWithValue("@clientBalType", clientBalanceType);
                cmd.Parameters.AddWithValue("@total", Math.Abs(data.Balance));
                cmd.Parameters.AddWithValue("@ledgerReference", $"JRN-{journalEntryNumber}");
                cmd.Parameters.AddWithValue("@isCancelled", false);

                cmd.ExecuteNonQuery();
            }
        }

        public (int rowsDeleted, string message) TruncateJournalData()
        {
            int totalDeleted = 0;

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                // Disable foreign key checks temporarily
                using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 0;", connection))
                {
                    cmd.ExecuteNonQuery();
                }

                try
                {
                    var tablesToTruncate = new[]
                    {
                        "tbl_acc_ledger_card_transactions",
                        "tbl_acc_client_bank_transactions",
                        "tbl_acc_journal_entry_lines",
                        "tbl_acc_journal_entry",
                        "tbl_acc_transactions"
                    };

                    foreach (var table in tablesToTruncate)
                    {
                        try
                        {
                            using (var cmd = new MySqlCommand($"DELETE FROM {table};", connection))
                            {
                                int deleted = cmd.ExecuteNonQuery();
                                totalDeleted += deleted;
                                _logAction($"Deleted {deleted} rows from {table}");
                            }
                        }
                        catch (MySqlException ex)
                        {
                            _logAction($"Warning: Could not delete from {table}: {ex.Message}");
                        }
                    }

                    // Reset auto-increment counters
                    foreach (var table in tablesToTruncate)
                    {
                        try
                        {
                            using (var cmd = new MySqlCommand($"ALTER TABLE {table} AUTO_INCREMENT = 1;", connection))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                        catch (MySqlException)
                        {
                            // Ignore if table doesn't have auto-increment
                        }
                    }
                }
                finally
                {
                    // Re-enable foreign key checks
                    using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 1;", connection))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            return (totalDeleted, $"Successfully deleted {totalDeleted} total rows from journal tables.");
        }
    }
}
