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

        // Fixed transaction date: 01/04/2026 05:00:00
        private readonly DateTime _transactionDate = new DateTime(2026, 4, 1, 5, 0, 0);

        public JournalEntryImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _logAction = logAction;
        }

        public List<JournalExcelData> ReadExcelData(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            
            if (extension == ".csv")
            {
                return ReadCsvData(filePath);
            }
            else
            {
                return ReadExcelFileData(filePath);
            }
        }

        private List<JournalExcelData> ReadCsvData(string filePath)
        {
            var data = new List<JournalExcelData>();
            var lines = File.ReadAllLines(filePath);

            if (lines.Length < 2)
            {
                _logAction("No data found in CSV file.");
                return data;
            }

            _logAction($"Found {lines.Length - 1} data rows in CSV file.");
            _logAction($"Headers: {lines[0]}");

            // CSV format: Client,Matter,Client Name,Matter Description,F/E,W/T,Client (balance)
            for (int i = 1; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrEmpty(line)) continue;

                var columns = ParseCsvLine(line);
                if (columns.Length < 7) continue;

                var rowData = new JournalExcelData
                {
                    ClientCode = columns[0].Trim(),           // Client code (e.g., "2DE0001")
                    Matter = columns[1].Trim(),               // Matter number (e.g., "1")
                    ClientName = columns[2].Trim(),           // Client Name
                    MatterDescription = columns[3].Trim(),    // Matter Description
                    FeeEarner = columns[4].Trim(),            // F/E
                    WorkType = columns[5].Trim(),             // W/T
                    Amount = ParseAmount(columns[6])          // Balance (last column)
                };

                // Only include rows with ClientCode and Matter
                if (!string.IsNullOrEmpty(rowData.ClientCode) && !string.IsNullOrEmpty(rowData.Matter))
                {
                    data.Add(rowData);
                }
            }

            return data;
        }

        private string[] ParseCsvLine(string line)
        {
            var result = new List<string>();
            var current = "";
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current);
                    current = "";
                }
                else
                {
                    current += c;
                }
            }
            result.Add(current);

            return result.ToArray();
        }

        private decimal ParseAmount(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;
            
            // Remove quotes, spaces, commas used as thousand separators
            var cleaned = value.Trim().Trim('"').Trim();
            cleaned = cleaned.Replace(",", "").Replace(" ", "");
            
            if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amount))
            {
                return amount;
            }
            return 0;
        }

        private List<JournalExcelData> ReadExcelFileData(string filePath)
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

                // Read data rows (assuming same column order as CSV)
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new JournalExcelData
                    {
                        ClientCode = worksheet.Cell(row, 1).GetString().Trim(),
                        Matter = worksheet.Cell(row, 2).GetString().Trim(),
                        ClientName = worksheet.Cell(row, 3).GetString().Trim(),
                        MatterDescription = worksheet.Cell(row, 4).GetString().Trim(),
                        FeeEarner = worksheet.Cell(row, 5).GetString().Trim(),
                        WorkType = worksheet.Cell(row, 6).GetString().Trim(),
                        Amount = ParseAmount(worksheet.Cell(row, 7).GetString())
                    };

                    if (!string.IsNullOrEmpty(rowData.ClientCode) && !string.IsNullOrEmpty(rowData.Matter))
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
                    // Use combined reference: ClientCode-Matter (e.g., "2DE0001-1")
                    var caseReference = row.CaseReference;
                    var caseInfo = GetCaseIdByReference(connection, caseReference);

                    var importData = new JournalImportData
                    {
                        CaseId = caseInfo?.caseId,
                        LedgerCardId = caseInfo?.ledgerCardId,
                        CaseReference = caseReference,
                        ClientName = row.ClientName ?? "Unknown Client",
                        Balance = row.Amount,
                        Description = $"Opening Balance - {caseReference}",
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
                        
                        if (!caseInfo.Value.ledgerCardId.HasValue)
                        {
                            summary.MissingLedgerCards++;
                            _logAction($"⚠️ Case '{caseReference}' found but NO LEDGER CARD exists!");
                        }
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

        private (int caseId, int? ledgerCardId)? GetCaseIdByReference(MySqlConnection connection, string caseReference)
        {
            try
            {
                string sql = @"
                    SELECT c.case_id, l.ledger_card_id
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
                            var caseId = reader.GetInt32("case_id");
                            var ledgerCardId = reader.IsDBNull(reader.GetOrdinal("ledger_card_id")) 
                                ? (int?)null 
                                : reader.GetInt32("ledger_card_id");
                            return (caseId, ledgerCardId);
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

            // Filter only found cases WITH ledger cards
            var foundRecords = records.Where(r => r.IsFound && r.HasLedgerCard).ToList();
            var missingLedgerCards = records.Where(r => r.IsFound && !r.HasLedgerCard).ToList();

            if (missingLedgerCards.Count > 0)
            {
                _logAction($"❌ ERROR: {missingLedgerCards.Count} cases found but missing ledger cards:");
                foreach (var rec in missingLedgerCards)
                {
                    _logAction($"   - {rec.CaseReference} (CaseId: {rec.CaseId})");
                }
                _logAction("Please create ledger cards for these cases before importing.");
                return (0, missingLedgerCards.Count);
            }

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
                                // Insert journal entry line (Credit) - use LedgerCardId (validated above)
                                InsertJournalEntryLine(connection, transaction, journalEntryId, lineNumber++,
                                    "ledger", data.LedgerCardId!.Value, 0, data.Balance,
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
                cmd.Parameters.AddWithValue("@ledgerCardId", data.LedgerCardId!.Value);
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

        /// <summary>
        /// Export journal entries as SQL INSERT statements for production database
        /// </summary>
        public string ExportAsSql(List<JournalImportData> records)
        {
            var sql = new System.Text.StringBuilder();
            var foundRecords = records.Where(r => r.IsFound && r.HasLedgerCard).ToList();
            var missingLedgerCards = records.Where(r => r.IsFound && !r.HasLedgerCard).ToList();

            if (missingLedgerCards.Count > 0)
            {
                sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
                sql.AppendLine("-- ERROR: Cannot export - the following cases are missing ledger cards:");
                sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
                foreach (var rec in missingLedgerCards)
                {
                    sql.AppendLine($"--   Case: {rec.CaseReference} (CaseId: {rec.CaseId})");
                }
                sql.AppendLine("-- Please create ledger cards for these cases before exporting.");
                return sql.ToString();
            }

            if (foundRecords.Count == 0)
            {
                return "-- No valid records to export.";
            }

            decimal totalAmount = foundRecords.Sum(r => r.Balance);
            string transactionDateStr = _transactionDate.ToString("yyyy-MM-dd HH:mm:ss");
            string nowStr = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");

            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine("-- JOURNAL ENTRY IMPORT - Opening Balances");
            sql.AppendLine($"-- Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sql.AppendLine($"-- Transaction Date: {_transactionDate:dd/MM/yyyy}");
            sql.AppendLine($"-- Total Records: {foundRecords.Count}");
            sql.AppendLine($"-- Total Amount: {totalAmount:N2}");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine();
            sql.AppendLine("SET FOREIGN_KEY_CHECKS = 0;");
            sql.AppendLine();

            // Use variables for IDs
            sql.AppendLine("-- Variables for generated IDs");
            sql.AppendLine("SET @transaction_id = 0;");
            sql.AppendLine("SET @journal_entry_id = 0;");
            sql.AppendLine("SET @journal_entry_number = 0;");
            sql.AppendLine();

            // Get next journal entry number
            sql.AppendLine("-- Get next journal entry number");
            sql.AppendLine("SELECT @journal_entry_number := COALESCE(MAX(journal_entry_number), 0) + 1 FROM tbl_acc_journal_entry;");
            sql.AppendLine();

            // 1. Insert Transaction
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine("-- 1. INSERT TRANSACTION");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine($@"INSERT INTO tbl_acc_transactions (
    fk_branch_id, fk_transaction_type_id, fk_transaction_sub_type_id, 
    transaction_details, transaction_reference, transaction_amount,
    is_cancelled, post_by, post_date_time, transaction_date, transaction_type
) VALUES (
    1, 1, 4, 'Opening Balance', CONCAT('JRN ', @journal_entry_number), {Math.Abs(totalAmount):F2},
    0, 1, '{nowStr}', '{transactionDateStr}', 'Add'
);");
            sql.AppendLine("SET @transaction_id = LAST_INSERT_ID();");
            sql.AppendLine();

            // 2. Insert Journal Entry Header
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine("-- 2. INSERT JOURNAL ENTRY HEADER");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine($@"INSERT INTO tbl_acc_journal_entry (
    fk_branch_id, fk_transaction_id, journal_entry_number, reference,
    journal_entry_description, journal_entry_date, current_date_time,
    staff_id, total, is_canceled
) VALUES (
    1, @transaction_id, @journal_entry_number, CONCAT('JRN-', @journal_entry_number),
    'Opening Balance', '{transactionDateStr}', '{nowStr}',
    1, {Math.Abs(totalAmount):F2}, 0
);");
            sql.AppendLine("SET @journal_entry_id = LAST_INSERT_ID();");
            sql.AppendLine();

            // 3. Insert Bank Line (Debit)
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine("-- 3. INSERT BANK LINE (DEBIT SIDE)");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine($@"INSERT INTO tbl_acc_journal_entry_lines (
    fk_journal_entry_id, fk_branch_id, line_number, account_type,
    fk_account_id, description, debit, credit, date_time, fk_user_id, is_deleted
) VALUES (
    @journal_entry_id, 1, 1, 'bank', 1, 'Opening Balance', {Math.Abs(totalAmount):F2}, 0, '{nowStr}', 1, 0
);");
            sql.AppendLine();

            // 4. Insert Bank Transaction
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine("-- 4. INSERT CLIENT BANK TRANSACTION");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine($@"INSERT INTO tbl_acc_client_bank_transactions (
    fk_branch_id, fk_transaction_id, fk_client_bank_id, transaction_date_time,
    transaction, reference, details, dr_amount, cr_amount,
    balance_pre, balance_post, is_cancelled, is_reconciled, fk_bank_reconciliation_id
) VALUES (
    1, @transaction_id, 1, '{transactionDateStr}',
    'Opening Balance', CONCAT('JRN-', @journal_entry_number), 'Opening Balance', {Math.Abs(totalAmount):F2}, 0,
    0, {totalAmount:F2}, 0, 0, NULL
);");
            sql.AppendLine();

            // 5. Insert Journal Entry Lines (Credit) for each case
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine($"-- 5. INSERT JOURNAL ENTRY LINES (CREDIT SIDE) - {foundRecords.Count} records");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            
            int lineNumber = 2;
            foreach (var data in foundRecords)
            {
                var desc = $"Opening Balance - {data.CaseReference}".Replace("'", "''");
                sql.AppendLine($@"INSERT INTO tbl_acc_journal_entry_lines (
    fk_journal_entry_id, fk_branch_id, line_number, account_type,
    fk_account_id, description, debit, credit, date_time, fk_user_id, is_deleted
) VALUES (
    @journal_entry_id, 1, {lineNumber++}, 'ledger', {data.LedgerCardId!.Value}, '{desc}', 0, {data.Balance:F2}, '{nowStr}', 1, 0
);");
            }
            sql.AppendLine();

            // 6. Insert Ledger Card Transactions
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine($"-- 6. INSERT LEDGER CARD TRANSACTIONS - {foundRecords.Count} records");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            
            foreach (var data in foundRecords)
            {
                decimal clientBalance = Math.Abs(data.Balance);
                string clientBalanceType = data.Balance > 0 ? "CR" : "DR";

                sql.AppendLine($@"INSERT INTO tbl_acc_ledger_card_transactions (
    fk_branch_id, fk_transaction_id, fk_ledger_card_id, transaction_date_time,
    details, office_dr, office_cr, office_bal, office_bal_type,
    client_dr, client_cr, client_bal, client_bal_type, total,
    ledger_reference, is_cancelled
) VALUES (
    1, @transaction_id, {data.LedgerCardId!.Value}, '{transactionDateStr}',
    'Opening Balance', 0, 0, 0, 'DR',
    0, {data.Balance:F2}, {clientBalance:F2}, '{clientBalanceType}', {Math.Abs(data.Balance):F2},
    CONCAT('JRN-', @journal_entry_number), 0
);");
            }

            sql.AppendLine();
            sql.AppendLine("SET FOREIGN_KEY_CHECKS = 1;");
            sql.AppendLine();
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");
            sql.AppendLine("-- IMPORT COMPLETE");
            sql.AppendLine("-- ═══════════════════════════════════════════════════════════════════");

            return sql.ToString();
        }
    }
}
