using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using LeapMergeDoc.Models;
using MySql.Data.MySqlClient;
using System.DirectoryServices;
using System.Globalization;
using System.IO;

namespace LeapMergeDoc.Services
{
    public class AccountUpdateImportService
    {
        private readonly string _connectionString;
        private readonly Action<string> _logAction;
        private readonly int _branchId = 1;
        private readonly int _userId = 1;
        private readonly int _defaultPaymentTypeId = 6; // Bank Transfer (default since we don't know exact types)

        public AccountUpdateImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _logAction = logAction;
        }

        #region Excel Reading

        /// <summary>
        /// Reads Excel data in the Bank Account Register format:
        /// Column B: Transaction Date
        /// Column D: Transaction Details (Matter No.)
        /// Column E: Reason (Receipt/Payment)
        /// Column H: Reason Description
        /// Column J: Withdrawal (Payment amount)
        /// Column K: Deposit Balance (Receipt amount)
        /// Column O: Receipts/Payment type
        /// Column P: Client to office
        /// Column Q: Case Number (case_reference_auto)
        /// </summary>
        public List<AccountUpdateExcelData> ReadExcelData(string filePath)
        {
            var data = new List<AccountUpdateExcelData>();

            // Copy file to temp location to avoid file lock issues
            string tempFilePath = Path.Combine(Path.GetTempPath(), $"AccountUpdate_{Guid.NewGuid()}.xlsx");
            try
            {
                File.Copy(filePath, tempFilePath, true);
                _logAction($"Copied Excel file to temp: {tempFilePath}");
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ Could not copy file: {ex.Message}. Trying to read original...");
                tempFilePath = filePath;
            }

            try
            {
                using (var workbook = new XLWorkbook(tempFilePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var range = worksheet.RangeUsed();

                    if (range == null)
                    {
                        _logAction("No data found in Excel file.");
                        return data;
                    }

                    var rowCount = range.RowCount();
                    _logAction($"Found {rowCount - 1} data rows in Excel file.");

                    // Auto-detect header row by scanning rows 1 and 2 for known header keywords
                    var headers = new Dictionary<string, int>();
                    var maxCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 20;
                    var knownHeaders = new[] { "transaction date", "transaction details", "withdrawal", "deposit", "case number", "reason", "balance" };

                    int row1Matches = 0, row2Matches = 0;
                    for (int col = 1; col <= maxCol; col++)
                    {
                        var val1 = worksheet.Cell(1, col).GetString().Trim().ToLower();
                        var val2 = worksheet.Cell(2, col).GetString().Trim().ToLower();
                        if (knownHeaders.Any(h => val1.Contains(h))) row1Matches++;
                        if (knownHeaders.Any(h => val2.Contains(h))) row2Matches++;
                    }
                    var headerRow = row2Matches > row1Matches ? 2 : 1;
                    _logAction($"Auto-detected header row: {headerRow} (row1 matches: {row1Matches}, row2 matches: {row2Matches})");

                    for (int col = 1; col <= maxCol; col++)
                    {
                        var headerValue = worksheet.Cell(headerRow, col).GetString().Trim().ToLower();
                        if (!string.IsNullOrEmpty(headerValue))
                        {
                            headers[headerValue] = col;
                        }
                    }
                    
                    _logAction($"Found {headers.Count} headers in row {headerRow}");

                    // Map column positions based on header names or fixed positions
                    int colTransactionDate = GetColumnIndex(headers, "transaction date", "transactiondate", "date") ?? 1;
                    int colEntryDate = GetColumnIndex(headers, "entry date", "entrydate") ?? -1;
                    int colTransactionDetails = GetColumnIndex(headers, "transaction details", "transactiondetails", "details") ?? 3;
                    int colComments = GetColumnIndex(headers, "comments", "comment", "notes") ?? -1;
                    int colPaymentType = GetColumnIndex(headers, "payment type", "paymenttype") ?? -1;
                    int colReceivedFromPaidTo = GetColumnIndex(headers, "received from/paid to", "received from", "receivedfrom/paidto", "receivedfrom", "paid to", "paidto") ?? -1;
                    int colReason = GetColumnIndex(headers, "reason") ?? 7;
                    int? colReasonDesc = GetColumnIndex(headers, "reason:", "reason description", "reason desc");
                    int colWithdrawal = GetColumnIndex(headers, "withdrawal") ?? 8;
                    int colDepositBalance = GetColumnIndex(headers, "deposit balance", "depositbalance", "deposit") ?? 9;
                    int colReceiptsPayment = GetColumnIndex(headers, "receipts /payment", "receipts/payment", "receipts", "receipt") ?? 13;
                    int colClientToOffice = GetColumnIndex(headers, "client to office", "clienttooffice", "c2o") ?? 14;
                    int colCaseNumber = GetColumnIndex(headers, "case number", "casenumber", "case") ?? 15;

                    _logAction($"Column mapping - Date:{colTransactionDate}, EntryDate:{colEntryDate}, Details:{colTransactionDetails}, Comments:{colComments}, PaymentType:{colPaymentType}, ReceivedFrom:{colReceivedFromPaidTo}, Reason:{colReason}, ReasonDesc:{colReasonDesc?.ToString() ?? "N/A"}, Withdrawal:{colWithdrawal}, Deposit:{colDepositBalance}, Receipts/Payment:{colReceiptsPayment}, C2O:{colClientToOffice}, Case:{colCaseNumber}");

                    // Read data rows starting from the row after headers
                    for (int row = headerRow + 1; row <= rowCount; row++)
                    {
                        try
                        {
                            var caseNumber = worksheet.Cell(row, colCaseNumber).GetString().Trim();
                            
                            // Skip rows without case number
                            if (string.IsNullOrEmpty(caseNumber))
                                continue;

                            var rowData = new AccountUpdateExcelData { RowNumber = row };

                            // Case Reference (from Case Number column - this is case_reference_auto)
                            rowData.CaseReference = caseNumber;

                            // Transaction Date
                            var dateCell = worksheet.Cell(row, colTransactionDate);
                        rowData.TransactionDate = ParseDateFromCell(dateCell);

                        // Transaction Details
                        rowData.Description = worksheet.Cell(row, colTransactionDetails).GetString().Trim();

                        // Reason (Receipt/Payment)
                        var reason = worksheet.Cell(row, colReason).GetString().Trim();

                        // Reason Description (only if that specific column was found in headers)
                        if (colReasonDesc.HasValue)
                        {
                            var reasonDesc = worksheet.Cell(row, colReasonDesc.Value).GetString().Trim();
                            if (!string.IsNullOrEmpty(reasonDesc))
                            {
                                rowData.Description = reasonDesc;
                            }
                        }

                        // Comments
                        if (colComments > 0)
                        {
                            rowData.Comments = worksheet.Cell(row, colComments).GetString().Trim();
                        }

                        // Payment Type
                        if (colPaymentType > 0)
                        {
                            var paymentTypeStr = worksheet.Cell(row, colPaymentType).GetString().Trim();
                            if (!string.IsNullOrEmpty(paymentTypeStr))
                            {
                                if (int.TryParse(paymentTypeStr, out int ptId))
                                    rowData.PaymentTypeId = ptId;
                                else
                                    rowData.PaymentTypeId = MapPaymentTypeName(paymentTypeStr);
                            }
                        }

                        // Received From / Paid To
                        if (colReceivedFromPaidTo > 0)
                        {
                            var receivedPaid = worksheet.Cell(row, colReceivedFromPaidTo).GetString().Trim();
                            if (!string.IsNullOrEmpty(receivedPaid))
                            {
                                rowData.ReceivedFrom = receivedPaid;
                                rowData.PaidTo = receivedPaid;
                            }
                        }

                        // Withdrawal amount (Payment)
                        var withdrawalStr = worksheet.Cell(row, colWithdrawal).GetString().Trim();
                        decimal withdrawal = ParseAmount(withdrawalStr);

                        // Deposit amount (Receipt)
                        var depositStr = worksheet.Cell(row, colDepositBalance).GetString().Trim();
                        decimal deposit = ParseAmount(depositStr);

                        // Receipts/Payment type indicator
                        var typeIndicator = worksheet.Cell(row, colReceiptsPayment).GetString().Trim();

                        // Client to Office indicator
                        var c2oIndicator = worksheet.Cell(row, colClientToOffice).GetString().Trim();

                        // Determine transaction type and amount
                        if (!string.IsNullOrEmpty(c2oIndicator) && 
                            (c2oIndicator.ToLower().Contains("c2o") || 
                             c2oIndicator.ToLower().Contains("client to office") ||
                             c2oIndicator.ToLower().Contains("transfer")))
                        {
                            rowData.TransactionType = "C2O";
                            rowData.Amount = deposit > 0 ? deposit : withdrawal;
                            rowData.InvoiceAmount = rowData.Amount;
                        }
                        else if (reason.ToLower() == "payment" || withdrawal > 0)
                        {
                            rowData.TransactionType = "Payment";
                            rowData.Amount = withdrawal > 0 ? withdrawal : deposit;
                        }
                        else if (reason.ToLower() == "receipt" || deposit > 0)
                        {
                            rowData.TransactionType = "Receipt";
                            rowData.Amount = deposit > 0 ? deposit : withdrawal;
                        }
                        else if (!string.IsNullOrEmpty(typeIndicator))
                        {
                            // Use type indicator column
                            if (typeIndicator.ToLower().Contains("receipt"))
                            {
                                rowData.TransactionType = "Receipt";
                                rowData.Amount = deposit > 0 ? deposit : withdrawal;
                            }
                            else if (typeIndicator.ToLower().Contains("payment"))
                            {
                                rowData.TransactionType = "Payment";
                                rowData.Amount = withdrawal > 0 ? withdrawal : deposit;
                            }
                        }

                        // Default payment type
                        rowData.PaymentTypeId = 1; // BACS

                        // Only add if we have valid data
                        if (!string.IsNullOrEmpty(rowData.TransactionType) && 
                            !string.IsNullOrEmpty(rowData.CaseReference) &&
                            rowData.Amount > 0)
                        {
                            data.Add(rowData);
                            _logAction($"Row {row}: {rowData.TransactionType} - {rowData.CaseReference} - {rowData.Amount:C}");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logAction($"⚠️ Error reading row {row}: {ex.Message}");
                    }
                }
                }
            }
            finally
            {
                // Cleanup temp file
                if (tempFilePath != filePath && File.Exists(tempFilePath))
                {
                    try { File.Delete(tempFilePath); } catch { }
                }
            }

            _logAction($"Successfully read {data.Count} valid records from Excel.");
            return data;
        }

        private int? GetColumnIndex(Dictionary<string, int> headers, params string[] possibleNames)
        {
            foreach (var name in possibleNames)
            {
                // Exact match
                if (headers.TryGetValue(name.ToLower(), out int col))
                    return col;

                // Partial match
                var match = headers.FirstOrDefault(h => h.Key.Contains(name.ToLower()));
                if (match.Value > 0)
                    return match.Value;
            }
            return null;
        }

        private DateTime? ParseDateFromCell(IXLCell cell)
        {
            try
            {
                // Try to get as DateTime directly
                if (cell.DataType == XLDataType.DateTime)
                {
                    return cell.GetDateTime();
                }

                // Try to parse string value
                var dateStr = cell.GetString().Trim();
                return ParseDate(dateStr);
            }
            catch
            {
                return null;
            }
        }

        private decimal ParseAmount(string amountStr)
        {
            if (string.IsNullOrEmpty(amountStr))
                return 0;

            // Remove currency symbols, commas, spaces
            amountStr = amountStr.Replace("£", "").Replace("$", "").Replace(",", "").Replace(" ", "").Trim();

            if (decimal.TryParse(amountStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amount))
            {
                return Math.Abs(amount);
            }

            // Try UK culture
            if (decimal.TryParse(amountStr, NumberStyles.Any, new CultureInfo("en-GB"), out amount))
            {
                return Math.Abs(amount);
            }

            return 0;
        }

        private void MapExcelColumn(AccountUpdateExcelData rowData, string header, string cellValue)
        {
            switch (header)
            {
                case "transactiontype":
                case "type":
                case "txntype":
                case "reason":
                    rowData.TransactionType = cellValue;
                    break;
                case "casereference":
                case "case":
                case "matter":
                case "caseref":
                case "casenumber":
                    rowData.CaseReference = cellValue;
                    break;
                case "clientname":
                case "client":
                    rowData.ClientName = cellValue;
                    break;
                case "transactiondate":
                case "date":
                case "txndate":
                    rowData.TransactionDate = ParseDate(cellValue);
                    break;
                case "amount":
                case "transactionamount":
                case "txnamount":
                    if (decimal.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amount))
                    {
                        rowData.Amount = amount;
                    }
                    break;
                case "description":
                case "desc":
                case "transactiondescription":
                case "transactiondetails":
                    rowData.Description = cellValue;
                    break;
                case "paymentreference":
                case "reference":
                case "ref":
                    rowData.PaymentReference = cellValue;
                    break;
                case "receivedfrom":
                case "from":
                case "payer":
                    rowData.ReceivedFrom = cellValue;
                    break;
                case "paidto":
                case "to":
                case "payee":
                    rowData.PaidTo = cellValue;
                    break;
                case "paymenttypeid":
                case "paymenttype":
                    if (int.TryParse(cellValue, out int paymentType))
                    {
                        rowData.PaymentTypeId = paymentType;
                    }
                    else
                    {
                        rowData.PaymentTypeId = MapPaymentTypeName(cellValue);
                    }
                    break;
                case "comments":
                case "comment":
                case "notes":
                    rowData.Comments = cellValue;
                    break;
                case "invoicenumber":
                case "invoice":
                case "invno":
                    rowData.InvoiceNumber = cellValue;
                    break;
                case "invoiceamount":
                case "invamount":
                    if (decimal.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal invAmount))
                    {
                        rowData.InvoiceAmount = invAmount;
                    }
                    break;
                case "bankaccount":
                case "bank":
                case "bankname":
                    rowData.BankAccountName = cellValue;
                    break;
                case "clientbankid":
                    if (int.TryParse(cellValue, out int clientBankId))
                    {
                        rowData.ClientBankId = clientBankId;
                    }
                    break;
                case "officebankid":
                    if (int.TryParse(cellValue, out int officeBankId))
                    {
                        rowData.OfficeBankId = officeBankId;
                    }
                    break;
            }
        }

        private int MapPaymentTypeName(string paymentTypeName)
        {
            return paymentTypeName?.ToLower() switch
            {
                "bacs" => 1,
                "cash" => 2,
                "cheque" or "check" => 3,
                "card" => 4,
                "client cheque" => 5,
                "transfer" => 6,
                _ => 1 // Default to BACS
            };
        }

        private DateTime? ParseDate(string? dateString)
        {
            if (string.IsNullOrEmpty(dateString))
                return null;

            string[] formats = {
                "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd",
                "dd-MM-yyyy", "MM-dd-yyyy", "d/M/yyyy",
                "M/d/yyyy", "yyyy/MM/dd", "dd/MM/yyyy HH:mm:ss"
            };

            foreach (string format in formats)
            {
                if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, 
                    DateTimeStyles.None, out DateTime result))
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

        #endregion

        #region Data Processing

        public (List<AccountUpdateImportData> records, AccountUpdateImportSummary summary) ProcessExcelData(
            List<AccountUpdateExcelData> excelData, int defaultClientBankId, int defaultOfficeBankId)
        {
            var records = new List<AccountUpdateImportData>();
            var summary = new AccountUpdateImportSummary();

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                foreach (var row in excelData)
                {
                    var importData = ProcessSingleRecord(connection, row, defaultClientBankId, defaultOfficeBankId);
                    records.Add(importData);

                    // Update summary
                    summary.TotalRecords++;

                    if (importData.IsValid)
                    {
                        summary.ValidRecords++;

                        if (importData.IsFound)
                            summary.FoundCases++;
                        else
                            summary.NotFoundCases++;

                        switch (importData.TransactionType)
                        {
                            case AccountTransactionType.BankReceipt:
                                summary.ReceiptCount++;
                                summary.ReceiptTotal += importData.Amount;
                                break;
                            case AccountTransactionType.BankPayment:
                                summary.PaymentCount++;
                                summary.PaymentTotal += importData.Amount;
                                break;
                            case AccountTransactionType.ClientToOffice:
                                summary.ClientToOfficeCount++;
                                summary.ClientToOfficeTotal += importData.Amount;
                                break;
                        }
                    }
                    else
                    {
                        summary.InvalidRecords++;
                        if (!string.IsNullOrEmpty(importData.ValidationError))
                        {
                            summary.Errors.Add($"Row {importData.RowNumber}: {importData.ValidationError}");
                        }
                    }
                }
            }

            return (records, summary);
        }

        private AccountUpdateImportData ProcessSingleRecord(MySqlConnection connection, 
            AccountUpdateExcelData row, int defaultClientBankId, int defaultOfficeBankId)
        {
            var importData = new AccountUpdateImportData
            {
                RowNumber = row.RowNumber,
                TransactionType = ParseTransactionType(row.TransactionType),
                CaseReference = row.CaseReference,
                ClientName = row.ClientName,
                TransactionDate = row.TransactionDate ?? DateTime.UtcNow,
                Amount = row.Amount,
                Description = row.Description ?? $"Imported transaction - {row.CaseReference}",
                PaymentReference = row.PaymentReference,
                ReceivedFrom = row.ReceivedFrom,
                PaidTo = row.PaidTo,
                PaymentTypeId = row.PaymentTypeId > 0 ? row.PaymentTypeId : 1,
                Comments = row.Comments,
                InvoiceNumber = row.InvoiceNumber,
                InvoiceAmount = row.InvoiceAmount ?? row.Amount,
                ClientBankId = row.ClientBankId ?? defaultClientBankId,
                OfficeBankId = row.OfficeBankId ?? defaultOfficeBankId,
                IsValid = true
            };

            // Lookup case
            var caseInfo = GetCaseInfo(connection, row.CaseReference!);
            if (caseInfo != null)
            {
                importData.CaseId = caseInfo.Value.caseId;
                importData.ClientId = caseInfo.Value.clientId;
                importData.LedgerCardId = caseInfo.Value.ledgerCardId;
                importData.IsFound = true;
            }
            else
            {
                //importData.IsFound = false;
                //importData.ValidationError = $"Case not found: {row.CaseReference}";
                //_logAction($"⚠️ Case not found: '{row.CaseReference}'");

                _logAction($"⚠️ Case not found: '{row.CaseReference}' — auto-creating...");
                try
                {
                    var created = AutoCreateCase(connection, row);
                    importData.CaseId = created.caseId;
                    importData.ClientId = created.clientId;
                    importData.LedgerCardId = created.ledgerCardId;
                    importData.IsFound = true;
                    importData.ValidationError = null;
                    _logAction($"✅ Auto-created case '{row.CaseReference}' (CaseId={created.caseId})");
                }
                catch (Exception ex)
                {
                    importData.IsFound = false;
                    importData.IsValid = false;
                    importData.ValidationError = $"Case not found and auto-create failed: {ex.Message}";
                    _logAction($"❌ Auto-create failed for '{row.CaseReference}': {ex.Message}");
                }
            }

            // Validate
            ValidateImportData(importData);

            return importData;
        }

        private (int caseId, int clientId, int ledgerCardId) AutoCreateCase( MySqlConnection connection, AccountUpdateExcelData row)
        {
            using var transaction = connection.BeginTransaction();
            try
            {
                int clientId = ResolveClientId(connection, transaction, row.ClientName);


                ParseCaseReference(
                    row.CaseReference!,
                    out int? aopId,
                    out int? caseTypeId,
                    out int? caseSubTypeId,
                    out int caseNumber,
                    connection,
                    transaction);

                // ── 3. tbl_case_details_general ─────────────────────────────────────
                int caseId = InsertCaseDetailsGeneral(
                    connection, transaction, row, clientId,
                    aopId, caseTypeId, caseSubTypeId, caseNumber);


                // ── 5. tbl_acc_ledger_cards ──────────────────────────────────────────
                int ledgerCardId = InsertLedgerCard(connection, transaction, caseId, clientId);

                // ── 6. tbl_case_permissions ──────────────────────────────────────────
                InsertCasePermissions(connection, transaction, caseId);

                // ── 7. tbl_case_clients ──────────────────────────────────────────────
                InsertCaseClient(connection, transaction, caseId, clientId, row);

                // ── 8. tbl_case_client_greeting ──────────────────────────────────────
                InsertCaseClientGreeting(connection, transaction, caseId, row.ClientName);

                transaction.Commit();
                return (caseId, clientId, ledgerCardId);
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        private int ResolveClientId(MySqlConnection conn, MySqlTransaction tx, string? clientName)
        {
            if (!string.IsNullOrWhiteSpace(clientName))
            {
                var name = clientName.Trim();

                // ── 1. Check tbl_client_individual (given_names + last_name) ─────────
                const string individualSql = @"
            SELECT fk_client_id
            FROM tbl_client_individual
            WHERE LOWER(TRIM(CONCAT(given_names, ' ', last_name))) = LOWER(@fullName)
               OR LOWER(TRIM(given_names)) = LOWER(@fullName)
               OR LOWER(TRIM(last_name))   = LOWER(@fullName)
            LIMIT 1;";

                using (var cmd = new MySqlCommand(individualSql, conn, tx))
                {
                    cmd.Parameters.AddWithValue("@fullName", name);
                    var result = cmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        _logAction($"✅ Client matched in tbl_client_individual: '{clientName}'");
                        return Convert.ToInt32(result);
                    }
                }

                // ── 2. Check tbl_client_company (company_name) ───────────────────────
                const string companySql = @"
            SELECT fk_client_id
            FROM tbl_client_company
            WHERE LOWER(TRIM(company_name)) = LOWER(@fullName)
            LIMIT 1;";

                using (var cmd = new MySqlCommand(companySql, conn, tx))
                {
                    cmd.Parameters.AddWithValue("@fullName", name);
                    var result = cmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        _logAction($"✅ Client matched in tbl_client_company: '{clientName}'");
                        return Convert.ToInt32(result);
                    }
                }

                // ── 3. Check tbl_client_organisation (org_name) ──────────────────────
                const string orgSql = @"
            SELECT fk_client_id
            FROM tbl_client_organisation
            WHERE LOWER(TRIM(org_name)) = LOWER(@fullName)
            LIMIT 1;";

                using (var cmd = new MySqlCommand(orgSql, conn, tx))
                {
                    cmd.Parameters.AddWithValue("@fullName", name);
                    var result = cmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        _logAction($"✅ Client matched in tbl_client_organisation: '{clientName}'");
                        return Convert.ToInt32(result);
                    }
                }
            }

            // ── Not found — pick a random existing fk_client_id from any client table ─
            const string randomSql = @"
        SELECT fk_client_id FROM (
            SELECT fk_client_id FROM tbl_client_individual  WHERE fk_client_id IS NOT NULL
            UNION
            SELECT fk_client_id FROM tbl_client_company     WHERE fk_client_id IS NOT NULL
            UNION
            SELECT fk_client_id FROM tbl_client_organisation WHERE fk_client_id IS NOT NULL
        ) AS all_clients
        ORDER BY RAND()
        LIMIT 1;";

            using (var randomCmd = new MySqlCommand(randomSql, conn, tx))
            {
                var randomResult = randomCmd.ExecuteScalar();
                if (randomResult != null && randomResult != DBNull.Value)
                {
                    var randomId = Convert.ToInt32(randomResult);
                    _logAction($"⚠️ Client '{clientName}' not found — assigned random client id={randomId}");
                    return randomId;
                }
            }

            throw new InvalidOperationException("No clients exist in the database — cannot resolve a client id.");
        }

        private void ParseCaseReference(
        string caseRef,
        out int? aopId,
        out int? caseTypeId,
        out int? caseSubTypeId,
        out int caseNumber,
        MySqlConnection conn,
        MySqlTransaction tx)
        {
            aopId = null;
            caseTypeId = null;
            caseSubTypeId = null;
            caseNumber = 0;

            // ── Get next unique case_number (MAX + 1) ─────────────────────────────────
            const string maxCaseNumberSql = @"
        SELECT COALESCE(MAX(case_number), 0) + 1
        FROM tbl_case_details_general;";

            using (var cmd = new MySqlCommand(maxCaseNumberSql, conn, tx))
            {
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                    caseNumber = Convert.ToInt32(result);
            }

            // ── Split reference into segments ─────────────────────────────────────────
            var parts = caseRef.Split(new[] { '/', '.' }, StringSplitOptions.RemoveEmptyEntries);

            string? apCode = parts.Length >= 2 ? parts[1] : null;
            string? caseTypeCode = parts.Length >= 3 ? parts[2] : null;

            // ── 1. Lookup aopId by ap_code (parts[1]) ────────────────────────────────
            if (!string.IsNullOrWhiteSpace(apCode))
            {
                const string aopSql = @"
            SELECT areas_of_practice_id
            FROM tbl_areas_of_practice
            WHERE UPPER(TRIM(ap_code)) = UPPER(TRIM(@code))
              AND (is_active = 1 OR is_active IS NULL)
            LIMIT 1;";

                using var cmd = new MySqlCommand(aopSql, conn, tx);
                cmd.Parameters.AddWithValue("@code", apCode);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    aopId = Convert.ToInt32(result);
                    _logAction($"✅ AoP matched by ap_code '{apCode}': aopId={aopId}");
                }
            }

            // Not matched → pick a random active aopId
            if (!aopId.HasValue)
            {
                const string randomAopSql = @"
            SELECT areas_of_practice_id
            FROM tbl_areas_of_practice
            WHERE (is_active = 1 OR is_active IS NULL)
            ORDER BY RAND()
            LIMIT 1;";

                using var cmd = new MySqlCommand(randomAopSql, conn, tx);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    aopId = Convert.ToInt32(result);
                    _logAction($"⚠️ AoP code '{apCode}' not found — assigned random aopId={aopId}");
                }
            }

            // ── 2. Lookup caseTypeId by case_type_code (parts[2]) + aopId ────────────
            if (!string.IsNullOrWhiteSpace(caseTypeCode) && aopId.HasValue)
            {
                const string caseTypeSql = @"
            SELECT case_type_id
            FROM tbl_case_type
            WHERE UPPER(TRIM(case_type_code)) = UPPER(TRIM(@code))
              AND fk_areas_of_practice_id = @aopId
              AND (is_active = 1 OR is_active IS NULL)
            LIMIT 1;";

                using var cmd = new MySqlCommand(caseTypeSql, conn, tx);
                cmd.Parameters.AddWithValue("@code", caseTypeCode);
                cmd.Parameters.AddWithValue("@aopId", aopId.Value);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    caseTypeId = Convert.ToInt32(result);
                    _logAction($"✅ CaseType matched by code '{caseTypeCode}' + aopId={aopId}: caseTypeId={caseTypeId}");
                }
            }

            // Not matched → pick a random active caseTypeId
            if (!caseTypeId.HasValue)
            {
                const string randomCaseTypeSql = @"
            SELECT case_type_id
            FROM tbl_case_type
            WHERE (is_active = 1 OR is_active IS NULL)
            ORDER BY RAND()
            LIMIT 1;";

                using var cmd = new MySqlCommand(randomCaseTypeSql, conn, tx);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    caseTypeId = Convert.ToInt32(result);
                    _logAction($"⚠️ CaseType code '{caseTypeCode}' not found — assigned random caseTypeId={caseTypeId}");
                }
            }

            // ── 3. Lookup caseSubTypeId by caseTypeId (parts[2] is sub_type_code) ─────
            if (caseTypeId.HasValue)
            {
                const string subTypeSql = @"
            SELECT case_sub_type_id
            FROM tbl_case_sub_type
            WHERE fk_case_type_id = @caseTypeId
              AND (is_active = 1 OR is_active IS NULL)
            ORDER BY case_sub_type_id
            LIMIT 1;";

                using var cmd = new MySqlCommand(subTypeSql, conn, tx);
                cmd.Parameters.AddWithValue("@caseTypeId", caseTypeId.Value);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    caseSubTypeId = Convert.ToInt32(result);
                    _logAction($"✅ SubCaseType matched: caseSubTypeId={caseSubTypeId}");
                }
            }

            _logAction($"📋 '{caseRef}' → caseNumber={caseNumber} | " +
                       $"aopId={aopId?.ToString() ?? "null"} | " +
                       $"caseTypeId={caseTypeId?.ToString() ?? "null"} | " +
                       $"caseSubTypeId={caseSubTypeId?.ToString() ?? "null"}");
        }


        private int InsertCaseDetailsGeneral(
        MySqlConnection conn, MySqlTransaction tx,
        AccountUpdateExcelData row, int clientId,
        int? aopId, int? caseTypeId, int? caseSubTypeId, int caseNumber)
        {
            string clientName = ResolveClientNameById(conn, tx, clientId)
                        ?? row.ClientName
                        ?? "Unknown Client";

            const string sql = @"
        INSERT INTO tbl_case_details_general (
            fk_branch_id, fk_area_of_practice_id, fk_case_type_id, fk_case_sub_type_id,
            case_reference_auto, case_number, case_name,
            case_clients, date_opened, person_opened,
            is_case_active, is_case_archived, is_case_not_proceeding, mnl_check, conf_search
        ) VALUES (
            @branchId, @aopId, @caseTypeId, @caseSubTypeId,
            @caseRef, @caseNumber, @caseName,
            @caseClients, @dateOpened, @personOpened,
            0, 1, 0, 1, 1
        );
        SELECT LAST_INSERT_ID();";

            using var cmd = new MySqlCommand(sql, conn, tx);
            cmd.Parameters.AddWithValue("@branchId", _branchId);
            cmd.Parameters.AddWithValue("@aopId", (object?)aopId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@caseTypeId", (object?)caseTypeId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@caseSubTypeId", (object?)caseSubTypeId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@caseRef", row.CaseReference ?? "");
            cmd.Parameters.AddWithValue("@caseNumber", caseNumber);
            cmd.Parameters.AddWithValue("@caseName", row.ClientName ?? row.CaseReference ?? "Auto-created");
            cmd.Parameters.AddWithValue("@caseClients", clientName ?? "");
            cmd.Parameters.AddWithValue("@dateOpened", row.TransactionDate);
            cmd.Parameters.AddWithValue("@personOpened", _userId);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        private string? ResolveClientNameById(MySqlConnection conn, MySqlTransaction tx, int clientId)
        {
            // 1. Individual → CONCAT given_names + last_name
            const string individualSql = @"
        SELECT TRIM(CONCAT(COALESCE(given_names, ''), ' ', COALESCE(last_name, '')))
        FROM tbl_client_individual
        WHERE fk_client_id = @clientId
        LIMIT 1;";

            using (var cmd = new MySqlCommand(individualSql, conn, tx))
            {
                cmd.Parameters.AddWithValue("@clientId", clientId);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    var name = result.ToString()?.Trim();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        _logAction($"✅ Client name resolved from tbl_client_individual: '{name}'");
                        return name;
                    }
                }
            }

            // 2. Company → company_name
            const string companySql = @"
        SELECT company_name
        FROM tbl_client_company
        WHERE fk_client_id = @clientId
        LIMIT 1;";

            using (var cmd = new MySqlCommand(companySql, conn, tx))
            {
                cmd.Parameters.AddWithValue("@clientId", clientId);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    var name = result.ToString()?.Trim();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        _logAction($"✅ Client name resolved from tbl_client_company: '{name}'");
                        return name;
                    }
                }
            }

            // 3. Organisation → org_name
            const string orgSql = @"
        SELECT org_name
        FROM tbl_client_organisation
        WHERE fk_client_id = @clientId
        LIMIT 1;";

            using (var cmd = new MySqlCommand(orgSql, conn, tx))
            {
                cmd.Parameters.AddWithValue("@clientId", clientId);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    var name = result.ToString()?.Trim();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        _logAction($"✅ Client name resolved from tbl_client_organisation: '{name}'");
                        return name;
                    }
                }
            }

            _logAction($"⚠️ Could not resolve client name for clientId={clientId}");
            return null;
        }


        private int InsertLedgerCard(MySqlConnection conn, MySqlTransaction tx, int caseId, int clientId)
        {
            const string sql = @"
        INSERT INTO tbl_acc_ledger_cards (fk_branch_id, fk_client_ids, fk_case_id, is_deleted, is_archived)
        VALUES (@branchId, @clientIds, @caseId, 0, 0);
        SELECT LAST_INSERT_ID();";

            using var cmd = new MySqlCommand(sql, conn, tx);
            cmd.Parameters.AddWithValue("@branchId", _branchId);
            cmd.Parameters.AddWithValue("@clientIds", clientId.ToString());
            cmd.Parameters.AddWithValue("@caseId", caseId);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        private void InsertCasePermissions(MySqlConnection conn, MySqlTransaction tx, int caseId)
        {
            const string sql = @"
        INSERT INTO tbl_case_permissions (
            fk_case_id, person_opened, person_responsible,
            person_acting, person_assisting,
            is_everyone_view, is_everyone_add_edit
        ) VALUES (
            @caseId, @userId, @userId,
            @userId, NULL,
            1, 0
        );";

            using var cmd = new MySqlCommand(sql, conn, tx);
            cmd.Parameters.AddWithValue("@caseId", caseId);
            cmd.Parameters.AddWithValue("@userId", _userId);
            cmd.ExecuteNonQuery();
        }

        private void InsertCaseClient(MySqlConnection conn, MySqlTransaction tx, int caseId, int clientId, AccountUpdateExcelData row)
        {
            const string sql = @"
        INSERT INTO tbl_case_clients (fk_case_id, fk_client_id, client_order, date_added, user_added)
        VALUES (@caseId, @clientId, 1, @dateAdded, @userId);";

            using var cmd = new MySqlCommand(sql, conn, tx);
            cmd.Parameters.AddWithValue("@caseId", caseId);
            cmd.Parameters.AddWithValue("@clientId", clientId);
            cmd.Parameters.AddWithValue("@dateAdded", row.TransactionDate);
            cmd.Parameters.AddWithValue("@userId", _userId);
            cmd.ExecuteNonQuery();
        }

        private void InsertCaseClientGreeting( MySqlConnection conn, MySqlTransaction tx, int caseId, string? clientName)
        {
            const string sql = @"
        INSERT INTO tbl_case_client_greeting (fk_case_id, greeting_type, greeting_all)
        VALUES (@caseId, 'all', @greeting);";

            using var cmd = new MySqlCommand(sql, conn, tx);
            cmd.Parameters.AddWithValue("@caseId", caseId);
            cmd.Parameters.AddWithValue("@greeting", $"Dear {clientName ?? "Client"}");
            cmd.ExecuteNonQuery();
        }





        private AccountTransactionType ParseTransactionType(string? typeString)
        {
            if (string.IsNullOrEmpty(typeString))
                return AccountTransactionType.Unknown;

            return typeString.ToLower() switch
            {
                "receipt" or "bankreceipt" or "rec" or "r" => AccountTransactionType.BankReceipt,
                "payment" or "bankpayment" or "pay" or "p" => AccountTransactionType.BankPayment,
                "c2o" or "clienttooffice" or "cto" or "transfer" or "t" => AccountTransactionType.ClientToOffice,
                _ => AccountTransactionType.Unknown
            };
        }

        private void ValidateImportData(AccountUpdateImportData data)
        {
            if (data.TransactionType == AccountTransactionType.Unknown)
            {
                data.IsValid = false;
                data.ValidationError = "Unknown transaction type";
                return;
            }

            if (data.Amount <= 0)
            {
                data.IsValid = false;
                data.ValidationError = "Invalid amount (must be greater than 0)";
                return;
            }

            if (!data.IsFound)
            {
                data.IsValid = false;
                return;
            }

            if (data.ClientBankId <= 0)
            {
                data.IsValid = false;
                data.ValidationError = "Client bank ID is required";
                return;
            }

            if (data.TransactionType == AccountTransactionType.ClientToOffice && data.OfficeBankId <= 0)
            {
                data.IsValid = false;
                data.ValidationError = "Office bank ID is required for Client to Office transfers";
                return;
            }
        }

        private (int caseId, int clientId, int ledgerCardId)? GetCaseInfo(MySqlConnection connection, string caseReference)
        {
            try
            {
                // Primary lookup by case_reference_auto (this matches the Case Number column in Excel)
                string sql = @"
                    SELECT c.case_id, COALESCE(cc.fk_client_id, 0) as client_id, 
                           COALESCE(l.ledger_card_id, 0) as ledger_card_id
                    FROM tbl_case_details_general c
                    LEFT JOIN tbl_case_clients cc ON cc.fk_case_id = c.case_id
                    LEFT JOIN tbl_acc_ledger_cards l ON l.fk_case_id = c.case_id
                    WHERE c.case_reference_auto = @caseReference 
                    LIMIT 1";

                using (var cmd = new MySqlCommand(sql, connection))
                {
                    cmd.Parameters.AddWithValue("@caseReference", caseReference);

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var caseId = reader.GetInt32("case_id");
                            var clientId = reader.GetInt32("client_id");
                            var ledgerCardId = reader.GetInt32("ledger_card_id");
                            
                            _logAction($"✓ Found case: {caseReference} -> CaseId:{caseId}, ClientId:{clientId}, LedgerId:{ledgerCardId}");
                            return (caseId, clientId, ledgerCardId);
                        }
                    }
                }

                // Fallback: try case_reference_manual
                sql = @"
                    SELECT c.case_id, COALESCE(cc.fk_client_id, 0) as client_id, 
                           COALESCE(l.ledger_card_id, 0) as ledger_card_id
                    FROM tbl_case_details_general c
                    LEFT JOIN tbl_case_clients cc ON cc.fk_case_id = c.case_id
                    LEFT JOIN tbl_acc_ledger_cards l ON l.fk_case_id = c.case_id
                    WHERE c.case_reference_manual = @caseReference
                    LIMIT 1";

                using (var cmd = new MySqlCommand(sql, connection))
                {
                    cmd.Parameters.AddWithValue("@caseReference", caseReference);

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var caseId = reader.GetInt32("case_id");
                            var clientId = reader.GetInt32("client_id");
                            var ledgerCardId = reader.GetInt32("ledger_card_id");
                            
                            _logAction($"✓ Found case (manual ref): {caseReference} -> CaseId:{caseId}, ClientId:{clientId}, LedgerId:{ledgerCardId}");
                            return (caseId, clientId, ledgerCardId);
                        }
                    }
                }

                // Final fallback: partial match
                sql = @"
                    SELECT c.case_id, COALESCE(cc.fk_client_id, 0) as client_id, 
                           COALESCE(l.ledger_card_id, 0) as ledger_card_id
                    FROM tbl_case_details_general c
                    LEFT JOIN tbl_case_clients cc ON cc.fk_case_id = c.case_id
                    LEFT JOIN tbl_acc_ledger_cards l ON l.fk_case_id = c.case_id
                    WHERE c.case_reference_auto LIKE @casePattern
                       OR c.case_reference_manual LIKE @casePattern
                    LIMIT 1";

                using (var cmd = new MySqlCommand(sql, connection))
                {
                    cmd.Parameters.AddWithValue("@casePattern", $"%{caseReference}%");

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var caseId = reader.GetInt32("case_id");
                            var clientId = reader.GetInt32("client_id");
                            var ledgerCardId = reader.GetInt32("ledger_card_id");
                            
                            _logAction($"✓ Found case (partial match): {caseReference} -> CaseId:{caseId}, ClientId:{clientId}, LedgerId:{ledgerCardId}");
                            return (caseId, clientId, ledgerCardId);
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

        #endregion

        #region Import Operations

        public AccountUpdateImportResult ImportAccountUpdates(List<AccountUpdateImportData> records)
        {
            var result = new AccountUpdateImportResult();

            // Filter only valid and found records
            var validRecords = records.Where(r => r.IsValid && r.IsFound).ToList();

            if (validRecords.Count == 0)
            {
                _logAction("No valid records to import.");
                return result;
            }

            // Group by transaction type for processing order
            // Process in order: Receipts, Payments, C2O (which needs invoice first)
            var receipts = validRecords.Where(r => r.TransactionType == AccountTransactionType.BankReceipt).ToList();
            var payments = validRecords.Where(r => r.TransactionType == AccountTransactionType.BankPayment).ToList();
            var c2oTransfers = validRecords.Where(r => r.TransactionType == AccountTransactionType.ClientToOffice).ToList();

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                // Process Bank Receipts
                foreach (var receipt in receipts)
                {
                    try
                    {
                        ImportBankReceipt(connection, receipt);
                        result.ReceiptsCreated++;
                        result.SuccessCount++;
                        _logAction($"✅ Receipt imported: {receipt.CaseReference} - {receipt.Amount:C}");
                    }
                    catch (Exception ex)
                    {
                        result.ErrorCount++;
                        result.Errors.Add($"Row {receipt.RowNumber} (Receipt): {ex.Message}");
                        _logAction($"❌ Error importing receipt {receipt.CaseReference}: {ex.Message}");
                    }
                }

                // Process Bank Payments
                foreach (var payment in payments)
                {
                    try
                    {
                        ImportBankPayment(connection, payment);
                        result.PaymentsCreated++;
                        result.SuccessCount++;
                        _logAction($"✅ Payment imported: {payment.CaseReference} - {payment.Amount:C}");
                    }
                    catch (Exception ex)
                    {
                        result.ErrorCount++;
                        result.Errors.Add($"Row {payment.RowNumber} (Payment): {ex.Message}");
                        _logAction($"❌ Error importing payment {payment.CaseReference}: {ex.Message}");
                    }
                }

                // Process Client to Office Transfers (create invoice first if needed)
                foreach (var c2o in c2oTransfers)
                {
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            // Create Quick Invoice first
                            int invoiceId = CreateQuickInvoice(connection, transaction, c2o);
                            c2o.InvoiceId = invoiceId;
                            result.InvoicesCreated++;
                            _logAction($"✅ Invoice created: {c2o.InvoiceNumber} for {c2o.CaseReference}");

                            // Then create Client to Office transfer
                            ImportClientToOffice(connection, transaction, c2o);
                            result.ClientToOfficeCreated++;
                            result.SuccessCount++;
                            _logAction($"✅ C2O imported: {c2o.CaseReference} - {c2o.Amount:C}");

                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            result.ErrorCount++;
                            result.Errors.Add($"Row {c2o.RowNumber} (C2O): {ex.Message}");
                            _logAction($"❌ Error importing C2O {c2o.CaseReference}: {ex.Message}");
                        }
                    }
                }
            }

            return result;
        }

        #endregion

        #region Bank Receipt

        /// <summary>
        /// Bank Receipt:
        /// - Insert into Bank Receipt table
        /// - DR Client Bank (money in)
        /// - CR Client Ledger
        /// </summary>
        private void ImportBankReceipt(MySqlConnection connection, AccountUpdateImportData data)
        {
            using (var transaction = connection.BeginTransaction())
            {
                try
                {
                    // Generate receipt number
                    string receiptNumber = GenerateReceiptNumber(connection, transaction);

                    // Create transaction (TransactionTypeId=1, TransactionSubTypeId=1 for Bank Receipt)
                    int transactionId = CreateTransaction(connection, transaction,
                        transactionTypeId: 1,
                        transactionSubTypeId: 1, // Bank Receipt sub type
                        amount: data.Amount,
                        date: data.TransactionDate,
                        details: data.Description ?? "Bank Receipt",
                        reference: "REP " + receiptNumber);

                    // Insert client receipt (Bank Receipt Table)
                    int clientReceiptId = InsertClientReceipt(connection, transaction, data, transactionId, receiptNumber);

                    // Get bank current balance
                    decimal bankBalance = GetBankOpeningBalance(connection, transaction, data.ClientBankId);

                    // DR Client Bank (Debit - money in)
                    InsertClientBankTransaction(connection, transaction, transactionId, data.ClientBankId,
                        data.TransactionDate, data.Description ?? "Bank Receipt",
                        "REC " + receiptNumber, data.Amount, 0, bankBalance, bankBalance + data.Amount);

                    // CR Client Ledger
                    if (data.LedgerCardId > 0)
                    {
                        decimal prevClientBalance = GetPrevClientBalance(connection, transaction, data.LedgerCardId.Value);
                        decimal newClientBalance = prevClientBalance + data.Amount;

                        InsertLedgerCardTransaction(connection, transaction, transactionId, data.LedgerCardId.Value,
                            data.TransactionDate, data.Description ?? "Bank Receipt",
                            officeDr: 0, officeCr: 0, officeBal: 0,
                            clientDr: 0, clientCr: data.Amount, clientBal: newClientBalance,
                            data.PaymentReference ?? "REC " + receiptNumber);
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
        }

        private int InsertClientReceipt(MySqlConnection connection, MySqlTransaction transaction,
            AccountUpdateImportData data, int transactionId, string receiptNumber)
        {
            string sql = @"
                INSERT INTO tbl_acc_client_receipt (
                    receipt_number, receipt_date, fk_client_bank_id, fk_case_id, fk_client_id,
                    received_from, transaction_description, amount, authorised_by, payment_type_id,
                    payment_reference, comments, receipt_reference, fk_transaction_id, fk_branch_id,
                    entry_date, staff_id, is_cancelled, is_received_from_other, received_from_other, is_client
                ) VALUES (
                    @receiptNumber, @receiptDate, @clientBankId, @caseId, @clientId,
                    @receivedFrom, @description, @amount, @authorisedBy, @paymentTypeId,
                    @paymentReference, @comments, @receiptReference, @transactionId, @branchId,
                    @entryDate, @staffId, @isCancelled, @isReceivedFromOther, @receivedFromOther, @isClient
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@receiptNumber", receiptNumber);
                cmd.Parameters.AddWithValue("@receiptDate", data.TransactionDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@clientBankId", data.ClientBankId);
                cmd.Parameters.AddWithValue("@caseId", data.CaseId);
                cmd.Parameters.AddWithValue("@clientId", data.ClientId);
                cmd.Parameters.AddWithValue("@receivedFrom", 0); // Not from other
                cmd.Parameters.AddWithValue("@description", data.Description ?? "Bank Receipt");
                cmd.Parameters.AddWithValue("@amount", data.Amount);
                cmd.Parameters.AddWithValue("@authorisedBy", _userId);
                cmd.Parameters.AddWithValue("@paymentTypeId", _defaultPaymentTypeId); // Bank Transfer
                cmd.Parameters.AddWithValue("@paymentReference", data.PaymentReference ?? "");
                cmd.Parameters.AddWithValue("@comments", data.Comments ?? "Imported from Excel");
                cmd.Parameters.AddWithValue("@receiptReference", data.PaymentReference ?? "");
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@entryDate", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@staffId", _userId);
                cmd.Parameters.AddWithValue("@isCancelled", false);
                cmd.Parameters.AddWithValue("@isReceivedFromOther", false);
                cmd.Parameters.AddWithValue("@receivedFromOther", "");
                cmd.Parameters.AddWithValue("@isClient", true);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private string GenerateReceiptNumber(MySqlConnection connection, MySqlTransaction transaction)
        {
            string sql = "SELECT COALESCE(MAX(CAST(receipt_number AS UNSIGNED)), 0) + 1 FROM tbl_acc_client_receipt WHERE receipt_number REGEXP '^[0-9]+$'";
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                return cmd.ExecuteScalar()?.ToString() ?? "1";
            }
        }

        #endregion

        #region Bank Payment

        /// <summary>
        /// Bank Payment:
        /// - Insert into Bank Payment table
        /// - DR Client Ledger
        /// - CR Client Bank (money out)
        /// </summary>
        private void ImportBankPayment(MySqlConnection connection, AccountUpdateImportData data)
        {
            using (var transaction = connection.BeginTransaction())
            {
                try
                {
                    // Generate payment number
                    string paymentNumber = GeneratePaymentNumber(connection, transaction);

                    // Create transaction (TransactionTypeId=1, TransactionSubTypeId=5 for Bank Payment)
                    int transactionId = CreateTransaction(connection, transaction,
                        transactionTypeId: 1,
                        transactionSubTypeId: 5, // Bank Payment sub type
                        amount: data.Amount,
                        date: data.TransactionDate,
                        details: data.Description ?? "Bank Payment",
                        reference: "PAY " + paymentNumber);

                    // Insert client payment (Bank Payment Table)
                    int clientPaymentId = InsertClientPayment(connection, transaction, data, transactionId, paymentNumber);

                    // Get bank current balance
                    decimal bankBalance = GetBankOpeningBalance(connection, transaction, data.ClientBankId);

                    // CR Client Bank (Credit - money out)
                    InsertClientBankTransaction(connection, transaction, transactionId, data.ClientBankId,
                        data.TransactionDate, data.Description ?? "Bank Payment",
                        "PAY " + paymentNumber, 0, data.Amount, bankBalance, bankBalance - data.Amount);

                    // DR Client Ledger
                    if (data.LedgerCardId > 0)
                    {
                        decimal prevClientBalance = GetPrevClientBalance(connection, transaction, data.LedgerCardId.Value);
                        decimal newClientBalance = prevClientBalance - data.Amount;

                        InsertLedgerCardTransaction(connection, transaction, transactionId, data.LedgerCardId.Value,
                            data.TransactionDate, data.Description ?? "Bank Payment",
                            officeDr: 0, officeCr: 0, officeBal: 0,
                            clientDr: data.Amount, clientCr: 0, clientBal: newClientBalance,
                            data.PaymentReference ?? "PAY " + paymentNumber);
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
        }

        private int InsertClientPayment(MySqlConnection connection, MySqlTransaction transaction,
            AccountUpdateImportData data, int transactionId, string paymentNumber)
        {
            string sql = @"
                INSERT INTO tbl_acc_client_payment (
                    payment_number, payment_create_date, fk_client_bank_id, fk_case_id,
                    pay_to, transaction_description, amount, authorised_by, payment_method,
                    payment_reference, comments, fk_transaction_id, fk_branch_id,
                    entry_date, paid_by, is_cancelled, is_client
                ) VALUES (
                    @paymentNumber, @paymentDate, @clientBankId, @caseId,
                    @paidTo, @description, @amount, @authorisedBy, @paymentTypeId,
                    @paymentReference, @comments, @transactionId, @branchId,
                    @entryDate, @staffId, @isCancelled, @isClient
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@paymentNumber", paymentNumber);
                cmd.Parameters.AddWithValue("@paymentDate", data.TransactionDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@clientBankId", data.ClientBankId);
                cmd.Parameters.AddWithValue("@caseId", data.CaseId);
                cmd.Parameters.AddWithValue("@clientId", data.ClientId);
                cmd.Parameters.AddWithValue("@paidTo", data.PaidTo ?? "");
                cmd.Parameters.AddWithValue("@description", data.Description ?? "Bank Payment");
                cmd.Parameters.AddWithValue("@amount", data.Amount);
                cmd.Parameters.AddWithValue("@authorisedBy", _userId);
                cmd.Parameters.AddWithValue("@paymentTypeId", _defaultPaymentTypeId); // Bank Transfer
                cmd.Parameters.AddWithValue("@paymentReference", data.PaymentReference ?? "");
                cmd.Parameters.AddWithValue("@comments", data.Comments ?? "Imported from Excel");
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@entryDate", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@staffId", _userId);
                cmd.Parameters.AddWithValue("@isCancelled", false);
                cmd.Parameters.AddWithValue("@isClient", true);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private string GeneratePaymentNumber(MySqlConnection connection, MySqlTransaction transaction)
        {
            string sql = "SELECT COALESCE(MAX(CAST(payment_number AS UNSIGNED)), 0) + 1 FROM tbl_acc_client_payment WHERE payment_number REGEXP '^[0-9]+$'";
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                return cmd.ExecuteScalar()?.ToString() ?? "1";
            }
        }

        #endregion

        #region Client To Office

        /// <summary>
        /// Client to Office (C2O):
        /// - Create Simple Invoice (Professional Fee, NO VAT, NO split categories)
        ///   - tax_amount = 0
        ///   - total_amount = net_amount = invoice amount
        ///   - balance_due = 0 (immediately paid)
        ///   - status = 9 (fully paid)
        /// - NO nominal ledger transactions (tax related - skip)
        /// - NO correspondence
        /// - Insert C2O table record
        /// - Ledger Card Transaction (Client DR, Office CR)
        /// - NO office bank transaction (money stays in client bank)
        /// </summary>
        private int CreateQuickInvoice(MySqlConnection connection, MySqlTransaction transaction,
            AccountUpdateImportData data)
        {
            // Generate invoice number
            string invoiceNumber = data.InvoiceNumber ?? GenerateInvoiceNumber(connection, transaction);

            // Create transaction for invoice (TransactionTypeId=1, TransactionSubTypeId=6 for Invoice)
            int invoiceTransactionId = CreateTransaction(connection, transaction,
                transactionTypeId: 1,
                transactionSubTypeId: 6, // Invoice sub type
                amount: data.InvoiceAmount,
                date: data.TransactionDate,
                details: data.Description ?? "Quick Invoice",
                reference: invoiceNumber);

            // Get income account ID (default to 1 if not found)
            //int incomeAccountId = data.IncomeAccountId ?? GetDefaultIncomeAccountId(connection, transaction);
            int incomeAccountId = data.IncomeAccountId ?? 0;

            // Insert invoice
            string sql = @"
                INSERT INTO tbl_acc_invoice (
                    fk_branch_id, fk_case_id, invoice_number, finalised_on, due_date,
                    total, tax, amount, fk_transaction_id,
                    fk_current_status, entry_date, authorised_on, transaction_date, invoice_to, authorised_by, is_cancelled, authorisation_type,
                    finalised_comments
                ) VALUES (
                    @branchId, @caseId, @invoiceNumber, @invoiceDate, @dueDate,
                    @totalAmount, @taxAmount, @netAmount, @transactionId,
                    @currentStatus, @entryDate, @entryDate, @entryDate, @invoiceTo, @staffId, @isDeleted, @invoiceType,
                    @comments
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@caseId", data.CaseId);
                cmd.Parameters.AddWithValue("@invoiceNumber", invoiceNumber);
                cmd.Parameters.AddWithValue("@invoiceDate", data.TransactionDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@dueDate", data.TransactionDate.AddDays(30).ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@totalAmount", data.InvoiceAmount);
                cmd.Parameters.AddWithValue("@taxAmount", 0);
                cmd.Parameters.AddWithValue("@netAmount", data.InvoiceAmount);
                cmd.Parameters.AddWithValue("@balanceDue", 0); // Will be paid immediately by C2O
                cmd.Parameters.AddWithValue("@transactionId", invoiceTransactionId);
                cmd.Parameters.AddWithValue("@currentStatus", 9); // Paid status
                cmd.Parameters.AddWithValue("@entryDate", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@staffId", _userId);
                cmd.Parameters.AddWithValue("@isDeleted", false);
                cmd.Parameters.AddWithValue("@invoiceType", "Approved");
                cmd.Parameters.AddWithValue("@invoiceTo", data.PaidTo);
                //cmd.Parameters.AddWithValue("@incomeAccountId", incomeAccountId);
                cmd.Parameters.AddWithValue("@comments", data.Comments ?? "Imported invoice");

                int invoiceId = Convert.ToInt32(cmd.ExecuteScalar());

                // Save invoice status
                SaveInvoiceStatus(connection, transaction, invoiceId, 9);

                // Update data with invoice info
                data.InvoiceId = invoiceId;
                data.InvoiceNumber = invoiceNumber;

                return invoiceId;
            }
        }

        private void ImportClientToOffice(MySqlConnection connection, MySqlTransaction transaction,
            AccountUpdateImportData data)
        {
            // Generate C2O number
            string c2oNumber = GenerateC2ONumber(connection, transaction);

            // Create transaction for C2O (TransactionTypeId=1, TransactionSubTypeId=8 for Client to Office)
            int transactionId = CreateTransaction(connection, transaction,
                transactionTypeId: 1,
                transactionSubTypeId: 8, // Client to Office sub type
                amount: data.Amount,
                date: data.TransactionDate,
                details: $"INV {data.InvoiceNumber} Paid",
                reference: $"BIL {c2oNumber} INV {data.InvoiceNumber}");

            // Insert C2O transfer record
            string sql = @"
                INSERT INTO tbl_acc_client_to_office_transactions (
                    client_to_office_transfer_number, fk_client_bank_id, fk_office_bank_id,
                    transaction_date, amount_transfer, transaction_description, authorised_by,
                    transfer_method, transfer_method_reference_number, comments, fk_branch_id,
                    fk_transaction_id, fk_case_id, entry_date, transfer_type, fk_invoice_id,
                    fk_user_id, is_reversed
                ) VALUES (
                    @transferNumber, @clientBankId, @officeBankId,
                    @transactionDate, @amount, @description, @authorisedBy,
                    @transferMethod, @transferMethodRef, @comments, @branchId,
                    @transactionId, @caseId, @entryDate, @transferType, @invoiceId,
                    @userId, @isReversed
                );
                SELECT LAST_INSERT_ID();";

            int c2oId;
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@transferNumber", c2oNumber);
                cmd.Parameters.AddWithValue("@clientBankId", data.ClientBankId);
                cmd.Parameters.AddWithValue("@officeBankId", data.OfficeBankId);
                cmd.Parameters.AddWithValue("@transactionDate", data.TransactionDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@amount", data.Amount);
                cmd.Parameters.AddWithValue("@description", data.Description ?? "Client to Office Transfer");
                cmd.Parameters.AddWithValue("@authorisedBy", _userId);
                cmd.Parameters.AddWithValue("@transferMethod", "BACS");
                cmd.Parameters.AddWithValue("@transferMethodRef", data.PaymentReference ?? "");
                cmd.Parameters.AddWithValue("@comments", data.Comments ?? "");
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@caseId", data.CaseId);
                cmd.Parameters.AddWithValue("@entryDate", DateTime.UtcNow.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@transferType", "Invoice");
                cmd.Parameters.AddWithValue("@invoiceId", data.InvoiceId);
                cmd.Parameters.AddWithValue("@userId", _userId);
                cmd.Parameters.AddWithValue("@isReversed", false);

                c2oId = Convert.ToInt32(cmd.ExecuteScalar());
            }

            // Get client bank balance for the transaction
            decimal clientBankBalance = GetBankOpeningBalance(connection, transaction, data.ClientBankId);

            // Insert client bank transaction (Credit - money out from client bank)
            InsertClientBankTransaction(connection, transaction, transactionId, data.ClientBankId,
                data.TransactionDate, $"INV {data.InvoiceNumber} Paid",
                $"BIL {c2oId}", 0, data.Amount, clientBankBalance, clientBankBalance - data.Amount);

            // NOTE: No office bank transaction - we don't transfer money to office bank at any stage
            // The C2O only moves money from client to office ledger, not to office bank account

            // Insert ledger card transaction (Client DR, Office CR - moves money from client ledger to office ledger)
            if (data.LedgerCardId > 0)
            {
                decimal prevClientBalance = GetPrevClientBalance(connection, transaction, data.LedgerCardId.Value);
                decimal prevOfficeBalance = GetPrevOfficeBalance(connection, transaction, data.LedgerCardId.Value);
                
                decimal newClientBalance = prevClientBalance - data.Amount;  // Client balance decreases
                decimal newOfficeBalance = prevOfficeBalance + data.Amount;  // Office balance increases (CR)

                InsertLedgerCardTransaction(connection, transaction, transactionId, data.LedgerCardId.Value,
                    data.TransactionDate, $"INV {data.InvoiceNumber} Paid",
                    officeDr: 0, officeCr: data.Amount, officeBal: newOfficeBalance,
                    clientDr: data.Amount, clientCr: 0, clientBal: newClientBalance,
                    data.PaymentReference ?? $"BIL {c2oId}");
            }
        }

        private string GenerateC2ONumber(MySqlConnection connection, MySqlTransaction transaction)
        {
            string sql = "SELECT COALESCE(MAX(CAST(client_to_office_transfer_number AS UNSIGNED)), 0) + 1 FROM tbl_acc_client_to_office_transactions WHERE client_to_office_transfer_number REGEXP '^[0-9]+$'";
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                return cmd.ExecuteScalar()?.ToString() ?? "1";
            }
        }

        private string GenerateInvoiceNumber(MySqlConnection connection, MySqlTransaction transaction)
        {
            string sql = "SELECT COALESCE(MAX(CAST(invoice_number AS UNSIGNED)), 0) + 1 FROM tbl_acc_invoice WHERE invoice_number REGEXP '^[0-9]+$'";
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                return cmd.ExecuteScalar()?.ToString() ?? "1";
            }
        }

        private int GetDefaultIncomeAccountId(MySqlConnection connection, MySqlTransaction transaction)
        {
            string sql = "SELECT account_id FROM tbl_acc_nominal_accounts WHERE account_name LIKE '%Income%' OR account_name LIKE '%Revenue%' LIMIT 1";
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                var result = cmd.ExecuteScalar();
                return result != null ? Convert.ToInt32(result) : 1;
            }
        }

        private void SaveInvoiceStatus(MySqlConnection connection, MySqlTransaction transaction, int invoiceId, int statusId)
        {
            string sql = @"
                INSERT INTO tbl_acc_invoice_status (
                    fk_invoice_id, fk_branch_id, invoice_status_type_id, date_time, fk_user_id
                ) VALUES (
                    @invoiceId, @branchId, @statusId, @statusDate, @userId
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@invoiceId", invoiceId);
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@statusId", statusId);
                cmd.Parameters.AddWithValue("@statusDate", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@userId", _userId);
                cmd.ExecuteNonQuery();
            }
        }

        #endregion

        #region Common Database Operations

        private int CreateTransaction(MySqlConnection connection, MySqlTransaction transaction,
            int transactionTypeId, int transactionSubTypeId, decimal amount,
            DateTime date, string details, string reference)
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
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@typeId", transactionTypeId);
                cmd.Parameters.AddWithValue("@subTypeId", transactionSubTypeId);
                cmd.Parameters.AddWithValue("@details", details);
                cmd.Parameters.AddWithValue("@reference", reference);
                cmd.Parameters.AddWithValue("@amount", Math.Abs(amount));
                cmd.Parameters.AddWithValue("@isCancelled", false);
                cmd.Parameters.AddWithValue("@postBy", _userId);
                cmd.Parameters.AddWithValue("@postDateTime", DateTime.UtcNow);
                cmd.Parameters.AddWithValue("@transactionDate", date);
                cmd.Parameters.AddWithValue("@transactionType", "Add");

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void InsertClientBankTransaction(MySqlConnection connection, MySqlTransaction transaction,
            int transactionId, int clientBankId, DateTime transactionDate, string description,
            string reference, decimal drAmount, decimal crAmount, decimal balancePre, decimal balancePost)
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
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@clientBankId", clientBankId);
                cmd.Parameters.AddWithValue("@transactionDateTime", transactionDate);
                cmd.Parameters.AddWithValue("@transaction", description);
                cmd.Parameters.AddWithValue("@reference", reference);
                cmd.Parameters.AddWithValue("@details", description);
                cmd.Parameters.AddWithValue("@drAmount", drAmount);
                cmd.Parameters.AddWithValue("@crAmount", crAmount);
                cmd.Parameters.AddWithValue("@balancePre", balancePre);
                cmd.Parameters.AddWithValue("@balancePost", balancePost);
                cmd.Parameters.AddWithValue("@isCancelled", false);
                cmd.Parameters.AddWithValue("@isReconciled", false);
                cmd.Parameters.AddWithValue("@bankReconciliationId", DBNull.Value);

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertOfficeBankTransaction(MySqlConnection connection, MySqlTransaction transaction,
            int transactionId, int officeBankId, DateTime transactionDate, string description,
            string reference, decimal drAmount, decimal crAmount, decimal balancePre, decimal balancePost)
        {
            string sql = @"
                INSERT INTO tbl_acc_office_bank_transactions (
                    fk_branch_id, fk_transaction_id, fk_office_bank_id, transaction_date_time,
                    transaction, reference, details, dr_amount, cr_amount,
                    balance_pre, balance_post, is_cancelled, is_reconciled, fk_bank_reconciliation_id
                ) VALUES (
                    @branchId, @transactionId, @officeBankId, @transactionDateTime,
                    @transaction, @reference, @details, @drAmount, @crAmount,
                    @balancePre, @balancePost, @isCancelled, @isReconciled, @bankReconciliationId
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@officeBankId", officeBankId);
                cmd.Parameters.AddWithValue("@transactionDateTime", transactionDate);
                cmd.Parameters.AddWithValue("@transaction", description);
                cmd.Parameters.AddWithValue("@reference", reference);
                cmd.Parameters.AddWithValue("@details", description);
                cmd.Parameters.AddWithValue("@drAmount", drAmount);
                cmd.Parameters.AddWithValue("@crAmount", crAmount);
                cmd.Parameters.AddWithValue("@balancePre", balancePre);
                cmd.Parameters.AddWithValue("@balancePost", balancePost);
                cmd.Parameters.AddWithValue("@isCancelled", false);
                cmd.Parameters.AddWithValue("@isReconciled", false);
                cmd.Parameters.AddWithValue("@bankReconciliationId", DBNull.Value);

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertLedgerCardTransaction(MySqlConnection connection, MySqlTransaction transaction,
            int transactionId, int ledgerCardId, DateTime transactionDate, string details,
            decimal officeDr, decimal officeCr, decimal officeBal,
            decimal clientDr, decimal clientCr, decimal clientBal, string ledgerReference)
        {
            string officeBalType = officeBal >= 0 ? "CR" : "DR";
            string clientBalType = clientBal >= 0 ? "CR" : "DR";

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
                cmd.Parameters.AddWithValue("@branchId", _branchId);
                cmd.Parameters.AddWithValue("@transactionId", transactionId);
                cmd.Parameters.AddWithValue("@ledgerCardId", ledgerCardId);
                cmd.Parameters.AddWithValue("@transactionDateTime", transactionDate);
                cmd.Parameters.AddWithValue("@details", details);
                cmd.Parameters.AddWithValue("@officeDr", officeDr);
                cmd.Parameters.AddWithValue("@officeCr", officeCr);
                cmd.Parameters.AddWithValue("@officeBal", Math.Abs(officeBal));
                cmd.Parameters.AddWithValue("@officeBalType", officeBalType);
                cmd.Parameters.AddWithValue("@clientDr", clientDr);
                cmd.Parameters.AddWithValue("@clientCr", clientCr);
                cmd.Parameters.AddWithValue("@clientBal", Math.Abs(clientBal));
                cmd.Parameters.AddWithValue("@clientBalType", clientBalType);
                cmd.Parameters.AddWithValue("@total", Math.Abs(clientBal) + Math.Abs(officeBal));
                cmd.Parameters.AddWithValue("@ledgerReference", ledgerReference);
                cmd.Parameters.AddWithValue("@isCancelled", false);

                cmd.ExecuteNonQuery();
            }
        }

        private decimal GetBankOpeningBalance(MySqlConnection connection, MySqlTransaction transaction, int bankId)
        {
            //string sql = @"
            //    SELECT COALESCE(
            //        (SELECT balance_post FROM tbl_acc_client_bank_transactions 
            //         WHERE fk_client_bank_id = @bankId AND is_cancelled = 0 
            //         ORDER BY client_bank_transaction_id DESC LIMIT 1),
            //        (SELECT  FROM tbl_acc_bank_account WHERE bank_account_id = @bankId),
            //        0
            //    ) as balance";

            string sql = @"
                SELECT COALESCE(
                    (SELECT balance_post FROM tbl_acc_client_bank_transactions 
                     WHERE fk_client_bank_id = @bankId AND is_cancelled = 0 
                     ORDER BY client_bank_transaction_id DESC LIMIT 1),
                    0
                ) as balance";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@bankId", bankId);
                var result = cmd.ExecuteScalar();
                return result != null && result != DBNull.Value ? Convert.ToDecimal(result) : 0;
            }
        }

        private decimal GetOfficeBankOpeningBalance(MySqlConnection connection, MySqlTransaction transaction, int bankId)
        {
            string sql = @"
                SELECT COALESCE(
                    (SELECT balance_post FROM tbl_acc_office_bank_transactions 
                     WHERE fk_office_bank_id = @bankId AND is_cancelled = 0 
                     ORDER BY office_bank_transaction_id DESC LIMIT 1),
                    (SELECT opening_balance FROM tbl_acc_office_bank_accounts WHERE office_bank_id = @bankId),
                    0
                ) as balance";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@bankId", bankId);
                var result = cmd.ExecuteScalar();
                return result != null && result != DBNull.Value ? Convert.ToDecimal(result) : 0;
            }
        }

        private decimal GetPrevClientBalance(MySqlConnection connection, MySqlTransaction transaction, int ledgerCardId)
        {
            string sql = @"
                SELECT COALESCE(
                    (SELECT client_bal FROM tbl_acc_ledger_card_transactions 
                     WHERE fk_ledger_card_id = @ledgerCardId AND is_cancelled = 0 
                     ORDER BY ledger_card_transaction_id DESC LIMIT 1),
                    0
                ) as balance";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@ledgerCardId", ledgerCardId);
                var result = cmd.ExecuteScalar();
                return result != null && result != DBNull.Value ? Convert.ToDecimal(result) : 0;
            }
        }

        private decimal GetPrevOfficeBalance(MySqlConnection connection, MySqlTransaction transaction, int ledgerCardId)
        {
            string sql = @"
                SELECT COALESCE(
                    (SELECT office_bal FROM tbl_acc_ledger_card_transactions 
                     WHERE fk_ledger_card_id = @ledgerCardId AND is_cancelled = 0 
                     ORDER BY ledger_card_transaction_id DESC LIMIT 1),
                    0
                ) as balance";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@ledgerCardId", ledgerCardId);
                var result = cmd.ExecuteScalar();
                return result != null && result != DBNull.Value ? Convert.ToDecimal(result) : 0;
            }
        }

        private void HandlePaymentType(MySqlConnection connection, MySqlTransaction transaction,
            int transactionId, int clientReceiptId, AccountUpdateImportData data)
        {
            // Only handle specific payment types if needed
            if (data.PaymentTypeId == 2) // Cash
            {
                string sql = @"
                    INSERT INTO tbl_acc_payment_cash (
                        fk_branch_id, fk_transaction_id, fk_client_receipt_id, memo
                    ) VALUES (
                        @branchId, @transactionId, @clientReceiptId, @memo
                    )";

                using (var cmd = new MySqlCommand(sql, connection, transaction))
                {
                    cmd.Parameters.AddWithValue("@branchId", _branchId);
                    cmd.Parameters.AddWithValue("@transactionId", transactionId);
                    cmd.Parameters.AddWithValue("@clientReceiptId", clientReceiptId);
                    cmd.Parameters.AddWithValue("@memo", data.Comments ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        #endregion

        #region Bank Account Lookup

        public List<BankAccountInfo> GetClientBankAccounts()
        {
            var accounts = new List<BankAccountInfo>();

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                // Using tbl_acc_bank_account with bank_account_type = 'Client'
                string sql = @"
                    SELECT bank_account_id, institution, account_name, account_number, account_sort_code
                    FROM tbl_acc_bank_account
                    WHERE bank_account_type = 'Client' AND is_active = 1
                    ORDER BY account_name";

                using (var cmd = new MySqlCommand(sql, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        accounts.Add(new BankAccountInfo
                        {
                            BankId = reader.GetInt32("bank_account_id"),
                            BankName = reader.IsDBNull(reader.GetOrdinal("account_name")) ? "" : reader.GetString("account_name"),
                            AccountNumber = reader.IsDBNull(reader.GetOrdinal("account_number")) ? "" : reader.GetString("account_number"),
                            SortCode = reader.IsDBNull(reader.GetOrdinal("account_sort_code")) ? "" : reader.GetString("account_sort_code"),
                            Institution = reader.IsDBNull(reader.GetOrdinal("institution")) ? "" : reader.GetString("institution"),
                            IsClientBank = true,
                            OpeningBalance = 0
                        });
                    }
                }
            }

            _logAction($"Found {accounts.Count} client bank accounts");
            return accounts;
        }

        public List<BankAccountInfo> GetOfficeBankAccounts()
        {
            var accounts = new List<BankAccountInfo>();

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                // Using tbl_acc_bank_account with bank_account_type = 'Office'
                string sql = @"
                    SELECT bank_account_id, institution, account_name, account_number, account_sort_code
                    FROM tbl_acc_bank_account
                    WHERE bank_account_type = 'Office' AND is_active = 1
                    ORDER BY account_name";

                using (var cmd = new MySqlCommand(sql, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        accounts.Add(new BankAccountInfo
                        {
                            BankId = reader.GetInt32("bank_account_id"),
                            BankName = reader.IsDBNull(reader.GetOrdinal("account_name")) ? "" : reader.GetString("account_name"),
                            AccountNumber = reader.IsDBNull(reader.GetOrdinal("account_number")) ? "" : reader.GetString("account_number"),
                            SortCode = reader.IsDBNull(reader.GetOrdinal("account_sort_code")) ? "" : reader.GetString("account_sort_code"),
                            Institution = reader.IsDBNull(reader.GetOrdinal("institution")) ? "" : reader.GetString("institution"),
                            IsClientBank = false,
                            OpeningBalance = 0
                        });
                    }
                }
            }

            _logAction($"Found {accounts.Count} office bank accounts");
            return accounts;
        }

        #endregion
    }
}
