using ClosedXML.Excel;
using LeapMergeDoc.Models;
using MySql.Data.MySqlClient;
using System.Globalization;
using System.IO;

namespace LeapMergeDoc.Services
{
    public class DatabaseImportService
    {
        private readonly string _connectionString;
        private readonly MatterTypeMatcher _matterTypeMatcher;
        private readonly Action<string> _logAction;
        private Dictionary<string, ClientMasterRecord>? _clientMasterData;

        public DatabaseImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _matterTypeMatcher = new MatterTypeMatcher();
            _logAction = logAction;
        }

        /// <summary>
        /// Loads the client master CSV file that contains the actual client for each Client No.
        /// The CSV should have Client No in the first column and client details.
        /// </summary>
        public int LoadClientMasterCsv(string filePath)
        {
            _clientMasterData = new Dictionary<string, ClientMasterRecord>(StringComparer.OrdinalIgnoreCase);
            int loaded = 0;

            using (var reader = new StreamReader(filePath))
            {
                // Read header
                string? header = reader.ReadLine();
                if (header == null) return 0;

                // Parse: Client No,Title,Initials,Forename,Surname,...
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var values = ParseCsvLine(line);
                    if (values.Count < 5) continue;

                    string clientNo = values[0].Trim();
                    if (string.IsNullOrEmpty(clientNo)) continue;

                    // Only take the first entry per Client No (the actual client)
                    if (_clientMasterData.ContainsKey(clientNo)) continue;

                    var record = new ClientMasterRecord
                    {
                        ClientNo = clientNo,
                        Title = values.Count > 1 ? values[1].Trim() : "",
                        Initials = values.Count > 2 ? values[2].Trim() : "",
                        Forename = values.Count > 3 ? values[3].Trim() : "",
                        Surname = values.Count > 4 ? values[4].Trim() : ""
                    };

                    _clientMasterData[clientNo] = record;
                    loaded++;
                }
            }

            _logAction($"Loaded {loaded} client master records from {Path.GetFileName(filePath)}");
            return loaded;
        }

        /// <summary>
        /// Gets the client master record for a given Client No.
        /// </summary>
        public ClientMasterRecord? GetClientMaster(string clientNo)
        {
            if (_clientMasterData == null || string.IsNullOrEmpty(clientNo))
                return null;

            _clientMasterData.TryGetValue(clientNo, out var record);
            return record;
        }

        public (int rowsDeleted, string message) TruncateImportedData()
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
                    // Delete from related tables first (child tables)
                    var tablesToTruncate = new[]
                    {
                        "tbl_case_client_greeting",
                        "tbl_case_clients",
                        "tbl_case_permissions",
                        "tbl_acc_ledger_cards",
                        "tbl_case_details_general"
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

            return (totalDeleted, $"Successfully deleted {totalDeleted} total rows from case tables.");
        }

        public List<CaseExcelData> ReadExcelData(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLower();
            if (ext == ".csv")
            {
                return ReadCsvData(filePath);
            }
            return ReadExcelDataInternal(filePath);
        }

        private List<CaseExcelData> ReadCsvData(string filePath)
        {
            var data = new List<CaseExcelData>();

            var lines = File.ReadAllLines(filePath);
            if (lines.Length == 0)
            {
                _logAction("No data found in CSV file.");
                return data;
            }

            _logAction($"Found {lines.Length - 1} data rows in CSV file.");

            // Parse headers
            var headers = ParseCsvLine(lines[0]).Select(h => h.Trim()).ToList();
            _logAction($"Headers: {string.Join(", ", headers)}");

            var normalizedHeaders = headers.Select(h => h.ToLower().Replace("_", "").Replace(" ", "").Replace(".", "")).ToList();
            _logAction($"Normalized: {string.Join(", ", normalizedHeaders)}");

            // Parse data rows
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;

                var values = ParseCsvLine(lines[i]);
                var rowData = new CaseExcelData();

                for (int col = 0; col < Math.Min(headers.Count, values.Count); col++)
                {
                    var cellValue = values[col].Trim();
                    var header = normalizedHeaders[col];

                    MapCellToRowData(rowData, header, cellValue);
                }

                if (!string.IsNullOrEmpty(rowData.ClientNo) ||
                    !string.IsNullOrEmpty(rowData.CaseName) ||
                    !string.IsNullOrEmpty(rowData.MatterNo) ||
                    !string.IsNullOrEmpty(rowData.Surname))
                {
                    data.Add(rowData);
                }
            }

            return data;
        }

        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var current = new System.Text.StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
            result.Add(current.ToString());

            return result;
        }

        private void MapCellToRowData(CaseExcelData rowData, string header, string cellValue)
        {
            switch (header)
            {
                // New LEAP format columns
                case "clientno":
                case "client no":
                    rowData.ClientNo = cellValue;
                    break;
                case "matter":
                    rowData.MatterNumber = cellValue;
                    break;
                case "surname":
                    rowData.Surname = cellValue;
                    rowData.ClientLastName = cellValue;
                    if (string.IsNullOrEmpty(rowData.ClientName))
                        rowData.ClientName = cellValue;
                    break;
                case "forename":
                    rowData.Forename = cellValue;
                    rowData.ClientFirstName = cellValue;
                    if (!string.IsNullOrEmpty(rowData.Surname))
                        rowData.ClientName = $"{cellValue} {rowData.Surname}".Trim();
                    break;
                case "fe":
                case "f/e":
                    rowData.FeeEarnerCode = cellValue;
                    break;
                case "workid":
                case "work id":
                    rowData.WorkId = cellValue;
                    break;

                // Original mappings
                case "matterno":
                case "matternumber":
                case "matterref":
                    rowData.MatterNo = cellValue;
                    break;
                case "casename":
                case "client":
                case "clientname":
                    rowData.CaseName = cellValue;
                    rowData.ClientName = cellValue;
                    break;
                case "instructiondate":
                case "dateopened":
                case "opened":
                case "opendate":
                    rowData.DateOpened = ParseDate(cellValue);
                    break;
                case "mattertype":
                case "areaofpractice":
                case "area":
                case "type":
                    rowData.MatterType = cellValue;
                    break;
                case "matterdescription":
                case "description":
                case "matterdesc":
                    rowData.MatterDescription = cellValue;
                    break;
                case "archivedate":
                case "archived":
                case "archiveddate":
                    rowData.ArchiveDate = ParseDate(cellValue);
                    break;
                case "staffresp":
                case "personresponsible":
                case "responsible":
                case "feeearner":
                    rowData.StaffRespName = cellValue;
                    break;
                case "staffact":
                case "personacting":
                case "acting":
                    rowData.StaffActName = cellValue;
                    break;
                case "staffassist":
                case "personassisting":
                case "assisting":
                case "assistant":
                    rowData.StaffAssistName = cellValue;
                    break;
                case "credit":
                case "casecredit":
                case "broughtby":
                    rowData.CreditName = cellValue;
                    break;
            }
        }

        private List<CaseExcelData> ReadExcelDataInternal(string filePath)
        {
            var data = new List<CaseExcelData>();

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

                // Log normalized headers for debugging
                var normalizedHeaders = headers.Select(h => h.ToLower().Replace("_", "").Replace(" ", "").Replace(".", "")).ToList();
                _logAction($"Normalized: {string.Join(", ", normalizedHeaders)}");

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new CaseExcelData();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetString().Trim();
                        var header = normalizedHeaders[col - 1];

                        MapCellToRowData(rowData, header, cellValue);
                    }

                    // Include row if it has ClientNo, MatterNo, CaseName, or Surname
                    if (!string.IsNullOrEmpty(rowData.ClientNo) ||
                        !string.IsNullOrEmpty(rowData.CaseName) ||
                        !string.IsNullOrEmpty(rowData.MatterNo) ||
                        !string.IsNullOrEmpty(rowData.Surname))
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

        public List<ProcessedCaseData> ProcessExcelData(List<CaseExcelData> excelData)
        {
            var processedData = new List<ProcessedCaseData>();

            // Track not found items for summary
            var notFoundFeeEarners = new HashSet<string>();
            var notFoundWorkIds = new HashSet<string>();

            // Load users from database for staff mapping
            var userLookup = LoadUsersFromDatabase();
            _logAction($"Loaded {userLookup.Count} users from database for staff mapping.");

            // Load case types from database for Work Id mapping
            var caseTypeLookup = LoadCaseTypesFromDatabase();
            _logAction($"Loaded {caseTypeLookup.Count} case types from database for Work Id mapping.");

            // Group rows by Client No + Matter Number (same case can have client + contacts)
            var caseGroups = excelData
                .Where(r => !string.IsNullOrEmpty(r.ClientNo) && !string.IsNullOrEmpty(r.MatterNumber))
                .GroupBy(r => $"{r.ClientNo}-{r.MatterNumber}")
                .ToList();

            // Also include rows without ClientNo/MatterNumber as individual cases
            var standaloneRows = excelData
                .Where(r => string.IsNullOrEmpty(r.ClientNo) || string.IsNullOrEmpty(r.MatterNumber))
                .ToList();

            _logAction($"Found {caseGroups.Count} case groups + {standaloneRows.Count} standalone rows");

            int caseNumber = 1;

            // Process grouped cases (with Client No + Matter Number)
            foreach (var group in caseGroups)
            {
                var rows = group.ToList();
                
                // Find the client row (empty Forename = company client)
                // If no company row, use the first row
                var clientRow = rows.FirstOrDefault(r => string.IsNullOrEmpty(r.Forename)) ?? rows.First();
                
                // Other rows are contacts
                var contactRows = rows.Where(r => r != clientRow && !string.IsNullOrEmpty(r.Forename)).ToList();

                string caseReferenceAuto = group.Key;

                // Extract contacts (note: Title/Email/Phone might not be in case CSV)
                var contacts = contactRows.Select(c => new CaseContactInfo
                {
                    Forename = c.Forename,
                    Surname = c.Surname
                }).ToList();

                if (contactRows.Count > 0)
                {
                    _logAction($"Case {caseReferenceAuto}: Client='{clientRow.Surname}', Contacts={string.Join(", ", contactRows.Select(c => $"{c.Forename} {c.Surname}"))}");
                }

                var caseData = ProcessSingleCase(clientRow, caseReferenceAuto, caseNumber, userLookup, caseTypeLookup, notFoundFeeEarners, notFoundWorkIds);
                caseData.Contacts = contacts;
                processedData.Add(caseData);
                caseNumber++;
            }

            // Process standalone rows
            foreach (var row in standaloneRows)
            {
                string? caseReferenceAuto = row.MatterNo;
                var caseData = ProcessSingleCase(row, caseReferenceAuto, caseNumber, userLookup, caseTypeLookup, notFoundFeeEarners, notFoundWorkIds);
                processedData.Add(caseData);
                caseNumber++;
            }

            // Log summary of not found items for manual review
            if (notFoundFeeEarners.Count > 0)
            {
                _logAction($"\n========== NOT FOUND FEE EARNER CODES ({notFoundFeeEarners.Count}) ==========");
                foreach (var code in notFoundFeeEarners.OrderBy(x => x))
                {
                    _logAction($"  F/E Code: '{code}'");
                }
                _logAction("=======================================================\n");
            }

            if (notFoundWorkIds.Count > 0)
            {
                _logAction($"\n========== NOT FOUND WORK IDs ({notFoundWorkIds.Count}) ==========");
                foreach (var workId in notFoundWorkIds.OrderBy(x => x))
                {
                    _logAction($"  Work Id: '{workId}'");
                }
                _logAction("=======================================================\n");
            }

            return processedData;
        }

        private ProcessedCaseData ProcessSingleCase(
            CaseExcelData row,
            string? caseReferenceAuto,
            int caseNumber,
            Dictionary<string, int> userLookup,
            Dictionary<string, (int CaseTypeId, int AreaOfPracticeId)> caseTypeLookup,
            HashSet<string> notFoundFeeEarners,
            HashSet<string> notFoundWorkIds)
        {
            // Lookup fee earner by code (F/E column) or by name
            int? feeEarnerId = null;
            if (!string.IsNullOrEmpty(row.FeeEarnerCode))
            {
                feeEarnerId = LookupUserId(row.FeeEarnerCode, userLookup);
                if (feeEarnerId == null)
                {
                    notFoundFeeEarners.Add(row.FeeEarnerCode);
                    _logAction($"⚠️ Fee Earner code not found: '{row.FeeEarnerCode}' for case {caseReferenceAuto}");
                }
            }
            else if (!string.IsNullOrEmpty(row.StaffRespName))
            {
                feeEarnerId = LookupUserId(row.StaffRespName, userLookup);
            }

            // F/E maps to both PersonResponsible and PersonActing
            int? personResponsible = feeEarnerId;
            int? personActing = feeEarnerId;
            int? personAssisting = LookupUserId(row.StaffAssistName, userLookup);
            int? caseCredit = LookupUserId(row.CreditName, userLookup);

            // Lookup Work Id to get AreaOfPractice and CaseType
            int? areaOfPracticeId = null;
            int? caseTypeId = null;

            if (!string.IsNullOrEmpty(row.WorkId) && caseTypeLookup.TryGetValue(row.WorkId.Trim(), out var caseTypeInfo))
            {
                caseTypeId = caseTypeInfo.CaseTypeId;
                areaOfPracticeId = caseTypeInfo.AreaOfPracticeId > 0 ? caseTypeInfo.AreaOfPracticeId : null;
            }
            else if (!string.IsNullOrEmpty(row.WorkId))
            {
                notFoundWorkIds.Add(row.WorkId);
                _logAction($"⚠️ Work Id not found in database: '{row.WorkId}' for case {caseReferenceAuto}");
                var matterMatch = _matterTypeMatcher.MatchMatterType(row.MatterType);
                areaOfPracticeId = matterMatch.AreaOfPracticeId;
                caseTypeId = matterMatch.CaseTypeId;
            }
            else
            {
                var matterMatch = _matterTypeMatcher.MatchMatterType(row.MatterType);
                areaOfPracticeId = matterMatch.AreaOfPracticeId;
                caseTypeId = matterMatch.CaseTypeId;
            }

            return new ProcessedCaseData
            {
                OriginalData = row,
                FkBranchId = 1,
                FkAreaOfPracticeId = areaOfPracticeId,
                FkCaseTypeId = caseTypeId,
                FkCaseSubTypeId = null,
                CaseReferenceAuto = caseReferenceAuto,
                CaseNumber = caseNumber,
                CaseName = !string.IsNullOrEmpty(row.MatterDescription) ? row.MatterDescription : row.CaseName,
                DateOpened = row.DateOpened,
                PersonOpened = 1,
                PersonResponsible = personResponsible,
                PersonActing = personActing,
                PersonAssisting = personAssisting,
                CaseCredit = caseCredit,
                IsCaseActive = row.ArchiveDate == null,
                IsCaseArchived = row.ArchiveDate != null,
                IsCaseNotProceeding = false,
                MnlCheck = true,
                ConfSearch = true
            };
        }

        private Dictionary<string, int> LoadUsersFromDatabase()
        {
            var userLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                string sql = @"
                    SELECT user_id, user_code, first_name, last_name, 
                           CONCAT(COALESCE(first_name, ''), ' ', COALESCE(last_name, '')) as full_name
                    FROM tbl_user 
                    WHERE is_deleted = 0 OR is_deleted IS NULL";

                using (var cmd = new MySqlCommand(sql, connection))
                {
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int userId = reader.GetInt32("user_id");
                            string userCode = reader.IsDBNull(reader.GetOrdinal("user_code")) ? "" : reader.GetString("user_code");
                            string firstName = reader.IsDBNull(reader.GetOrdinal("first_name")) ? "" : reader.GetString("first_name");
                            string lastName = reader.IsDBNull(reader.GetOrdinal("last_name")) ? "" : reader.GetString("last_name");
                            string fullName = reader.IsDBNull(reader.GetOrdinal("full_name")) ? "" : reader.GetString("full_name");

                            // Add user_code as lookup key (for F/E column)
                            if (!string.IsNullOrEmpty(userCode) && !userLookup.ContainsKey(userCode.Trim()))
                                userLookup[userCode.Trim()] = userId;

                            // Add multiple lookup keys for the same user
                            if (!string.IsNullOrEmpty(fullName) && !userLookup.ContainsKey(fullName.Trim()))
                                userLookup[fullName.Trim()] = userId;

                            // Also add first name + last name without extra spaces
                            string normalizedName = $"{firstName} {lastName}".Trim();
                            if (!string.IsNullOrEmpty(normalizedName) && !userLookup.ContainsKey(normalizedName))
                                userLookup[normalizedName] = userId;

                            // Add last name only as fallback
                            if (!string.IsNullOrEmpty(lastName) && !userLookup.ContainsKey(lastName.Trim()))
                                userLookup[lastName.Trim()] = userId;
                        }
                    }
                }
            }

            // Add special mappings
            userLookup["Admin Admin"] = 3;  // Yara KHALIFA
            userLookup["Admin"] = 3;
            userLookup["BAIS Office"] = 3;  // Yara KHALIFA

            return userLookup;
        }

        private int? LookupUserId(string? staffName, Dictionary<string, int> userLookup)
        {
            if (string.IsNullOrWhiteSpace(staffName))
                return null;

            string trimmedName = staffName.Trim();

            // Check for Unassigned
            if (trimmedName.Equals("Unassigned", StringComparison.OrdinalIgnoreCase) ||
                trimmedName.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
                trimmedName.Equals("-", StringComparison.OrdinalIgnoreCase))
                return null;

            // Try exact match
            if (userLookup.TryGetValue(trimmedName, out int userId))
                return userId;

            // Try matching last name only
            var nameParts = trimmedName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (nameParts.Length >= 1)
            {
                string lastName = nameParts[^1]; // Last word
                if (userLookup.TryGetValue(lastName, out userId))
                    return userId;
            }

            // Try matching first name + last name (handles middle names)
            if (nameParts.Length >= 2)
            {
                string firstLast = $"{nameParts[0]} {nameParts[^1]}";
                if (userLookup.TryGetValue(firstLast, out userId))
                    return userId;
            }

            _logAction($"⚠️ Staff not found: '{trimmedName}'");
            return null;
        }

        private Dictionary<string, (int CaseTypeId, int AreaOfPracticeId)> LoadCaseTypesFromDatabase()
        {
            var caseTypeLookup = new Dictionary<string, (int CaseTypeId, int AreaOfPracticeId)>(StringComparer.OrdinalIgnoreCase);

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                string sql = @"
                    SELECT case_type_id, fk_areas_of_practice_id, case_type_code
                    FROM tbl_case_type 
                    WHERE is_active = 1 OR is_active IS NULL";

                using (var cmd = new MySqlCommand(sql, connection))
                {
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int caseTypeId = reader.GetInt32("case_type_id");
                            int areaOfPracticeId = reader.IsDBNull(reader.GetOrdinal("fk_areas_of_practice_id")) ? 0 : reader.GetInt32("fk_areas_of_practice_id");
                            string caseTypeCode = reader.IsDBNull(reader.GetOrdinal("case_type_code")) ? "" : reader.GetString("case_type_code");

                            if (!string.IsNullOrEmpty(caseTypeCode) && !caseTypeLookup.ContainsKey(caseTypeCode.Trim()))
                            {
                                caseTypeLookup[caseTypeCode.Trim()] = (caseTypeId, areaOfPracticeId);
                            }
                        }
                    }
                }
            }

            return caseTypeLookup;
        }


        public (int found, int notFound) CheckClientMappings(List<ProcessedCaseData> processedData)
        {
            int foundClients = 0;
            int notFoundClients = 0;

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                foreach (var caseData in processedData)
                {
                    // Use client master lookup if available
                    ClientInfo? clientInfo = null;
                    
                    if (_clientMasterData != null && caseData.OriginalData != null)
                    {
                        // Extract Client No from case reference (e.g., "SSN0001-1" -> "SSN0001")
                        string clientNo = ExtractClientNoFromCaseRef(caseData.CaseReferenceAuto ?? caseData.OriginalData.ClientNo ?? "");
                        
                        if (!string.IsNullOrEmpty(clientNo))
                        {
                            var masterRecord = GetClientMaster(clientNo);
                            if (masterRecord != null)
                            {
                                // Search for the actual client from master record
                                clientInfo = FindClientInfoByMasterRecord(connection, masterRecord);
                                _logAction($"Client master lookup for {clientNo}: {(clientInfo != null ? clientInfo.FullName : "NOT FOUND")} (IsCompany: {masterRecord.IsCompany})");
                            }
                        }
                    }
                    
                    // Fallback to original logic if no client master or not found
                    if (clientInfo == null)
                    {
                        clientInfo = FindClientInfo(connection, caseData.OriginalData!);
                    }
                    
                    caseData.LinkedClientId = clientInfo?.ClientId;
                    caseData.ClientFullName = clientInfo?.FullName;

                    if (clientInfo != null)
                        foundClients++;
                    else
                        notFoundClients++;
                }
            }

            return (foundClients, notFoundClients);
        }

        /// <summary>
        /// Extracts Client No from case reference. E.g., "SSN0001-1" -> "SSN0001"
        /// </summary>
        private string ExtractClientNoFromCaseRef(string caseRef)
        {
            if (string.IsNullOrEmpty(caseRef))
                return "";

            // Case reference format: {ClientNo}-{MatterNumber}
            int dashIndex = caseRef.LastIndexOf('-');
            if (dashIndex > 0)
            {
                return caseRef.Substring(0, dashIndex).Trim();
            }
            return caseRef.Trim();
        }

        /// <summary>
        /// Finds client in database using the client master record.
        /// </summary>
        private ClientInfo? FindClientInfoByMasterRecord(MySqlConnection connection, ClientMasterRecord masterRecord)
        {
            if (masterRecord.IsCompany)
            {
                // Company client - search by company name (Surname holds company name when no Forename)
                if (string.IsNullOrEmpty(masterRecord.Surname))
                    return null;
                return SearchCompanyByName(connection, masterRecord.Surname);
            }
            else
            {
                // Individual client - search by forename and surname
                if (string.IsNullOrEmpty(masterRecord.Forename) || string.IsNullOrEmpty(masterRecord.Surname))
                    return null;

                var result = SearchIndividualByNames(connection, masterRecord.Forename, masterRecord.Surname);
                if (result != null) return result;

                // Fallback to last name only
                result = SearchIndividualByLastName(connection, masterRecord.Surname);
                if (result != null) return result;

                // Final fallback - search by full name
                string fullName = $"{masterRecord.Forename} {masterRecord.Surname}".Trim();
                result = SearchIndividualByFullName(connection, fullName);
                return result;
            }
        }

        private ClientInfo? FindClientInfo(MySqlConnection connection, CaseExcelData caseData)
        {
            ClientInfo? result = null;

            // If Surname is provided directly (new LEAP format), use it
            if (!string.IsNullOrEmpty(caseData.Surname))
            {
                string surname = caseData.Surname.Trim();
                string? forename = caseData.Forename?.Trim();

                // If no forename, it's a company
                if (string.IsNullOrEmpty(forename))
                {
                    // Search for company by name
                    result = SearchCompanyByName(connection, surname);
                    if (result != null) return result;
                    
                    // Also try individual search by last name (surname might be person without forename)
                    result = SearchIndividualByLastName(connection, surname);
                    if (result != null) return result;
                }
                else
                {
                    // Individual client - search by forename and surname
                    result = SearchIndividualByNames(connection, forename, surname);
                    if (result != null) return result;

                    // Fallback to last name only
                    result = SearchIndividualByLastName(connection, surname);
                    if (result != null) return result;
                }

                // Final fallback - search by full name
                string fullName = string.IsNullOrEmpty(forename) ? surname : $"{forename} {surname}";
                result = SearchIndividualByFullName(connection, fullName);
                if (result != null) return result;

                result = SearchCompanyByName(connection, fullName);
                return result;
            }

            // Fallback to original logic using ClientName
            string clientName = caseData.ClientName ?? "";

            if (string.IsNullOrEmpty(clientName))
                return null;

            // Strip title from client name for better matching
            var (givenNames, lastName) = ExtractNamesFromClientName(clientName);
            string nameWithoutTitle = $"{givenNames} {lastName}".Trim();

            // Try multiple search strategies
            // Strategy 1: Search by given names AND last name (most accurate)
            if (!string.IsNullOrEmpty(givenNames) && !string.IsNullOrEmpty(lastName))
            {
                result = SearchIndividualByNames(connection, givenNames, lastName);
                if (result != null) return result;
            }

            // Strategy 2: Search by last name only
            if (!string.IsNullOrEmpty(lastName))
            {
                result = SearchIndividualByLastName(connection, lastName);
                if (result != null) return result;
            }

            // Strategy 3: Search by full name without title (LIKE match)
            if (!string.IsNullOrEmpty(nameWithoutTitle))
            {
                result = SearchIndividualByFullName(connection, nameWithoutTitle);
                if (result != null) return result;
            }

            // Strategy 4: Search in company clients
            result = SearchCompanyByName(connection, clientName);

            return result;
        }

        private (string givenNames, string lastName) ExtractNamesFromClientName(string clientName)
        {
            if (string.IsNullOrWhiteSpace(clientName))
                return ("", "");

            // Known titles to strip
            var titles = new[] {
                "MR.", "MRS.", "MISS.", "MS.", "DR.", "PROF.", "MASTER.", "MX.", "REV.",
                "MR", "MRS", "MISS", "MS", "DR", "PROF", "MASTER", "MX", "REV",
                "SIR", "LADY", "LORD", "DAME"
            };

            var parts = clientName.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
                return ("", "");

            int startIndex = 0;

            // Check if first word is a title
            var firstWord = parts[0].ToUpper();
            foreach (var title in titles)
            {
                if (firstWord == title)
                {
                    startIndex = 1;
                    break;
                }
            }

            var remainingParts = parts.Skip(startIndex).ToArray();

            if (remainingParts.Length == 0)
            {
                return ("", "");
            }
            else if (remainingParts.Length == 1)
            {
                // Only one name - treat as last name
                return ("", remainingParts[0]);
            }
            else
            {
                // Check for "&" indicating multiple people - take first person only
                var ampersandIndex = Array.FindIndex(remainingParts, p => p == "&" || p == "AND");
                if (ampersandIndex > 0)
                {
                    remainingParts = remainingParts.Take(ampersandIndex).ToArray();
                }

                // Last word is last name, rest are given names
                var givenNames = string.Join(" ", remainingParts.Take(remainingParts.Length - 1));
                var lastName = remainingParts.Last();
                return (givenNames, lastName);
            }
        }

        private ClientInfo? SearchIndividualByNames(MySqlConnection connection, string givenNames, string lastName)
        {
            string sql = @"
                SELECT c.client_id, CONCAT(COALESCE(ci.given_names, ''), ' ', COALESCE(ci.last_name, '')) as full_name
                FROM tbl_client c
                INNER JOIN tbl_client_individual ci ON c.client_id = ci.fk_client_id
                WHERE CONVERT(ci.last_name USING utf8mb4) LIKE @lastName
                  AND (CONVERT(ci.given_names USING utf8mb4) LIKE @givenNames OR CONVERT(ci.given_names USING utf8mb4) LIKE @givenNamesStart)
                LIMIT 1";

            using (var cmd = new MySqlCommand(sql, connection))
            {
                cmd.Parameters.AddWithValue("@lastName", $"%{lastName}%");
                cmd.Parameters.AddWithValue("@givenNames", $"%{givenNames}%");
                cmd.Parameters.AddWithValue("@givenNamesStart", $"{givenNames.Split(' ')[0]}%");

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return new ClientInfo
                        {
                            ClientId = reader.GetInt32("client_id"),
                            FullName = reader.GetString("full_name")
                        };
                    }
                }
            }

            return null;
        }

        private ClientInfo? SearchIndividualByLastName(MySqlConnection connection, string lastName)
        {
            string sql = @"
                SELECT c.client_id, CONCAT(COALESCE(ci.given_names, ''), ' ', COALESCE(ci.last_name, '')) as full_name
                FROM tbl_client c
                INNER JOIN tbl_client_individual ci ON c.client_id = ci.fk_client_id
                WHERE CONVERT(ci.last_name USING utf8mb4) = CONVERT(@lastName USING utf8mb4)
                LIMIT 1";

            using (var cmd = new MySqlCommand(sql, connection))
            {
                cmd.Parameters.AddWithValue("@lastName", lastName);

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return new ClientInfo
                        {
                            ClientId = reader.GetInt32("client_id"),
                            FullName = reader.GetString("full_name")
                        };
                    }
                }
            }

            return null;
        }

        private ClientInfo? SearchIndividualByFullName(MySqlConnection connection, string nameWithoutTitle)
        {
            string sql = @"
                SELECT c.client_id, CONCAT(COALESCE(ci.given_names, ''), ' ', COALESCE(ci.last_name, '')) as full_name
                FROM tbl_client c
                INNER JOIN tbl_client_individual ci ON c.client_id = ci.fk_client_id
                WHERE CONCAT(COALESCE(CONVERT(ci.given_names USING utf8mb4), ''), ' ', COALESCE(CONVERT(ci.last_name USING utf8mb4), '')) LIKE @fullName
                LIMIT 1";

            using (var cmd = new MySqlCommand(sql, connection))
            {
                cmd.Parameters.AddWithValue("@fullName", $"%{nameWithoutTitle}%");

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return new ClientInfo
                        {
                            ClientId = reader.GetInt32("client_id"),
                            FullName = reader.GetString("full_name")
                        };
                    }
                }
            }

            return null;
        }

        private ClientInfo? SearchCompanyByName(MySqlConnection connection, string clientName)
        {
            string sql = @"
                SELECT c.client_id, cc.company_name as full_name
                FROM tbl_client c
                INNER JOIN tbl_client_company cc ON c.client_id = cc.fk_client_id
                WHERE CONVERT(cc.company_name USING utf8mb4) LIKE @clientName
                LIMIT 1";

            using (var cmd = new MySqlCommand(sql, connection))
            {
                cmd.Parameters.AddWithValue("@clientName", $"%{clientName}%");

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return new ClientInfo
                        {
                            ClientId = reader.GetInt32("client_id"),
                            FullName = reader.GetString("full_name")
                        };
                    }
                }
            }

            return null;
        }

        public (int success, int errors) ImportCasesToDatabase(List<ProcessedCaseData> processedData)
        {
            int successCount = 0;
            int errorCount = 0;

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        foreach (var item in processedData)
                        {
                            try
                            {
                                if (!string.IsNullOrEmpty(item.ClientFullName))
                                {
                                    item.CaseNameWithClient = $"{item.CaseName} - {item.ClientFullName}";
                                }
                                else
                                {
                                    item.CaseNameWithClient = item.CaseName;
                                }

                                var caseId = InsertCase(connection, transaction, item);

                                if (caseId > 0)
                                {
                                    InsertCasePermissions(connection, transaction, caseId, item);
                                    InsertCaseLedgerIds(connection, transaction, caseId, item);

                                if (item.LinkedClientId.HasValue)
                                    {
                                        InsertCaseClientRelationship(connection, transaction, caseId, item.LinkedClientId.Value);
                                        
                                        // Update ledger with client ID (like CaseRepository does)
                                        UpdateLedgerWithClientId(connection, transaction, caseId, item.LinkedClientId.Value);

                                        // Get client info for case_clients field
                                        var clientInfo = GetClientInfoById(connection, item.LinkedClientId.Value);
                                        string clientNameDisplay = string.Empty;
                                        
                                        if (clientInfo != null)
                                        {
                                            if (clientInfo.ClientType == "Individual")
                                            {
                                                clientNameDisplay = $"{clientInfo.Title} {clientInfo.FullName}".Trim();
                                            }
                                            else
                                            {
                                                clientNameDisplay = clientInfo.CompanyName ?? "";
                                            }
                                            
                                            // Update company contacts if this is a company client and we have contacts
                                            if (clientInfo.ClientType == "Company" && item.Contacts.Count > 0)
                                            {
                                                UpdateCompanyContacts(connection, transaction, item.LinkedClientId.Value, item.Contacts);
                                            }
                                        }

                                        // Greeting can be inserted manually later
                                        UpdateCaseDetails(connection, transaction, caseId, clientNameDisplay);
                                    }

                                    successCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                _logAction($"Error processing case '{item.CaseName}': {ex.Message}");
                                errorCount++;
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception($"Transaction failed: {ex.Message}", ex);
                    }
                }
            }

            return (successCount, errorCount);
        }

        private int InsertCase(MySqlConnection connection, MySqlTransaction transaction, ProcessedCaseData item)
        {
            string sql = @"
                INSERT INTO tbl_case_details_general (
                    fk_branch_id, fk_area_of_practice_id, fk_case_type_id, fk_case_sub_type_id,
                    case_reference_auto, case_number, case_name, date_opened, person_opened,
                    is_case_active, is_case_archived, is_case_not_proceeding,
                    mnl_check, conf_search, case_credit
                ) VALUES (
                    @fk_branch_id, @fk_area_of_practice_id, @fk_case_type_id, @fk_case_sub_type_id,
                    @case_reference_auto, @case_number, @case_name, @date_opened, @person_opened,
                    @is_case_active, @is_case_archived, @is_case_not_proceeding,
                    @mnl_check, @conf_search, @case_credit
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_branch_id", item.FkBranchId);
                cmd.Parameters.AddWithValue("@fk_area_of_practice_id", item.FkAreaOfPracticeId);
                cmd.Parameters.AddWithValue("@fk_case_type_id", item.FkCaseTypeId);
                cmd.Parameters.AddWithValue("@fk_case_sub_type_id", item.FkCaseSubTypeId);
                cmd.Parameters.AddWithValue("@case_reference_auto", SanitizeString(item.CaseReferenceAuto));
                cmd.Parameters.AddWithValue("@case_number", item.CaseNumber);
                cmd.Parameters.AddWithValue("@case_name", SanitizeString(item.CaseName ?? item.ClientFullName));  // MatterDescription first, then ClientFullName
                cmd.Parameters.AddWithValue("@date_opened", item.DateOpened);
                cmd.Parameters.AddWithValue("@person_opened", item.PersonOpened);
                cmd.Parameters.AddWithValue("@is_case_active", item.IsCaseActive);
                cmd.Parameters.AddWithValue("@is_case_archived", item.IsCaseArchived);
                cmd.Parameters.AddWithValue("@is_case_not_proceeding", item.IsCaseNotProceeding);
                cmd.Parameters.AddWithValue("@mnl_check", item.MnlCheck);
                cmd.Parameters.AddWithValue("@conf_search", item.ConfSearch);
                cmd.Parameters.AddWithValue("@case_credit", item.CaseCredit.HasValue ? (object)item.CaseCredit.Value : DBNull.Value);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private string SanitizeString(string? input)
        {
            if (string.IsNullOrEmpty(input)) return input ?? "";
            // Remove BOM and other problematic characters
            return input.Replace("\uFEFF", "").Replace("\uFFFD", "").Replace("�", "").Trim();
        }

        private void InsertCasePermissions(MySqlConnection connection, MySqlTransaction transaction, int caseId, ProcessedCaseData item)
        {
            string sql = @"
                INSERT INTO tbl_case_permissions (
                    fk_case_id, person_opened, person_responsible, person_assisting, person_acting,
                    is_everyone_view, is_everyone_add_edit
                ) VALUES (
                    @fk_case_id, @person_opened, @person_responsible, @person_assisting, @person_acting,
                    @is_everyone_view, @is_everyone_add_edit
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_case_id", caseId);
                cmd.Parameters.AddWithValue("@person_opened", item.PersonOpened);
                cmd.Parameters.AddWithValue("@person_responsible", item.PersonResponsible.HasValue ? (object)item.PersonResponsible.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@person_acting", item.PersonActing.HasValue ? (object)item.PersonActing.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@person_assisting", item.PersonAssisting.HasValue ? (object)item.PersonAssisting.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@is_everyone_view", true);
                cmd.Parameters.AddWithValue("@is_everyone_add_edit", true);

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertCaseLedgerIds(MySqlConnection connection, MySqlTransaction transaction, int caseId, ProcessedCaseData item)
        {
            string sql = @"
                INSERT INTO tbl_acc_ledger_cards (
                    fk_branch_id, fk_client_ids, fk_case_id, is_deleted, is_archived
                ) VALUES (
                    @fk_branch_id, @fk_client_ids, @fk_case_id, @is_deleted, @is_archived
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_branch_id", item.FkBranchId);
                // Set client ID if available
                if (item.LinkedClientId.HasValue)
                    cmd.Parameters.AddWithValue("@fk_client_ids", item.LinkedClientId.Value.ToString());
                else
                    cmd.Parameters.AddWithValue("@fk_client_ids", DBNull.Value);
                cmd.Parameters.AddWithValue("@fk_case_id", caseId);
                cmd.Parameters.AddWithValue("@is_deleted", false);
                cmd.Parameters.AddWithValue("@is_archived", item.IsCaseArchived);

                cmd.ExecuteNonQuery();
            }
        }

        private void UpdateLedgerWithClientId(MySqlConnection connection, MySqlTransaction transaction, int caseId, int clientId)
        {
            string sql = @"UPDATE tbl_acc_ledger_cards SET fk_client_ids = @fk_client_ids WHERE fk_case_id = @fk_case_id";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_client_ids", clientId.ToString());
                cmd.Parameters.AddWithValue("@fk_case_id", caseId);
                cmd.ExecuteNonQuery();
            }
        }

        private void InsertCaseClientRelationship(MySqlConnection connection, MySqlTransaction transaction, int caseId, int clientId)
        {
            string sql = @"
                INSERT INTO tbl_case_clients (fk_case_id, fk_client_id, date_added, user_added)
                VALUES (@fk_case_id, @fk_client_id, @date_added, @user_added)
                ON DUPLICATE KEY UPDATE date_added = VALUES(date_added)";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_case_id", caseId);
                cmd.Parameters.AddWithValue("@fk_client_id", clientId);
                cmd.Parameters.AddWithValue("@date_added", DateTime.Now);
                cmd.Parameters.AddWithValue("@user_added", 1);

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    if (ex.Number != 1146) throw;
                }
            }
        }

        private int InsertCaseClientGreeting(MySqlConnection connection, MySqlTransaction transaction, int caseId, string greetingType, string greetingAll)
        {
            string sql = @"
                INSERT INTO tbl_case_client_greeting (
                    fk_case_id, greeting_type, greeting_all
                ) VALUES (
                    @fk_case_id, @greeting_type, @greeting_all
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_case_id", caseId);
                cmd.Parameters.AddWithValue("@greeting_type", greetingType);
                cmd.Parameters.AddWithValue("@greeting_all", greetingAll ?? (object)DBNull.Value);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void UpdateCaseDetails(MySqlConnection connection, MySqlTransaction transaction, int caseId, string clientName)
        {
            string sql = @"
                UPDATE tbl_case_details_general 
                SET case_clients = @case_clients
                WHERE case_id = @case_id";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@case_id", caseId);
                cmd.Parameters.AddWithValue("@case_clients", clientName);
                cmd.ExecuteNonQuery();
            }
        }

        private void UpdateCompanyContacts(MySqlConnection connection, MySqlTransaction transaction, int clientId, List<CaseContactInfo> contacts)
        {
            if (contacts.Count == 0) return;

            // Update tbl_client_company with contact info (max 2 contacts)
            var updateParts = new List<string>();
            var parameters = new Dictionary<string, object>();

            if (contacts.Count > 0)
            {
                var contact1 = contacts[0];
                updateParts.Add("comp_contact1_given_names = @c1_forename");
                updateParts.Add("comp_contact1_last_name = @c1_surname");
                parameters["@c1_forename"] = SanitizeString(contact1.Forename) ?? "";
                parameters["@c1_surname"] = SanitizeString(contact1.Surname) ?? "";

                if (!string.IsNullOrEmpty(contact1.Email))
                {
                    updateParts.Add("comp_contact1_email = @c1_email");
                    parameters["@c1_email"] = contact1.Email;
                }
                if (!string.IsNullOrEmpty(contact1.Phone))
                {
                    updateParts.Add("comp_contact1_mobile = @c1_phone");
                    parameters["@c1_phone"] = contact1.Phone;
                }
            }

            if (contacts.Count > 1)
            {
                var contact2 = contacts[1];
                updateParts.Add("comp_contact2_given_names = @c2_forename");
                updateParts.Add("comp_contact2_last_name = @c2_surname");
                parameters["@c2_forename"] = SanitizeString(contact2.Forename) ?? "";
                parameters["@c2_surname"] = SanitizeString(contact2.Surname) ?? "";

                if (!string.IsNullOrEmpty(contact2.Email))
                {
                    updateParts.Add("comp_contact2_email = @c2_email");
                    parameters["@c2_email"] = contact2.Email;
                }
                if (!string.IsNullOrEmpty(contact2.Phone))
                {
                    updateParts.Add("comp_contact2_mobile = @c2_phone");
                    parameters["@c2_phone"] = contact2.Phone;
                }
            }

            if (updateParts.Count == 0) return;

            string sql = $@"
                UPDATE tbl_client_company 
                SET {string.Join(", ", updateParts)}
                WHERE fk_client_id = @clientId";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@clientId", clientId);
                foreach (var param in parameters)
                {
                    cmd.Parameters.AddWithValue(param.Key, param.Value);
                }
                int updated = cmd.ExecuteNonQuery();
                if (updated > 0)
                {
                    var contactNames = contacts.Take(2).Select(c => $"{c.Forename} {c.Surname}").ToList();
                    _logAction($"Updated company contacts for client {clientId}: {string.Join(", ", contactNames)}");
                }
            }
        }

        private ClientInfo? GetClientInfoById(MySqlConnection connection, int clientId)
        {
            string sqlIndividual = @"
                SELECT c.client_id, 
                       ci.given_names as first_name,
                       ci.last_name,
                       t.title,
                       'Individual' as client_type
                FROM tbl_client c
                INNER JOIN tbl_client_individual ci ON c.client_id = ci.fk_client_id
                LEFT JOIN tbl_title t ON ci.fk_title_id = t.title_id
                WHERE c.client_id = @clientId
                LIMIT 1";

            using (var cmd = new MySqlCommand(sqlIndividual, connection))
            {
                cmd.Parameters.AddWithValue("@clientId", clientId);
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return new ClientInfo
                        {
                            ClientId = reader.GetInt32("client_id"),
                            FirstName = reader.IsDBNull(reader.GetOrdinal("first_name")) ? null : reader.GetString("first_name"),
                            LastName = reader.IsDBNull(reader.GetOrdinal("last_name")) ? null : reader.GetString("last_name"),
                            Title = reader.IsDBNull(reader.GetOrdinal("title")) ? null : reader.GetString("title"),
                            FullName = $"{(reader.IsDBNull(reader.GetOrdinal("first_name")) ? "" : reader.GetString("first_name"))} {(reader.IsDBNull(reader.GetOrdinal("last_name")) ? "" : reader.GetString("last_name"))}".Trim(),
                            ClientType = "Individual",
                            CompanyName = null
                        };
                    }
                }
            }

            string sqlCompany = @"
                SELECT c.client_id, 
                       cc.company_name,
                       'Company' as client_type
                FROM tbl_client c
                INNER JOIN tbl_client_company cc ON c.client_id = cc.fk_client_id
                WHERE c.client_id = @clientId
                LIMIT 1";

            using (var cmd = new MySqlCommand(sqlCompany, connection))
            {
                cmd.Parameters.AddWithValue("@clientId", clientId);
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return new ClientInfo
                        {
                            ClientId = reader.GetInt32("client_id"),
                            FullName = reader.GetString("company_name"),
                            ClientType = "Company",
                            CompanyName = reader.GetString("company_name")
                        };
                    }
                }
            }

            return null;
        }
    }
}
