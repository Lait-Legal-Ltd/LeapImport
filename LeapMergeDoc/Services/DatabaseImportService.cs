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

        public DatabaseImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _matterTypeMatcher = new MatterTypeMatcher();
            _logAction = logAction;
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
                        var header = headers[col - 1].ToLower().Replace("_", "").Replace(" ", "").Replace(".", "");

                        switch (header)
                        {
                            case "matterno":
                            case "matternumber":
                            case "matterref":
                                // Matter No is the unique reference - store as string for CaseReferenceAuto
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
                                // Store as name for later lookup
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

                    // Include row if it has MatterNo OR CaseName
                    if (!string.IsNullOrEmpty(rowData.CaseName) || !string.IsNullOrEmpty(rowData.MatterNo))
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

            // Load users from database for staff mapping
            var userLookup = LoadUsersFromDatabase();
            _logAction($"Loaded {userLookup.Count} users from database for staff mapping.");

            int caseNumber = 1; // Start from 1 for sequential case numbers

            foreach (var row in excelData)
            {
                // Use MatterTypeMatcher to get area of practice and case type
                var matterMatch = _matterTypeMatcher.MatchMatterType(row.MatterType);

                // Lookup staff by name
                int? personResponsible = LookupUserId(row.StaffRespName, userLookup);
                int? personActing = LookupUserId(row.StaffActName, userLookup);
                int? personAssisting = LookupUserId(row.StaffAssistName, userLookup);
                int? caseCredit = LookupUserId(row.CreditName, userLookup);

                var caseData = new ProcessedCaseData
                {
                    OriginalData = row,
                    FkBranchId = 1,
                    FkAreaOfPracticeId = matterMatch.AreaOfPracticeId,
                    FkCaseTypeId = matterMatch.CaseTypeId,
                    FkCaseSubTypeId = null,
                    CaseReferenceAuto = row.MatterNo,  // Matter No from Excel is the case reference
                    CaseNumber = caseNumber,  // Sequential case number
                    CaseName = !string.IsNullOrEmpty(row.MatterDescription) ? row.MatterDescription : row.CaseName,  // MatterDescription is case name
                    DateOpened = row.DateOpened,
                    PersonOpened = 7,
                    PersonResponsible = personResponsible,  // null if user not found
                    PersonActing = personActing,
                    PersonAssisting = personAssisting,
                    CaseCredit = caseCredit,  // Who brought the case
                    IsCaseActive = row.ArchiveDate == null,
                    IsCaseArchived = row.ArchiveDate != null,
                    IsCaseNotProceeding = false,
                    MnlCheck = false,
                    ConfSearch = false
                };

                processedData.Add(caseData);
                caseNumber++; // Increment for next case
            }

            return processedData;
        }

        private Dictionary<string, int> LoadUsersFromDatabase()
        {
            var userLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                string sql = @"
                    SELECT user_id, first_name, last_name, 
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
                            string firstName = reader.IsDBNull(reader.GetOrdinal("first_name")) ? "" : reader.GetString("first_name");
                            string lastName = reader.IsDBNull(reader.GetOrdinal("last_name")) ? "" : reader.GetString("last_name");
                            string fullName = reader.IsDBNull(reader.GetOrdinal("full_name")) ? "" : reader.GetString("full_name");

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


        public (int found, int notFound) CheckClientMappings(List<ProcessedCaseData> processedData)
        {
            int foundClients = 0;
            int notFoundClients = 0;

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                foreach (var caseData in processedData)
                {
                    var clientInfo = FindClientInfo(connection, caseData.OriginalData!);
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

        private ClientInfo? FindClientInfo(MySqlConnection connection, CaseExcelData caseData)
        {
            string clientName = caseData.ClientName ?? "";

            if (string.IsNullOrEmpty(clientName))
                return null;

            // Strip title from client name for better matching
            var (givenNames, lastName) = ExtractNamesFromClientName(clientName);
            string nameWithoutTitle = $"{givenNames} {lastName}".Trim();

            // Try multiple search strategies
            ClientInfo? result = null;

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
                WHERE ci.last_name LIKE @lastName
                  AND (ci.given_names LIKE @givenNames OR ci.given_names LIKE @givenNamesStart)
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
                WHERE ci.last_name = @lastName
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
                WHERE CONCAT(COALESCE(ci.given_names, ''), ' ', COALESCE(ci.last_name, '')) LIKE @fullName
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
                WHERE cc.company_name LIKE @clientName
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
                                        InsertCaseClientGreeting(connection, transaction, caseId, "individual", null!);

                                        var clientInfo = GetClientInfoById(connection, item.LinkedClientId.Value);
                                        string clientname = string.Empty;
                                        if (clientInfo != null)
                                        {
                                            clientname = (clientInfo.ClientType == "Individual")
                                                ? clientInfo.Title + clientInfo.FullName
                                                : clientInfo.CompanyName ?? "";
                                        }
                                        UpdateCaseDetails(connection, transaction, caseId, clientname);
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
                cmd.Parameters.AddWithValue("@case_reference_auto", item.CaseReferenceAuto);
                cmd.Parameters.AddWithValue("@case_number", item.CaseNumber);
                cmd.Parameters.AddWithValue("@case_name", item.CaseName ?? item.ClientFullName);  // MatterDescription first, then ClientFullName
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
                cmd.Parameters.AddWithValue("@fk_branch_id", 1);
                cmd.Parameters.AddWithValue("@fk_client_ids", DBNull.Value);
                cmd.Parameters.AddWithValue("@fk_case_id", caseId);
                cmd.Parameters.AddWithValue("@is_deleted", false);
                cmd.Parameters.AddWithValue("@is_archived", false);

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
