using LeapMergeDoc.Models;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;

namespace LeapMergeDoc.Services
{
    public class ClientImportService
    {
        private readonly string _connectionString;
        private readonly Action<string> _logAction;

        private static readonly Dictionary<string, int> TitleMappings = new Dictionary<string, int>
        {
            {"MR", 1}, {"MR.", 1},
            {"MRS", 2}, {"MRS.", 2},
            {"MISS", 3}, {"MISS.", 3},
            {"DR", 4}, {"DR.", 4},
            {"PROF", 5}, {"PROF.", 5},
            {"MS", 6}, {"MS.", 6},
            {"MASTER", 7}, {"MASTER.", 7},
            {"MX", 8}, {"MX.", 8},
            {"REV", 9}, {"REV.", 9}
        };

        public ClientImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _logAction = logAction;
        }

        public (int rowsDeleted, string message) TruncateClientData()
        {
            int totalDeleted = 0;

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 0;", connection))
                {
                    cmd.ExecuteNonQuery();
                }

                try
                {
                    var tablesToTruncate = new[]
                    {
                        "tbl_client_individual",
                        "tbl_client_company",
                        "tbl_client"
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

                    foreach (var table in tablesToTruncate)
                    {
                        try
                        {
                            using (var cmd = new MySqlCommand($"ALTER TABLE {table} AUTO_INCREMENT = 1;", connection))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                        catch (MySqlException) { }
                    }
                }
                finally
                {
                    using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 1;", connection))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            return (totalDeleted, $"Successfully deleted {totalDeleted} total rows from client tables.");
        }

        public List<ExcelRowData> ReadExcelData(string filePath)
        {
            var data = new List<ExcelRowData>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                _logAction($"Found {rowCount - 1} data rows in Excel file.");

                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Value?.ToString() ?? "");
                }

                _logAction($"Headers: {string.Join(", ", headers.Take(10))}...");

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new ExcelRowData();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value?.ToString()?.Trim() ?? "";
                        var header = headers[col - 1].ToLower();

                        switch (header)
                        {
                            case "short name":
                                rowData.ShortName = cellValue;
                                break;
                            case "client name":
                                rowData.ClientName = cellValue;
                                ExtractTitleAndNames(rowData, cellValue);
                                break;
                            case "first email address":
                                rowData.FirstEmailAddress = cellValue;
                                break;
                            case "date of birth":
                                if (DateTime.TryParse(cellValue, out DateTime dob))
                                    rowData.DateOfBirth = dob;
                                break;
                            case "building name":
                                rowData.BuildingName = cellValue;
                                break;
                            case "street level":
                                rowData.StreetLevel = cellValue;
                                break;
                            case "number":
                                rowData.Number = cellValue;
                                break;
                            case "street":
                                rowData.Street = cellValue;
                                break;
                            case "town/city":
                                rowData.TownCity = cellValue;
                                break;
                            case "county":
                                rowData.County = cellValue;
                                break;
                            case "postcode":
                                rowData.Postcode = cellValue;
                                break;
                            case "country":
                                rowData.Country = cellValue;
                                break;
                            case "phone":
                                rowData.Phone = cellValue;
                                break;
                            case "home":
                                rowData.Home = cellValue;
                                break;
                            case "work":
                                rowData.Work = cellValue;
                                break;
                            case "mobile":
                                rowData.Mobile = cellValue;
                                break;
                            case "fax":
                                rowData.Fax = cellValue;
                                break;
                            case "pobox instructions":
                                rowData.POBoxInstructions = cellValue;
                                break;
                            case "pobox type":
                                rowData.POBoxType = cellValue;
                                break;
                            case "pobox number":
                                rowData.POBoxNumber = cellValue;
                                break;
                            case "pobox town/city":
                                rowData.POBoxTownCity = cellValue;
                                break;
                            case "pobox county":
                                rowData.POBoxCounty = cellValue;
                                break;
                            case "pobox postcode":
                                rowData.POBoxPostcode = cellValue;
                                break;
                            case "dx instructions":
                                rowData.DxInstructions = cellValue;
                                break;
                            case "dx number":
                                rowData.DxNumber = cellValue;
                                break;
                            case "exchange":
                                rowData.Exchange = cellValue;
                                break;
                            case "mkt consent?":
                                rowData.MktConsent = cellValue?.ToLower() == "yes" || cellValue?.ToLower() == "true" || cellValue == "1";
                                break;
                        }
                    }

                    if (!string.IsNullOrEmpty(rowData.ClientName))
                    {
                        data.Add(rowData);
                    }
                }
            }

            return data;
        }

        public List<ProcessedClientData> ProcessExcelData(List<ExcelRowData> excelData)
        {
            var processedData = new List<ProcessedClientData>();

            foreach (var row in excelData)
            {
                var clientData = new ProcessedClientData
                {
                    OriginalData = row,
                    ClientType = DetermineClientType(row.LastNames),
                    TitleId = GetTitleId(row.Title),
                    FkBranchId = 1,
                    FkUserId = 1,
                    DateTimeCreated = DateTime.Now,
                    IsArchived = false,
                    IsActive = true
                };

                processedData.Add(clientData);
            }

            return processedData;
        }

        private string DetermineClientType(string? lastNames)
        {
            if (string.IsNullOrEmpty(lastNames))
                return "Individual";

            var upperLastNames = lastNames.ToUpper().Trim();

            if (upperLastNames.Contains("LTD") || upperLastNames.Contains("LIMITED") ||
                upperLastNames.Contains("PLC") || upperLastNames.Contains("LLC") ||
                upperLastNames.Contains("INC") || upperLastNames.Contains("CORP"))
            {
                return "Company";
            }

            return "Individual";
        }

        private int? GetTitleId(string? title)
        {
            if (string.IsNullOrEmpty(title))
                return null;

            var upperTitle = title.ToUpper().Trim();
            return TitleMappings.ContainsKey(upperTitle) ? TitleMappings[upperTitle] : (int?)null;
        }

        private void ExtractTitleAndNames(ExcelRowData rowData, string clientName)
        {
            if (string.IsNullOrWhiteSpace(clientName))
                return;

            var titles = new[] { "MR.", "MRS.", "MISS.", "MS.", "DR.", "PROF.", "MASTER.", "MX.", "REV.",
                                 "MR", "MRS", "MISS", "MS", "DR", "PROF", "MASTER", "MX", "REV" };

            var parts = clientName.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
                return;

            int startIndex = 0;
            string? foundTitle = null;

            var firstWord = parts[0].ToUpper();
            foreach (var title in titles)
            {
                if (firstWord == title)
                {
                    foundTitle = parts[0];
                    startIndex = 1;
                    break;
                }
            }

            rowData.Title = foundTitle;

            var remainingParts = parts.Skip(startIndex).ToArray();

            if (remainingParts.Length == 0)
            {
                return;
            }
            else if (remainingParts.Length == 1)
            {
                rowData.LastNames = remainingParts[0];
            }
            else if (remainingParts.Length == 2)
            {
                rowData.GivenNames = remainingParts[0];
                rowData.LastNames = remainingParts[1];
            }
            else
            {
                var ampersandIndex = Array.FindIndex(remainingParts, p => p == "&");

                if (ampersandIndex > 0)
                {
                    var firstPersonParts = remainingParts.Take(ampersandIndex).ToArray();
                    if (firstPersonParts.Length >= 2)
                    {
                        rowData.GivenNames = string.Join(" ", firstPersonParts.Take(firstPersonParts.Length - 1));
                        rowData.LastNames = firstPersonParts.Last();
                    }
                    else if (firstPersonParts.Length == 1)
                    {
                        rowData.LastNames = firstPersonParts[0];
                    }
                }
                else
                {
                    rowData.GivenNames = string.Join(" ", remainingParts.Take(remainingParts.Length - 1));
                    rowData.LastNames = remainingParts.Last();
                }
            }
        }

        public (int success, int errors) ImportToDatabase(List<ProcessedClientData> processedData)
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
                                var clientId = InsertClient(connection, transaction, item);

                                if (clientId > 0)
                                {
                                    if (item.ClientType == "Company")
                                    {
                                        InsertClientCompany(connection, transaction, clientId, item);
                                    }
                                    else
                                    {
                                        InsertClientIndividual(connection, transaction, clientId, item);
                                    }

                                    successCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                _logAction($"Error: {item.OriginalData?.ClientName}: {ex.Message}");
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

        private int InsertClient(MySqlConnection connection, MySqlTransaction transaction, ProcessedClientData item)
        {
            string sql = @"
                INSERT INTO tbl_client (fk_branch_id, client_type, date_time_created, fk_user_id, is_archived, is_active)
                VALUES (@fk_branch_id, @client_type, @date_time_created, @fk_user_id, @is_archived, @is_active);
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_branch_id", item.FkBranchId);
                cmd.Parameters.AddWithValue("@client_type", item.ClientType);
                cmd.Parameters.AddWithValue("@date_time_created", item.DateTimeCreated);
                cmd.Parameters.AddWithValue("@fk_user_id", item.FkUserId);
                cmd.Parameters.AddWithValue("@is_archived", item.IsArchived);
                cmd.Parameters.AddWithValue("@is_active", item.IsActive);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void InsertClientCompany(MySqlConnection connection, MySqlTransaction transaction, int clientId, ProcessedClientData item)
        {
            string sql = @"
                INSERT INTO tbl_client_company (
                    fk_client_id, company_name, 
                    comp_email_address, comp_fax,
                    comp_contact_phone1, comp_contact_phone2,
                    comp_current_add_line1, comp_current_add_line2,
                    comp_current_add_city, comp_current_add_county, comp_current_add_post_code,
                    comp_current_add_dx_number, comp_current_add_dx_exchange,
                    comp_current_add_pobox_number, comp_current_add_pobox_county, 
                    comp_current_add_pobox_town, comp_current_add_pobox_post_code,
                    comp_is_active
                ) VALUES (
                    @fk_client_id, @company_name, 
                    @comp_email_address, @comp_fax,
                    @comp_contact_phone1, @comp_contact_phone2,
                    @comp_current_add_line1, @comp_current_add_line2,
                    @comp_current_add_city, @comp_current_add_county, @comp_current_add_post_code,
                    @comp_current_add_dx_number, @comp_current_add_dx_exchange,
                    @comp_current_add_pobox_number, @comp_current_add_pobox_county, 
                    @comp_current_add_pobox_town, @comp_current_add_pobox_post_code,
                    @comp_is_active
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                var data = item.OriginalData;

                cmd.Parameters.AddWithValue("@fk_client_id", clientId);
                cmd.Parameters.AddWithValue("@company_name", data?.ClientName ?? "");
                cmd.Parameters.AddWithValue("@comp_email_address", data?.FirstEmailAddress ?? "");
                cmd.Parameters.AddWithValue("@comp_fax", data?.Fax ?? "");
                cmd.Parameters.AddWithValue("@comp_contact_phone1", data?.Phone ?? data?.Work ?? "");
                cmd.Parameters.AddWithValue("@comp_contact_phone2", data?.Mobile ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_line1", data?.AddressLine1 ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_line2", data?.AddressLine2 ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_city", data?.TownCity ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_county", data?.County ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_post_code", data?.Postcode ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_dx_number", data?.DxNumber ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_dx_exchange", data?.Exchange ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_pobox_number", data?.POBoxNumber ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_pobox_county", data?.POBoxCounty ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_pobox_town", data?.POBoxTownCity ?? "");
                cmd.Parameters.AddWithValue("@comp_current_add_pobox_post_code", data?.POBoxPostcode ?? "");
                cmd.Parameters.AddWithValue("@comp_is_active", true);

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertClientIndividual(MySqlConnection connection, MySqlTransaction transaction, int clientId, ProcessedClientData item)
        {
            string sql = @"
                INSERT INTO tbl_client_individual (
                    fk_client_id, fk_title_id, given_names, last_name, 
                    dob, email, 
                    home_phone, mobile_phone, other_phone,
                    current_add_line1, current_add_line2, 
                    current_add_city, current_add_county, current_add_post_code,
                    current_add_dx_number, current_add_dx_exchange,
                    current_add_pobox_number, current_add_pobox_county, 
                    current_add_pobox_town, current_add_pobox_post_code,
                    ind_is_active
                ) VALUES (
                    @fk_client_id, @fk_title_id, @given_names, @last_name, 
                    @dob, @email, 
                    @home_phone, @mobile_phone, @other_phone,
                    @current_add_line1, @current_add_line2, 
                    @current_add_city, @current_add_county, @current_add_post_code,
                    @current_add_dx_number, @current_add_dx_exchange,
                    @current_add_pobox_number, @current_add_pobox_county, 
                    @current_add_pobox_town, @current_add_pobox_post_code,
                    @ind_is_active
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                var data = item.OriginalData;

                cmd.Parameters.AddWithValue("@fk_client_id", clientId);
                cmd.Parameters.AddWithValue("@fk_title_id", item.TitleId.HasValue ? (object)item.TitleId.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@given_names", data?.GivenNames ?? "");
                cmd.Parameters.AddWithValue("@last_name", data?.LastNames ?? "");
                cmd.Parameters.AddWithValue("@dob", data?.DateOfBirth.HasValue == true ? (object)data.DateOfBirth.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@email", data?.FirstEmailAddress ?? "");
                cmd.Parameters.AddWithValue("@home_phone", data?.Home ?? "");
                cmd.Parameters.AddWithValue("@mobile_phone", data?.Mobile ?? "");
                cmd.Parameters.AddWithValue("@other_phone", data?.Work ?? data?.Phone ?? "");
                cmd.Parameters.AddWithValue("@current_add_line1", data?.AddressLine1 ?? "");
                cmd.Parameters.AddWithValue("@current_add_line2", data?.AddressLine2 ?? "");
                cmd.Parameters.AddWithValue("@current_add_city", data?.TownCity ?? "");
                cmd.Parameters.AddWithValue("@current_add_county", data?.County ?? "");
                cmd.Parameters.AddWithValue("@current_add_post_code", data?.Postcode ?? "");
                cmd.Parameters.AddWithValue("@current_add_dx_number", data?.DxNumber ?? "");
                cmd.Parameters.AddWithValue("@current_add_dx_exchange", data?.Exchange ?? "");
                cmd.Parameters.AddWithValue("@current_add_pobox_number", data?.POBoxNumber ?? "");
                cmd.Parameters.AddWithValue("@current_add_pobox_county", data?.POBoxCounty ?? "");
                cmd.Parameters.AddWithValue("@current_add_pobox_town", data?.POBoxTownCity ?? "");
                cmd.Parameters.AddWithValue("@current_add_pobox_post_code", data?.POBoxPostcode ?? "");
                cmd.Parameters.AddWithValue("@ind_is_active", true);

                cmd.ExecuteNonQuery();
            }
        }
    }
}
