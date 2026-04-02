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
            var extension = Path.GetExtension(filePath).ToLower();
            
            if (extension == ".csv")
            {
                return ReadCsvData(filePath);
            }
            
            return ReadExcelFileData(filePath);
        }

        private List<ExcelRowData> ReadCsvData(string filePath)
        {
            var data = new List<ExcelRowData>();
            var lines = File.ReadAllLines(filePath);

            if (lines.Length < 2)
            {
                _logAction("CSV file is empty or has no data rows.");
                return data;
            }

            // Parse headers
            var headers = ParseCsvLine(lines[0]);
            _logAction($"Found {lines.Length - 1} data rows in CSV file.");
            _logAction($"Headers: {string.Join(", ", headers.Take(10))}...");

            // Detect format
            bool isNewFormat = headers.Any(h => h.Equals("Forename", StringComparison.OrdinalIgnoreCase) ||
                                                h.Equals("Surname", StringComparison.OrdinalIgnoreCase));

            if (isNewFormat)
            {
                _logAction("Detected new CSV format (Forename/Surname columns)");
            }

            // Parse data rows
            for (int i = 1; i < lines.Length; i++)
            {
                var values = ParseCsvLine(lines[i]);
                var rowData = new ExcelRowData();

                for (int col = 0; col < Math.Min(headers.Count, values.Count); col++)
                {
                    var cellValue = values[col]?.Trim() ?? "";
                    var header = headers[col].ToLower().Trim();

                    MapCellToRowData(rowData, header, cellValue);
                }

                // For new format, build ClientName from parts
                if (isNewFormat)
                {
                    if (!string.IsNullOrEmpty(rowData.Forename))
                    {
                        rowData.ClientName = string.Join(" ", new[] { rowData.Title, rowData.Forename, rowData.Surname }
                            .Where(s => !string.IsNullOrWhiteSpace(s)));
                    }
                    else if (!string.IsNullOrEmpty(rowData.Surname))
                    {
                        rowData.ClientName = rowData.Surname;
                    }
                }

                // Add row if it has valid data
                if (!string.IsNullOrEmpty(rowData.ClientName) ||
                    !string.IsNullOrEmpty(rowData.Surname) ||
                    !string.IsNullOrEmpty(rowData.Forename))
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
            var currentValue = new System.Text.StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // Escaped quote
                        currentValue.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(currentValue.ToString());
                    currentValue.Clear();
                }
                else
                {
                    currentValue.Append(c);
                }
            }

            result.Add(currentValue.ToString());
            return result;
        }

        private void MapCellToRowData(ExcelRowData rowData, string header, string cellValue)
        {
            switch (header)
            {
                // New format mappings
                case "title":
                    rowData.Title = cellValue;
                    break;
                case "initials":
                    rowData.Initials = cellValue;
                    break;
                case "forename":
                    rowData.Forename = cellValue;
                    rowData.GivenNames = cellValue;
                    break;
                case "surname":
                    rowData.Surname = cellValue;
                    rowData.LastNames = cellValue;
                    break;
                case "house":
                    rowData.House = cellValue;
                    break;
                case "area":
                    rowData.Area = cellValue;
                    break;
                case "town":
                    rowData.TownCity = cellValue;
                    break;
                case "post code":
                    rowData.Postcode = cellValue;
                    break;
                case "email":
                    rowData.FirstEmailAddress = cellValue;
                    break;
                case "phone no":
                    rowData.Phone = cellValue;
                    break;

                // Original format mappings
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

        private List<ExcelRowData> ReadExcelFileData(string filePath)
        {
            var data = new List<ExcelRowData>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("Excel file contains no worksheets.");
                }

                var worksheet = package.Workbook.Worksheets.First();
                if (worksheet.Dimension == null)
                {
                    _logAction("Worksheet is empty.");
                    return data;
                }

                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                _logAction($"Found {rowCount - 1} data rows in Excel file.");

                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Value?.ToString() ?? "");
                }

                _logAction($"Headers: {string.Join(", ", headers.Take(10))}...");

                // Detect format: new format has "Forename" and "Surname" columns
                bool isNewFormat = headers.Any(h => h.Equals("Forename", StringComparison.OrdinalIgnoreCase) ||
                                                    h.Equals("Surname", StringComparison.OrdinalIgnoreCase));

                if (isNewFormat)
                {
                    _logAction("Detected new format (Forename/Surname columns)");
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new ExcelRowData();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value?.ToString()?.Trim() ?? "";
                        var header = headers[col - 1].ToLower().Trim();

                        MapCellToRowData(rowData, header, cellValue);
                    }

                    // For new format, build ClientName from parts
                    if (isNewFormat)
                    {
                        if (!string.IsNullOrEmpty(rowData.Forename))
                        {
                            // Individual: has forename
                            rowData.ClientName = string.Join(" ", new[] { rowData.Title, rowData.Forename, rowData.Surname }
                                .Where(s => !string.IsNullOrWhiteSpace(s)));
                        }
                        else if (!string.IsNullOrEmpty(rowData.Surname))
                        {
                            // Company: only surname (company name)
                            rowData.ClientName = rowData.Surname;
                        }
                    }

                    // Add row if it has valid data
                    if (!string.IsNullOrEmpty(rowData.ClientName) || 
                        !string.IsNullOrEmpty(rowData.Surname) || 
                        !string.IsNullOrEmpty(rowData.Forename))
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
                    ClientType = DetermineClientType(row),
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

        /// <summary>
        /// Determines client type based on:
        /// - New format: If Forename exists → Individual, otherwise Company
        /// - Old format: Check for company keywords like LTD, LIMITED, PLC, etc.
        /// </summary>
        private string DetermineClientType(ExcelRowData row)
        {
            // New format: Check Forename field
            // If Forename has a value, it's an Individual
            // If only Surname (no Forename), it's a Company
            if (!string.IsNullOrEmpty(row.Forename))
            {
                return "Individual";
            }

            // If we have Surname but no Forename, it's a Company
            if (!string.IsNullOrEmpty(row.Surname) && string.IsNullOrEmpty(row.Forename))
            {
                return "Company";
            }

            // Fallback to old format logic for backward compatibility
            var lastName = row.LastNames;
            if (string.IsNullOrEmpty(lastName))
                return "Individual";

            var upperLastNames = lastName.ToUpper().Trim();

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

        /// <summary>
        /// Exports processed client data to an Excel file
        /// </summary>
        public string ExportToExcel(List<ProcessedClientData> processedData, string outputFilePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Clients");

                // Headers
                var headers = new[]
                {
                    "Client Type", "Title", "Given Names", "Last Name / Company Name",
                    "Email", "Phone", "Date of Birth",
                    "Address Line 1", "Address Line 2", "Town/City", "County", "Post Code",
                    "Mobile", "Home Phone", "Work Phone", "Fax"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                }

                // Data rows
                int row = 2;
                foreach (var item in processedData)
                {
                    var data = item.OriginalData;

                    worksheet.Cells[row, 1].Value = item.ClientType;
                    worksheet.Cells[row, 2].Value = data?.Title ?? "";
                    worksheet.Cells[row, 3].Value = data?.GivenNames ?? "";
                    worksheet.Cells[row, 4].Value = item.ClientType == "Company" ? data?.ClientName : data?.LastNames;
                    worksheet.Cells[row, 5].Value = data?.FirstEmailAddress ?? "";
                    worksheet.Cells[row, 6].Value = data?.PrimaryContactNumber ?? "";
                    worksheet.Cells[row, 7].Value = data?.DateOfBirth?.ToString("yyyy-MM-dd") ?? "";
                    worksheet.Cells[row, 8].Value = data?.AddressLine1 ?? "";
                    worksheet.Cells[row, 9].Value = data?.AddressLine2 ?? "";
                    worksheet.Cells[row, 10].Value = data?.TownCity ?? "";
                    worksheet.Cells[row, 11].Value = data?.County ?? "";
                    worksheet.Cells[row, 12].Value = data?.Postcode ?? "";
                    worksheet.Cells[row, 13].Value = data?.Mobile ?? "";
                    worksheet.Cells[row, 14].Value = data?.Home ?? "";
                    worksheet.Cells[row, 15].Value = data?.Work ?? "";
                    worksheet.Cells[row, 16].Value = data?.Fax ?? "";

                    row++;
                }

                // Auto-fit columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Save file
                package.SaveAs(new FileInfo(outputFilePath));

                _logAction($"Exported {processedData.Count} clients to: {outputFilePath}");
                return outputFilePath;
            }
        }

        /// <summary>
        /// Exports processed client data to separate worksheets for Individuals and Companies
        /// </summary>
        public string ExportToExcelSeparated(List<ProcessedClientData> processedData, string outputFilePath)
        {
            using (var package = new ExcelPackage())
            {
                // Create Individual clients worksheet
                var individuals = processedData.Where(p => p.ClientType == "Individual").ToList();
                CreateClientWorksheet(package, "Individuals", individuals, false);

                // Create Company clients worksheet
                var companies = processedData.Where(p => p.ClientType == "Company").ToList();
                CreateClientWorksheet(package, "Companies", companies, true);

                // Create Summary worksheet
                var summarySheet = package.Workbook.Worksheets.Add("Summary");
                summarySheet.Cells[1, 1].Value = "Export Summary";
                summarySheet.Cells[1, 1].Style.Font.Bold = true;
                summarySheet.Cells[1, 1].Style.Font.Size = 14;

                summarySheet.Cells[3, 1].Value = "Total Clients:";
                summarySheet.Cells[3, 2].Value = processedData.Count;
                summarySheet.Cells[4, 1].Value = "Individuals:";
                summarySheet.Cells[4, 2].Value = individuals.Count;
                summarySheet.Cells[5, 1].Value = "Companies:";
                summarySheet.Cells[5, 2].Value = companies.Count;
                summarySheet.Cells[6, 1].Value = "Export Date:";
                summarySheet.Cells[6, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                summarySheet.Cells[summarySheet.Dimension.Address].AutoFitColumns();

                // Save file
                package.SaveAs(new FileInfo(outputFilePath));

                _logAction($"Exported {individuals.Count} individuals and {companies.Count} companies to: {outputFilePath}");
                return outputFilePath;
            }
        }

        private void CreateClientWorksheet(ExcelPackage package, string sheetName, List<ProcessedClientData> clients, bool isCompany)
        {
            var worksheet = package.Workbook.Worksheets.Add(sheetName);

            string[] headers;
            if (isCompany)
            {
                headers = new[]
                {
                    "Company Name", "Email", "Phone", "Fax",
                    "Address Line 1", "Address Line 2", "Town/City", "County", "Post Code"
                };
            }
            else
            {
                headers = new[]
                {
                    "Title", "Given Names", "Last Name", "Email", "Date of Birth",
                    "Phone", "Mobile", "Home Phone", "Work Phone",
                    "Address Line 1", "Address Line 2", "Town/City", "County", "Post Code"
                };
            }

            // Write headers
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(
                    isCompany ? System.Drawing.Color.LightGreen : System.Drawing.Color.LightBlue);
            }

            // Write data
            int row = 2;
            foreach (var item in clients)
            {
                var data = item.OriginalData;

                if (isCompany)
                {
                    worksheet.Cells[row, 1].Value = data?.ClientName ?? data?.Surname ?? "";
                    worksheet.Cells[row, 2].Value = data?.FirstEmailAddress ?? "";
                    worksheet.Cells[row, 3].Value = data?.PrimaryContactNumber ?? "";
                    worksheet.Cells[row, 4].Value = data?.Fax ?? "";
                    worksheet.Cells[row, 5].Value = data?.AddressLine1 ?? "";
                    worksheet.Cells[row, 6].Value = data?.AddressLine2 ?? "";
                    worksheet.Cells[row, 7].Value = data?.TownCity ?? "";
                    worksheet.Cells[row, 8].Value = data?.County ?? "";
                    worksheet.Cells[row, 9].Value = data?.Postcode ?? "";
                }
                else
                {
                    worksheet.Cells[row, 1].Value = data?.Title ?? "";
                    worksheet.Cells[row, 2].Value = data?.GivenNames ?? "";
                    worksheet.Cells[row, 3].Value = data?.LastNames ?? "";
                    worksheet.Cells[row, 4].Value = data?.FirstEmailAddress ?? "";
                    worksheet.Cells[row, 5].Value = data?.DateOfBirth?.ToString("yyyy-MM-dd") ?? "";
                    worksheet.Cells[row, 6].Value = data?.PrimaryContactNumber ?? "";
                    worksheet.Cells[row, 7].Value = data?.Mobile ?? "";
                    worksheet.Cells[row, 8].Value = data?.Home ?? "";
                    worksheet.Cells[row, 9].Value = data?.Work ?? "";
                    worksheet.Cells[row, 10].Value = data?.AddressLine1 ?? "";
                    worksheet.Cells[row, 11].Value = data?.AddressLine2 ?? "";
                    worksheet.Cells[row, 12].Value = data?.TownCity ?? "";
                    worksheet.Cells[row, 13].Value = data?.County ?? "";
                    worksheet.Cells[row, 14].Value = data?.Postcode ?? "";
                }

                row++;
            }

            // Auto-fit columns
            if (worksheet.Dimension != null)
            {
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            }
        }
    }
}
