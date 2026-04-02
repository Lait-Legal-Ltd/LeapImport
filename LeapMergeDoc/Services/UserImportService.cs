using LeapMergeDoc.Models;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;

namespace LeapMergeDoc.Services
{
    public class UserImportService
    {
        private readonly string _connectionString;
        private readonly Action<string> _logAction;

        private static readonly Dictionary<string, int> TitleMappings = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
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

        public UserImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _logAction = logAction;
        }

        #region Read Data

        public List<UserExcelRowData> ReadExcelData(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLower();

            if (extension == ".csv")
            {
                return ReadCsvData(filePath);
            }

            return ReadExcelFileData(filePath);
        }

        private List<UserExcelRowData> ReadCsvData(string filePath)
        {
            var data = new List<UserExcelRowData>();
            var lines = File.ReadAllLines(filePath);

            if (lines.Length < 2)
            {
                _logAction("CSV file is empty or has no data rows.");
                return data;
            }

            var headers = ParseCsvLine(lines[0]);
            _logAction($"Found {lines.Length - 1} data rows in CSV file.");
            _logAction($"Headers: {string.Join(", ", headers.Take(10))}...");

            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                    continue;

                var values = ParseCsvLine(lines[i]);
                var rowData = new UserExcelRowData();

                for (int col = 0; col < Math.Min(headers.Count, values.Count); col++)
                {
                    var cellValue = values[col]?.Trim() ?? "";
                    var header = headers[col].ToLower().Trim();

                    MapCellToRowData(rowData, header, cellValue);
                }

                // Add row if it has valid data (at least UserCode or Name)
                if (!string.IsNullOrEmpty(rowData.UserCode) ||
                    !string.IsNullOrEmpty(rowData.FirstName) ||
                    !string.IsNullOrEmpty(rowData.LastName))
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

        private List<UserExcelRowData> ReadExcelFileData(string filePath)
        {
            var data = new List<UserExcelRowData>();

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

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new UserExcelRowData();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value?.ToString()?.Trim() ?? "";
                        var header = headers[col - 1].ToLower().Trim();

                        MapCellToRowData(rowData, header, cellValue);
                    }

                    // Add row if it has valid data
                    if (!string.IsNullOrEmpty(rowData.UserCode) ||
                        !string.IsNullOrEmpty(rowData.FirstName) ||
                        !string.IsNullOrEmpty(rowData.LastName))
                    {
                        data.Add(rowData);
                    }
                }
            }

            return data;
        }

        private void ParseFeeEarnerName(UserExcelRowData rowData, string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return;

            var parts = fullName.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length >= 2)
            {
                // First part is FirstName, last part is LastName, anything in between is MiddleName
                rowData.FirstName = parts[0];
                rowData.LastName = parts[^1]; // Last element

                if (parts.Length > 2)
                {
                    rowData.MiddleName = string.Join(" ", parts.Skip(1).Take(parts.Length - 2));
                }
            }
            else if (parts.Length == 1)
            {
                // Only one name - use as LastName
                rowData.LastName = parts[0];
            }
        }

        private void MapCellToRowData(UserExcelRowData rowData, string header, string cellValue)
        {
            switch (header)
            {
                // New format: F/E, Fee Earner Description, Fee Earner Status, In Use, Email
                case "f/e":
                case "fe":
                    rowData.UserCode = cellValue;
                    break;
                case "fee earner description":
                    rowData.FeeEarnerDescription = cellValue;
                    ParseFeeEarnerName(rowData, cellValue);
                    break;
                case "fee earner status":
                    rowData.Designation = cellValue;
                    break;
                case "in use":
                    rowData.InUse = cellValue?.ToLower() == "true" || cellValue == "1" || cellValue?.ToLower() == "yes";
                    break;

                // Standard mappings
                case "title":
                    rowData.Title = cellValue;
                    break;
                case "first name":
                case "firstname":
                case "forename":
                case "given name":
                    rowData.FirstName = cellValue;
                    break;
                case "middle name":
                case "middlename":
                    rowData.MiddleName = cellValue;
                    break;
                case "last name":
                case "lastname":
                case "surname":
                case "family name":
                    rowData.LastName = cellValue;
                    break;
                case "user code":
                case "usercode":
                case "code":
                case "fee earner":
                case "feeearner":
                case "fee earner code":
                case "staff code":
                case "initials":
                    rowData.UserCode = cellValue;
                    break;
                case "email":
                case "email address":
                    rowData.Email = cellValue;
                    break;
                case "home phone":
                case "homephone":
                case "phone":
                case "telephone":
                    rowData.HomePhone = cellValue;
                    break;
                case "mobile":
                case "mobile phone":
                case "cell":
                case "cell phone":
                    rowData.Mobile = cellValue;
                    break;
                case "address":
                    rowData.Address = cellValue;
                    break;
                case "qualifications":
                case "qualification":
                    rowData.Qualifications = cellValue;
                    break;
                case "designation":
                case "job title":
                case "position":
                case "role":
                    rowData.Designation = cellValue;
                    break;
                case "ni number":
                case "ninumber":
                case "national insurance":
                case "ni":
                    rowData.NiNumber = cellValue;
                    break;
                case "date of birth":
                case "dob":
                case "birth date":
                    if (DateTime.TryParse(cellValue, out DateTime dob))
                        rowData.DateOfBirth = dob;
                    break;
                case "sex":
                case "gender":
                    rowData.Sex = cellValue?.ToLower() == "male" || cellValue?.ToLower() == "m" || cellValue == "1";
                    break;
                case "notes":
                case "comments":
                    rowData.Notes = cellValue;
                    break;
            }
        }

        #endregion

        #region Process Data

        public List<ProcessedUserData> ProcessExcelData(List<UserExcelRowData> excelData)
        {
            var processedData = new List<ProcessedUserData>();

            // Get existing users from database for duplicate checking
            var existingUsers = GetExistingUsers();
            _logAction($"Found {existingUsers.Count} existing users in database.");

            foreach (var row in excelData)
            {
                var userData = new ProcessedUserData
                {
                    OriginalData = row,
                    TitleId = GetTitleId(row.Title),
                    FkUserRoleId = null,  // Will be set separately if needed
                    FkBranchId = 1,       // Default branch
                    IsActive = row.InUse ?? true,  // Use InUse from Excel, default to true
                    IsDeleted = false
                };

                // Check for duplicates based on Email (primary check)
                if (!string.IsNullOrEmpty(row.Email))
                {
                    var existingUser = existingUsers.FirstOrDefault(u =>
                        !string.IsNullOrEmpty(u.Email) &&
                        u.Email.Equals(row.Email, StringComparison.OrdinalIgnoreCase));

                    if (existingUser != null)
                    {
                        userData.IsDuplicate = true;
                        userData.ExistingUserId = existingUser.UserId;

                        if (existingUser.IsDeleted == true)
                        {
                            userData.DuplicateReason = $"Email '{row.Email}' exists but user is DELETED (User ID: {existingUser.UserId}, Name: {existingUser.FirstName} {existingUser.LastName})";
                        }
                        else if (existingUser.IsActive == false)
                        {
                            userData.DuplicateReason = $"Email '{row.Email}' exists but user is INACTIVE (User ID: {existingUser.UserId}, Name: {existingUser.FirstName} {existingUser.LastName})";
                        }
                        else
                        {
                            userData.DuplicateReason = $"Email '{row.Email}' already exists (User ID: {existingUser.UserId}, Name: {existingUser.FirstName} {existingUser.LastName})";
                        }
                    }
                }

                // Generate UserCode if not provided
                if (string.IsNullOrEmpty(row.UserCode) && !string.IsNullOrEmpty(row.FirstName) && !string.IsNullOrEmpty(row.LastName))
                {
                    row.UserCode = GenerateUserCode(row.FirstName, row.LastName, existingUsers.Select(u => u.UserCode).ToList());
                    _logAction($"Generated UserCode: {row.UserCode} for {row.FirstName} {row.LastName}");
                }

                processedData.Add(userData);
            }

            return processedData;
        }

        private List<ExistingUserInfo> GetExistingUsers()
        {
            var users = new List<ExistingUserInfo>();

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                string sql = @"
                    SELECT user_id, user_code, first_name, last_name, email, is_active, is_deleted 
                    FROM tbl_user";

                using (var cmd = new MySqlCommand(sql, connection))
                {
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            users.Add(new ExistingUserInfo
                            {
                                UserId = reader.GetInt32("user_id"),
                                UserCode = reader.IsDBNull(reader.GetOrdinal("user_code")) ? null : reader.GetString("user_code"),
                                FirstName = reader.IsDBNull(reader.GetOrdinal("first_name")) ? null : reader.GetString("first_name"),
                                LastName = reader.IsDBNull(reader.GetOrdinal("last_name")) ? null : reader.GetString("last_name"),
                                Email = reader.IsDBNull(reader.GetOrdinal("email")) ? null : reader.GetString("email"),
                                IsActive = reader.IsDBNull(reader.GetOrdinal("is_active")) ? null : reader.GetBoolean("is_active"),
                                IsDeleted = reader.IsDBNull(reader.GetOrdinal("is_deleted")) ? null : reader.GetBoolean("is_deleted")
                            });
                        }
                    }
                }
            }

            return users;
        }

        private string GenerateUserCode(string firstName, string lastName, List<string?> existingCodes)
        {
            if (string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(lastName))
                return "";

            // Generate base code from initials (first letter of first name + first two letters of last name)
            string baseCode = (firstName[0].ToString() + lastName.Substring(0, Math.Min(2, lastName.Length))).ToUpper();

            // Check if code already exists
            if (!existingCodes.Any(c => c?.Equals(baseCode, StringComparison.OrdinalIgnoreCase) == true))
            {
                return baseCode;
            }

            // If exists, append numbers until unique
            int counter = 1;
            string newCode;
            do
            {
                newCode = baseCode + counter;
                counter++;
            } while (existingCodes.Any(c => c?.Equals(newCode, StringComparison.OrdinalIgnoreCase) == true));

            return newCode;
        }

        private int? GetTitleId(string? title)
        {
            if (string.IsNullOrEmpty(title))
                return null;

            var upperTitle = title.ToUpper().Trim();
            return TitleMappings.ContainsKey(upperTitle) ? TitleMappings[upperTitle] : (int?)null;
        }

        #endregion

        #region Import to Database

        public (int success, int skipped, int errors) ImportToDatabase(List<ProcessedUserData> processedData, bool skipDuplicates = true)
        {
            int successCount = 0;
            int skippedCount = 0;
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
                                // Skip duplicates if flag is set
                                if (item.IsDuplicate && skipDuplicates)
                                {
                                    _logAction($"⏭️ Skipped: {item.OriginalData?.FullName} - {item.DuplicateReason}");
                                    skippedCount++;
                                    continue;
                                }

                                var userId = InsertUser(connection, transaction, item);

                                if (userId > 0)
                                {
                                    successCount++;
                                    _logAction($"✅ Inserted: {item.OriginalData?.FullName} (UserCode: {item.OriginalData?.UserCode}, ID: {userId})");
                                }
                            }
                            catch (Exception ex)
                            {
                                _logAction($"❌ Error: {item.OriginalData?.FullName}: {ex.Message}");
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

            return (successCount, skippedCount, errorCount);
        }

        private int InsertUser(MySqlConnection connection, MySqlTransaction transaction, ProcessedUserData item)
        {
            string sql = @"
                INSERT INTO tbl_user (
                    fk_title_id, first_name, middle_name, last_name, user_code,
                    email, home_phone, mobile, address, qualifications,
                    designation, ni_number, dob, sex, notes,
                    fk_user_role_id, fk_branch_id, is_active, is_deleted,
                    is_firsttime_enabled
                ) VALUES (
                    @fk_title_id, @first_name, @middle_name, @last_name, @user_code,
                    @email, @home_phone, @mobile, @address, @qualifications,
                    @designation, @ni_number, @dob, @sex, @notes,
                    @fk_user_role_id, @fk_branch_id, @is_active, @is_deleted,
                    @is_firsttime_enabled
                );
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                var data = item.OriginalData;

                cmd.Parameters.AddWithValue("@fk_title_id", item.TitleId.HasValue ? (object)item.TitleId.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@first_name", data?.FirstName ?? "");
                cmd.Parameters.AddWithValue("@middle_name", data?.MiddleName ?? "");
                cmd.Parameters.AddWithValue("@last_name", data?.LastName ?? "");
                cmd.Parameters.AddWithValue("@user_code", data?.UserCode ?? "");
                cmd.Parameters.AddWithValue("@email", data?.Email ?? "");
                cmd.Parameters.AddWithValue("@home_phone", data?.HomePhone ?? "");
                cmd.Parameters.AddWithValue("@mobile", data?.Mobile ?? "");
                cmd.Parameters.AddWithValue("@address", data?.Address ?? "");
                cmd.Parameters.AddWithValue("@qualifications", data?.Qualifications ?? "");
                cmd.Parameters.AddWithValue("@designation", data?.Designation ?? "");
                cmd.Parameters.AddWithValue("@ni_number", data?.NiNumber ?? "");
                cmd.Parameters.AddWithValue("@dob", data?.DateOfBirth.HasValue == true ? (object)DateOnly.FromDateTime(data.DateOfBirth.Value) : DBNull.Value);
                cmd.Parameters.AddWithValue("@sex", data?.Sex.HasValue == true ? (object)data.Sex.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@notes", data?.Notes ?? "");
                cmd.Parameters.AddWithValue("@fk_user_role_id", item.FkUserRoleId.HasValue ? (object)item.FkUserRoleId.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@fk_branch_id", item.FkBranchId.HasValue ? (object)item.FkBranchId.Value.ToString() : DBNull.Value);
                cmd.Parameters.AddWithValue("@is_active", item.IsActive);
                cmd.Parameters.AddWithValue("@is_deleted", item.IsDeleted);
                cmd.Parameters.AddWithValue("@is_firsttime_enabled", true);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        #endregion

        #region Export

        public string ExportToExcel(List<ProcessedUserData> processedData, string outputFilePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Users");

                // Headers
                var headers = new[]
                {
                    "Status", "User Code", "Title", "First Name", "Middle Name", "Last Name",
                    "Email", "Home Phone", "Mobile", "Designation", "Address",
                    "Qualifications", "NI Number", "Date of Birth", "Gender", "Notes", "Duplicate Info"
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

                    worksheet.Cells[row, 1].Value = item.IsDuplicate ? "DUPLICATE" : "NEW";
                    worksheet.Cells[row, 2].Value = data?.UserCode ?? "";
                    worksheet.Cells[row, 3].Value = data?.Title ?? "";
                    worksheet.Cells[row, 4].Value = data?.FirstName ?? "";
                    worksheet.Cells[row, 5].Value = data?.MiddleName ?? "";
                    worksheet.Cells[row, 6].Value = data?.LastName ?? "";
                    worksheet.Cells[row, 7].Value = data?.Email ?? "";
                    worksheet.Cells[row, 8].Value = data?.HomePhone ?? "";
                    worksheet.Cells[row, 9].Value = data?.Mobile ?? "";
                    worksheet.Cells[row, 10].Value = data?.Designation ?? "";
                    worksheet.Cells[row, 11].Value = data?.Address ?? "";
                    worksheet.Cells[row, 12].Value = data?.Qualifications ?? "";
                    worksheet.Cells[row, 13].Value = data?.NiNumber ?? "";
                    worksheet.Cells[row, 14].Value = data?.DateOfBirth?.ToString("yyyy-MM-dd") ?? "";
                    worksheet.Cells[row, 15].Value = data?.Sex == true ? "Male" : (data?.Sex == false ? "Female" : "");
                    worksheet.Cells[row, 16].Value = data?.Notes ?? "";
                    worksheet.Cells[row, 17].Value = item.DuplicateReason ?? "";

                    // Highlight duplicates
                    if (item.IsDuplicate)
                    {
                        for (int col = 1; col <= headers.Length; col++)
                        {
                            worksheet.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
                        }
                    }

                    row++;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(outputFilePath));

                _logAction($"Exported {processedData.Count} users to: {outputFilePath}");
                return outputFilePath;
            }
        }

        #endregion

        #region Truncate

        public (int rowsDeleted, string message) TruncateUserData()
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
                    using (var cmd = new MySqlCommand("DELETE FROM tbl_user;", connection))
                    {
                        totalDeleted = cmd.ExecuteNonQuery();
                        _logAction($"Deleted {totalDeleted} rows from tbl_user");
                    }

                    using (var cmd = new MySqlCommand("ALTER TABLE tbl_user AUTO_INCREMENT = 1;", connection))
                    {
                        cmd.ExecuteNonQuery();
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

            return (totalDeleted, $"Successfully deleted {totalDeleted} rows from tbl_user.");
        }

        #endregion
    }
}
