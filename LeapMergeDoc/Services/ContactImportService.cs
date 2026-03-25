using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using LeapMergeDoc.Models;

namespace LeapMergeDoc.Services
{
    public class ContactImportService
    {
        private readonly string _connectionString;
        private readonly Action<string> _logAction;

        // Company indicators
        private static readonly string[] CompanyIndicators = new[] {
            "LTD", "LIMITED", "LLP", "PLC", "INC", "CORP", "CORPORATION",
            "SOLICITORS", "ASSOCIATES", "PARTNERS", "SERVICES", "GROUP",
            "COUNCIL", "AUTHORITY", "OFFICE", "COURT", "TRIBUNAL",
            "REGISTRY", "AGENCY", "DEPARTMENT", "MINISTRY", "BANK"
        };

        public ContactImportService(string connectionString, Action<string> logAction)
        {
            _connectionString = connectionString;
            _logAction = logAction;
        }

        public List<ContactExcelData> ReadExcelData(string filePath)
        {
            var data = new List<ContactExcelData>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var range = worksheet.RangeUsed();
                if (range == null) return data;

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

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new ContactExcelData();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetString().Trim();
                        var header = headers[col - 1].ToLower().Replace(" ", "").Replace("/", "");

                        switch (header)
                        {
                            case "cardname":
                                rowData.CardName = cellValue;
                                break;
                            case "firstemailaddress":
                                rowData.Email = cellValue;
                                break;
                            case "dateofbirth":
                                rowData.DateOfBirth = ParseDate(cellValue);
                                break;
                            case "buildingname":
                                rowData.BuildingName = cellValue;
                                break;
                            case "streetlevel":
                                rowData.StreetLevel = cellValue;
                                break;
                            case "number":
                                rowData.Number = cellValue;
                                break;
                            case "street":
                                rowData.Street = cellValue;
                                break;
                            case "towncity":
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
                            case "poboxinstructions":
                                rowData.POBoxInstructions = cellValue;
                                break;
                            case "poboxtype":
                                rowData.POBoxType = cellValue;
                                break;
                            case "poboxnumber":
                                rowData.POBoxNumber = cellValue;
                                break;
                            case "poboxtowncity":
                                rowData.POBoxTownCity = cellValue;
                                break;
                            case "poboxcounty":
                                rowData.POBoxCounty = cellValue;
                                break;
                            case "poboxpostcode":
                                rowData.POBoxPostcode = cellValue;
                                break;
                            case "dxinstructions":
                                rowData.DxInstructions = cellValue;
                                break;
                            case "dxnumber":
                                rowData.DxNumber = cellValue;
                                break;
                            case "exchange":
                                rowData.Exchange = cellValue;
                                break;
                            case "mktconsent?":
                            case "mktconsent":
                                rowData.MktConsent = cellValue;
                                break;
                        }
                    }

                    if (!string.IsNullOrEmpty(rowData.CardName))
                    {
                        data.Add(rowData);
                    }
                }
            }

            return data;
        }

        private DateTime? ParseDate(string? dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString)) return null;
            if (DateTime.TryParse(dateString, out DateTime result)) return result;
            return null;
        }

        public List<ProcessedContactData> ProcessExcelData(List<ContactExcelData> excelData)
        {
            var processedData = new List<ProcessedContactData>();
            var clientNames = LoadExistingClientNames();
            int skippedCount = 0;

            foreach (var row in excelData)
            {
                if (string.IsNullOrEmpty(row.CardName)) continue;

                var processed = new ProcessedContactData
                {
                    OriginalData = row,
                    IsCompany = IsCompanyName(row.CardName)
                };

                // Build address lines
                var addrParts = new List<string>();
                if (!string.IsNullOrEmpty(row.BuildingName)) addrParts.Add(row.BuildingName);
                if (!string.IsNullOrEmpty(row.Number)) addrParts.Add(row.Number);
                if (!string.IsNullOrEmpty(row.StreetLevel)) addrParts.Add(row.StreetLevel);
                processed.AddressLine1 = string.Join(" ", addrParts);
                processed.AddressLine2 = row.Street;

                if (processed.IsCompany)
                {
                    processed.CompanyName = row.CardName;
                    processed.FkContactTypeId = 3;  // Company (from tbl_contact_type)
                    processed.IsExistingClient = clientNames.Companies.Contains(row.CardName.ToUpper());
                }
                else
                {
                    // Parse personal name
                    var nameParts = ParsePersonName(row.CardName);
                    processed.GivenNames = nameParts.givenNames;
                    processed.LastName = nameParts.lastName;
                    processed.FkContactTypeId = 1;  // Individual (from tbl_contact_type)

                    // Check if exists in clients
                    var fullName = $"{processed.GivenNames} {processed.LastName}".Trim().ToUpper();
                    processed.IsExistingClient = clientNames.Individuals.Contains(fullName);
                }

                if (processed.IsExistingClient)
                {
                    skippedCount++;
                }

                processedData.Add(processed);
            }

            _logAction($"Processed {processedData.Count} records. Skipping {skippedCount} existing clients.");
            return processedData;
        }

        private bool IsCompanyName(string name)
        {
            var upperName = name.ToUpper();
            return CompanyIndicators.Any(ind => upperName.Contains(ind));
        }

        private (string givenNames, string lastName) ParsePersonName(string cardName)
        {
            // Strip titles
            var titles = new[] { "MR.", "MRS.", "MS.", "MISS.", "DR.", "PROF.", "MR", "MRS", "MS", "MISS", "DR", "PROF" };
            var name = cardName.Trim();

            foreach (var title in titles)
            {
                if (name.ToUpper().StartsWith(title + " "))
                {
                    name = name.Substring(title.Length + 1).Trim();
                    break;
                }
            }

            var parts = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return ("", "");
            if (parts.Length == 1) return ("", parts[0]);

            var lastName = parts[parts.Length - 1];
            var givenNames = string.Join(" ", parts.Take(parts.Length - 1));
            return (givenNames, lastName);
        }

        private (HashSet<string> Individuals, HashSet<string> Companies) LoadExistingClientNames()
        {
            var individuals = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var companies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                // Load individual clients
                string sqlInd = "SELECT CONCAT(COALESCE(given_names, ''), ' ', COALESCE(last_name, '')) as full_name FROM tbl_client_individual";
                using (var cmd = new MySqlCommand(sqlInd, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var name = reader.GetString("full_name").Trim().ToUpper();
                        if (!string.IsNullOrEmpty(name))
                            individuals.Add(name);
                    }
                }

                // Load company clients
                string sqlComp = "SELECT company_name FROM tbl_client_company WHERE company_name IS NOT NULL";
                using (var cmd = new MySqlCommand(sqlComp, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var name = reader.GetString("company_name").Trim().ToUpper();
                        if (!string.IsNullOrEmpty(name))
                            companies.Add(name);
                    }
                }
            }

            _logAction($"Loaded {individuals.Count} individual clients and {companies.Count} company clients to check against.");
            return (individuals, companies);
        }

        public (int success, int skipped, int errors) ImportContactsToDatabase(List<ProcessedContactData> processedData)
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
                            if (item.IsExistingClient)
                            {
                                skippedCount++;
                                continue;
                            }

                            try
                            {
                                // Insert into tbl_contact
                                var contactId = InsertContact(connection, transaction, item);

                                if (contactId > 0)
                                {
                                    // Insert into company or personal table
                                    if (item.IsCompany)
                                    {
                                        InsertContactCompany(connection, transaction, contactId, item);
                                    }
                                    else
                                    {
                                        InsertContactPersonal(connection, transaction, contactId, item);
                                    }
                                    successCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                _logAction($"❌ Error importing {item.OriginalData?.CardName}: {ex.Message}");
                                errorCount++;
                            }
                        }

                        transaction.Commit();
                        _logAction($"✅ Import complete. Success: {successCount}, Skipped: {skippedCount}, Errors: {errorCount}");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logAction($"❌ Transaction failed: {ex.Message}");
                        throw;
                    }
                }
            }

            return (successCount, skippedCount, errorCount);
        }

        private int InsertContact(MySqlConnection connection, MySqlTransaction transaction, ProcessedContactData item)
        {
            string sql = @"
                INSERT INTO tbl_contact (fk_branch_id, fk_contact_type_id, is_active, date_created, fk_user_id, is_supplier)
                VALUES (@fk_branch_id, @fk_contact_type_id, @is_active, @date_created, @fk_user_id, @is_supplier);
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_branch_id", 1);
                cmd.Parameters.AddWithValue("@fk_contact_type_id", item.FkContactTypeId);
                cmd.Parameters.AddWithValue("@is_active", true);
                cmd.Parameters.AddWithValue("@date_created", DateTime.Now);
                cmd.Parameters.AddWithValue("@fk_user_id", 7);
                cmd.Parameters.AddWithValue("@is_supplier", false);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void InsertContactCompany(MySqlConnection connection, MySqlTransaction transaction, int contactId, ProcessedContactData item)
        {
            string sql = @"
                INSERT INTO tbl_contact_company (
                    fk_contact_id, fk_company_contacts_category_id, company_name, company_email, company_fax,
                    company_contact_phone1, company_contact_phone2,
                    company_address_line1, company_address_line2, company_city, company_county, company_postcode,
                    company_dx_number, company_dx_exchange,
                    company_pobox_number, company_pobox_town, company_pobox_county, company_pobox_postcode,
                    company_contact1_mobile, company_is_active
                ) VALUES (
                    @fk_contact_id, @fk_company_contacts_category_id, @company_name, @company_email, @company_fax,
                    @company_contact_phone1, @company_contact_phone2,
                    @company_address_line1, @company_address_line2, @company_city, @company_county, @company_postcode,
                    @company_dx_number, @company_dx_exchange,
                    @company_pobox_number, @company_pobox_town, @company_pobox_county, @company_pobox_postcode,
                    @company_contact1_mobile, @company_is_active
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_contact_id", contactId);
                cmd.Parameters.AddWithValue("@fk_company_contacts_category_id", 48);  // Default: SOLICITORS
                cmd.Parameters.AddWithValue("@company_name", item.CompanyName);
                cmd.Parameters.AddWithValue("@company_email", item.OriginalData?.Email ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_fax", item.OriginalData?.Fax ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_contact_phone1", item.OriginalData?.Phone ?? item.OriginalData?.Home ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_contact_phone2", item.OriginalData?.Work ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_address_line1", item.AddressLine1 ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_address_line2", item.AddressLine2 ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_city", item.OriginalData?.TownCity ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_county", item.OriginalData?.County ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_postcode", item.OriginalData?.Postcode ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_dx_number", item.OriginalData?.DxNumber ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_dx_exchange", item.OriginalData?.Exchange ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_pobox_number", item.OriginalData?.POBoxNumber ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_pobox_town", item.OriginalData?.POBoxTownCity ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_pobox_county", item.OriginalData?.POBoxCounty ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_pobox_postcode", item.OriginalData?.POBoxPostcode ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_contact1_mobile", item.OriginalData?.Mobile ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@company_is_active", true);

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertContactPersonal(MySqlConnection connection, MySqlTransaction transaction, int contactId, ProcessedContactData item)
        {
            string sql = @"
                INSERT INTO tbl_contact_personal (
                    fk_contact_id, fk_personal_contacts_category_id, given_names, last_name, date_of_birth,
                    personal_cont_email, personal_cont_home_phone, personal_cont_mobile,
                    personal_cont_address1, personal_cont_address2, personal_cont_city, personal_cont_county, personal_cont_post_code,
                    employment_cont_office_phone, employment_cont_fax
                ) VALUES (
                    @fk_contact_id, @fk_personal_contacts_category_id, @given_names, @last_name, @date_of_birth,
                    @personal_cont_email, @personal_cont_home_phone, @personal_cont_mobile,
                    @personal_cont_address1, @personal_cont_address2, @personal_cont_city, @personal_cont_county, @personal_cont_post_code,
                    @employment_cont_office_phone, @employment_cont_fax
                )";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@fk_contact_id", contactId);
                cmd.Parameters.AddWithValue("@fk_personal_contacts_category_id", 41);  // Default: OTHER SIDES
                cmd.Parameters.AddWithValue("@given_names", item.GivenNames ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@last_name", item.LastName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@date_of_birth", item.OriginalData?.DateOfBirth ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_email", item.OriginalData?.Email ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_home_phone", item.OriginalData?.Home ?? item.OriginalData?.Phone ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_mobile", item.OriginalData?.Mobile ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_address1", item.AddressLine1 ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_address2", item.AddressLine2 ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_city", item.OriginalData?.TownCity ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_county", item.OriginalData?.County ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@personal_cont_post_code", item.OriginalData?.Postcode ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@employment_cont_office_phone", item.OriginalData?.Work ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@employment_cont_fax", item.OriginalData?.Fax ?? (object)DBNull.Value);

                cmd.ExecuteNonQuery();
            }
        }

        public void TruncateContactData()
        {
            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        var tables = new[] { "tbl_contact_company", "tbl_contact_personal", "tbl_contact" };

                        using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 0", connection, transaction))
                            cmd.ExecuteNonQuery();

                        foreach (var table in tables)
                        {
                            using (var cmd = new MySqlCommand($"TRUNCATE TABLE {table}", connection, transaction))
                            {
                                cmd.ExecuteNonQuery();
                                _logAction($"Truncated {table}");
                            }
                        }

                        using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 1", connection, transaction))
                            cmd.ExecuteNonQuery();

                        transaction.Commit();
                        _logAction("✅ All contact tables truncated successfully.");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logAction($"❌ Error truncating tables: {ex.Message}");
                        throw;
                    }
                }
            }
        }
    }
}
