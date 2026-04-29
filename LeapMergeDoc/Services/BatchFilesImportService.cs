using LeapMergeDoc.Models;
using MySql.Data.MySqlClient;
using System.IO;
using System.Text.RegularExpressions;

namespace LeapMergeDoc.Services
{
    /// <summary>
    /// Service to batch import files from folders named by case reference
    /// Parses folder names like "AE.74.23" or "AE.225.25 ALNASSAR" to find case IDs
    /// </summary>
    public class BatchFilesImportService
    {
        private readonly string _connectionString;
        private readonly S3Configuration _s3Config;
        private readonly Action<string> _logAction;

        public BatchFilesImportService(string connectionString, S3Configuration s3Config, Action<string> logAction)
        {
            _connectionString = connectionString;
            _s3Config = s3Config;
            _logAction = logAction;
        }

        /// <summary>
        /// Represents a case folder to import
        /// </summary>
        public class CaseFolderInfo
        {
            public string FolderPath { get; set; } = "";
            public string FolderName { get; set; } = "";
            public string ParsedCaseReference { get; set; } = "";
            public int? CaseId { get; set; }
            public string? CaseName { get; set; }
            public int FileCount { get; set; }
            public bool IsSelected { get; set; } = true;
            public string Status { get; set; } = "Pending";
        }

        public class MissingFileInfo
        {
            public string FolderName { get; set; } = string.Empty;
            public string RelativePath { get; set; } = string.Empty;
            public string Reason { get; set; } = string.Empty;
        }

        /// <summary>
        /// Scan root folder and return list of case folders with parsed references
        /// </summary>
        public List<CaseFolderInfo> ScanRootFolder(string rootFolderPath)
        {
            var caseFolders = new List<CaseFolderInfo>();

            if (!Directory.Exists(rootFolderPath))
            {
                _logAction($"❌ Root folder does not exist: {rootFolderPath}");
                return caseFolders;
            }

            _logAction($"📁 Scanning root folder: {rootFolderPath}");

            var directories = Directory.GetDirectories(rootFolderPath);
            _logAction($"Found {directories.Length} subfolders");

            foreach (var dir in directories)
            {
                var folderName = Path.GetFileName(dir);
                var caseReference = ParseCaseReference(folderName);
                var fileCount = CountFilesRecursive(dir);

                var folderInfo = new CaseFolderInfo
                {
                    FolderPath = dir,
                    FolderName = folderName,
                    ParsedCaseReference = caseReference,
                    FileCount = fileCount
                };

                caseFolders.Add(folderInfo);
                _logAction($"  📂 {folderName} → {caseReference} ({fileCount} files)");
            }

            return caseFolders;
        }

        /// <summary>
        /// Parse folder name to case reference format
        /// Examples:
        ///   "AE.74.23" → "AE/74/23"
        ///   "AE.225.25 ALNASSAR" → "AE/225/25"
        ///   "AS.166.24 AbdelAziz" → "AS/166/24"
        /// </summary>
        public string ParseCaseReference(string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
                return "";

            // First, split by space and take the first part (handles "AE.225.25 ALNASSAR")
            var firstPart = folderName.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? folderName;

            // Replace dots with slashes
            var caseReference = firstPart.Replace('.', '/');

            return caseReference;
        }

        /// <summary>
        /// Look up case IDs in database for all case folders
        /// </summary>
        public void LookupCaseIds(List<CaseFolderInfo> caseFolders)
        {
            _logAction("🔍 Looking up case IDs in database...");

            using var connection = new MySqlConnection(_connectionString);
            connection.Open();

            foreach (var folder in caseFolders)
            {
                if (string.IsNullOrEmpty(folder.ParsedCaseReference))
                {
                    folder.Status = "Invalid reference";
                    continue;
                }

                var caseInfo = FindCaseByReference(connection, folder.ParsedCaseReference);
                
                if (caseInfo.HasValue)
                {
                    folder.CaseId = caseInfo.Value.caseId;
                    folder.CaseName = caseInfo.Value.caseName;
                    folder.Status = "Found";
                    _logAction($"  ✅ {folder.ParsedCaseReference} → Case ID: {folder.CaseId} ({folder.CaseName})");
                }
                else
                {
                    folder.Status = "Not found";
                    _logAction($"  ❌ {folder.ParsedCaseReference} → Not found in database");
                }
            }

            var found = caseFolders.Count(f => f.CaseId.HasValue);
            var notFound = caseFolders.Count(f => !f.CaseId.HasValue);
            _logAction($"");
            _logAction($"═══════════════════════════════════════");
            _logAction($"📊 SUMMARY: {found} cases found, {notFound} not found");
            _logAction($"═══════════════════════════════════════");

            // List all NOT FOUND folders for manual review
            if (notFound > 0)
            {
                _logAction($"");
                _logAction($"❌ NOT FOUND FOLDERS (for manual review):");
                _logAction($"───────────────────────────────────────");
                foreach (var folder in caseFolders.Where(f => !f.CaseId.HasValue))
                {
                    _logAction($"   📁 {folder.FolderName} → {folder.ParsedCaseReference}");
                }
                _logAction($"───────────────────────────────────────");
            }
        }

        /// <summary>
        /// Find case by reference number
        /// </summary>
        private (int caseId, string caseName)? FindCaseByReference(MySqlConnection connection, string caseReference)
        {
            // Search by case_reference_auto column
            var sql = @"SELECT case_id, case_name 
                        FROM tbl_case_details_general 
                        WHERE case_reference_auto = @reference 
                        LIMIT 1";

            using var cmd = new MySqlCommand(sql, connection);
            cmd.Parameters.AddWithValue("@reference", caseReference);

            using var reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                var caseId = reader.GetInt32("case_id");
                var caseName = reader.IsDBNull(reader.GetOrdinal("case_name")) ? "" : reader.GetString("case_name");
                return (caseId, caseName);
            }

            return null;
        }

        /// <summary>
        /// Import files from selected case folders
        /// </summary>
        /// <param name="caseFolders">List of case folders to import</param>
        /// <param name="moveToCompletedFolder">If set, move successfully uploaded folders here</param>
        public (int totalSuccess, int totalErrors, int casesProcessed, List<MissingFileInfo> missingFiles) ImportSelectedFolders(
            List<CaseFolderInfo> caseFolders, 
            string? moveToCompletedFolder = null)
        {
            int totalSuccess = 0;
            int totalErrors = 0;
            int casesProcessed = 0;
            var missingFiles = new List<MissingFileInfo>();

            var selectedFolders = caseFolders
                .Where(f => f.IsSelected && f.CaseId.HasValue)
                .ToList();

            // Create completed folder if specified
            if (!string.IsNullOrEmpty(moveToCompletedFolder) && !Directory.Exists(moveToCompletedFolder))
            {
                Directory.CreateDirectory(moveToCompletedFolder);
                _logAction($"📁 Created completed folder: {moveToCompletedFolder}");
            }

            _logAction($"═══════════════════════════════════════");
            _logAction($"📦 Starting batch import of {selectedFolders.Count} case folders...");
            if (!string.IsNullOrEmpty(moveToCompletedFolder))
            {
                _logAction($"📂 Move on success: {moveToCompletedFolder}");
            }
            _logAction($"═══════════════════════════════════════");

            foreach (var folder in selectedFolders)
            {
                _logAction($"");
                _logAction($"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                _logAction($"📁 Processing: {folder.FolderName}");
                _logAction($"   Case: {folder.CaseName} (ID: {folder.CaseId})");
                _logAction($"   Files: {folder.FileCount}");

                try
                {
                    folder.Status = "Importing...";

                    var importService = new FilesCollectionImportService(
                        _connectionString,
                        _s3Config,
                        _logAction);

                    // Set completed folder path for file-by-file moving
                    if (!string.IsNullOrEmpty(moveToCompletedFolder))
                    {
                        // Create subfolder for this case folder
                        importService.CompletedFolderPath = Path.Combine(moveToCompletedFolder, folder.FolderName);
                    }

                    var (success, errors) = importService.UploadFilesFromFolder(folder.FolderPath, folder.CaseId!.Value);

                    foreach (var failedFile in importService.LastImportResults.Where(r => r.Status == "Failed"))
                    {
                        missingFiles.Add(new MissingFileInfo
                        {
                            FolderName = folder.FolderName,
                            RelativePath = failedFile.RelativePath,
                            Reason = string.IsNullOrWhiteSpace(failedFile.ErrorMessage) ? "Failed to import" : failedFile.ErrorMessage
                        });
                    }

                    totalSuccess += success;
                    totalErrors += errors;
                    casesProcessed++;

                    // Files are moved individually during upload now
                    // Just update status
                    if (errors == 0)
                    {
                        folder.Status = $"✅ Done ({success} files)";
                    }
                    else
                    {
                        folder.Status = $"⚠️ Done ({success} ok, {errors} errors)";
                    }
                    
                    _logAction($"   Result: {success} uploaded, {errors} errors");
                }
                catch (Exception ex)
                {
                    folder.Status = $"❌ Error: {ex.Message}";
                    _logAction($"   ❌ Error: {ex.Message}");
                    totalErrors++;
                }
            }

            _logAction($"");
            _logAction($"═══════════════════════════════════════");
            _logAction($"📊 BATCH IMPORT COMPLETE");
            _logAction($"   Cases processed: {casesProcessed}");
            _logAction($"   Total files uploaded: {totalSuccess}");
            _logAction($"   Total errors: {totalErrors}");
            _logAction($"═══════════════════════════════════════");

            return (totalSuccess, totalErrors, casesProcessed, missingFiles);
        }

        /// <summary>
        /// Count files recursively in a folder (including ZIP files)
        /// </summary>
        private int CountFilesRecursive(string folderPath)
        {
            try
            {
                return Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories).Length;
            }
            catch
            {
                return 0;
            }
        }
    }
}
