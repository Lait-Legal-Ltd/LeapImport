using Amazon;
using Amazon.Extensions.NETCore.Setup;
using Amazon.S3;
using Amazon.S3.Model;
using LeapMergeDoc.Models;
using MySql.Data.MySqlClient;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace LeapMergeDoc.Services
{
    public class FilesCollectionImportService
    {
        public class FileImportResult
        {
            public int CaseId { get; set; }
            public string FolderPath { get; set; } = string.Empty;
            public string RelativePath { get; set; } = string.Empty;
            public string FileName { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string? ErrorMessage { get; set; }
        }

        private readonly string _connectionString;
        private readonly S3Configuration _s3Config;
        private readonly Action<string> _logAction;
        private IAmazonS3? _s3Client;
        private readonly List<string> _tempExtractionFolders = new List<string>();

        public FilesCollectionImportService(string connectionString, S3Configuration s3Config, Action<string> logAction)
        {
            _connectionString = connectionString;
            _s3Config = s3Config;
            _logAction = logAction;
        }

        private IAmazonS3 GetS3Client()
        {
            if (_s3Client == null)
            {
                var profilesLocation = _s3Config.ProfilePath;
                
                if (string.IsNullOrEmpty(profilesLocation) || !File.Exists(profilesLocation))
                {
                    profilesLocation = @"C:\AWS_Profile\LAITLegal.txt";
                    
                    if (!File.Exists(profilesLocation))
                    {
                        profilesLocation = Environment.GetEnvironmentVariable("AWS_SHARED_CREDENTIALS_FILE") 
                            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".aws", "credentials");
                    }
                }

                var awsOptions = new AWSOptions
                {
                    Profile = "aws_profile",
                    ProfilesLocation = profilesLocation,
                    Region = RegionEndpoint.GetBySystemName("eu-west-2")
                };

                _s3Client = awsOptions.CreateServiceClient<IAmazonS3>();
            }
            return _s3Client;
        }

        /// <summary>
        /// Path to move successfully uploaded files to (set before calling UploadFilesFromFolder)
        /// </summary>
        public string? CompletedFolderPath { get; set; }

        /// <summary>
        /// Per-file outcomes from the most recent UploadFilesFromFolder run.
        /// </summary>
        public List<FileImportResult> LastImportResults { get; } = new List<FileImportResult>();

        public (int success, int errors) UploadFilesFromFolder(string rootFolderPath, int caseId)
        {
            int successCount = 0;
            int errorCount = 0;
            LastImportResults.Clear();

            try
            {
                _logAction("Scanning folder structure...");
                var folderGroups = ScanFolderStructure(rootFolderPath, caseId);
                _logAction($"Found {folderGroups.Count} folders with files");

                _logAction("Creating folders in database...");
                var folderIdMap = new Dictionary<string, int?>();
                foreach (var folderGroup in folderGroups)
                {
                    if (!string.IsNullOrEmpty(folderGroup.FolderPath))
                    {
                        if (!folderIdMap.ContainsKey(folderGroup.FolderPath))
                        {
                            var folderId = CreateOrGetFolderWithParents(caseId, folderGroup.FolderPath);
                            folderIdMap[folderGroup.FolderPath] = folderId;
                            _logAction($"Folder created/retrieved: {folderGroup.FolderPath} (ID: {folderId})");
                        }
                    }
                }

                _logAction("Initializing S3 client...");
                var s3Client = GetS3Client();

                _logAction("Starting upload process...");
                foreach (var folderGroup in folderGroups)
                {
                    try
                    {
                        _logAction($"Processing folder: {folderGroup.FolderName} ({folderGroup.Files.Count} files)");

                        int? folderId = null;
                        if (!string.IsNullOrEmpty(folderGroup.FolderPath) && folderIdMap.ContainsKey(folderGroup.FolderPath))
                        {
                            folderId = folderIdMap[folderGroup.FolderPath];
                        }

                        // Process each file individually - upload, save to DB, then move
                        foreach (var fileInfo in folderGroup.Files)
                        {
                            try
                            {
                                var relativePath = Path.GetRelativePath(rootFolderPath, fileInfo.LocalFilePath);

                                // Check for duplicate before upload (by case + folder + filename)
                                if (IsFileDuplicate(caseId, folderId, fileInfo.FileName))
                                {
                                    _logAction($"⏭️ Skipped (duplicate): {fileInfo.FileName}");
                                    LastImportResults.Add(new FileImportResult
                                    {
                                        CaseId = caseId,
                                        FolderPath = folderGroup.FolderPath,
                                        RelativePath = relativePath,
                                        FileName = fileInfo.FileName,
                                        Status = "SkippedDuplicate"
                                    });

                                    // Move to completed since it's already uploaded
                                    MoveFileToCompleted(fileInfo.LocalFilePath, rootFolderPath);
                                    continue;
                                }

                                _logAction($"Uploading: {fileInfo.FileName}");

                                // 1. Upload to S3
                                var s3Key = UploadFileToS3(s3Client, fileInfo, caseId);
                                fileInfo.S3Key = s3Key;

                                // 2. Save to DB immediately (single file)
                                SaveSingleFileToDatabase(caseId, folderId, folderGroup.FolderName, fileInfo);

                                _logAction($"✅ Uploaded: {fileInfo.FileName} -> {s3Key}");
                                successCount++;
                                LastImportResults.Add(new FileImportResult
                                {
                                    CaseId = caseId,
                                    FolderPath = folderGroup.FolderPath,
                                    RelativePath = relativePath,
                                    FileName = fileInfo.FileName,
                                    Status = "Uploaded"
                                });

                                // 3. Move file to completed folder immediately
                                MoveFileToCompleted(fileInfo.LocalFilePath, rootFolderPath);
                            }
                            catch (Exception ex)
                            {
                                _logAction($"❌ Failed to upload {fileInfo.FileName}: {ex.Message}");
                                errorCount++;
                                LastImportResults.Add(new FileImportResult
                                {
                                    CaseId = caseId,
                                    FolderPath = folderGroup.FolderPath,
                                    RelativePath = Path.GetRelativePath(rootFolderPath, fileInfo.LocalFilePath),
                                    FileName = fileInfo.FileName,
                                    Status = "Failed",
                                    ErrorMessage = ex.Message
                                });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logAction($"❌ Error processing folder {folderGroup.FolderName}: {ex.Message}");
                        errorCount += folderGroup.Files.Count;
                    }
                }

                // After all files processed, try to clean up empty folders in source
                CleanupEmptyFolders(rootFolderPath);
            }
            catch (Exception ex)
            {
                _logAction($"❌ Fatal error: {ex.Message}");
                throw;
            }
            finally
            {
                _s3Client?.Dispose();
                
                // Clean up temporary extraction folders
                CleanupTempFolders();
            }

            return (successCount, errorCount);
        }

        /// <summary>
        /// Move a single file to the completed folder, preserving relative path structure
        /// </summary>
        private void MoveFileToCompleted(string sourceFilePath, string rootFolderPath)
        {
            if (string.IsNullOrEmpty(CompletedFolderPath))
                return;

            try
            {
                // Get relative path from root folder
                var relativePath = Path.GetRelativePath(rootFolderPath, sourceFilePath);
                var destPath = Path.Combine(CompletedFolderPath, relativePath);
                var destDir = Path.GetDirectoryName(destPath);

                // Create destination directory if needed
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Move file
                if (File.Exists(sourceFilePath))
                {
                    // Handle if destination already exists
                    if (File.Exists(destPath))
                    {
                        var fileName = Path.GetFileNameWithoutExtension(destPath);
                        var ext = Path.GetExtension(destPath);
                        destPath = Path.Combine(destDir!, $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
                    }
                    
                    File.Move(sourceFilePath, destPath);
                }
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ Could not move file: {ex.Message}");
            }
        }

        /// <summary>
        /// Remove empty directories after files have been moved
        /// </summary>
        private void CleanupEmptyFolders(string rootFolderPath)
        {
            try
            {
                // Recursively delete empty subdirectories
                foreach (var dir in Directory.GetDirectories(rootFolderPath, "*", SearchOption.AllDirectories)
                    .OrderByDescending(d => d.Length)) // Process deepest first
                {
                    try
                    {
                        if (Directory.Exists(dir) && !Directory.EnumerateFileSystemEntries(dir).Any())
                        {
                            Directory.Delete(dir);
                        }
                    }
                    catch { }
                }

                // Delete root if empty
                if (Directory.Exists(rootFolderPath) && !Directory.EnumerateFileSystemEntries(rootFolderPath).Any())
                {
                    Directory.Delete(rootFolderPath);
                    _logAction($"📁 Removed empty source folder: {rootFolderPath}");
                }
            }
            catch { }
        }

        /// <summary>
        /// Check if a file with the same name already exists in the database for this case/folder
        /// Checks: caseId + folderId (if present) + fileName
        /// Note: S3 uses GUID filenames, so we check the database which stores original filenames
        /// </summary>
        private bool IsFileDuplicate(int caseId, int? folderId, string fileName)
        {
            try
            {
                using (var connection = new MySqlConnection(_connectionString))
                {
                    connection.Open();

                    // Check in tbl_case_correspondence_file for existing file with same name
                    string sql;
                    if (folderId.HasValue)
                    {
                        // Check within the specific folder AND case
                        sql = @"
                            SELECT COUNT(*) FROM tbl_case_correspondence_file 
                            WHERE fk_case_id = @CaseId 
                              AND fk_folder_id = @FolderId 
                              AND file_name = @FileName 
                              AND is_deleted = 0";

                        using (var cmd = new MySqlCommand(sql, connection))
                        {
                            cmd.Parameters.AddWithValue("@CaseId", caseId);
                            cmd.Parameters.AddWithValue("@FolderId", folderId.Value);
                            cmd.Parameters.AddWithValue("@FileName", fileName);
                            var count = Convert.ToInt32(cmd.ExecuteScalar());
                            return count > 0;
                        }
                    }
                    else
                    {
                        // Check for file across the entire case (root level, folder_id = 0)
                        sql = @"
                            SELECT COUNT(*) FROM tbl_case_correspondence_file 
                            WHERE fk_case_id = @CaseId 
                              AND fk_folder_id = 0 
                              AND file_name = @FileName 
                              AND is_deleted = 0";

                        using (var cmd = new MySqlCommand(sql, connection))
                        {
                            cmd.Parameters.AddWithValue("@CaseId", caseId);
                            cmd.Parameters.AddWithValue("@FileName", fileName);
                            var count = Convert.ToInt32(cmd.ExecuteScalar());
                            return count > 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ Duplicate check failed for {fileName}: {ex.Message}");
                return false; // If check fails, allow upload to proceed
            }
        }

        /// <summary>
        /// Save a single file to database (called immediately after S3 upload)
        /// </summary>
        private void SaveSingleFileToDatabase(int caseId, int? folderId, string folderName, FileUploadInfo fileInfo)
        {
            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        var documentFormatMap = GetDocumentFormats(connection, transaction);

                        // Create correspondence record
                        var correspondenceId = InsertCaseCorrespondence(connection, transaction, caseId);

                        // Create document record - use file name as title for individual tracking
                        var documentTitle = fileInfo.FileName;
                        var caseDocumentId = InsertCaseCorrespondenceDocument(
                            connection, transaction, caseId, correspondenceId, documentTitle);

                        // Create upload record
                        var uploadDocumentId = InsertCaseCorrespondenceDocumentUpload(
                            connection, transaction, caseDocumentId);

                        // Get format ID
                        var formatId = documentFormatMap.GetValueOrDefault(fileInfo.Extension, 0);
                        
                        // Create upload_file record
                        InsertCaseCorrespondenceDocumentUploadFile(
                            connection, transaction, uploadDocumentId, caseDocumentId, 
                            fileInfo, formatId);

                        // Link to folder
                        if (folderId.HasValue)
                        {
                            AddFileToFolder(connection, transaction, folderId.Value, 
                                caseId, correspondenceId, caseDocumentId, fileInfo.FileName, fileInfo.S3Key);
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception($"DB save failed for {fileInfo.FileName}: {ex.Message}");
                    }
                }
            }
        }

        private void CleanupTempFolders()
        {
            foreach (var tempFolder in _tempExtractionFolders)
            {
                try
                {
                    if (Directory.Exists(tempFolder))
                    {
                        Directory.Delete(tempFolder, true);
                        _logAction($"✅ Cleaned up temporary folder: {tempFolder}");
                    }
                }
                catch (Exception ex)
                {
                    _logAction($"⚠️ Warning: Could not delete temp folder {tempFolder}: {ex.Message}");
                }
            }
            _tempExtractionFolders.Clear();
        }

        private List<FolderFileGroup> ScanFolderStructure(string rootPath, int caseId)
        {
            var folderGroupsDict = new Dictionary<string, FolderFileGroup>();
            _tempExtractionFolders.Clear();

            ScanDirectory(rootPath, rootPath, folderGroupsDict);

            return folderGroupsDict.Values.ToList();
        }

        private void ScanDirectory(string currentPath, string rootPath, Dictionary<string, FolderFileGroup> folderGroupsDict)
        {
            try
            {
                var files = Directory.GetFiles(currentPath);
                foreach (var filePath in files)
                {
                    var fileInfo = new FileInfo(filePath);
                    var extension = fileInfo.Extension.TrimStart('.').ToLower();

                    // Skip ZIP files - don't process or extract them
                    if (extension == "zip")
                    {
                        _logAction($"⏭️ Skipping ZIP file: {fileInfo.Name}");
                        continue;
                    }
                    
                    // Regular file - process normally
                    var relativePath = Path.GetRelativePath(rootPath, filePath);
                    var folderPath = Path.GetDirectoryName(relativePath) ?? "";
                    var normalizedFolderPath = NormalizePath(folderPath);

                    var folderKey = string.IsNullOrEmpty(normalizedFolderPath) ? "ROOT" : normalizedFolderPath;
                    
                    if (!folderGroupsDict.ContainsKey(folderKey))
                    {
                        var folderName = string.IsNullOrEmpty(normalizedFolderPath) 
                            ? "" 
                            : Path.GetFileName(normalizedFolderPath);
                        
                        folderGroupsDict[folderKey] = new FolderFileGroup
                        {
                            FolderName = folderName,
                            FolderPath = normalizedFolderPath,
                            Files = new List<FileUploadInfo>()
                        };
                    }

                    folderGroupsDict[folderKey].Files.Add(new FileUploadInfo
                    {
                        LocalFilePath = filePath,
                        FileName = fileInfo.Name,
                        FolderPath = normalizedFolderPath,
                        Extension = extension
                    });
                }

                var subdirectories = Directory.GetDirectories(currentPath);
                foreach (var subDir in subdirectories)
                {
                    ScanDirectory(subDir, rootPath, folderGroupsDict);
                }
            }
            catch (Exception ex)
            {
                _logAction($"Warning: Error scanning directory {currentPath}: {ex.Message}");
            }
        }

        private string ExtractZipFile(string zipFilePath)
        {
            try
            {
                var zipFileInfo = new FileInfo(zipFilePath);
                
                // Create temp extraction folder in system temp directory to avoid conflicts
                var tempBasePath = Path.Combine(Path.GetTempPath(), "LeapMergeDoc_Extractions");
                if (!Directory.Exists(tempBasePath))
                {
                    Directory.CreateDirectory(tempBasePath);
                }
                
                var extractPath = Path.Combine(tempBasePath, $"extract_{Guid.NewGuid():N}");
                
                // Extract ZIP file
                ZipFile.ExtractToDirectory(zipFilePath, extractPath);
                
                _logAction($"✅ Extracted {zipFileInfo.Name} to {extractPath}");
                
                return extractPath;
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error extracting ZIP file {zipFilePath}: {ex.Message}");
                return string.Empty;
            }
        }

        private void ProcessExtractedZipContents(string extractPath, string rootPath, string zipParentFolderPath, string zipFolderName, Dictionary<string, FolderFileGroup> folderGroupsDict)
        {
            try
            {
                // Check if extracted ZIP has a root folder with the same name as ZIP
                // If so, skip that level to avoid duplication (e.g., Finance.zip contains Finance/ folder)
                var subdirectories = Directory.GetDirectories(extractPath);
                string actualExtractRoot = extractPath;
                bool skipDuplicateLevel = false;
                
                // Check if there's a subdirectory with the same name as the ZIP
                foreach (var subDir in subdirectories)
                {
                    var dirInfo = new DirectoryInfo(subDir);
                    if (dirInfo.Name.Equals(zipFolderName, StringComparison.OrdinalIgnoreCase))
                    {
                        // ZIP contains a folder with same name as ZIP - skip that level for files inside it
                        skipDuplicateLevel = true;
                        _logAction($"⚠️ ZIP '{zipFolderName}' contains folder '{zipFolderName}' - will skip duplicate level for nested files");
                        break;
                    }
                }
                
                // Get all files from extracted ZIP (recursively)
                var allFiles = Directory.GetFiles(extractPath, "*", SearchOption.AllDirectories);
                
                foreach (var extractedFile in allFiles)
                {
                    var fileInfo = new FileInfo(extractedFile);
                    var extension = fileInfo.Extension.TrimStart('.').ToLower();
                    
                    // Skip nested ZIP files
                    if (extension == "zip")
                    {
                        _logAction($"⚠️ Found nested ZIP in extracted contents: {fileInfo.Name} (skipping)");
                        continue;
                    }
                    
                    // Calculate relative path from extraction root
                    var relativeFromExtract = Path.GetRelativePath(extractPath, extractedFile);
                    var fileDirectory = Path.GetDirectoryName(relativeFromExtract) ?? "";
                    
                    // If we detected a duplicate level and this file is inside that duplicate folder, skip that level
                    if (skipDuplicateLevel && !string.IsNullOrEmpty(fileDirectory))
                    {
                        var dirParts = fileDirectory.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
                        if (dirParts.Length > 0 && dirParts[0].Equals(zipFolderName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Remove the duplicate level from the path
                            if (dirParts.Length > 1)
                            {
                                fileDirectory = string.Join("/", dirParts.Skip(1));
                            }
                            else
                            {
                                fileDirectory = ""; // File is directly in the duplicate folder, so it's at root of ZIP folder
                            }
                        }
                    }
                    
                    // Build final folder path: treat ZIP contents as if in a folder named after ZIP at root level
                    // If ZIP is in a subfolder, preserve that path, then add ZIP folder name
                    string finalFolderPath;
                    var normalizedZipParentPath = NormalizePath(zipParentFolderPath);
                    var normalizedFileDirectory = NormalizePath(fileDirectory);
                    
                    // Build path: zipParentFolderPath/zipFolderName/fileDirectory (if fileDirectory exists)
                    // But if we skipped a duplicate level, the fileDirectory already starts with the folder name
                    if (string.IsNullOrEmpty(normalizedZipParentPath))
                    {
                        // ZIP is at root, so folder is just zipFolderName/fileDirectory
                        if (string.IsNullOrEmpty(normalizedFileDirectory))
                        {
                            finalFolderPath = zipFolderName;
                        }
                        else
                        {
                            finalFolderPath = $"{zipFolderName}/{normalizedFileDirectory}";
                        }
                    }
                    else
                    {
                        // ZIP is in a subfolder, preserve that path
                        if (string.IsNullOrEmpty(normalizedFileDirectory))
                        {
                            finalFolderPath = $"{normalizedZipParentPath}/{zipFolderName}";
                        }
                        else
                        {
                            finalFolderPath = $"{normalizedZipParentPath}/{zipFolderName}/{normalizedFileDirectory}";
                        }
                    }
                    
                    finalFolderPath = NormalizePath(finalFolderPath);
                    var folderKey = string.IsNullOrEmpty(finalFolderPath) ? "ROOT" : finalFolderPath;
                    
                    // Check if folder already exists to avoid duplicates
                    if (!folderGroupsDict.ContainsKey(folderKey))
                    {
                        var folderName = string.IsNullOrEmpty(finalFolderPath) 
                            ? "" 
                            : Path.GetFileName(finalFolderPath);
                        
                        folderGroupsDict[folderKey] = new FolderFileGroup
                        {
                            FolderName = folderName,
                            FolderPath = finalFolderPath,
                            Files = new List<FileUploadInfo>()
                        };
                    }
                    
                    folderGroupsDict[folderKey].Files.Add(new FileUploadInfo
                    {
                        LocalFilePath = extractedFile,
                        FileName = fileInfo.Name,
                        FolderPath = finalFolderPath,
                        Extension = extension
                    });
                }
                
                _logAction($"✅ Processed {allFiles.Length} files from extracted ZIP");
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error processing extracted ZIP contents: {ex.Message}");
            }
        }

        private string UploadFileToS3(IAmazonS3 s3Client, FileUploadInfo fileInfo, int caseId)
        {
            var generatedFileName = GenerateUniqueFileName(fileInfo.FileName);
            var s3Key = $"casecorrdoc/{caseId}/{generatedFileName}";

            using (var fileStream = new FileStream(fileInfo.LocalFilePath, FileMode.Open, FileAccess.Read))
            {
                var request = new PutObjectRequest
                {
                    BucketName = _s3Config.BucketName,
                    Key = s3Key,
                    InputStream = fileStream,
                    ContentType = GetContentType(fileInfo.Extension)
                };

                var response = s3Client.PutObjectAsync(request).Result;
            }

            return s3Key;
        }

        private static string GenerateUniqueFileName(string filename)
        {
            string extension = Path.GetExtension(filename);
            return $"{Guid.NewGuid()}{extension}";
        }

        private string GetContentType(string extension)
        {
            return extension.ToLower() switch
            {
                "pdf" => "application/pdf",
                "doc" => "application/msword",
                "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "xls" => "application/vnd.ms-excel",
                "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "jpg" or "jpeg" => "image/jpeg",
                "png" => "image/png",
                "gif" => "image/gif",
                "txt" => "text/plain",
                "html" => "text/html",
                _ => "application/octet-stream"
            };
        }

        private int? CreateOrGetFolderWithParents(int caseId, string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
                return null;

            var normalizedPath = NormalizePath(folderPath);
            var pathParts = normalizedPath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (pathParts.Length == 0)
                return null;

            int? parentFolderId = null;
            string currentPath = "";

            foreach (var part in pathParts)
            {
                currentPath = string.IsNullOrEmpty(currentPath) ? part : $"{currentPath}/{part}";
                parentFolderId = CreateOrGetFolder(caseId, part, currentPath, parentFolderId);
                if (!parentFolderId.HasValue)
                {
                    _logAction($"⚠️ Warning: Failed to create folder '{part}' in path '{currentPath}'");
                    break;
                }
            }

            return parentFolderId;
        }

        private int? CreateOrGetFolder(int caseId, string folderName, string folderPath, int? parentFolderId)
        {
            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                string normalizedPath = NormalizePath(folderPath);
                
                // Check by folder_name + parent_folder_id to avoid duplicates
                string checkSql = @"
                    SELECT folder_id 
                    FROM tbl_case_correspondence_folder 
                    WHERE fk_case_id = @caseId 
                      AND folder_name = @folderName
                      AND (parent_folder_id = @parentFolderId OR (parent_folder_id IS NULL AND @parentFolderId IS NULL))
                      AND is_deleted = 0
                    LIMIT 1";

                using (var cmd = new MySqlCommand(checkSql, connection))
                {
                    cmd.Parameters.AddWithValue("@caseId", caseId);
                    cmd.Parameters.AddWithValue("@folderName", folderName);
                    if (parentFolderId.HasValue)
                    {
                        cmd.Parameters.AddWithValue("@parentFolderId", parentFolderId.Value);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@parentFolderId", DBNull.Value);
                    }
                    var existingId = cmd.ExecuteScalar();
                    if (existingId != null && existingId != DBNull.Value)
                    {
                        var folderId = Convert.ToInt32(existingId);
                        // Update folder_path if it's different
                        UpdateFolderPathIfNeeded(connection, folderId, normalizedPath);
                        return folderId;
                    }
                }

                string maxOrderSql = @"
                    SELECT COALESCE(MAX(folder_order), 0) 
                    FROM tbl_case_correspondence_folder 
                    WHERE fk_case_id = @caseId 
                      AND (parent_folder_id = @parentFolderId OR (parent_folder_id IS NULL AND @parentFolderId IS NULL))
                      AND is_deleted = 0";

                int folderOrder = 1;
                using (var cmd = new MySqlCommand(maxOrderSql, connection))
                {
                    cmd.Parameters.AddWithValue("@caseId", caseId);
                    if (parentFolderId.HasValue)
                    {
                        cmd.Parameters.AddWithValue("@parentFolderId", parentFolderId.Value);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@parentFolderId", DBNull.Value);
                    }
                    var maxOrder = cmd.ExecuteScalar();
                    if (maxOrder != null && maxOrder != DBNull.Value)
                    {
                        folderOrder = Convert.ToInt32(maxOrder) + 1;
                    }
                }

                string insertSql = @"
                    INSERT INTO tbl_case_correspondence_folder 
                    (fk_case_id, folder_name, folder_path, parent_folder_id, folder_order, created_by, created_date, is_deleted)
                    VALUES (@caseId, @folderName, @folderPath, @parentFolderId, @folderOrder, @createdBy, @createdDate, 0);
                    SELECT LAST_INSERT_ID();";

                using (var cmd = new MySqlCommand(insertSql, connection))
                {
                    cmd.Parameters.AddWithValue("@caseId", caseId);
                    cmd.Parameters.AddWithValue("@folderName", folderName);
                    cmd.Parameters.AddWithValue("@folderPath", normalizedPath);
                    if (parentFolderId.HasValue)
                    {
                        cmd.Parameters.AddWithValue("@parentFolderId", parentFolderId.Value);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@parentFolderId", DBNull.Value);
                    }
                    cmd.Parameters.AddWithValue("@folderOrder", folderOrder);
                    cmd.Parameters.AddWithValue("@createdBy", 1);
                    cmd.Parameters.AddWithValue("@createdDate", DateTime.UtcNow);

                    var folderId = cmd.ExecuteScalar();
                    return folderId != null ? Convert.ToInt32(folderId) : null;
                }
            }
        }

        private string NormalizePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;
            
            return path.Replace('\\', '/').Trim('/');
        }

        private void UpdateFolderPathIfNeeded(MySqlConnection connection, int folderId, string normalizedPath)
        {
            try
            {
                string updateSql = @"
                    UPDATE tbl_case_correspondence_folder 
                    SET folder_path = @folderPath 
                    WHERE folder_id = @folderId 
                      AND (folder_path IS NULL OR folder_path != @folderPath)";
                
                using (var cmd = new MySqlCommand(updateSql, connection))
                {
                    cmd.Parameters.AddWithValue("@folderId", folderId);
                    cmd.Parameters.AddWithValue("@folderPath", normalizedPath);
                    cmd.ExecuteNonQuery();
                }
            }
            catch
            {
            }
        }

        private void SaveFilesToDatabase(int caseId, int? folderId, string folderName, List<FileUploadInfo> files)
        {
            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        var documentFormatMap = GetDocumentFormats(connection, transaction);

                        // Create ONE correspondence record for the folder
                        var correspondenceId = InsertCaseCorrespondence(connection, transaction, caseId);

                        // Create ONE document record for the folder (document_title = folder name)
                        var documentTitle = string.IsNullOrEmpty(folderName) ? "Root Files" : folderName;
                        var caseDocumentId = InsertCaseCorrespondenceDocument(
                            connection, transaction, caseId, correspondenceId, documentTitle);

                        // Create ONE upload record for the folder
                        var uploadDocumentId = InsertCaseCorrespondenceDocumentUpload(
                            connection, transaction, caseDocumentId);

                        // Now process each file in the folder
                        foreach (var fileInfo in files)
                        {
                            var formatId = documentFormatMap.GetValueOrDefault(fileInfo.Extension, 0);
                            
                            // Create upload_file record for each file
                            InsertCaseCorrespondenceDocumentUploadFile(
                                connection, transaction, uploadDocumentId, caseDocumentId, 
                                fileInfo, formatId);

                            // Create file record linking to folder (if folder exists)
                            if (folderId.HasValue)
                            {
                                AddFileToFolder(connection, transaction, folderId.Value, 
                                    caseId, correspondenceId, caseDocumentId, fileInfo.FileName, fileInfo.S3Key);
                            }
                            else
                            {
                                _logAction($"⚠️ Warning: File {fileInfo.FileName} has no folderId, skipping folder mapping (file is at root level)");
                            }
                        }

                        transaction.Commit();
                        _logAction($"✅ Transaction committed for folder '{folderName}' with {files.Count} files");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logAction($"❌ Transaction rolled back: {ex.Message}");
                        throw;
                    }
                }
            }
        }

        private Dictionary<string, int> GetDocumentFormats(MySqlConnection connection, MySqlTransaction transaction)
        {
            var formatMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            string sql = "SELECT document_format_id, document_format_name FROM tbl_document_format";
            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var formatName = reader.GetString("document_format_name");
                        var formatId = reader.GetInt32("document_format_id");
                        formatMap[formatName] = formatId;
                    }
                }
            }

            return formatMap;
        }

        private int InsertCaseCorrespondence(MySqlConnection connection, MySqlTransaction transaction, int caseId)
        {
            string sql = @"
                INSERT INTO tbl_case_correspondence 
                (fk_case_id, fk_correspondence_item_type_id, fk_user_id, date_time, is_deleted)
                VALUES (@caseId, 1, 1, @dateTime, 0);
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@caseId", caseId);
                cmd.Parameters.AddWithValue("@dateTime", DateTime.UtcNow);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private int InsertCaseCorrespondenceDocument(MySqlConnection connection, MySqlTransaction transaction, 
            int caseId, int correspondenceId, string documentTitle)
        {
            string sql = @"
                INSERT INTO tbl_case_correspondence_document 
                (fk_case_id, fk_correspondence_id, document_title, document_category, is_draft, is_emailed, is_posted, is_printed)
                VALUES (@caseId, @correspondenceId, @documentTitle, 'upload', 0, 0, 0, 0);
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@caseId", caseId);
                cmd.Parameters.AddWithValue("@correspondenceId", correspondenceId);
                cmd.Parameters.AddWithValue("@documentTitle", documentTitle);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private int InsertCaseCorrespondenceDocumentUpload(MySqlConnection connection, MySqlTransaction transaction, int caseDocumentId)
        {
            string sql = @"
                INSERT INTO tbl_case_correspondence_document_upload 
                (fk_case_document_id, is_anti_money_laundering_measure, is_title_document, is_search_document, 
                 description, method, date_sent_received, document_reference)
                VALUES (@caseDocumentId, 0, 0, 0, NULL, 'other', @dateSentReceived, 'other');
                SELECT LAST_INSERT_ID();";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@caseDocumentId", caseDocumentId);
                var dateParam = new MySqlParameter("@dateSentReceived", MySqlDbType.Date)
                {
                    Value = DateTime.UtcNow.Date
                };
                cmd.Parameters.Add(dateParam);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void InsertCaseCorrespondenceDocumentUploadFile(MySqlConnection connection, MySqlTransaction transaction,
            int uploadDocumentId, int caseDocumentId, FileUploadInfo fileInfo, int formatId)
        {
            string sql = @"
                INSERT INTO tbl_case_correspondence_document_upload_file 
                (fk_correspondence_document_upload_id, document_name, fk_case_document_id, 
                 fk_document_format_id, is_editable, is_renamable, document_path)
                VALUES (@uploadDocumentId, @documentName, @caseDocumentId, @formatId, 0, 0, @documentPath)";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@uploadDocumentId", uploadDocumentId);
                cmd.Parameters.AddWithValue("@documentName", fileInfo.FileName);
                cmd.Parameters.AddWithValue("@caseDocumentId", caseDocumentId);
                cmd.Parameters.AddWithValue("@formatId", formatId);
                cmd.Parameters.AddWithValue("@documentPath", fileInfo.S3Key);
                cmd.ExecuteNonQuery();
            }
        }

        private void AddFileToFolder(MySqlConnection connection, MySqlTransaction transaction, 
            int folderId, int caseId, int correspondenceId, int caseDocumentId, string fileName, string filepath)
        {
            // Check for duplicate by S3 path (af1) and folder, not by fk_source_id
            // because all files in a folder share the same caseDocumentId
            string checkSql = @"
                SELECT file_id 
                FROM tbl_case_correspondence_file 
                WHERE fk_case_id = @caseId 
                  AND fk_folder_id = @folderId 
                  AND af1 = @filepath
                  AND is_deleted = 0
                LIMIT 1";

            using (var checkCmd = new MySqlCommand(checkSql, connection, transaction))
            {
                checkCmd.Parameters.AddWithValue("@caseId", caseId);
                checkCmd.Parameters.AddWithValue("@folderId", folderId);
                checkCmd.Parameters.AddWithValue("@filepath", filepath);
                var existingFileId = checkCmd.ExecuteScalar();
                if (existingFileId != null && existingFileId != DBNull.Value)
                {
                    _logAction($"⚠️ File already mapped to folder: {fileName} (file_id: {existingFileId})");
                    return;
                }
            }

            string maxOrderSql = @"
                SELECT COALESCE(MAX(file_order), 0) 
                FROM tbl_case_correspondence_file 
                WHERE fk_case_id = @caseId AND fk_folder_id = @folderId AND is_deleted = 0";

            int fileOrder = 1;
            using (var cmd = new MySqlCommand(maxOrderSql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@caseId", caseId);
                cmd.Parameters.AddWithValue("@folderId", folderId);
                var maxOrder = cmd.ExecuteScalar();
                if (maxOrder != null && maxOrder != DBNull.Value)
                {
                    fileOrder = Convert.ToInt32(maxOrder) + 1;
                }
            }

            string sql = @"
                INSERT INTO tbl_case_correspondence_file 
                (fk_case_id, fk_folder_id, fk_correspondence_id, file_name, file_order, 
                 source_type, fk_source_id, uploaded_by, uploaded_date, is_deleted, af1)
                VALUES (@caseId, @folderId, @correspondenceId, @fileName, @fileOrder, 
                        'DOCUMENT', @fkSourceId, @uploadedBy, @uploadedDate, 0, @filepath)";

            using (var cmd = new MySqlCommand(sql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@caseId", caseId);
                cmd.Parameters.AddWithValue("@folderId", folderId);
                cmd.Parameters.AddWithValue("@correspondenceId", correspondenceId);
                cmd.Parameters.AddWithValue("@fileName", fileName);
                cmd.Parameters.AddWithValue("@filepath", filepath);
                cmd.Parameters.AddWithValue("@fileOrder", fileOrder);
                cmd.Parameters.AddWithValue("@fkSourceId", caseDocumentId);
                cmd.Parameters.AddWithValue("@uploadedBy", 1);
                cmd.Parameters.AddWithValue("@uploadedDate", DateTime.UtcNow);
                cmd.ExecuteNonQuery();
                _logAction($"✅ Mapped file to folder: {fileName} -> folder_id: {folderId}");
            }
        }

        public (int filesDeleted, int errors) TruncateUploadedFiles(int caseId)
        {
            int filesDeleted = 0;
            int errors = 0;
            var s3KeysToDelete = new List<string>();

            using (var connection = new MySqlConnection(_connectionString))
            {
                connection.Open();

                _logAction("Fetching uploaded files from database...");
                string selectSql = @"
                    SELECT DISTINCT f.document_path
                    FROM tbl_case_correspondence_document_upload_file f
                    INNER JOIN tbl_case_correspondence_document d ON f.fk_case_document_id = d.case_document_id
                    WHERE d.fk_case_id = @caseId
                      AND f.document_path IS NOT NULL
                      AND f.document_path != ''";

                using (var cmd = new MySqlCommand(selectSql, connection))
                {
                    cmd.Parameters.AddWithValue("@caseId", caseId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var documentPath = reader.GetString("document_path");
                            if (!string.IsNullOrEmpty(documentPath))
                            {
                                s3KeysToDelete.Add(documentPath);
                            }
                        }
                    }
                }

                _logAction($"Found {s3KeysToDelete.Count} files to delete from S3");

                if (s3KeysToDelete.Count > 0)
                {
                    var s3Client = GetS3Client();
                    _logAction("Deleting files from S3 bucket...");

                    foreach (var s3Key in s3KeysToDelete)
                    {
                        try
                        {
                            var deleteRequest = new DeleteObjectRequest
                            {
                                BucketName = _s3Config.BucketName,
                                Key = s3Key
                            };

                            s3Client.DeleteObjectAsync(deleteRequest).Wait();
                            _logAction($"✅ Deleted from S3: {s3Key}");
                            filesDeleted++;
                        }
                        catch (Exception ex)
                        {
                            _logAction($"❌ Failed to delete {s3Key} from S3: {ex.Message}");
                            errors++;
                        }
                    }
                }

                        _logAction("Deleting records from database...");
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 0;", connection, transaction))
                        {
                            cmd.ExecuteNonQuery();
                        }

                        int totalDeleted = 0;

                        string deleteFileSql = @"
                            DELETE FROM tbl_case_correspondence_file
                            WHERE fk_case_id = @caseId";
                        using (var cmd = new MySqlCommand(deleteFileSql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@caseId", caseId);
                            int deleted = cmd.ExecuteNonQuery();
                            _logAction($"Deleted {deleted} rows from tbl_case_correspondence_file");
                            totalDeleted += deleted;
                        }

                        string deleteFolderSql = @"
                            DELETE FROM tbl_case_correspondence_folder
                            WHERE fk_case_id = @caseId";
                        using (var cmd = new MySqlCommand(deleteFolderSql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@caseId", caseId);
                            int deleted = cmd.ExecuteNonQuery();
                            _logAction($"Deleted {deleted} rows from tbl_case_correspondence_folder");
                            totalDeleted += deleted;
                        }

                        string deleteUploadFileSql = @"
                            DELETE f FROM tbl_case_correspondence_document_upload_file f
                            INNER JOIN tbl_case_correspondence_document d ON f.fk_case_document_id = d.case_document_id
                            WHERE d.fk_case_id = @caseId";
                        using (var cmd = new MySqlCommand(deleteUploadFileSql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@caseId", caseId);
                            int deleted = cmd.ExecuteNonQuery();
                            _logAction($"Deleted {deleted} rows from tbl_case_correspondence_document_upload_file");
                            totalDeleted += deleted;
                        }

                        string deleteUploadSql = @"
                            DELETE u FROM tbl_case_correspondence_document_upload u
                            INNER JOIN tbl_case_correspondence_document d ON u.fk_case_document_id = d.case_document_id
                            WHERE d.fk_case_id = @caseId";
                        using (var cmd = new MySqlCommand(deleteUploadSql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@caseId", caseId);
                            int deleted = cmd.ExecuteNonQuery();
                            _logAction($"Deleted {deleted} rows from tbl_case_correspondence_document_upload");
                            totalDeleted += deleted;
                        }

                        string deleteDocumentSql = @"
                            DELETE FROM tbl_case_correspondence_document
                            WHERE fk_case_id = @caseId";
                        using (var cmd = new MySqlCommand(deleteDocumentSql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@caseId", caseId);
                            int deleted = cmd.ExecuteNonQuery();
                            _logAction($"Deleted {deleted} rows from tbl_case_correspondence_document");
                            totalDeleted += deleted;
                        }

                        string deleteCorrespondenceSql = @"
                            DELETE FROM tbl_case_correspondence
                            WHERE fk_case_id = @caseId";
                        using (var cmd = new MySqlCommand(deleteCorrespondenceSql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@caseId", caseId);
                            int deleted = cmd.ExecuteNonQuery();
                            _logAction($"Deleted {deleted} rows from tbl_case_correspondence");
                            totalDeleted += deleted;
                        }

                        using (var cmd = new MySqlCommand("SET FOREIGN_KEY_CHECKS = 1;", connection, transaction))
                        {
                            cmd.ExecuteNonQuery();
                        }

                        transaction.Commit();
                        _logAction($"✅ Database records deleted successfully. Total: {totalDeleted} rows");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logAction($"❌ Database deletion failed: {ex.Message}");
                        throw;
                    }
                }
            }

            return (filesDeleted, errors);
        }
    }
}
