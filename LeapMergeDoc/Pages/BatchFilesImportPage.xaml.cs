using LeapMergeDoc.Models;
using LeapMergeDoc.Services;
using MySql.Data.MySqlClient;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class BatchFilesImportPage : Page
    {
        private string? _rootFolderPath;
        private string? _lastMissingReportPath;
        private ObservableCollection<CaseFolderViewModel> _caseFolders = new();

        public BatchFilesImportPage()
        {
            InitializeComponent();
            dgCaseFolders.ItemsSource = _caseFolders;
        }

        private void Log(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                txtLog.Text = $"[{DateTime.Now:HH:mm:ss}] {message}\n{txtLog.Text}";
            }
            else
            {
                Dispatcher.Invoke(() => txtLog.Text = $"[{DateTime.Now:HH:mm:ss}] {message}\n{txtLog.Text}");
            }
        }

        private void BtnTestConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using var connection = new MySqlConnection(txtConnectionString.Text);
                connection.Open();
                Log("✅ Database connection successful!");
                MessageBox.Show("Connection successful!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Log($"❌ Connection failed: {ex.Message}");
                MessageBox.Show($"Connection failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnBrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select the root folder containing case subfolders",
                ShowNewFolderButton = false
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _rootFolderPath = dialog.SelectedPath;
                txtRootFolder.Text = _rootFolderPath;
                txtRootFolder.Foreground = System.Windows.Media.Brushes.Black;
                btnScanFolder.IsEnabled = true;
                btnDownloadAllFiles.IsEnabled = true;
                Log($"📁 Selected folder: {_rootFolderPath}");
                
                // Auto-suggest completed folder path
                if (string.IsNullOrEmpty(txtCompletedFolder.Text))
                {
                    var parentDir = Path.GetDirectoryName(_rootFolderPath);
                    if (!string.IsNullOrEmpty(parentDir))
                    {
                        txtCompletedFolder.Text = Path.Combine(parentDir, "Completed");
                        txtCompletedFolder.Foreground = System.Windows.Media.Brushes.Black;
                    }
                }
            }
        }

        private void BtnBrowseCompletedFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select folder to move successfully uploaded folders to",
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtCompletedFolder.Text = dialog.SelectedPath;
                txtCompletedFolder.Foreground = System.Windows.Media.Brushes.Black;
                Log($"📂 Completed folder: {dialog.SelectedPath}");
            }
        }

        private async void BtnScanFolder_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_rootFolderPath) || !Directory.Exists(_rootFolderPath))
            {
                MessageBox.Show("Please select a valid root folder.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnScanFolder.IsEnabled = false;
            _caseFolders.Clear();

            try
            {
                var s3Config = new S3Configuration
                {
                    BucketName = txtBucketName.Text,
                    ProfilePath = txtProfilePath.Text
                };

                var service = new BatchFilesImportService(txtConnectionString.Text, s3Config, Log);

                // Scan folders
                Log("📂 Scanning folders...");
                var folders = await Task.Run(() => service.ScanRootFolder(_rootFolderPath));

                // Convert to view models
                foreach (var folder in folders)
                {
                    _caseFolders.Add(new CaseFolderViewModel
                    {
                        FolderPath = folder.FolderPath,
                        FolderName = folder.FolderName,
                        ParsedCaseReference = folder.ParsedCaseReference,
                        FileCount = folder.FileCount,
                        IsSelected = true,
                        Status = "Pending lookup"
                    });
                }

                // Lookup case IDs
                Log("🔍 Looking up cases in database...");
                await Task.Run(() =>
                {
                    var folderInfos = _caseFolders.Select(vm => new BatchFilesImportService.CaseFolderInfo
                    {
                        FolderPath = vm.FolderPath,
                        FolderName = vm.FolderName,
                        ParsedCaseReference = vm.ParsedCaseReference,
                        FileCount = vm.FileCount
                    }).ToList();

                    service.LookupCaseIds(folderInfos);

                    // Update view models with results
                    Dispatcher.Invoke(() =>
                    {
                        for (int i = 0; i < folderInfos.Count; i++)
                        {
                            _caseFolders[i].CaseId = folderInfos[i].CaseId;
                            _caseFolders[i].CaseName = folderInfos[i].CaseName;
                            _caseFolders[i].Status = folderInfos[i].Status;
                            _caseFolders[i].IsSelected = folderInfos[i].CaseId.HasValue;
                        }
                    });
                });

                // Update summary
                var found = _caseFolders.Count(f => f.CaseId.HasValue);
                var notFound = _caseFolders.Count - found;
                var totalFiles = _caseFolders.Where(f => f.CaseId.HasValue).Sum(f => f.FileCount);
                var totalScannedFiles = _caseFolders.Sum(f => f.FileCount);
                txtSummary.Text = $"📊 {found} cases found, {notFound} not found | {totalScannedFiles} files scanned ({totalFiles} in matched cases)";

                // Enable buttons
                btnSelectAll.IsEnabled = true;
                btnSelectNone.IsEnabled = true;
                btnSelectFound.IsEnabled = true;
                btnImport.IsEnabled = found > 0;

                // Generate initial missing-files report right after scan.
                _lastMissingReportPath = ExportMissingFilesReport();
                btnDownloadReport.IsEnabled = !string.IsNullOrWhiteSpace(_lastMissingReportPath);
                btnDownloadAllFiles.IsEnabled = true;
                if (!string.IsNullOrWhiteSpace(_lastMissingReportPath))
                {
                    Log($"📄 Missing files report ready: {_lastMissingReportPath}");
                }

                dgCaseFolders.Items.Refresh();
            }
            catch (Exception ex)
            {
                Log($"❌ Error scanning: {ex.Message}");
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnScanFolder.IsEnabled = true;
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _caseFolders)
            {
                folder.IsSelected = true;
            }
            dgCaseFolders.Items.Refresh();
            UpdateSelectionSummary();
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _caseFolders)
            {
                folder.IsSelected = false;
            }
            dgCaseFolders.Items.Refresh();
            UpdateSelectionSummary();
        }

        private void BtnSelectFound_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _caseFolders)
            {
                folder.IsSelected = folder.CaseId.HasValue;
            }
            dgCaseFolders.Items.Refresh();
            UpdateSelectionSummary();
        }

        private void UpdateSelectionSummary()
        {
            var selected = _caseFolders.Count(f => f.IsSelected && f.CaseId.HasValue);
            var totalFiles = _caseFolders.Where(f => f.IsSelected && f.CaseId.HasValue).Sum(f => f.FileCount);
            txtSummary.Text = $"📊 {selected} cases selected | {totalFiles} files to import";
            btnImport.IsEnabled = selected > 0;
        }

        private async void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var selectedFolders = _caseFolders.Where(f => f.IsSelected && f.CaseId.HasValue).ToList();
            
            if (selectedFolders.Count == 0)
            {
                MessageBox.Show("No valid cases selected for import.", "Nothing to Import", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var totalFiles = selectedFolders.Sum(f => f.FileCount);
            var result = MessageBox.Show(
                $"Are you sure you want to import files from {selectedFolders.Count} case folders?\n\n" +
                $"Total files to upload: {totalFiles}\n\n" +
                "This may take a while for large files.",
                "Confirm Import",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                Log("Import cancelled by user.");
                return;
            }

            btnImport.IsEnabled = false;
            btnScanFolder.IsEnabled = false;

            try
            {
                var s3Config = new S3Configuration
                {
                    BucketName = txtBucketName.Text,
                    ProfilePath = txtProfilePath.Text
                };

                var service = new BatchFilesImportService(txtConnectionString.Text, s3Config, Log);

                // Get completed folder path (if enabled)
                string? completedFolderPath = null;
                if (chkMoveOnSuccess.IsChecked == true && !string.IsNullOrWhiteSpace(txtCompletedFolder.Text))
                {
                    completedFolderPath = txtCompletedFolder.Text;
                }

                // Convert view models to service model
                var folderInfos = selectedFolders.Select(vm => new BatchFilesImportService.CaseFolderInfo
                {
                    FolderPath = vm.FolderPath,
                    FolderName = vm.FolderName,
                    ParsedCaseReference = vm.ParsedCaseReference,
                    CaseId = vm.CaseId,
                    CaseName = vm.CaseName,
                    FileCount = vm.FileCount,
                    IsSelected = vm.IsSelected,
                    Status = vm.Status
                }).ToList();

                await Task.Run(() =>
                {
                    var (totalSuccess, totalErrors, casesProcessed, missingImportFiles) = service.ImportSelectedFolders(folderInfos, completedFolderPath);

                    // Update view model statuses from results
                    Dispatcher.Invoke(() =>
                    {
                        for (int i = 0; i < folderInfos.Count; i++)
                        {
                            var vm = selectedFolders[i];
                            vm.Status = folderInfos[i].Status;
                        }
                        dgCaseFolders.Items.Refresh();

                        var reportPath = ExportMissingFilesReport(missingImportFiles);
                        _lastMissingReportPath = reportPath;
                        btnDownloadReport.IsEnabled = !string.IsNullOrWhiteSpace(_lastMissingReportPath);
                        if (!string.IsNullOrWhiteSpace(reportPath))
                        {
                            Log($"📄 Missing files report saved: {reportPath}");
                        }
                    });
                });

                MessageBox.Show("Batch import completed! Check the activity log for details.", "Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Log($"❌ Import error: {ex.Message}");
                MessageBox.Show($"Import error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnImport.IsEnabled = true;
                btnScanFolder.IsEnabled = true;
            }
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtLog.Text = "";
        }

        private void BtnDownloadReport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_lastMissingReportPath) || !File.Exists(_lastMissingReportPath))
            {
                MessageBox.Show("No report found yet. Please run Scan & Lookup first.", "Report Not Available", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _lastMissingReportPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Log($"❌ Could not open report: {ex.Message}");
                MessageBox.Show($"Could not open report: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnDownloadAllFiles_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_rootFolderPath) || !Directory.Exists(_rootFolderPath))
            {
                MessageBox.Show("Please select a valid root folder first.", "Folder Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var allFiles = Directory.GetFiles(_rootFolderPath, "*", SearchOption.AllDirectories)
                    .OrderBy(f => f)
                    .ToList();

                if (allFiles.Count == 0)
                {
                    MessageBox.Show("No files found to package.", "Nothing to Download", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var parentDir = Path.GetDirectoryName(_rootFolderPath) ?? _rootFolderPath;
                var zipPath = Path.Combine(parentDir, $"all-files-{DateTime.Now:yyyyMMdd-HHmmss}.zip");

                using (var archive = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                {
                    foreach (var file in allFiles)
                    {
                        var relativePath = Path.GetRelativePath(_rootFolderPath, file);
                        archive.CreateEntryFromFile(file, relativePath, CompressionLevel.Fastest);
                    }
                }

                Log($"📦 All files ZIP created: {zipPath} ({allFiles.Count} files)");

                Process.Start(new ProcessStartInfo
                {
                    FileName = zipPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Log($"❌ Could not create ZIP: {ex.Message}");
                MessageBox.Show($"Could not create ZIP: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string? ExportMissingFilesReport(List<BatchFilesImportService.MissingFileInfo>? missingImportFiles = null)
        {
            if (string.IsNullOrWhiteSpace(_rootFolderPath))
            {
                return null;
            }

            try
            {
                var lines = new List<string>
                {
                    $"Missing Files Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                    $"Root Folder: {_rootFolderPath}",
                    ""
                };

                // Full scan view so user can see every file, including ZIP files.
                lines.Add("=== All Scanned Files (includes ZIP) ===");
                var allScannedFiles = Directory.GetFiles(_rootFolderPath, "*", SearchOption.AllDirectories)
                    .OrderBy(f => f)
                    .ToList();
                if (allScannedFiles.Count == 0)
                {
                    lines.Add("None");
                }
                else
                {
                    foreach (var file in allScannedFiles)
                    {
                        var rel = Path.GetRelativePath(_rootFolderPath, file);
                        var type = file.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) ? "ZIP" : "FILE";
                        lines.Add($"{rel} | {type}");
                    }
                }

                // Case folders that were scanned but not found in DB.
                lines.Add("");
                lines.Add("=== Case Not Found (from scan) ===");
                var notFoundFolders = _caseFolders.Where(f => !f.CaseId.HasValue).ToList();
                if (notFoundFolders.Count == 0)
                {
                    lines.Add("None");
                }
                else
                {
                    foreach (var folder in notFoundFolders)
                    {
                        if (!Directory.Exists(folder.FolderPath))
                        {
                            lines.Add($"{folder.FolderName} | [folder missing] {folder.FolderPath}");
                            continue;
                        }

                        var files = Directory.GetFiles(folder.FolderPath, "*", SearchOption.AllDirectories)
                            .OrderBy(f => f)
                            .ToList();

                        if (files.Count == 0)
                        {
                            lines.Add($"{folder.FolderName} | [no files]");
                            continue;
                        }

                        foreach (var file in files)
                        {
                            var rel = Path.GetRelativePath(_rootFolderPath, file);
                            lines.Add($"{folder.FolderName} | {rel}");
                        }
                    }
                }

                lines.Add("");
                lines.Add("=== Import Failed (after Import Selected) ===");
                var importFailures = missingImportFiles ?? new List<BatchFilesImportService.MissingFileInfo>();
                if (importFailures.Count == 0)
                {
                    lines.Add("None");
                }
                else
                {
                    foreach (var item in importFailures.OrderBy(m => m.RelativePath))
                    {
                        lines.Add($"{item.FolderName} | {item.RelativePath} | {item.Reason}");
                    }
                }

                var reportFileName = $"missing-files-{DateTime.Now:yyyyMMdd-HHmmss}.txt";
                var reportPath = Path.Combine(_rootFolderPath, reportFileName);
                File.WriteAllLines(reportPath, lines);
                return reportPath;
            }
            catch (Exception ex)
            {
                Log($"⚠️ Could not create missing files report: {ex.Message}");
                return null;
            }
        }
    }

    /// <summary>
    /// View model for case folders in the DataGrid
    /// </summary>
    public class CaseFolderViewModel : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _status = "";

        public string FolderPath { get; set; } = "";
        public string FolderName { get; set; } = "";
        public string ParsedCaseReference { get; set; } = "";
        public int? CaseId { get; set; }
        public string? CaseName { get; set; }
        public int FileCount { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged(nameof(Status));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
