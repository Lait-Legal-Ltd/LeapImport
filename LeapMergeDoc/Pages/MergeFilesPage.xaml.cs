using Microsoft.Win32;
using OfficeOpenXml;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class MergeFilesPage : Page
    {
        private string? _matterOpenFilePath;
        private string? _archivedFilePath;
        private string? _matterListFilePath;

        public MergeFilesPage()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }

        private void BtnSelectMatterOpen_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Matter Open Excel File",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _matterOpenFilePath = dialog.FileName;
                txtMatterOpenPath.Text = _matterOpenFilePath;
                txtMatterOpenPath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                UpdateMergeButtonState();
                UpdateStatus($"Matter Open file selected: {Path.GetFileName(_matterOpenFilePath)}");
            }
        }

        private void BtnSelectArchived_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Archived Cases Excel File",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _archivedFilePath = dialog.FileName;
                txtArchivedPath.Text = _archivedFilePath;
                txtArchivedPath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                UpdateMergeButtonState();
                UpdateStatus($"Archived Cases file selected: {Path.GetFileName(_archivedFilePath)}");
            }
        }

        private void BtnSelectMatterList_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Matter List Excel File",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _matterListFilePath = dialog.FileName;
                txtMatterListPath.Text = _matterListFilePath;
                txtMatterListPath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                UpdateMergeButtonState();
                UpdateStatus($"Matter List file selected: {Path.GetFileName(_matterListFilePath)}");
            }
        }

        private void UpdateMergeButtonState()
        {
            btnMerge.IsEnabled = !string.IsNullOrEmpty(_matterOpenFilePath) &&
                                 !string.IsNullOrEmpty(_archivedFilePath) &&
                                 !string.IsNullOrEmpty(_matterListFilePath);
        }

        private void UpdateStatus(string message)
        {
            txtStatus.Text = $"[{DateTime.Now:HH:mm:ss}] {message}\n{txtStatus.Text}";
        }

        private async void BtnMerge_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_matterOpenFilePath) ||
                string.IsNullOrEmpty(_archivedFilePath) ||
                string.IsNullOrEmpty(_matterListFilePath))
            {
                MessageBox.Show("Please select all three Excel files.", "Files Required",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnMerge.IsEnabled = false;
            UpdateStatus("Starting merge process...");

            try
            {
                await Task.Run(() => MergeExcelFiles());
                UpdateStatus("✅ Merge completed successfully!");
                MessageBox.Show("Files merged successfully!", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Error: {ex.Message}");
                MessageBox.Show($"Error during merge: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnMerge.IsEnabled = true;
            }
        }

        private void MergeExcelFiles()
        {
            // Step 1: Read Archived Cases
            Dispatcher.Invoke(() => UpdateStatus("Reading Archived Cases file..."));
            var archivedData = ReadArchivedCases(_archivedFilePath!);
            int totalArchiveRecords = archivedData.Count;
            Dispatcher.Invoke(() => UpdateStatus($"Found {totalArchiveRecords} archived records."));

            // Step 2: Read Matter List
            Dispatcher.Invoke(() => UpdateStatus("Reading Matter List file..."));
            var matterListData = ReadMatterList(_matterListFilePath!);
            int totalMatterListRecords = matterListData.Count;
            Dispatcher.Invoke(() => UpdateStatus($"Found {totalMatterListRecords} Matter List records."));

            // Step 3: Process Matter Open file
            Dispatcher.Invoke(() => UpdateStatus("Processing Matter Open file..."));

            var outputFileName = Path.Combine(
                Path.GetDirectoryName(_matterOpenFilePath)!,
                $"{Path.GetFileNameWithoutExtension(_matterOpenFilePath)}_Merged_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            );

            File.Copy(_matterOpenFilePath!, outputFileName, overwrite: true);

            using var package = new ExcelPackage(new FileInfo(outputFileName));
            var worksheet = package.Workbook.Worksheets[0];

            int matterNoCol = FindColumnByHeader(worksheet, "Matter No");
            if (matterNoCol == -1)
            {
                throw new Exception("Could not find 'Matter No' column in Matter Open file.");
            }

            int lastCol = worksheet.Dimension?.End.Column ?? 1;
            int matterDescCol = lastCol + 1;
            int archiveDateCol = lastCol + 2;

            worksheet.Cells[1, matterDescCol].Value = "Matter Description";
            worksheet.Cells[1, archiveDateCol].Value = "Archive Date";
            worksheet.Cells[1, matterDescCol].Style.Font.Bold = true;
            worksheet.Cells[1, archiveDateCol].Style.Font.Bold = true;

            int lastRow = worksheet.Dimension?.End.Row ?? 1;
            int totalMatterOpenRecords = lastRow - 1;

            int matchedFromArchive = 0;
            int matchedFromMatterList = 0;
            int unmatchedCount = 0;

            for (int row = 2; row <= lastRow; row++)
            {
                var matterNo = worksheet.Cells[row, matterNoCol].Text?.Trim();
                string? matterDescription = null;
                object? archiveDate = null;
                bool foundInArchive = false;
                bool foundInMatterList = false;

                if (!string.IsNullOrEmpty(matterNo) && archivedData.TryGetValue(matterNo, out var archiveInfo))
                {
                    matterDescription = archiveInfo.MatterDescription;
                    archiveDate = archiveInfo.ArchiveDate;
                    foundInArchive = true;
                    matchedFromArchive++;
                }

                if (string.IsNullOrEmpty(matterDescription) &&
                    !string.IsNullOrEmpty(matterNo) &&
                    matterListData.TryGetValue(matterNo, out var matterListDesc))
                {
                    matterDescription = matterListDesc;
                    foundInMatterList = true;
                    matchedFromMatterList++;
                }

                worksheet.Cells[row, matterDescCol].Value = matterDescription ?? "";
                worksheet.Cells[row, archiveDateCol].Value = archiveDate ?? "";

                if (archiveDate is DateTime)
                {
                    worksheet.Cells[row, archiveDateCol].Style.Numberformat.Format = "dd/MM/yyyy";
                }

                if (!foundInArchive && !foundInMatterList)
                {
                    unmatchedCount++;
                }
            }

            worksheet.Column(matterDescCol).AutoFit();
            worksheet.Column(archiveDateCol).AutoFit();

            package.Save();

            Dispatcher.Invoke(() =>
            {
                UpdateStatus("═══════════════════════════════════════");
                UpdateStatus($"📊 SUMMARY:");
                UpdateStatus($"   Total Archive records: {totalArchiveRecords}");
                UpdateStatus($"   Total Matter List records: {totalMatterListRecords}");
                UpdateStatus($"   Total Matter Open records: {totalMatterOpenRecords}");
                UpdateStatus($"   ───────────────────────────────────");
                UpdateStatus($"   ✅ Matched from Archive: {matchedFromArchive}");
                UpdateStatus($"   ✅ Matched from Matter List: {matchedFromMatterList}");
                UpdateStatus($"   ⬜ No match found (empty): {unmatchedCount}");
                UpdateStatus("═══════════════════════════════════════");
                UpdateStatus($"Output saved to: {Path.GetFileName(outputFileName)}");
            });
        }

        private Dictionary<string, ArchivedCaseInfo> ReadArchivedCases(string filePath)
        {
            var result = new Dictionary<string, ArchivedCaseInfo>(StringComparer.OrdinalIgnoreCase);

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];

            int matterNoCol = FindColumnByHeader(worksheet, "Matter No");
            int matterDescCol = FindColumnByHeader(worksheet, "Matter Description");
            int archiveDateCol = FindColumnByHeader(worksheet, "Archive Date");

            if (matterNoCol == -1)
                throw new Exception("Could not find 'Matter No' column in Archived Cases file.");
            if (matterDescCol == -1)
                throw new Exception("Could not find 'Matter Description' column in Archived Cases file.");
            if (archiveDateCol == -1)
                throw new Exception("Could not find 'Archive Date' column in Archived Cases file.");

            int lastRow = worksheet.Dimension?.End.Row ?? 1;

            for (int row = 2; row <= lastRow; row++)
            {
                var matterNo = worksheet.Cells[row, matterNoCol].Text?.Trim();

                if (!string.IsNullOrEmpty(matterNo) && !result.ContainsKey(matterNo))
                {
                    result[matterNo] = new ArchivedCaseInfo
                    {
                        MatterDescription = worksheet.Cells[row, matterDescCol].Text,
                        ArchiveDate = worksheet.Cells[row, archiveDateCol].Value
                    };
                }
            }

            return result;
        }

        private Dictionary<string, string> ReadMatterList(string filePath)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];

            int matterNoCol = FindColumnByHeader(worksheet, "Matter No");
            int matterDescCol = FindColumnByHeader(worksheet, "Matter Description");

            if (matterNoCol == -1)
                throw new Exception("Could not find 'Matter No' column in Matter List file.");
            if (matterDescCol == -1)
                throw new Exception("Could not find 'Matter Description' column in Matter List file.");

            int lastRow = worksheet.Dimension?.End.Row ?? 1;

            for (int row = 2; row <= lastRow; row++)
            {
                var matterNo = worksheet.Cells[row, matterNoCol].Text?.Trim();
                var matterDesc = worksheet.Cells[row, matterDescCol].Text?.Trim();

                if (!string.IsNullOrEmpty(matterNo) && !string.IsNullOrEmpty(matterDesc) && !result.ContainsKey(matterNo))
                {
                    result[matterNo] = matterDesc;
                }
            }

            return result;
        }

        private static int FindColumnByHeader(ExcelWorksheet worksheet, string headerName)
        {
            if (worksheet.Dimension == null) return -1;

            int lastCol = worksheet.Dimension.End.Column;

            for (int col = 1; col <= lastCol; col++)
            {
                var cellValue = worksheet.Cells[1, col].Text?.Trim();

                if (!string.IsNullOrEmpty(cellValue) &&
                    (cellValue.Equals(headerName, StringComparison.OrdinalIgnoreCase) ||
                     cellValue.StartsWith(headerName, StringComparison.OrdinalIgnoreCase)))
                {
                    return col;
                }
            }

            return -1;
        }
    }

    public class ArchivedCaseInfo
    {
        public string? MatterDescription { get; set; }
        public object? ArchiveDate { get; set; }
    }
}
