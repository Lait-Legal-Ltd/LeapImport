using LeapMergeDoc.Models;
using LeapMergeDoc.Services;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class UserImportPage : Page
    {
        private string? _userFilePath;
        private List<UserExcelRowData>? _excelData;
        private List<ProcessedUserData>? _processedData;

        public UserImportPage()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }

        private void UpdateStatus(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                txtStatus.Text = $"[{DateTime.Now:HH:mm:ss}] {message}\n{txtStatus.Text}";
            }
            else
            {
                Dispatcher.Invoke(() => txtStatus.Text = $"[{DateTime.Now:HH:mm:ss}] {message}\n{txtStatus.Text}");
            }
        }

        private void BtnTestConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var connection = new MySqlConnection(txtConnectionString.Text))
                {
                    connection.Open();
                    UpdateStatus("✅ Database connection successful!");
                    MessageBox.Show("Connection successful!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Connection failed: {ex.Message}");
                MessageBox.Show($"Connection failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnSelectUserFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select User Excel/CSV File",
                Filter = "Excel/CSV Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _userFilePath = dialog.FileName;
                txtUserFilePath.Text = _userFilePath;
                txtUserFilePath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                btnPreview.IsEnabled = true;
                UpdateStatus($"File selected: {Path.GetFileName(_userFilePath)}");
            }
        }

        private async void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_userFilePath))
            {
                MessageBox.Show("Please select a user Excel/CSV file.", "File Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnPreview.IsEnabled = false;
            btnImport.IsEnabled = false;
            btnExport.IsEnabled = false;
            UpdateStatus("Reading file...");

            try
            {
                var importService = new UserImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    _excelData = importService.ReadExcelData(_userFilePath);
                    Dispatcher.Invoke(() => UpdateStatus($"Read {_excelData.Count} records from file."));

                    _processedData = importService.ProcessExcelData(_excelData);
                    Dispatcher.Invoke(() => UpdateStatus($"Processed {_processedData.Count} user records."));

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════════════════════");
                        UpdateStatus($"📊 USER PREVIEW SUMMARY:");
                        UpdateStatus($"═══════════════════════════════════════════════════════");

                        // Basic counts
                        UpdateStatus($"📁 FILE DATA:");
                        UpdateStatus($"   Total records: {_excelData.Count}");
                        UpdateStatus($"   Processed users: {_processedData.Count}");

                        // Duplicate analysis (by Email)
                        int duplicates = _processedData.Count(p => p.IsDuplicate);
                        int newUsers = _processedData.Count(p => !p.IsDuplicate);

                        UpdateStatus($"");
                        UpdateStatus($"🔍 DUPLICATE CHECK (by Email):");
                        UpdateStatus($"   New users (will be imported): {newUsers}");
                        UpdateStatus($"   Duplicates found (same email in DB): {duplicates}");

                        // Data quality
                        int withEmail = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.Email));
                        int withoutEmail = _processedData.Count(p => string.IsNullOrEmpty(p.OriginalData?.Email));
                        int withUserCode = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.UserCode));
                        int withPhone = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.Mobile) || !string.IsNullOrEmpty(p.OriginalData?.HomePhone));
                        int withTitle = _processedData.Count(p => p.TitleId.HasValue);

                        UpdateStatus($"");
                        UpdateStatus($"📋 DATA QUALITY:");
                        UpdateStatus($"   With Email: {withEmail} / {_processedData.Count} (used for duplicate check)");
                        if (withoutEmail > 0)
                        {
                            UpdateStatus($"   ⚠️ WITHOUT Email: {withoutEmail} (cannot check duplicates!)");
                        }
                        UpdateStatus($"   With UserCode/Initials: {withUserCode} / {_processedData.Count}");
                        UpdateStatus($"   With Title: {withTitle} / {_processedData.Count}");
                        UpdateStatus($"   With Phone: {withPhone} / {_processedData.Count}");

                        // Show duplicate details
                        if (duplicates > 0)
                        {
                            UpdateStatus($"");
                            UpdateStatus($"⚠️ DUPLICATES FOUND:");
                            foreach (var dup in _processedData.Where(p => p.IsDuplicate).Take(10))
                            {
                                UpdateStatus($"   • {dup.OriginalData?.FullName} - {dup.DuplicateReason}");
                            }
                            if (duplicates > 10)
                            {
                                UpdateStatus($"   ... and {duplicates - 10} more");
                            }
                        }

                        // Sample of first 5 new users
                        var newUsersList = _processedData.Where(p => !p.IsDuplicate).Take(5).ToList();
                        if (newUsersList.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"📄 SAMPLE NEW USERS (first 5):");
                            foreach (var user in newUsersList)
                            {
                                var name = user.OriginalData?.FullName ?? "Unknown";
                                var code = user.OriginalData?.UserCode ?? "No Code";
                                UpdateStatus($"   • {name} (Code: {code})");
                            }
                            if (newUsers > 5)
                            {
                                UpdateStatus($"   ... and {newUsers - 5} more");
                            }
                        }

                        UpdateStatus($"═══════════════════════════════════════════════════════");
                    });
                });

                btnImport.IsEnabled = true;
                btnExport.IsEnabled = true;
                UpdateStatus("✅ Preview complete. Ready to import or export.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Error: {ex.Message}");
                MessageBox.Show($"Error during preview: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnPreview.IsEnabled = true;
            }
        }

        private async void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            if (_processedData == null || _processedData.Count == 0)
            {
                MessageBox.Show("Please preview the data first.", "Preview Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool skipDuplicates = chkSkipDuplicates.IsChecked == true;
            int newUsersCount = _processedData.Count(p => !p.IsDuplicate);
            int duplicatesCount = _processedData.Count(p => p.IsDuplicate);

            string message = skipDuplicates
                ? $"Are you sure you want to import {newUsersCount} new users to the database?\n\n{duplicatesCount} duplicates will be skipped."
                : $"Are you sure you want to import {_processedData.Count} users to the database?\n\n⚠️ WARNING: {duplicatesCount} duplicates will be imported (may cause errors)!";

            var result = MessageBox.Show(
                message,
                "Confirm Import",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                UpdateStatus("Import cancelled by user.");
                return;
            }

            btnPreview.IsEnabled = false;
            btnImport.IsEnabled = false;
            btnExport.IsEnabled = false;
            UpdateStatus("Starting user import...");

            try
            {
                var importService = new UserImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var (success, skipped, errors) = importService.ImportToDatabase(_processedData, skipDuplicates);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"📊 IMPORT COMPLETE:");
                        UpdateStatus($"   ✅ Successfully imported: {success}");
                        UpdateStatus($"   ⏭️ Skipped (duplicates): {skipped}");
                        UpdateStatus($"   ❌ Errors: {errors}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("User import completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Import failed: {ex.Message}");
                MessageBox.Show($"Import failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnPreview.IsEnabled = true;
                btnImport.IsEnabled = true;
                btnExport.IsEnabled = true;
            }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_processedData == null || _processedData.Count == 0)
            {
                MessageBox.Show("Please preview the data first.", "Preview Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Title = "Export Users to Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                FileName = $"UserExport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (saveDialog.ShowDialog() != true)
            {
                UpdateStatus("Export cancelled by user.");
                return;
            }

            btnPreview.IsEnabled = false;
            btnImport.IsEnabled = false;
            btnExport.IsEnabled = false;
            UpdateStatus("Starting user export...");

            try
            {
                var importService = new UserImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    importService.ExportToExcel(_processedData, saveDialog.FileName);

                    Dispatcher.Invoke(() =>
                    {
                        int duplicates = _processedData.Count(p => p.IsDuplicate);
                        int newUsers = _processedData.Count(p => !p.IsDuplicate);

                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"📊 EXPORT COMPLETE:");
                        UpdateStatus($"   Total users: {_processedData.Count}");
                        UpdateStatus($"   New users: {newUsers}");
                        UpdateStatus($"   Duplicates: {duplicates}");
                        UpdateStatus($"   File: {saveDialog.FileName}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show($"Exported {_processedData.Count} users successfully!\n\nFile: {saveDialog.FileName}",
                    "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Export failed: {ex.Message}");
                MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnPreview.IsEnabled = true;
                btnImport.IsEnabled = true;
                btnExport.IsEnabled = true;
            }
        }

        private async void BtnTruncate_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "⚠️ WARNING: This will DELETE ALL DATA from tbl_user!\n\n" +
                "This action CANNOT be undone!\n\n" +
                "Are you sure you want to proceed?",
                "Confirm Truncate",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                UpdateStatus("Truncate cancelled by user.");
                return;
            }

            var result2 = MessageBox.Show(
                "FINAL CONFIRMATION: All user data will be permanently deleted.\n\n" +
                "Click Yes to confirm.",
                "Final Confirmation",
                MessageBoxButton.YesNo,
                MessageBoxImage.Stop);

            if (result2 != MessageBoxResult.Yes)
            {
                UpdateStatus("Truncate cancelled by user.");
                return;
            }

            btnTruncate.IsEnabled = false;
            UpdateStatus("Starting user data truncation...");

            try
            {
                var importService = new UserImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var (rowsDeleted, message) = importService.TruncateUserData();

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"🗑️ TRUNCATE COMPLETE:");
                        UpdateStatus($"   Total rows deleted: {rowsDeleted}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("User data truncated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Truncate failed: {ex.Message}");
                MessageBox.Show($"Truncate failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnTruncate.IsEnabled = true;
            }
        }
    }
}
