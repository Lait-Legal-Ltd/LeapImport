using LeapMergeDoc.Models;
using LeapMergeDoc.Services;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class FilesCollectionImportPage : Page
    {
        private string? _selectedFolderPath;
        private int? _caseId;

        public FilesCollectionImportPage()
        {
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

        private void BtnVerifyCase_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(txtCaseId.Text, out int caseId) || caseId <= 0)
            {
                MessageBox.Show("Please enter a valid Case ID.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using (var connection = new MySqlConnection(txtConnectionString.Text))
                {
                    connection.Open();
                    string sql = "SELECT case_id, case_name FROM tbl_case_details_general WHERE case_id = @caseId";
                    using (var cmd = new MySqlCommand(sql, connection))
                    {
                        cmd.Parameters.AddWithValue("@caseId", caseId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                _caseId = caseId;
                                string caseName = reader.IsDBNull(reader.GetOrdinal("case_name")) ? "Unknown" : reader.GetString("case_name");
                                UpdateStatus($"✅ Case verified: {caseName} (ID: {caseId})");
                                MessageBox.Show($"Case verified:\n{caseName}\n(ID: {caseId})", "Case Found", MessageBoxButton.OK, MessageBoxImage.Information);
                                CheckUploadEnabled();
                            }
                            else
                            {
                                _caseId = null;
                                UpdateStatus($"❌ Case ID {caseId} not found in database.");
                                MessageBox.Show($"Case ID {caseId} not found in database.", "Case Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Error verifying case: {ex.Message}");
                MessageBox.Show($"Error verifying case: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select the root folder containing files to upload",
                ShowNewFolderButton = false
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _selectedFolderPath = dialog.SelectedPath;
                txtFolderPath.Text = _selectedFolderPath;
                txtFolderPath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                UpdateStatus($"Folder selected: {_selectedFolderPath}");
                CheckUploadEnabled();
            }
        }

        private void CheckUploadEnabled()
        {
            bool canUpload = !string.IsNullOrEmpty(_selectedFolderPath) &&
                           _caseId.HasValue &&
                           !string.IsNullOrEmpty(txtBucketName.Text);

            btnUpload.IsEnabled = canUpload;
        }

        private void txtBucketName_TextChanged(object sender, TextChangedEventArgs e)
        {
            CheckUploadEnabled();
        }

        private void txtCaseId_TextChanged(object sender, TextChangedEventArgs e)
        {
            _caseId = null;
            CheckUploadEnabled();
        }

        private async void BtnUpload_Click(object sender, RoutedEventArgs e)
        {
            if (!_caseId.HasValue)
            {
                MessageBox.Show("Please verify a Case ID first.", "Case Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_selectedFolderPath) || !Directory.Exists(_selectedFolderPath))
            {
                MessageBox.Show("Please select a valid source folder.", "Folder Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtBucketName.Text))
            {
                MessageBox.Show("Please enter S3 Bucket Name.", "S3 Configuration Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                $"Are you sure you want to upload all files from:\n{_selectedFolderPath}\n\nTo Case ID: {_caseId.Value}?",
                "Confirm Upload",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                UpdateStatus("Upload cancelled by user.");
                return;
            }

            btnUpload.IsEnabled = false;
            UpdateStatus("Starting file upload process...");

            try
            {
                var s3Config = new S3Configuration
                {
                    BucketName = txtBucketName.Text,
                    ProfilePath = txtProfilePath.Text
                };

                var importService = new FilesCollectionImportService(
                    txtConnectionString.Text, 
                    s3Config, 
                    UpdateStatus);

                await Task.Run(() =>
                {
                    var (success, errors) = importService.UploadFilesFromFolder(_selectedFolderPath, _caseId.Value);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"📊 UPLOAD COMPLETE:");
                        UpdateStatus($"   ✅ Successfully uploaded: {success} files");
                        UpdateStatus($"   ❌ Errors: {errors} files");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("File upload completed!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Upload failed: {ex.Message}");
                MessageBox.Show($"Upload failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnUpload.IsEnabled = true;
            }
        }

        private async void BtnTruncate_Click(object sender, RoutedEventArgs e)
        {
            if (!_caseId.HasValue)
            {
                MessageBox.Show("Please verify a Case ID first.", "Case Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtBucketName.Text))
            {
                MessageBox.Show("Please enter S3 Bucket Name.", "S3 Configuration Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                "⚠️ WARNING: This will DELETE ALL UPLOADED FILES for Case ID: " + _caseId.Value + "\n\n" +
                "This will:\n" +
                "• Delete all files from S3 bucket\n" +
                "• Delete all records from database\n\n" +
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
                "FINAL CONFIRMATION: All uploaded files for Case ID " + _caseId.Value + " will be permanently deleted from S3 and database.\n\n" +
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
            UpdateStatus("Starting truncate process...");

            try
            {
                var s3Config = new S3Configuration
                {
                    BucketName = txtBucketName.Text,
                    ProfilePath = txtProfilePath.Text
                };

                var importService = new FilesCollectionImportService(
                    txtConnectionString.Text,
                    s3Config,
                    UpdateStatus);

                await Task.Run(() =>
                {
                    var (filesDeleted, errors) = importService.TruncateUploadedFiles(_caseId.Value);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"🗑️ TRUNCATE COMPLETE:");
                        UpdateStatus($"   ✅ Files deleted from S3: {filesDeleted}");
                        UpdateStatus($"   ❌ Errors: {errors}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("Truncate completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
