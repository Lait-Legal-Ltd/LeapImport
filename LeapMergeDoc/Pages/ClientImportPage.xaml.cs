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
    public partial class ClientImportPage : Page
    {
        private string? _clientFilePath;
        private List<ExcelRowData>? _excelData;
        private List<ProcessedClientData>? _processedData;

        public ClientImportPage()
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

        private void BtnSelectClientFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Client Excel File",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _clientFilePath = dialog.FileName;
                txtClientFilePath.Text = _clientFilePath;
                txtClientFilePath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                btnPreview.IsEnabled = true;
                UpdateStatus($"File selected: {Path.GetFileName(_clientFilePath)}");
            }
        }

        private async void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_clientFilePath))
            {
                MessageBox.Show("Please select a client Excel file.", "File Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnPreview.IsEnabled = false;
            btnImport.IsEnabled = false;
            UpdateStatus("Reading Excel file...");

            try
            {
                var importService = new ClientImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    _excelData = importService.ReadExcelData(_clientFilePath);
                    Dispatcher.Invoke(() => UpdateStatus($"Read {_excelData.Count} records from Excel."));

                    _processedData = importService.ProcessExcelData(_excelData);
                    Dispatcher.Invoke(() => UpdateStatus($"Processed {_processedData.Count} client records."));

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════════════════════");
                        UpdateStatus($"📊 CLIENT PREVIEW SUMMARY:");
                        UpdateStatus($"═══════════════════════════════════════════════════════");

                        // Basic counts
                        UpdateStatus($"📁 EXCEL DATA:");
                        UpdateStatus($"   Total Excel records: {_excelData.Count}");
                        UpdateStatus($"   Processed clients: {_processedData.Count}");

                        // Client type breakdown
                        int companies = _processedData.Count(p => p.ClientType == "Company");
                        int individuals = _processedData.Count(p => p.ClientType == "Individual");

                        UpdateStatus($"");
                        UpdateStatus($"👥 CLIENT TYPES:");
                        UpdateStatus($"   Individuals: {individuals}");
                        UpdateStatus($"   Companies: {companies}");

                        // Data quality
                        int withEmail = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.FirstEmailAddress));
                        int withDOB = _processedData.Count(p => p.OriginalData?.DateOfBirth.HasValue == true);
                        int withPhone = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.PrimaryContactNumber));
                        int withAddress = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.TownCity));
                        int withTitle = _processedData.Count(p => p.TitleId.HasValue);

                        UpdateStatus($"");
                        UpdateStatus($"📋 DATA QUALITY:");
                        UpdateStatus($"   With Title: {withTitle} / {_processedData.Count}");
                        UpdateStatus($"   With Email: {withEmail} / {_processedData.Count}");
                        UpdateStatus($"   With DOB: {withDOB} / {_processedData.Count}");
                        UpdateStatus($"   With Phone: {withPhone} / {_processedData.Count}");
                        UpdateStatus($"   With Town/City: {withAddress} / {_processedData.Count}");

                        // Title breakdown
                        var titleGroups = _processedData
                            .Where(p => p.TitleId.HasValue)
                            .GroupBy(p => p.TitleId)
                            .Select(g => new { TitleId = g.Key, Count = g.Count() })
                            .OrderByDescending(x => x.Count)
                            .ToList();

                        if (titleGroups.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"📝 TITLES:");
                            foreach (var t in titleGroups)
                            {
                                var titleName = GetTitleName(t.TitleId);
                                UpdateStatus($"   {titleName}: {t.Count}");
                            }
                        }

                        // Sample of first 5 clients
                        UpdateStatus($"");
                        UpdateStatus($"📄 SAMPLE CLIENTS (first 5):");
                        foreach (var client in _processedData.Take(5))
                        {
                            var name = client.OriginalData?.ClientName ?? "Unknown";
                            var type = client.ClientType;
                            UpdateStatus($"   • {name} ({type})");
                        }

                        if (_processedData.Count > 5)
                        {
                            UpdateStatus($"   ... and {_processedData.Count - 5} more");
                        }

                        UpdateStatus($"═══════════════════════════════════════════════════════");
                    });
                });

                btnImport.IsEnabled = true;
                UpdateStatus("✅ Preview complete. Ready to import.");
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

            var result = MessageBox.Show(
                $"Are you sure you want to import {_processedData.Count} clients to the database?",
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
            UpdateStatus("Starting client import...");

            try
            {
                var importService = new ClientImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var (success, errors) = importService.ImportToDatabase(_processedData);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"📊 IMPORT COMPLETE:");
                        UpdateStatus($"   ✅ Successfully imported: {success}");
                        UpdateStatus($"   ❌ Errors: {errors}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("Client import completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
            }
        }

        private async void BtnTruncate_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "⚠️ WARNING: This will DELETE ALL DATA from the following tables:\n\n" +
                "• tbl_client_individual\n" +
                "• tbl_client_company\n" +
                "• tbl_client\n\n" +
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
                "FINAL CONFIRMATION: All client data will be permanently deleted.\n\n" +
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
            UpdateStatus("Starting client data truncation...");

            try
            {
                var importService = new ClientImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var (rowsDeleted, message) = importService.TruncateClientData();

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"🗑️ TRUNCATE COMPLETE:");
                        UpdateStatus($"   Total rows deleted: {rowsDeleted}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("Client data truncated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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

        private string GetTitleName(int? titleId)
        {
            if (!titleId.HasValue) return "Unknown";

            var titleNames = new Dictionary<int, string>
            {
                { 1, "Mr" },
                { 2, "Mrs" },
                { 3, "Miss" },
                { 4, "Dr" },
                { 5, "Prof" },
                { 6, "Ms" },
                { 7, "Master" },
                { 8, "Mx" },
                { 9, "Rev" }
            };

            return titleNames.TryGetValue(titleId.Value, out var name) ? name : $"Title {titleId}";
        }
    }
}
