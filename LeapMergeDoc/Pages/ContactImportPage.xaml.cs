using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using LeapMergeDoc.Models;
using LeapMergeDoc.Services;

namespace LeapMergeDoc.Pages
{
    public partial class ContactImportPage : Page
    {
        private List<ContactExcelData>? _excelData;
        private List<ProcessedContactData>? _processedData;

        public ContactImportPage()
        {
            InitializeComponent();
            LogStatus("Ready. Select a contacts Excel file to begin.");
        }

        private void LogStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n";
                txtStatus.ScrollToEnd();
            });
        }

        private void TestConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var connection = new MySqlConnection(txtConnectionString.Text))
                {
                    connection.Open();
                    LogStatus("✅ Database connection successful!");
                }
            }
            catch (Exception ex)
            {
                LogStatus($"❌ Connection failed: {ex.Message}");
            }
        }

        private void BrowseContactFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select Contacts Excel File"
            };

            if (dialog.ShowDialog() == true)
            {
                txtContactFilePath.Text = dialog.FileName;
                btnPreview.IsEnabled = true;
                LogStatus($"Selected file: {dialog.FileName}");
            }
        }

        private void PreviewData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var service = new ContactImportService(txtConnectionString.Text, LogStatus);

                LogStatus("Reading Excel file...");
                _excelData = service.ReadExcelData(txtContactFilePath.Text);
                LogStatus($"Read {_excelData.Count} records from Excel.");

                LogStatus("Processing data and checking against existing clients...");
                _processedData = service.ProcessExcelData(_excelData);

                // Generate summary
                int companies = 0, individuals = 0, skipped = 0;
                foreach (var p in _processedData)
                {
                    if (p.IsExistingClient) skipped++;
                    else if (p.IsCompany) companies++;
                    else individuals++;
                }

                LogStatus("");
                LogStatus("=== PREVIEW SUMMARY ===");
                LogStatus($"Total records: {_processedData.Count}");
                LogStatus($"Companies to import: {companies}");
                LogStatus($"Individuals to import: {individuals}");
                LogStatus($"Skipped (existing clients): {skipped}");
                LogStatus("");

                // Show sample of what will be imported
                LogStatus("--- Sample contacts to import ---");
                int sampleCount = 0;
                foreach (var p in _processedData)
                {
                    if (!p.IsExistingClient && sampleCount < 10)
                    {
                        var name = p.IsCompany ? p.CompanyName : $"{p.GivenNames} {p.LastName}";
                        var type = p.IsCompany ? "Company" : "Personal";
                        LogStatus($"  [{type}] {name}");
                        sampleCount++;
                    }
                }

                // Show sample of what will be skipped
                LogStatus("");
                LogStatus("--- Sample clients being skipped ---");
                sampleCount = 0;
                foreach (var p in _processedData)
                {
                    if (p.IsExistingClient && sampleCount < 5)
                    {
                        var name = p.IsCompany ? p.CompanyName : $"{p.GivenNames} {p.LastName}";
                        LogStatus($"  [SKIP] {name} (exists in client table)");
                        sampleCount++;
                    }
                }

                btnImport.IsEnabled = true;
                LogStatus("");
                LogStatus("✅ Preview complete. Ready to import.");
            }
            catch (Exception ex)
            {
                LogStatus($"❌ Error during preview: {ex.Message}");
            }
        }

        private void ImportData_Click(object sender, RoutedEventArgs e)
        {
            if (_processedData == null)
            {
                LogStatus("❌ No data to import. Please preview first.");
                return;
            }

            var result = MessageBox.Show(
                "This will import contacts to the database.\n\nExisting clients will be skipped.\n\nContinue?",
                "Confirm Import",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes) return;

            try
            {
                var service = new ContactImportService(txtConnectionString.Text, LogStatus);

                LogStatus("");
                LogStatus("Starting import...");

                var (success, skipped, errors) = service.ImportContactsToDatabase(_processedData);

                LogStatus("");
                LogStatus("=== IMPORT RESULTS ===");
                LogStatus($"Successfully imported: {success}");
                LogStatus($"Skipped (existing clients): {skipped}");
                LogStatus($"Errors: {errors}");
            }
            catch (Exception ex)
            {
                LogStatus($"❌ Import failed: {ex.Message}");
            }
        }

        private void TruncateData_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "⚠️ WARNING: This will DELETE ALL contact data!\n\n" +
                "Tables to be cleared:\n" +
                "- tbl_contact\n" +
                "- tbl_contact_company\n" +
                "- tbl_contact_personal\n\n" +
                "This action CANNOT be undone.\n\nAre you sure?",
                "Confirm Delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes) return;

            var confirm = MessageBox.Show(
                "FINAL CONFIRMATION\n\nType 'YES' in your mind and click Yes to proceed.",
                "Final Confirmation",
                MessageBoxButton.YesNo,
                MessageBoxImage.Stop);

            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                var service = new ContactImportService(txtConnectionString.Text, LogStatus);
                service.TruncateContactData();

                btnImport.IsEnabled = false;
                btnPreview.IsEnabled = false;
            }
            catch (Exception ex)
            {
                LogStatus($"❌ Truncate failed: {ex.Message}");
            }
        }
    }
}
