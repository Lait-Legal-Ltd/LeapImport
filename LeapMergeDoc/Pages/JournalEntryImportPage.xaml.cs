using LeapMergeDoc.Models;
using LeapMergeDoc.Services;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class JournalEntryImportPage : Page
    {
        private string? _filePath;
        private List<JournalExcelData>? _excelData;
        private List<JournalImportData>? _processedData;
        private JournalImportSummary? _summary;

        public JournalEntryImportPage()
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

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Trial Balance File",
                Filter = "CSV Files (*.csv)|*.csv|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _filePath = dialog.FileName;
                txtFilePath.Text = _filePath;
                txtFilePath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                btnPreview.IsEnabled = true;
                UpdateStatus($"File selected: {Path.GetFileName(_filePath)}");
            }
        }

        private async void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                MessageBox.Show("Please select an Excel file.", "File Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnPreview.IsEnabled = false;
            btnImport.IsEnabled = false;
            UpdateStatus("Reading Excel file...");

            try
            {
                var importService = new JournalEntryImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    _excelData = importService.ReadExcelData(_filePath);
                    Dispatcher.Invoke(() => UpdateStatus($"Read {_excelData.Count} records from Excel."));

                    var result = importService.ProcessExcelData(_excelData);
                    _processedData = result.records;
                    _summary = result.summary;

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════════════════════");
                        UpdateStatus($"📊 PREVIEW SUMMARY:");
                        UpdateStatus($"═══════════════════════════════════════════════════════");

                        UpdateStatus($"📁 EXCEL DATA:");
                        UpdateStatus($"   Total records: {_summary.TotalRecords}");
                        UpdateStatus($"   Total amount: {_summary.TotalAmount:C}");

                        UpdateStatus($"");
                        UpdateStatus($"👥 CASE MATCHING:");
                        UpdateStatus($"   ✅ Cases found: {_summary.FoundCases} ({_summary.FoundAmount:C})");
                        UpdateStatus($"   ⬜ Cases NOT found: {_summary.NotFoundCases} ({_summary.NotFoundAmount:C})");
                        
                        if (_summary.MissingLedgerCards > 0)
                        {
                            UpdateStatus($"   ⚠️ MISSING LEDGER CARDS: {_summary.MissingLedgerCards}");
                        }

                        // Show found cases with ledger card info (first 10)
                        var foundCases = _processedData
                            .Where(p => p.IsFound)
                            .Take(10)
                            .ToList();

                        if (foundCases.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"✅ FOUND CASES (sample with LedgerCardId):");
                            foreach (var item in foundCases)
                            {
                                var ledgerInfo = item.HasLedgerCard 
                                    ? $"LedgerCardId: {item.LedgerCardId}" 
                                    : "⚠️ NO LEDGER CARD!";
                                UpdateStatus($"   • {item.CaseReference} - {item.Balance:C} [{ledgerInfo}]");
                            }
                            if (_summary.FoundCases > 10)
                            {
                                UpdateStatus($"   ... and {_summary.FoundCases - 10} more");
                            }
                        }

                        // Show not found cases (first 10)
                        var notFoundCases = _processedData
                            .Where(p => !p.IsFound)
                            .Take(10)
                            .ToList();

                        if (notFoundCases.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"⚠️ NOT FOUND CASES (sample):");
                            foreach (var item in notFoundCases)
                            {
                                UpdateStatus($"   • {item.CaseReference} - {item.ClientName} ({item.Balance:C})");
                            }
                            if (_summary.NotFoundCases > 10)
                            {
                                UpdateStatus($"   ... and {_summary.NotFoundCases - 10} more");
                            }
                        }

                        UpdateStatus($"");
                        UpdateStatus($"📅 IMPORT DATE: 01/04/2026");
                        UpdateStatus($"═══════════════════════════════════════════════════════");
                    });
                });

                // Only enable import/export if there are found cases WITH ledger cards and NO missing ledger cards
                var hasValidCases = _summary?.FoundCases > 0 && _summary?.MissingLedgerCards == 0;
                btnImport.IsEnabled = hasValidCases;
                btnExportSql.IsEnabled = hasValidCases;
                
                if (_summary?.MissingLedgerCards > 0)
                {
                    UpdateStatus($"⚠️ Cannot import: {_summary.MissingLedgerCards} cases are missing ledger cards!");
                    UpdateStatus("Please create ledger cards for these cases first.");
                }
                else
                {
                    UpdateStatus("✅ Preview complete. Ready to import or export SQL.");
                }
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

            var foundCount = _processedData.Count(p => p.IsFound);
            var totalAmount = _processedData.Where(p => p.IsFound).Sum(p => p.Balance);

            var result = MessageBox.Show(
                $"Are you sure you want to import {foundCount} journal entries?\n\n" +
                $"Total Amount: {totalAmount:C}\n" +
                $"Transaction Date: 01/04/2026",
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
            UpdateStatus("Starting journal entry import...");

            try
            {
                var importService = new JournalEntryImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var (success, errors) = importService.ImportJournalEntries(_processedData);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"📊 IMPORT COMPLETE:");
                        UpdateStatus($"   ✅ Successfully imported: {success}");
                        UpdateStatus($"   ❌ Errors: {errors}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("Import completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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

        private void BtnExportSql_Click(object sender, RoutedEventArgs e)
        {
            if (_processedData == null || _processedData.Count == 0)
            {
                MessageBox.Show("Please preview the data first.", "Preview Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var foundCount = _processedData.Count(p => p.IsFound);
            if (foundCount == 0)
            {
                MessageBox.Show("No matching cases found to export.", "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save SQL File",
                Filter = "SQL Files (*.sql)|*.sql|All Files (*.*)|*.*",
                FilterIndex = 1,
                FileName = $"journal_entry_import_{DateTime.Now:yyyyMMdd_HHmmss}.sql"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    var importService = new JournalEntryImportService(txtConnectionString.Text, UpdateStatus);
                    var sql = importService.ExportAsSql(_processedData);
                    
                    File.WriteAllText(saveDialog.FileName, sql);

                    UpdateStatus($"✅ SQL exported to: {saveDialog.FileName}");
                    UpdateStatus($"   Records: {foundCount}");
                    MessageBox.Show($"SQL file exported successfully!\n\n{saveDialog.FileName}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    UpdateStatus($"❌ Export failed: {ex.Message}");
                    MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void BtnTruncate_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "⚠️ WARNING: This will DELETE ALL DATA from the following tables:\n\n" +
                "• tbl_acc_ledger_card_transactions\n" +
                "• tbl_acc_client_bank_transactions\n" +
                "• tbl_acc_journal_entry_lines\n" +
                "• tbl_acc_journal_entry\n" +
                "• tbl_acc_transactions\n\n" +
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

            // Double confirmation
            var result2 = MessageBox.Show(
                "FINAL CONFIRMATION: All journal data will be permanently deleted.\n\n" +
                "Type 'YES' in your mind and click Yes to confirm.",
                "Final Confirmation",
                MessageBoxButton.YesNo,
                MessageBoxImage.Stop);

            if (result2 != MessageBoxResult.Yes)
            {
                UpdateStatus("Truncate cancelled by user.");
                return;
            }

            btnTruncate.IsEnabled = false;
            UpdateStatus("Starting data truncation...");

            try
            {
                var importService = new JournalEntryImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var (rowsDeleted, message) = importService.TruncateJournalData();

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"🗑️ TRUNCATE COMPLETE:");
                        UpdateStatus($"   Total rows deleted: {rowsDeleted}");
                        UpdateStatus("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show("Data truncated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
