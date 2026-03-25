using LeapMergeDoc.Models;
using LeapMergeDoc.Services;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class AccountUpdateImportPage : Page
    {
        private string? _filePath;
        private List<AccountUpdateExcelData>? _excelData;
        private List<AccountUpdateImportData>? _processedData;
        private AccountUpdateImportSummary? _summary;
        
        // Hardcoded connection string
        private const string DefaultConnectionString = "Server=localhost;Database=bais-live;User=root;Password=1qaz!QAZ";

        public AccountUpdateImportPage()
        {
            InitializeComponent();
            
            // Set default connection string
            txtConnectionString.Text = DefaultConnectionString;
            
            // Auto-load bank accounts on page load
            Loaded += AccountUpdateImportPage_Loaded;
        }
        
        private void AccountUpdateImportPage_Loaded(object sender, RoutedEventArgs e)
        {
            // Auto-load banks when page loads
            LoadBankAccounts();
        }
        
        private void LoadBankAccounts()
        {
            try
            {
                var importService = new AccountUpdateImportService(txtConnectionString.Text, UpdateStatus);

                // Load client bank accounts
                var clientBanks = importService.GetClientBankAccounts();
                
                // If no banks found from DB, add hardcoded default
                if (clientBanks.Count == 0)
                {
                    clientBanks.Add(new BankAccountInfo
                    {
                        BankId = 1,
                        BankName = "BAIS - Client",
                        AccountNumber = "93336948",
                        SortCode = "20-92-63",
                        Institution = "Barclays",
                        IsClientBank = true,
                        OpeningBalance = 0
                    });
                    UpdateStatus("⚠️ No client banks in DB, using hardcoded default (ID=1)");
                }
                
                cmbClientBank.ItemsSource = clientBanks;
                if (clientBanks.Count > 0)
                {
                    cmbClientBank.SelectedIndex = 0;
                }

                // Load office bank accounts
                var officeBanks = importService.GetOfficeBankAccounts();
                cmbOfficeBank.ItemsSource = officeBanks;
                if (officeBanks.Count > 0)
                {
                    cmbOfficeBank.SelectedIndex = 0;
                }

                UpdateStatus($"✅ Loaded {clientBanks.Count} client banks and {officeBanks.Count} office banks");
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Error loading banks: {ex.Message}");
                
                // Fallback: Add hardcoded client bank
                var fallbackBanks = new List<BankAccountInfo>
                {
                    new BankAccountInfo
                    {
                        BankId = 1,
                        BankName = "BAIS - Client",
                        AccountNumber = "93336948",
                        SortCode = "20-92-63",
                        Institution = "Barclays",
                        IsClientBank = true,
                        OpeningBalance = 0
                    }
                };
                cmbClientBank.ItemsSource = fallbackBanks;
                cmbClientBank.SelectedIndex = 0;
                UpdateStatus("⚠️ Using hardcoded fallback client bank (ID=1)");
            }
        }

        private void UpdateStatus(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                txtStatus.Text = $"[{DateTime.Now:HH:mm:ss}] {message}\n{txtStatus.Text}";
                logScrollViewer.ScrollToTop();
            }
            else
            {
                Dispatcher.Invoke(() => 
                {
                    txtStatus.Text = $"[{DateTime.Now:HH:mm:ss}] {message}\n{txtStatus.Text}";
                    logScrollViewer.ScrollToTop();
                });
            }
        }
        
        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtStatus.Text = "Log cleared. Ready...";
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

        private void BtnLoadBanks_Click(object sender, RoutedEventArgs e)
        {
            LoadBankAccounts();
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Account Transactions Excel File",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
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

            if (cmbClientBank.SelectedValue == null)
            {
                MessageBox.Show("Please select a default client bank account.", "Bank Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnPreview.IsEnabled = false;
            btnImport.IsEnabled = false;
            UpdateStatus("Reading Excel file...");

            try
            {
                int defaultClientBankId = (int)cmbClientBank.SelectedValue;
                int defaultOfficeBankId = cmbOfficeBank.SelectedValue != null ? (int)cmbOfficeBank.SelectedValue : 0;

                var importService = new AccountUpdateImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    _excelData = importService.ReadExcelData(_filePath);
                    Dispatcher.Invoke(() => UpdateStatus($"Read {_excelData.Count} records from Excel."));

                    var result = importService.ProcessExcelData(_excelData, defaultClientBankId, defaultOfficeBankId);
                    _processedData = result.records;
                    _summary = result.summary;

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════════════════════");
                        UpdateStatus($"📊 PREVIEW SUMMARY:");
                        UpdateStatus($"═══════════════════════════════════════════════════════");

                        UpdateStatus($"📁 EXCEL DATA:");
                        UpdateStatus($"   Total records: {_summary.TotalRecords}");
                        UpdateStatus($"   Valid records: {_summary.ValidRecords}");
                        UpdateStatus($"   Invalid records: {_summary.InvalidRecords}");

                        UpdateStatus($"");
                        UpdateStatus($"💰 TRANSACTION BREAKDOWN:");
                        UpdateStatus($"   📥 Bank Receipts: {_summary.ReceiptCount} ({_summary.ReceiptTotal:C})");
                        UpdateStatus($"   📤 Bank Payments: {_summary.PaymentCount} ({_summary.PaymentTotal:C})");
                        UpdateStatus($"   🔄 Client to Office: {_summary.ClientToOfficeCount} ({_summary.ClientToOfficeTotal:C})");

                        UpdateStatus($"");
                        UpdateStatus($"👥 CASE MATCHING:");
                        UpdateStatus($"   ✅ Cases found: {_summary.FoundCases}");
                        UpdateStatus($"   ⬜ Cases NOT found: {_summary.NotFoundCases}");

                        // Show not found cases
                        var notFoundCases = _processedData?
                            .Where(p => !p.IsFound)
                            .Take(10)
                            .ToList();

                        if (notFoundCases != null && notFoundCases.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"⚠️ NOT FOUND CASES (sample):");
                            foreach (var item in notFoundCases)
                            {
                                UpdateStatus($"   • Row {item.RowNumber}: {item.CaseReference} ({item.Amount:C})");
                            }
                            if (_summary.NotFoundCases > 10)
                            {
                                UpdateStatus($"   ... and {_summary.NotFoundCases - 10} more");
                            }
                        }

                        // Show validation errors
                        if (_summary.Errors.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"❌ VALIDATION ERRORS:");
                            foreach (var error in _summary.Errors.Take(10))
                            {
                                UpdateStatus($"   • {error}");
                            }
                            if (_summary.Errors.Count > 10)
                            {
                                UpdateStatus($"   ... and {_summary.Errors.Count - 10} more errors");
                            }
                        }

                        UpdateStatus($"");
                        UpdateStatus($"═══════════════════════════════════════════════════════");
                    });
                });

                btnImport.IsEnabled = _summary?.ValidRecords > 0 && _summary?.FoundCases > 0;
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

            var validCount = _processedData.Count(p => p.IsValid && p.IsFound);
            var totalAmount = _processedData.Where(p => p.IsValid && p.IsFound).Sum(p => p.Amount);

            var receipts = _processedData.Count(p => p.IsValid && p.IsFound && p.TransactionType == AccountTransactionType.BankReceipt);
            var payments = _processedData.Count(p => p.IsValid && p.IsFound && p.TransactionType == AccountTransactionType.BankPayment);
            var c2o = _processedData.Count(p => p.IsValid && p.IsFound && p.TransactionType == AccountTransactionType.ClientToOffice);

            var result = MessageBox.Show(
                $"Are you sure you want to import {validCount} transactions?\n\n" +
                $"Bank Receipts: {receipts}\n" +
                $"Bank Payments: {payments}\n" +
                $"Client to Office: {c2o}\n\n" +
                $"Total Amount: {totalAmount:C}",
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
            UpdateStatus("Starting account update import...");

            try
            {
                var importService = new AccountUpdateImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var importResult = importService.ImportAccountUpdates(_processedData);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════════════════════");
                        UpdateStatus($"📊 IMPORT COMPLETE:");
                        UpdateStatus($"   ✅ Successfully imported: {importResult.SuccessCount}");
                        UpdateStatus($"   📥 Receipts created: {importResult.ReceiptsCreated}");
                        UpdateStatus($"   📤 Payments created: {importResult.PaymentsCreated}");
                        UpdateStatus($"   📄 Invoices created: {importResult.InvoicesCreated}");
                        UpdateStatus($"   🔄 C2O transfers created: {importResult.ClientToOfficeCreated}");
                        UpdateStatus($"   ❌ Errors: {importResult.ErrorCount}");

                        if (importResult.Errors.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"❌ ERROR DETAILS:");
                            foreach (var error in importResult.Errors.Take(20))
                            {
                                UpdateStatus($"   • {error}");
                            }
                        }

                        UpdateStatus("═══════════════════════════════════════════════════════");
                    });
                });

                MessageBox.Show("Import completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Import error: {ex.Message}");
                MessageBox.Show($"Error during import: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnPreview.IsEnabled = true;
                btnImport.IsEnabled = false;
                
                // Clear data after import
                _excelData = null;
                _processedData = null;
                _summary = null;
            }
        }
    }
}
