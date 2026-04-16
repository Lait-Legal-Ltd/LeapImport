using LeapMergeDoc.Models;
using LeapMergeDoc.Services;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class DatabaseImportPage : Page
    {
        private string? _mergedFilePath;
        private string? _clientMasterFilePath;
        private List<CaseExcelData>? _excelData;
        private List<ProcessedCaseData>? _processedData;

        public DatabaseImportPage()
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

        private void BtnSelectMergedFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Case Data File",
                Filter = "Excel/CSV Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _mergedFilePath = dialog.FileName;
                txtMergedFilePath.Text = _mergedFilePath;
                txtMergedFilePath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                btnPreview.IsEnabled = true;
                var ext = Path.GetExtension(_mergedFilePath).ToLower();
                var fileType = ext == ".csv" ? "CSV" : "Excel";
                UpdateStatus($"{fileType} file selected: {Path.GetFileName(_mergedFilePath)}");
            }
        }

        private void BtnSelectClientMasterFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Client Master CSV File",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                _clientMasterFilePath = dialog.FileName;
                txtClientMasterPath.Text = _clientMasterFilePath;
                txtClientMasterPath.Foreground = System.Windows.Media.Brushes.DarkGreen;
                UpdateStatus($"Client master CSV selected: {Path.GetFileName(_clientMasterFilePath)}");
            }
        }

        private async void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_mergedFilePath))
            {
                MessageBox.Show("Please select a merged Excel file.", "File Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            btnPreview.IsEnabled = false;
            btnImport.IsEnabled = false;
            UpdateStatus("Reading data file...");

            try
            {
                var importService = new DatabaseImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    // Load client master CSV if provided
                    if (!string.IsNullOrEmpty(_clientMasterFilePath))
                    {
                        Dispatcher.Invoke(() => UpdateStatus("Loading client master CSV..."));
                        int clientMasterCount = importService.LoadClientMasterCsv(_clientMasterFilePath);
                        Dispatcher.Invoke(() => UpdateStatus($"Loaded {clientMasterCount} client master records."));
                    }

                    _excelData = importService.ReadExcelData(_mergedFilePath);
                    Dispatcher.Invoke(() => UpdateStatus($"Read {_excelData.Count} records from file."));

                    _processedData = importService.ProcessExcelData(_excelData);
                    Dispatcher.Invoke(() => UpdateStatus($"Processed {_processedData.Count} case records."));

                    // Check client mappings (will use client master lookup if loaded)
                    Dispatcher.Invoke(() => UpdateStatus("Checking client mappings..."));
                    var (found, notFound) = importService.CheckClientMappings(_processedData);

                    // Check for existing cases in database
                    Dispatcher.Invoke(() => UpdateStatus("Checking for existing cases in database..."));
                    var (existingCount, newCount, existingRefs) = importService.CheckExistingCases(_processedData);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════════════════════");
                        UpdateStatus($"📊 PREVIEW SUMMARY:");
                        UpdateStatus($"═══════════════════════════════════════════════════════");

                        // Basic counts
                        UpdateStatus($"📁 EXCEL DATA:");
                        UpdateStatus($"   Total Excel records: {_excelData.Count}");
                        UpdateStatus($"   Processed cases: {_processedData.Count}");

                        // Existing cases check
                        UpdateStatus($"");
                        UpdateStatus($"🔍 EXISTING CASES CHECK:");
                        UpdateStatus($"   ✅ New cases to import: {newCount}");
                        UpdateStatus($"   ⏭️ Already exist (will skip): {existingCount}");
                        if (existingCount > 0 && existingRefs.Count <= 10)
                        {
                            UpdateStatus($"   Existing case refs: {string.Join(", ", existingRefs.Take(10))}");
                        }
                        else if (existingCount > 10)
                        {
                            UpdateStatus($"   Sample existing: {string.Join(", ", existingRefs.Take(10))}...");
                        }

                        // Data quality checks
                        int withMatterNo = _processedData.Count(p => !string.IsNullOrEmpty(p.CaseReferenceAuto));
                        int withCaseName = _processedData.Count(p => !string.IsNullOrEmpty(p.CaseName));
                        int withDateOpened = _processedData.Count(p => p.DateOpened.HasValue);
                        int withMatterType = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.MatterType));
                        int withMatterDesc = _processedData.Count(p => !string.IsNullOrEmpty(p.OriginalData?.MatterDescription));
                        int withArchiveDate = _processedData.Count(p => p.OriginalData?.ArchiveDate.HasValue == true);

                        UpdateStatus($"");
                        UpdateStatus($"📋 DATA QUALITY:");
                        UpdateStatus($"   With Matter No: {withMatterNo} / {_processedData.Count}");
                        UpdateStatus($"   With Case Name/Client: {withCaseName} / {_processedData.Count}");
                        UpdateStatus($"   With Date Opened: {withDateOpened} / {_processedData.Count}");
                        UpdateStatus($"   With Matter Type: {withMatterType} / {_processedData.Count}");
                        UpdateStatus($"   With Matter Description: {withMatterDesc} / {_processedData.Count}");
                        UpdateStatus($"   With Archive Date: {withArchiveDate} / {_processedData.Count}");

                        // Client mapping summary
                        UpdateStatus($"");
                        UpdateStatus($"👥 CLIENT MATCHING:");
                        UpdateStatus($"   ✅ Clients found in database: {found}");
                        UpdateStatus($"   ⬜ Clients NOT found: {notFound}");

                        // Show ALL unmatched clients (not just sample)
                        var unmatchedClients = _processedData
                            .Where(p => !p.LinkedClientId.HasValue)
                            .Select(p => new { 
                                ClientNo = p.OriginalData?.ClientNo ?? "",
                                ClientName = p.OriginalData?.CaseName ?? p.OriginalData?.Surname ?? p.CaseName ?? "",
                                CaseRef = p.CaseReferenceAuto ?? ""
                            })
                            .Where(x => !string.IsNullOrEmpty(x.ClientName))
                            .GroupBy(x => x.ClientNo)  // Group by client code to avoid duplicates
                            .Select(g => g.First())
                            .ToList();  // Get ALL unmatched

                        if (unmatchedClients.Any())
                        {
                            UpdateStatus($"   ⚠️ ALL {unmatchedClients.Count} UNMATCHED CLIENTS:");
                            foreach (var client in unmatchedClients)
                            {
                                var display = !string.IsNullOrEmpty(client.ClientNo) 
                                    ? $"{client.ClientNo}: {client.ClientName}"
                                    : client.ClientName;
                                UpdateStatus($"      • {display}");
                            }
                            
                            // Generate SQL INSERT statements for missing clients
                            UpdateStatus($"");
                            UpdateStatus($"📝 SQL TO CREATE MISSING CLIENTS:");
                            UpdateStatus($"-- Run this SQL to create missing clients as Company type:");
                            foreach (var client in unmatchedClients)
                            {
                                var escapedName = client.ClientName.Replace("'", "''");
                                UpdateStatus($"-- {client.ClientNo}: {client.ClientName}");
                                UpdateStatus($"INSERT INTO tbl_client (client_type, fk_branch_id, is_archived, is_active, date_time_created, fk_user_id)");
                                UpdateStatus($"VALUES ('Company', 1, 0, 1, NOW(), 1);");
                                UpdateStatus($"SET @last_client_id = LAST_INSERT_ID();");
                                UpdateStatus($"INSERT INTO tbl_client_company (fk_client_id, company_name) VALUES (@last_client_id, '{escapedName}');");
                                UpdateStatus($"");
                            }
                        }

                        // Matter Type mapping - ALL
                        UpdateStatus($"");
                        UpdateStatus($"📂 MATTER TYPE MAPPING (ALL):");
                        var matterTypesByArea = _processedData
                            .GroupBy(p => p.FkAreaOfPracticeId)
                            .Select(g => new { AreaId = g.Key, Count = g.Count() })
                            .OrderByDescending(x => x.Count)
                            .ToList();

                        foreach (var mt in matterTypesByArea)
                        {
                            var areaName = GetAreaName(mt.AreaId);
                            UpdateStatus($"   {areaName}: {mt.Count} cases");
                        }

                        // Original Matter Types (from Excel) - show unique values
                        UpdateStatus($"");
                        UpdateStatus($"📝 ORIGINAL MATTER TYPES (from Excel):");
                        var originalMatterTypes = _processedData
                            .Where(p => !string.IsNullOrEmpty(p.OriginalData?.MatterType))
                            .GroupBy(p => p.OriginalData!.MatterType)
                            .Select(g => new { MatterType = g.Key, Count = g.Count() })
                            .OrderByDescending(x => x.Count)
                            .ToList();

                        foreach (var mt in originalMatterTypes)
                        {
                            UpdateStatus($"   {mt.MatterType}: {mt.Count}");
                        }

                        // Work Type (W/T) codes analysis
                        var workTypes = _processedData
                            .Where(p => !string.IsNullOrEmpty(p.OriginalData?.WorkId))
                            .GroupBy(p => p.OriginalData!.WorkId)
                            .Select(g => new { 
                                WorkId = g.Key, 
                                Count = g.Count(),
                                HasCaseType = g.Any(x => x.FkCaseTypeId.HasValue)
                            })
                            .OrderByDescending(x => x.Count)
                            .ToList();

                        if (workTypes.Any())
                        {
                            UpdateStatus($"");
                            UpdateStatus($"🏷️ WORK TYPE (W/T) CODES:");
                            foreach (var wt in workTypes)
                            {
                                var status = wt.HasCaseType ? "✅" : "⚠️";
                                UpdateStatus($"   {status} {wt.WorkId}: {wt.Count} cases");
                            }
                        }

                        // Case status
                        int activeCases = _processedData.Count(p => p.IsCaseActive);
                        int archivedCases = _processedData.Count(p => p.IsCaseArchived);
                        int withDateArchived = _processedData.Count(p => p.DateArchived.HasValue);

                        UpdateStatus($"");
                        UpdateStatus($"📊 CASE STATUS:");
                        UpdateStatus($"   Active cases: {activeCases}");
                        UpdateStatus($"   Archived cases: {archivedCases}");
                        UpdateStatus($"   With Archive Date: {withDateArchived}");

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
                $"Are you sure you want to import {_processedData.Count} cases to the database?",
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
            UpdateStatus("Starting database import...");

            try
            {
                var importService = new DatabaseImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    // Load client master CSV if provided
                    if (!string.IsNullOrEmpty(_clientMasterFilePath))
                    {
                        Dispatcher.Invoke(() => UpdateStatus("Loading client master CSV..."));
                        importService.LoadClientMasterCsv(_clientMasterFilePath);
                    }

                    var (success, errors, skipped) = importService.ImportCasesToDatabase(_processedData);

                    Dispatcher.Invoke(() =>
                    {
                        UpdateStatus("═══════════════════════════════════════");
                        UpdateStatus($"📊 IMPORT COMPLETE:");
                        UpdateStatus($"   ✅ Successfully imported: {success}");
                        UpdateStatus($"   ⏭️ Skipped (already exist): {skipped}");
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

        private string GetAreaName(int? areaId)
        {
            if (!areaId.HasValue) return "Unknown";

            var areaNames = new Dictionary<int, string>
            {
                { 2, "Civil litigation" },
                { 3, "Company commercial" },
                { 6, "Criminal justice" },
                { 8, "Employment" },
                { 9, "Family and children" },
                { 11, "Immigration" },
                { 13, "Intellectual property" },
                { 14, "Legal aid" },
                { 18, "Commercial Conveyancing" },
                { 19, "Residential Conveyancing" },
                { 22, "Wills and Probate" },
                { 23, "Miscellaneous" }
            };

            return areaNames.TryGetValue(areaId.Value, out var name) ? name : $"Area {areaId}";
        }

        private async void BtnTruncate_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "⚠️ WARNING: This will DELETE ALL DATA from the following tables:\n\n" +
                "• tbl_case_client_greeting\n" +
                "• tbl_case_clients\n" +
                "• tbl_case_permissions\n" +
                "• tbl_acc_ledger_cards\n" +
                "• tbl_case_details_general\n\n" +
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
                "FINAL CONFIRMATION: All case data will be permanently deleted.\n\n" +
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
                var importService = new DatabaseImportService(txtConnectionString.Text, UpdateStatus);

                await Task.Run(() =>
                {
                    var (rowsDeleted, message) = importService.TruncateImportedData();

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
