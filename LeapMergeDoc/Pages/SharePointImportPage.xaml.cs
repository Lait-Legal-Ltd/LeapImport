using LeapMergeDoc.Models;
using LeapMergeDoc.Services;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace LeapMergeDoc.Pages
{
    public partial class SharePointImportPage : Page
    {
        private SharePointImportService? _sharePointService;
        private string _downloadedFolderPath = string.Empty;
        private List<SharePointSite> _sites = new List<SharePointSite>();
        private List<SharePointDrive> _drives = new List<SharePointDrive>();
        private bool _isAuthenticated = false;

        public SharePointImportPage()
        {
            InitializeComponent();
            InitializeServices();
            
            // Load saved config on page load
            Loaded += SharePointImportPage_Loaded;
        }

        private async void SharePointImportPage_Loaded(object sender, RoutedEventArgs e)
        {
            // First, load saved credentials to populate UI
            if (_sharePointService!.HasSavedConfig())
            {
                var config = _sharePointService.LoadConfig();
                if (config != null)
                {
                    // Populate UI fields
                    txtTenantId.Text = config.TenantId;
                    txtClientId.Text = config.ClientId;
                    txtClientSecret.Password = config.ClientSecret;
                }
            }

            // Try to use cached token first (no API call needed)
            if (_sharePointService.IsAuthenticated())
            {
                Log("✅ Using cached token - already authenticated!");
                SetAuthenticatedState(true);
                return;
            }

            // If we have config but no valid token, auto-authenticate
            if (_sharePointService.HasSavedConfig())
            {
                Log("🔄 Token expired or missing, re-authenticating...");
                var success = await _sharePointService.AuthenticateAsync();
                if (success)
                {
                    SetAuthenticatedState(true);
                }
            }
        }

        private void InitializeServices()
        {
            _sharePointService = new SharePointImportService(Log);
        }

        private void Log(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n";
                scrollLog.ScrollToEnd();
            });
        }

        private void SetAuthenticatedState(bool authenticated)
        {
            _isAuthenticated = authenticated;
            
            if (authenticated)
            {
                txtAuthStatus.Text = "✅ Authenticated successfully";
                txtAuthStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(5, 150, 105)); // Green
                
                btnSignOut.IsEnabled = true;
                btnTestConnection.IsEnabled = true;
                btnListSites.IsEnabled = true;
                btnDownloadFromUrl.IsEnabled = true;
                btnBrowseSites.IsEnabled = true;
            }
            else
            {
                txtAuthStatus.Text = "❌ Not authenticated";
                txtAuthStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(220, 38, 38)); // Red
                
                btnSignOut.IsEnabled = false;
                btnTestConnection.IsEnabled = false;
                btnListSites.IsEnabled = false;
                btnDownloadFromUrl.IsEnabled = false;
                btnBrowseSites.IsEnabled = false;
            }
        }
        
        private async void BtnTestConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnTestConnection.IsEnabled = false;
                await _sharePointService!.TestConnectionAsync();
            }
            finally
            {
                btnTestConnection.IsEnabled = true;
            }
        }

        private async void BtnListSites_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnListSites.IsEnabled = false;
                await _sharePointService!.ListAllAccessibleSitesAsync();
            }
            finally
            {
                btnListSites.IsEnabled = true;
            }
        }

        private void BtnSaveConfig_Click(object sender, RoutedEventArgs e)
        {
            var config = new AzureAdConfig
            {
                TenantId = txtTenantId.Text.Trim(),
                ClientId = txtClientId.Text.Trim(),
                ClientSecret = txtClientSecret.Password
            };

            if (string.IsNullOrEmpty(config.TenantId) || 
                string.IsNullOrEmpty(config.ClientId) || 
                string.IsNullOrEmpty(config.ClientSecret))
            {
                MessageBox.Show("Please fill in all Azure AD fields", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_sharePointService!.SaveConfig(config))
            {
                MessageBox.Show("Configuration saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async void BtnAuthenticate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnAuthenticate.IsEnabled = false;
                
                var config = new AzureAdConfig
                {
                    TenantId = txtTenantId.Text.Trim(),
                    ClientId = txtClientId.Text.Trim(),
                    ClientSecret = txtClientSecret.Password
                };

                // Client Credentials requires all fields
                if (string.IsNullOrEmpty(config.TenantId) || 
                    string.IsNullOrEmpty(config.ClientId) || 
                    string.IsNullOrEmpty(config.ClientSecret))
                {
                    MessageBox.Show("Client Credentials auth requires Tenant ID, Client ID, and Client Secret.\n\nFor third-party SharePoint access, use 'Interactive Login' button instead.", 
                        "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Log("Starting Client Credentials authentication (Own Tenant)...");

                var success = await _sharePointService!.AuthenticateWithClientCredentialsAsync(config);
                SetAuthenticatedState(success);
                
                if (!success)
                {
                    txtAuthStatus.Text = "❌ Authentication failed - check credentials";
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Authentication error: {ex.Message}");
                txtAuthStatus.Text = $"❌ Error: {ex.Message}";
            }
            finally
            {
                btnAuthenticate.IsEnabled = true;
            }
        }

        private async void BtnInteractiveLogin_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnInteractiveLogin.IsEnabled = false;
                
                var config = new AzureAdConfig
                {
                    TenantId = txtTenantId.Text.Trim(),
                    ClientId = txtClientId.Text.Trim(),
                    ClientSecret = "" // Not needed for interactive
                };

                if (string.IsNullOrEmpty(config.ClientId))
                {
                    MessageBox.Show("Please enter the Client ID (Application ID) from your Azure AD App Registration.", 
                        "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Log("Starting Interactive authentication (for Third-Party SharePoint)...");
                Log("A browser window will open - sign in with your account that has access to the SharePoint.");

                var success = await _sharePointService!.AuthenticateInteractiveAsync(config);
                SetAuthenticatedState(success);
                
                if (!success)
                {
                    txtAuthStatus.Text = "❌ Interactive authentication failed";
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Interactive auth error: {ex.Message}");
                txtAuthStatus.Text = $"❌ Error: {ex.Message}";
            }
            finally
            {
                btnInteractiveLogin.IsEnabled = true;
            }
        }

        private void BtnSignOut_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _sharePointService!.SignOut();
                _isAuthenticated = false;
                
                txtAuthStatus.Text = "Not authenticated";
                txtAuthStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(220, 38, 38)); // Red
                
                btnSignOut.IsEnabled = false;
                btnDownloadFromUrl.IsEnabled = false;
                btnBrowseSites.IsEnabled = false;
                btnDownloadFromDrive.IsEnabled = false;
                borderSitesList.Visibility = Visibility.Collapsed;
                
                _sites.Clear();
                _drives.Clear();
                lstSites.ItemsSource = null;
                lstDrives.ItemsSource = null;
            }
            catch (Exception ex)
            {
                Log($"❌ Sign out error: {ex.Message}");
            }
        }

        private async void BtnDownloadFromUrl_Click(object sender, RoutedEventArgs e)
        {
            var url = txtSharePointUrl.Text.Trim();
            if (string.IsNullOrEmpty(url))
            {
                MessageBox.Show("Please enter a SharePoint URL", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                btnDownloadFromUrl.IsEnabled = false;
                Log($"Accessing shared URL: {url}");

                _downloadedFolderPath = await _sharePointService!.DownloadFromSharedLinkAsync(url);
                
                if (!string.IsNullOrEmpty(_downloadedFolderPath))
                {
                    txtDownloadPath.Text = $"📁 Downloaded to: {_downloadedFolderPath}";
                    btnImportToCase.IsEnabled = true;
                    Log($"✅ Files downloaded to: {_downloadedFolderPath}");
                }
                else
                {
                    Log("❌ Failed to download files from URL");
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Download error: {ex.Message}");
            }
            finally
            {
                btnDownloadFromUrl.IsEnabled = _isAuthenticated;
            }
        }

        private async void BtnBrowseSites_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnBrowseSites.IsEnabled = false;
                Log("Browsing accessible SharePoint sites...");

                _sites = await _sharePointService!.ListAccessibleSitesAsync();
                
                lstSites.ItemsSource = _sites;
                lstDrives.ItemsSource = null;
                _drives.Clear();
                
                borderSitesList.Visibility = Visibility.Visible;
                
                if (_sites.Count == 0)
                {
                    Log("⚠️ No accessible sites found. Try using a shared URL instead.");
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Browse error: {ex.Message}");
            }
            finally
            {
                btnBrowseSites.IsEnabled = _isAuthenticated;
            }
        }

        private async void LstSites_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstSites.SelectedItem is SharePointSite selectedSite)
            {
                try
                {
                    Log($"Loading document libraries for: {selectedSite.Name}");
                    _drives = await _sharePointService!.GetSiteDrivesAsync(selectedSite.Id);
                    lstDrives.ItemsSource = _drives;
                }
                catch (Exception ex)
                {
                    Log($"❌ Error loading drives: {ex.Message}");
                }
            }
        }

        private void LstDrives_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            btnDownloadFromDrive.IsEnabled = lstDrives.SelectedItem != null;
        }

        private async void BtnDownloadFromDrive_Click(object sender, RoutedEventArgs e)
        {
            if (lstDrives.SelectedItem is not SharePointDrive selectedDrive)
            {
                MessageBox.Show("Please select a document library", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                btnDownloadFromDrive.IsEnabled = false;
                Log($"Downloading files from: {selectedDrive.Name}");

                _downloadedFolderPath = await _sharePointService!.DownloadAllFilesAsync(selectedDrive.Id, "root", selectedDrive.Name);
                
                if (!string.IsNullOrEmpty(_downloadedFolderPath))
                {
                    txtDownloadPath.Text = $"📁 Downloaded to: {_downloadedFolderPath}";
                    btnImportToCase.IsEnabled = true;
                    Log($"✅ Files downloaded to: {_downloadedFolderPath}");
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Download error: {ex.Message}");
            }
            finally
            {
                btnDownloadFromDrive.IsEnabled = true;
            }
        }

        private async void BtnImportToCase_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_downloadedFolderPath) || !Directory.Exists(_downloadedFolderPath))
            {
                MessageBox.Show("No downloaded files to import. Please download files first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(txtCaseId.Text.Trim(), out int caseId) || caseId <= 0)
            {
                MessageBox.Show("Please enter a valid Case ID", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtConnectionString.Text.Trim()))
            {
                MessageBox.Show("Please enter a database connection string", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtBucketName.Text.Trim()))
            {
                MessageBox.Show("Please enter S3 bucket name", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                btnImportToCase.IsEnabled = false;
                Log($"Starting import to Case ID: {caseId}");
                Log($"Source folder: {_downloadedFolderPath}");

                var s3Config = new S3Configuration
                {
                    BucketName = txtBucketName.Text.Trim(),
                    ProfilePath = txtProfilePath.Text.Trim()
                };

                var filesImportService = new FilesCollectionImportService(
                    txtConnectionString.Text.Trim(), 
                    s3Config, 
                    Log);

                await Task.Run(() =>
                {
                    var (success, errors) = filesImportService.UploadFilesFromFolder(_downloadedFolderPath, caseId);

                    Dispatcher.Invoke(() =>
                    {
                        Log("═══════════════════════════════════════");
                        Log($"✅ Import completed!");
                        Log($"   Successful: {success}");
                        Log($"   Errors: {errors}");
                        Log("═══════════════════════════════════════");
                    });
                });

                MessageBox.Show($"Import completed!\n\nSuccessful files imported.", 
                    "Import Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Log($"❌ Import error: {ex.Message}");
                MessageBox.Show($"Import failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnImportToCase.IsEnabled = true;
            }
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtLog.Text = string.Empty;
        }
    }
}
