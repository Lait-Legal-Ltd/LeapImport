using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace LeapMergeDoc.Services
{
    /// <summary>
    /// Azure AD App Registration credentials for SharePoint access
    /// </summary>
    public class AzureAdConfig
    {
        public string TenantId { get; set; } = "";      // Directory (tenant) ID
        public string ClientId { get; set; } = "";      // Application (client) ID
        public string ClientSecret { get; set; } = "";  // Client secret value (for app-only auth)
    }

    public class SharePointImportService
    {
        private readonly Action<string> _logAction;
        private readonly string _downloadBasePath;
        private string? _accessToken;
        private DateTime _tokenExpiry = DateTime.MinValue;
        private IPublicClientApplication? _publicClientApp;

        // Config file path for Azure AD credentials (simple text file)
        private static readonly string ConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LeapMergeDoc",
            "sharepoint_credentials.txt");

        // Token cache file (stores access token to avoid re-auth)
        private static readonly string TokenCacheFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LeapMergeDoc",
            "sharepoint_token.txt");

        // Microsoft Graph API scopes
        private const string GraphScope = "https://graph.microsoft.com/.default";
        
        // Delegated scopes for interactive login (user permissions)
        private readonly string[] _delegatedScopes = new[]
        {
            "https://graph.microsoft.com/Files.Read.All",
            "https://graph.microsoft.com/Sites.Read.All",
            "https://graph.microsoft.com/User.Read"
        };

        public SharePointImportService(Action<string> logAction, string? downloadBasePath = null)
        {
            _logAction = logAction;
            _downloadBasePath = downloadBasePath ?? Path.Combine(Path.GetTempPath(), "SharePointDownloads");
            
            if (!Directory.Exists(_downloadBasePath))
            {
                Directory.CreateDirectory(_downloadBasePath);
            }

            _logAction($"📁 Credentials file: {ConfigFilePath}");
            _logAction($"📁 Token cache file: {TokenCacheFilePath}");
        }

        /// <summary>
        /// Load Azure AD config from text file
        /// Format: TenantId|ClientId|ClientSecret (one line)
        /// </summary>
        public AzureAdConfig? LoadConfig()
        {
            if (!File.Exists(ConfigFilePath))
            {
                _logAction("⚠️ Credentials file not found. Please save Azure AD credentials first.");
                return null;
            }

            try
            {
                var lines = File.ReadAllLines(ConfigFilePath);
                if (lines.Length >= 2)
                {
                    var config = new AzureAdConfig
                    {
                        TenantId = lines.Length > 0 ? lines[0].Trim() : "",
                        ClientId = lines.Length > 1 ? lines[1].Trim() : "",
                        ClientSecret = lines.Length > 2 ? lines[2].Trim() : ""
                    };
                    _logAction("✅ Credentials loaded from file");
                    return config;
                }
                else
                {
                    _logAction("❌ Invalid credentials file format. Expected at least 2 lines: TenantId, ClientId");
                    return null;
                }
            }
            catch (Exception ex)
            {
                _logAction($"❌ Failed to load credentials: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Save Azure AD config to text file
        /// Format: Each value on separate line
        /// </summary>
        public bool SaveConfig(AzureAdConfig config)
        {
            try
            {
                var dir = Path.GetDirectoryName(ConfigFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Write each value on separate line for easy editing
                var content = $"{config.TenantId}\n{config.ClientId}\n{config.ClientSecret}";
                File.WriteAllText(ConfigFilePath, content);
                _logAction($"✅ Credentials saved to: {ConfigFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                _logAction($"❌ Failed to save credentials: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Load cached access token from file
        /// Format: AccessToken|ExpiryDateTime
        /// </summary>
        private bool LoadCachedToken()
        {
            if (!File.Exists(TokenCacheFilePath))
            {
                return false;
            }

            try
            {
                var lines = File.ReadAllLines(TokenCacheFilePath);
                if (lines.Length >= 2)
                {
                    var token = lines[0].Trim();
                    var expiryStr = lines[1].Trim();
                    
                    if (DateTime.TryParse(expiryStr, out var expiry) && expiry > DateTime.UtcNow.AddMinutes(5))
                    {
                        _accessToken = token;
                        _tokenExpiry = expiry;
                        _logAction($"✅ Loaded cached token (expires: {_tokenExpiry:yyyy-MM-dd HH:mm:ss})");
                        return true;
                    }
                    else
                    {
                        _logAction("⚠️ Cached token expired, need to re-authenticate");
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ Could not load cached token: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Save access token to cache file
        /// </summary>
        private void SaveTokenToCache()
        {
            try
            {
                var dir = Path.GetDirectoryName(TokenCacheFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var content = $"{_accessToken}\n{_tokenExpiry:O}";
                File.WriteAllText(TokenCacheFilePath, content);
                _logAction($"✅ Token cached to file (expires: {_tokenExpiry:yyyy-MM-dd HH:mm:ss})");
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ Could not cache token: {ex.Message}");
            }
        }

        /// <summary>
        /// Check if config file exists
        /// </summary>
        public bool HasSavedConfig()
        {
            return File.Exists(ConfigFilePath);
        }

        /// <summary>
        /// Check if there's a valid cached token
        /// </summary>
        public bool HasValidCachedToken()
        {
            return LoadCachedToken();
        }

        /// <summary>
        /// Authenticate using Client Credentials Flow (App-only, no user interaction)
        /// NOTE: This only works for SharePoint sites in YOUR OWN tenant!
        /// </summary>
        public async Task<bool> AuthenticateWithClientCredentialsAsync(AzureAdConfig config)
        {
            try
            {
                _logAction("🔐 Authenticating with Azure AD (Client Credentials - Own Tenant Only)...");
                _logAction($"   Tenant ID: {config.TenantId}");
                _logAction($"   Client ID: {config.ClientId}");

                var app = ConfidentialClientApplicationBuilder
                    .Create(config.ClientId)
                    .WithClientSecret(config.ClientSecret)
                    .WithAuthority(new Uri($"https://login.microsoftonline.com/{config.TenantId}"))
                    .Build();

                var result = await app.AcquireTokenForClient(new[] { GraphScope })
                    .ExecuteAsync();

                _accessToken = result.AccessToken;
                _tokenExpiry = result.ExpiresOn.UtcDateTime;

                // Save token to cache file
                SaveTokenToCache();

                _logAction($"✅ Authentication successful!");
                _logAction($"   Token expires: {_tokenExpiry:yyyy-MM-dd HH:mm:ss} UTC");
                return true;
            }
            catch (MsalException ex)
            {
                _logAction($"❌ Authentication failed: {ex.Message}");
                _logAction($"   Error code: {ex.ErrorCode}");
                return false;
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error during authentication: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Authenticate using Interactive/Delegated Flow (User login via browser)
        /// This works for THIRD-PARTY SharePoint sites where you have guest access!
        /// </summary>
        public async Task<bool> AuthenticateInteractiveAsync(AzureAdConfig config)
        {
            try
            {
                _logAction("🔐 Authenticating with Azure AD (Interactive - for Third-Party Access)...");
                _logAction($"   Client ID: {config.ClientId}");
                _logAction("   Opening browser for login...");

                // Use "common" authority for multi-tenant / guest access
                _publicClientApp = PublicClientApplicationBuilder
                    .Create(config.ClientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, "common")
                    .WithRedirectUri("http://localhost")
                    .Build();

                AuthenticationResult? result = null;

                // Try silent auth first
                var accounts = await _publicClientApp.GetAccountsAsync();
                var firstAccount = accounts.FirstOrDefault();

                if (firstAccount != null)
                {
                    try
                    {
                        result = await _publicClientApp.AcquireTokenSilent(_delegatedScopes, firstAccount).ExecuteAsync();
                        _logAction("✅ Authenticated silently (cached credentials)");
                    }
                    catch (MsalUiRequiredException)
                    {
                        // Need interactive
                    }
                }

                if (result == null)
                {
                    // Interactive authentication - opens browser
                    result = await _publicClientApp.AcquireTokenInteractive(_delegatedScopes)
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();

                    _logAction($"✅ Authenticated as: {result.Account.Username}");
                }

                _accessToken = result.AccessToken;
                _tokenExpiry = result.ExpiresOn.UtcDateTime;

                // Save token to cache
                SaveTokenToCache();

                _logAction($"✅ Interactive authentication successful!");
                _logAction($"   User: {result.Account.Username}");
                _logAction($"   Token expires: {_tokenExpiry:yyyy-MM-dd HH:mm:ss} UTC");
                return true;
            }
            catch (MsalException ex)
            {
                _logAction($"❌ Interactive authentication failed: {ex.Message}");
                _logAction($"   Error code: {ex.ErrorCode}");
                return false;
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error during interactive authentication: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Test the current token and show what it can access
        /// </summary>
        public async Task TestConnectionAsync()
        {
            if (string.IsNullOrEmpty(_accessToken))
            {
                _logAction("❌ Not authenticated. Please authenticate first.");
                return;
            }

            _logAction("═══════════════════════════════════════");
            _logAction("🔍 Testing connection and token scope...");

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

            // Test 1: Get current user (works with both app and user tokens)
            try
            {
                var meResponse = await client.GetAsync("https://graph.microsoft.com/v1.0/me");
                if (meResponse.IsSuccessStatusCode)
                {
                    var meContent = await meResponse.Content.ReadAsStringAsync();
                    var meJson = JsonDocument.Parse(meContent);
                    var displayName = meJson.RootElement.TryGetProperty("displayName", out var dn) ? dn.GetString() : "N/A";
                    var email = meJson.RootElement.TryGetProperty("mail", out var em) ? em.GetString() : "N/A";
                    _logAction($"✅ Token type: USER (Delegated)");
                    _logAction($"   User: {displayName}");
                    _logAction($"   Email: {email}");
                }
                else if (meResponse.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    _logAction("✅ Token type: APP-ONLY (Client Credentials)");
                    _logAction("   ⚠️ App tokens cannot access third-party SharePoint!");
                    _logAction("   ⚠️ Use 'Interactive Login' for third-party access.");
                }
                else
                {
                    var error = await meResponse.Content.ReadAsStringAsync();
                    _logAction($"⚠️ /me endpoint: {meResponse.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ /me test failed: {ex.Message}");
            }

            // Test 2: List accessible sites
            try
            {
                var sitesResponse = await client.GetAsync("https://graph.microsoft.com/v1.0/sites?search=*&$top=5");
                if (sitesResponse.IsSuccessStatusCode)
                {
                    var sitesContent = await sitesResponse.Content.ReadAsStringAsync();
                    var sitesJson = JsonDocument.Parse(sitesContent);
                    
                    if (sitesJson.RootElement.TryGetProperty("value", out var sites))
                    {
                        var count = sites.GetArrayLength();
                        _logAction($"✅ Accessible SharePoint sites: {count}");
                        
                        foreach (var site in sites.EnumerateArray())
                        {
                            var siteName = site.TryGetProperty("displayName", out var sn) ? sn.GetString() : "Unknown";
                            var siteUrl = site.TryGetProperty("webUrl", out var su) ? su.GetString() : "";
                            _logAction($"   📁 {siteName}: {siteUrl}");
                        }
                    }
                }
                else
                {
                    var error = await sitesResponse.Content.ReadAsStringAsync();
                    _logAction($"⚠️ Cannot list sites: {sitesResponse.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ Sites test failed: {ex.Message}");
            }

            _logAction("═══════════════════════════════════════");
        }

        /// <summary>
        /// Authenticate using saved config (with token caching)
        /// </summary>
        public async Task<bool> AuthenticateAsync()
        {
            // First try to use cached token
            if (LoadCachedToken())
            {
                _logAction("✅ Using cached token (no re-authentication needed)");
                return true;
            }

            // Need to authenticate
            var config = LoadConfig();
            if (config == null)
            {
                _logAction("❌ No saved configuration. Please enter Azure AD credentials.");
                return false;
            }

            if (string.IsNullOrEmpty(config.ClientId))
            {
                _logAction("❌ Incomplete configuration. Please verify Client ID is filled.");
                return false;
            }

            // If we have client secret, use client credentials
            if (!string.IsNullOrEmpty(config.ClientSecret))
            {
                return await AuthenticateWithClientCredentialsAsync(config);
            }
            else
            {
                // Otherwise use interactive
                return await AuthenticateInteractiveAsync(config);
            }
        }

        /// <summary>
        /// Check if current token is valid
        /// </summary>
        public bool IsAuthenticated()
        {
            // First check in-memory token
            if (!string.IsNullOrEmpty(_accessToken) && DateTime.UtcNow < _tokenExpiry.AddMinutes(-5))
            {
                return true;
            }
            
            // Try to load from cache
            return LoadCachedToken();
        }

        /// <summary>
        /// Clear authentication (memory only, keep cached file)
        /// </summary>
        public void SignOut()
        {
            _accessToken = null;
            _tokenExpiry = DateTime.MinValue;
            _logAction("✅ Signed out (token cleared from memory)");
        }

        /// <summary>
        /// Delete all saved files (credentials and token cache)
        /// </summary>
        public void DeleteAllConfig()
        {
            if (File.Exists(ConfigFilePath))
            {
                File.Delete(ConfigFilePath);
                _logAction("🗑️ Credentials file deleted");
            }
            if (File.Exists(TokenCacheFilePath))
            {
                File.Delete(TokenCacheFilePath);
                _logAction("🗑️ Token cache deleted");
            }
            SignOut();
        }

        /// <summary>
        /// Get path to credentials file (for user reference)
        /// </summary>
        public string GetCredentialsFilePath()
        {
            return ConfigFilePath;
        }

        /// <summary>
        /// List all SharePoint sites the user has access to
        /// </summary>
        public async Task<List<SharePointSite>> ListAccessibleSitesAsync()
        {
            var sites = new List<SharePointSite>();

            if (string.IsNullOrEmpty(_accessToken))
            {
                _logAction("❌ Not authenticated. Please authenticate first.");
                return sites;
            }

            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                // Search for sites the user has access to
                var response = await client.GetAsync("https://graph.microsoft.com/v1.0/sites?search=*");
                
                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var json = JsonDocument.Parse(content);
                    
                    if (json.RootElement.TryGetProperty("value", out var sitesArray))
                    {
                        foreach (var site in sitesArray.EnumerateArray())
                        {
                            sites.Add(new SharePointSite
                            {
                                Id = site.GetProperty("id").GetString() ?? "",
                                Name = site.GetProperty("displayName").GetString() ?? "",
                                WebUrl = site.GetProperty("webUrl").GetString() ?? ""
                            });
                        }
                    }

                    _logAction($"✅ Found {sites.Count} accessible SharePoint sites");
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logAction($"❌ Failed to list sites: {response.StatusCode} - {error}");
                }
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error listing sites: {ex.Message}");
            }

            return sites;
        }

        /// <summary>
        /// Get drives (document libraries) in a SharePoint site
        /// </summary>
        public async Task<List<SharePointDrive>> GetSiteDrivesAsync(string siteId)
        {
            var drives = new List<SharePointDrive>();

            if (string.IsNullOrEmpty(_accessToken))
            {
                _logAction("❌ Not authenticated. Please authenticate first.");
                return drives;
            }

            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                var response = await client.GetAsync($"https://graph.microsoft.com/v1.0/sites/{siteId}/drives");
                
                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var json = JsonDocument.Parse(content);
                    
                    if (json.RootElement.TryGetProperty("value", out var drivesArray))
                    {
                        foreach (var drive in drivesArray.EnumerateArray())
                        {
                            drives.Add(new SharePointDrive
                            {
                                Id = drive.GetProperty("id").GetString() ?? "",
                                Name = drive.GetProperty("name").GetString() ?? "",
                                WebUrl = drive.TryGetProperty("webUrl", out var webUrl) ? webUrl.GetString() ?? "" : ""
                            });
                        }
                    }

                    _logAction($"✅ Found {drives.Count} document libraries");
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logAction($"❌ Failed to get drives: {response.StatusCode} - {error}");
                }
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error getting drives: {ex.Message}");
            }

            return drives;
        }

        /// <summary>
        /// List items (files and folders) in a drive or folder
        /// </summary>
        public async Task<List<SharePointItem>> ListItemsAsync(string driveId, string folderId = "root")
        {
            var items = new List<SharePointItem>();

            if (string.IsNullOrEmpty(_accessToken))
            {
                _logAction("❌ Not authenticated. Please authenticate first.");
                return items;
            }

            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                string url = folderId == "root" 
                    ? $"https://graph.microsoft.com/v1.0/drives/{driveId}/root/children"
                    : $"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{folderId}/children";

                await ListItemsRecursiveAsync(client, url, items, driveId, "");
                
                _logAction($"✅ Found {items.Count} total items");
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error listing items: {ex.Message}");
            }

            return items;
        }

        private async Task ListItemsRecursiveAsync(HttpClient client, string url, List<SharePointItem> items, string driveId, string parentPath)
        {
            var response = await client.GetAsync(url);
            
            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                var json = JsonDocument.Parse(content);
                
                if (json.RootElement.TryGetProperty("value", out var itemsArray))
                {
                    foreach (var item in itemsArray.EnumerateArray())
                    {
                        var itemId = item.GetProperty("id").GetString() ?? "";
                        var itemName = item.GetProperty("name").GetString() ?? "";
                        var isFolder = item.TryGetProperty("folder", out _);
                        var itemPath = string.IsNullOrEmpty(parentPath) ? itemName : $"{parentPath}/{itemName}";

                        var sharePointItem = new SharePointItem
                        {
                            Id = itemId,
                            Name = itemName,
                            IsFolder = isFolder,
                            Path = itemPath,
                            Size = item.TryGetProperty("size", out var size) ? size.GetInt64() : 0,
                            WebUrl = item.TryGetProperty("webUrl", out var webUrl) ? webUrl.GetString() ?? "" : "",
                            DownloadUrl = item.TryGetProperty("@microsoft.graph.downloadUrl", out var downloadUrl) 
                                ? downloadUrl.GetString() ?? "" 
                                : ""
                        };

                        items.Add(sharePointItem);
                        _logAction($"  {(isFolder ? "📁" : "📄")} {itemPath}");

                        // Recursively list folder contents
                        if (isFolder)
                        {
                            var childrenUrl = $"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{itemId}/children";
                            await ListItemsRecursiveAsync(client, childrenUrl, items, driveId, itemPath);
                        }
                    }
                }

                // Handle pagination
                if (json.RootElement.TryGetProperty("@odata.nextLink", out var nextLink))
                {
                    await ListItemsRecursiveAsync(client, nextLink.GetString()!, items, driveId, parentPath);
                }
            }
        }

        /// <summary>
        /// Download all files from a drive/folder to local path
        /// </summary>
        public async Task<string> DownloadAllFilesAsync(string driveId, string folderId = "root", string? targetFolderName = null)
        {
            var downloadPath = Path.Combine(_downloadBasePath, targetFolderName ?? $"Download_{DateTime.Now:yyyyMMdd_HHmmss}");
            
            if (!Directory.Exists(downloadPath))
            {
                Directory.CreateDirectory(downloadPath);
            }

            _logAction($"📥 Downloading files to: {downloadPath}");

            var items = await ListItemsAsync(driveId, folderId);
            var files = items.Where(i => !i.IsFolder).ToList();

            _logAction($"Found {files.Count} files to download");

            int downloaded = 0;
            int errors = 0;

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

            foreach (var file in files)
            {
                try
                {
                    // Create folder structure
                    var fileFolderPath = Path.GetDirectoryName(file.Path) ?? "";
                    var localFolderPath = Path.Combine(downloadPath, fileFolderPath);
                    
                    if (!string.IsNullOrEmpty(fileFolderPath) && !Directory.Exists(localFolderPath))
                    {
                        Directory.CreateDirectory(localFolderPath);
                    }

                    var localFilePath = Path.Combine(downloadPath, file.Path);

                    // Get download URL
                    string downloadUrl = file.DownloadUrl;
                    if (string.IsNullOrEmpty(downloadUrl))
                    {
                        // Need to fetch download URL
                        var itemResponse = await client.GetAsync($"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{file.Id}");
                        if (itemResponse.IsSuccessStatusCode)
                        {
                            var itemContent = await itemResponse.Content.ReadAsStringAsync();
                            var itemJson = JsonDocument.Parse(itemContent);
                            if (itemJson.RootElement.TryGetProperty("@microsoft.graph.downloadUrl", out var dlUrl))
                            {
                                downloadUrl = dlUrl.GetString() ?? "";
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(downloadUrl))
                    {
                        // Download file (download URL doesn't need auth header)
                        using var downloadClient = new HttpClient();
                        var fileBytes = await downloadClient.GetByteArrayAsync(downloadUrl);
                        await File.WriteAllBytesAsync(localFilePath, fileBytes);
                        
                        downloaded++;
                        _logAction($"✅ Downloaded: {file.Path} ({FormatFileSize(file.Size)})");
                    }
                    else
                    {
                        errors++;
                        _logAction($"❌ Could not get download URL for: {file.Path}");
                    }
                }
                catch (Exception ex)
                {
                    errors++;
                    _logAction($"❌ Error downloading {file.Path}: {ex.Message}");
                }
            }

            _logAction($"\n📊 Download complete: {downloaded} successful, {errors} errors");
            return downloadPath;
        }

        /// <summary>
        /// Download files from a shared link URL
        /// Tries multiple methods: sharing link API, then direct site access
        /// </summary>
        public async Task<string> DownloadFromSharedLinkAsync(string sharedUrl, string? targetFolderName = null)
        {
            if (string.IsNullOrEmpty(_accessToken))
            {
                _logAction("❌ Not authenticated. Please authenticate first.");
                return "";
            }

            _logAction($"📥 Accessing shared URL: {sharedUrl}");

            // Method 1: Try via sharing link API (works for "Anyone with the link")
            var result = await TryAccessViaSharingLinkAsync(sharedUrl, targetFolderName);
            if (!string.IsNullOrEmpty(result))
            {
                return result;
            }

            // Method 2: Try direct site access (works for guest users with direct access)
            _logAction("🔄 Trying direct site access (for guest accounts)...");
            result = await TryAccessDirectlyAsync(sharedUrl, targetFolderName);
            if (!string.IsNullOrEmpty(result))
            {
                return result;
            }

            _logAction("❌ Failed to download files from URL");
            return "";
        }

        /// <summary>
        /// Try to access via the /shares endpoint (for sharing links)
        /// </summary>
        private async Task<string> TryAccessViaSharingLinkAsync(string sharedUrl, string? targetFolderName)
        {
            try
            {
                // Encode the shared URL for Graph API
                var encodedUrl = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(sharedUrl))
                    .TrimEnd('=')
                    .Replace('/', '_')
                    .Replace('+', '-');

                var sharingUrl = $"u!{encodedUrl}";

                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                _logAction("   Trying /shares endpoint...");
                var response = await client.GetAsync($"https://graph.microsoft.com/v1.0/shares/{sharingUrl}/driveItem");
                
                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var json = JsonDocument.Parse(content);
                    
                    var driveId = json.RootElement.GetProperty("parentReference").GetProperty("driveId").GetString();
                    var itemId = json.RootElement.GetProperty("id").GetString();
                    var itemName = json.RootElement.GetProperty("name").GetString();
                    var isFolder = json.RootElement.TryGetProperty("folder", out _);

                    _logAction($"✅ Found shared item: {itemName}");

                    if (isFolder)
                    {
                        return await DownloadAllFilesAsync(driveId!, itemId!, targetFolderName ?? itemName);
                    }
                    else
                    {
                        var downloadPath = Path.Combine(_downloadBasePath, targetFolderName ?? "SharedFiles");
                        if (!Directory.Exists(downloadPath))
                        {
                            Directory.CreateDirectory(downloadPath);
                        }

                        if (json.RootElement.TryGetProperty("@microsoft.graph.downloadUrl", out var downloadUrl))
                        {
                            using var downloadClient = new HttpClient();
                            var fileBytes = await downloadClient.GetByteArrayAsync(downloadUrl.GetString());
                            var filePath = Path.Combine(downloadPath, itemName!);
                            await File.WriteAllBytesAsync(filePath, fileBytes);
                            _logAction($"✅ Downloaded: {itemName}");
                        }

                        return downloadPath;
                    }
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logAction($"⚠️ Sharing link API failed: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                _logAction($"⚠️ Sharing link method error: {ex.Message}");
            }

            return "";
        }

        /// <summary>
        /// Try to access SharePoint directly by parsing the URL and navigating the site
        /// This works for guest users who have direct access to the SharePoint site
        /// </summary>
        private async Task<string> TryAccessDirectlyAsync(string sharedUrl, string? targetFolderName)
        {
            try
            {
                // Parse SharePoint URL to extract host and site
                // Format: https://tenant.sharepoint.com/:f:/s/sitename/... or https://tenant.sharepoint.com/sites/sitename/...
                var uri = new Uri(sharedUrl);
                var host = uri.Host; // e.g., "baisuk.sharepoint.com"
                var pathSegments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);

                string? sitePath = null;
                
                // Handle /:f:/s/sitename/ format (sharing links)
                if (pathSegments.Length >= 3 && pathSegments[0].StartsWith(":") && pathSegments[1] == "s")
                {
                    sitePath = $"/sites/{pathSegments[2]}";
                }
                // Handle /sites/sitename/ format (direct links)
                else if (pathSegments.Length >= 2 && pathSegments[0] == "sites")
                {
                    sitePath = $"/sites/{pathSegments[1]}";
                }

                if (string.IsNullOrEmpty(sitePath))
                {
                    _logAction("⚠️ Could not parse site path from URL");
                    return "";
                }

                _logAction($"   Looking up site: {host}{sitePath}");

                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                // Get site by hostname and path
                var siteResponse = await client.GetAsync($"https://graph.microsoft.com/v1.0/sites/{host}:{sitePath}");
                
                if (siteResponse.IsSuccessStatusCode)
                {
                    var siteContent = await siteResponse.Content.ReadAsStringAsync();
                    var siteJson = JsonDocument.Parse(siteContent);
                    
                    var siteId = siteJson.RootElement.GetProperty("id").GetString();
                    var siteName = siteJson.RootElement.GetProperty("displayName").GetString();
                    
                    _logAction($"✅ Found site: {siteName}");
                    _logAction($"   Site ID: {siteId}");

                    // List document libraries (drives)
                    var drivesResponse = await client.GetAsync($"https://graph.microsoft.com/v1.0/sites/{siteId}/drives");
                    
                    if (drivesResponse.IsSuccessStatusCode)
                    {
                        var drivesContent = await drivesResponse.Content.ReadAsStringAsync();
                        var drivesJson = JsonDocument.Parse(drivesContent);
                        
                        if (drivesJson.RootElement.TryGetProperty("value", out var drives) && drives.GetArrayLength() > 0)
                        {
                            _logAction($"📁 Found {drives.GetArrayLength()} document libraries:");
                            
                            foreach (var drive in drives.EnumerateArray())
                            {
                                var driveName = drive.GetProperty("name").GetString();
                                var driveId = drive.GetProperty("id").GetString();
                                _logAction($"   📁 {driveName}");
                            }

                            // Download from first drive (usually "Documents")
                            var firstDrive = drives[0];
                            var firstDriveId = firstDrive.GetProperty("id").GetString();
                            var firstDriveName = firstDrive.GetProperty("name").GetString();
                            
                            _logAction($"\n📥 Downloading from: {firstDriveName}");
                            return await DownloadAllFilesAsync(firstDriveId!, "root", targetFolderName ?? siteName);
                        }
                        else
                        {
                            _logAction("⚠️ No document libraries found in site");
                        }
                    }
                    else
                    {
                        var error = await drivesResponse.Content.ReadAsStringAsync();
                        _logAction($"❌ Cannot access document libraries: {drivesResponse.StatusCode}");
                        _logAction($"   {error}");
                    }
                }
                else
                {
                    var error = await siteResponse.Content.ReadAsStringAsync();
                    _logAction($"❌ Cannot access site directly: {siteResponse.StatusCode}");
                    _logAction($"   This might mean your guest account doesn't have SharePoint site access.");
                    _logAction($"   Ask the SharePoint owner to grant you 'Site Member' or 'Site Visitor' access.");
                }
            }
            catch (Exception ex)
            {
                _logAction($"❌ Direct access error: {ex.Message}");
            }

            return "";
        }

        /// <summary>
        /// List all SharePoint sites you have access to (including where you're a guest)
        /// </summary>
        public async Task ListAllAccessibleSitesAsync()
        {
            if (string.IsNullOrEmpty(_accessToken))
            {
                _logAction("❌ Not authenticated. Please authenticate first.");
                return;
            }

            _logAction("═══════════════════════════════════════");
            _logAction("🔍 Searching for accessible SharePoint sites...");
            _logAction("   (This includes sites where you're a guest)");

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

            try
            {
                // Search all sites
                var response = await client.GetAsync("https://graph.microsoft.com/v1.0/sites?search=*");
                
                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var json = JsonDocument.Parse(content);
                    
                    if (json.RootElement.TryGetProperty("value", out var sites))
                    {
                        var count = sites.GetArrayLength();
                        _logAction($"\n✅ Found {count} accessible sites:\n");
                        
                        foreach (var site in sites.EnumerateArray())
                        {
                            var name = site.TryGetProperty("displayName", out var dn) ? dn.GetString() : "Unknown";
                            var webUrl = site.TryGetProperty("webUrl", out var wu) ? wu.GetString() : "";
                            var id = site.GetProperty("id").GetString();
                            
                            _logAction($"📁 {name}");
                            _logAction($"   URL: {webUrl}");
                            _logAction($"   ID: {id}\n");
                        }

                        if (count == 0)
                        {
                            _logAction("No sites found. Your guest account might need explicit SharePoint access.");
                            _logAction("Contact the SharePoint owner to add you as a Site Member.");
                        }
                    }
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logAction($"❌ Error searching sites: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                _logAction($"❌ Error: {ex.Message}");
            }

            _logAction("═══════════════════════════════════════");
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double size = bytes;
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }
            return $"{size:0.##} {sizes[order]}";
        }
    }

    public class SharePointSite
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string WebUrl { get; set; } = "";
    }

    public class SharePointDrive
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string WebUrl { get; set; } = "";
    }

    public class SharePointItem
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public bool IsFolder { get; set; }
        public string Path { get; set; } = "";
        public long Size { get; set; }
        public string WebUrl { get; set; } = "";
        public string DownloadUrl { get; set; } = "";
    }
}
