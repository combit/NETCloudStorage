using Azure.Core;
using combit.Reporting;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;

namespace combit.Reporting.CloudStorage
{
    public class MicrosoftOneDriveExportParameter
    {
        /// <summary>
        /// Destination file name in Microsoft OneDrive.
        /// </summary>
        public string CloudFileName { get; set; }

        /// <summary>
        /// Destination path in Microsoft OneDrive root.
        /// </summary>
        public string CloudPath { get; set; }

        /// <summary>
        /// Application Id of your Microsoft App.
        /// </summary>
        public string ApplicationId { get; set; }
    }

    public class MicrosoftOneDriveUploadParameter : MicrosoftOneDriveExportParameter
    {
        /// <summary>
        /// Content to upload.
        /// </summary>
        public FileStream UploadStream { get; set; }
    }

    public class MicrosoftOneDriveCredentials
    {
        /// <summary>
        /// Application Id of your Microsoft App.
        /// </summary>
        public string ApplicationId { get; set; }

        /// <summary>
        /// Application secret of your Microsoft App.
        /// </summary>
        public string ApplicationSecret { get; set; }

        /// <summary>
        /// Redirect uri of your Microsoft App.
        /// </summary>
        public string RedirectUri { get; set; }

        /// <summary>
        /// Required access scope.
        /// </summary>
        public string Scope { get; set; }

        /// <summary>
        /// The current valid refresh token.
        /// </summary>
        public string RefreshToken { get; set; }
    }

    public static class MicrosoftOneDrive
    {
        /// <summary>
        /// Uploads given content to a file in the Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadParameters">requied parameters for Microsoft OneDrive OAuth 2.0 upload.</param>
        public static async void Upload(this Reporting.ListLabel ll, MicrosoftOneDriveUploadParameter uploadParameters)
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient(uploadParameters.ApplicationId);

            if (graphClient != null)
            {
                _ = await graphClient.Me.GetAsync();

                var drive = await graphClient.Me.Drive.GetAsync();
                _ = await graphClient
                              .Drives[drive.Id]
                              .Items["root"]
                              .ItemWithPath($"{uploadParameters.CloudPath}/{uploadParameters.CloudFileName}")
                              .Content
                              .PutAsync(uploadParameters.UploadStream);
            }
        }

        /// <summary>
        /// Uploads given content to a file in the Microsoft OneDrive Cloud Storage with an existing refresh token.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadParameters">requied parameters for Microsoft OneDrive OAuth 2.0 upload silently.</param>
        /// <param name="credentials">requied parameters for Microsoft OneDrive OAuth 2.0 authentication.</param>
        public static void UploadSilently(this Reporting.ListLabel ll, MicrosoftOneDriveUploadParameter uploadParameters, MicrosoftOneDriveCredentials credentials)
        {
            // Initialize the GraphServiceClient.
            GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient(credentials.RefreshToken, uploadParameters.ApplicationId, credentials.ApplicationSecret, credentials.Scope, credentials.RedirectUri);

            // Add the file.
            UploadLargeFile(graphClient, uploadParameters.UploadStream, string.Concat(uploadParameters.CloudPath, "/", uploadParameters.CloudFileName)).Wait();
        }

        // Uploads a large file to the current user's root directory.
        private static async Task UploadLargeFile(GraphServiceClient graphClient, Stream fileStream, string fileName)
        {
            var uploadSessionRequestBody = new DriveUpload.CreateUploadSessionPostRequestBody();

            var drive = await graphClient.Me.Drive.GetAsync();
            var uploadSession = await graphClient
                .Drives[drive.Id]
                .Items["root"]
                .ItemWithPath(fileName)
                .CreateUploadSession
                .PostAsync(uploadSessionRequestBody);

            int maxSliceSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.

            var fileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize, graphClient.RequestAdapter);

            _ = await fileUploadTask.UploadAsync();
        }

        /// <summary>
        /// Check credentials to access Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="credentials">requied parameters for Microsoft OneDrive OAuth 2.0 authentication.</param>
        public static bool CheckCredentials(MicrosoftOneDriveCredentials credentials)
        {
            GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient(credentials.RefreshToken, credentials.ApplicationId, credentials.ApplicationSecret, credentials.Scope, credentials.RedirectUri);
            if (graphClient != null)
            {
                var user = graphClient.Me.GetAsync();
                return user.Result.DisplayName != string.Empty;
            }
            return false;
        }

        /// <summary>
        /// Export a report using current instance of ListLabel and upload it directly to the Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="exportConfiguration">required export configuration for native ListLabel Export method</param>
        /// <param name="exportParameters">requied parameters for Microsoft OneDrive OAuth 2.0 upload.</param>
        public static void Export(this Reporting.ListLabel ll, ExportConfiguration exportConfiguration, MicrosoftOneDriveExportParameter exportParameters)
        {
            ll.AutoShowSelectFile = false;
            ll.AutoShowPrintOptions = false;
            ll.AutoDestination = LlPrintMode.Export;
            ll.AutoProjectType = LlProject.List;
            ll.AutoBoxType = LlBoxType.None;
            exportConfiguration.ExportOptions.Add("Export.Quiet", "1");
            switch (exportConfiguration.ExportTarget)
            {
                case LlExportTarget.Pdf:
                    exportParameters.CloudFileName += ".pdf";
                    break;
                case LlExportTarget.Rtf:
                    exportParameters.CloudFileName += ".rtf";
                    break;
                case LlExportTarget.Xls:
                    exportParameters.CloudFileName += ".xls";
                    break;
                case LlExportTarget.Xlsx:
                    exportParameters.CloudFileName += ".xlsx";
                    break;
                case LlExportTarget.Docx:
                    exportParameters.CloudFileName += ".docx";
                    break;
                case LlExportTarget.Xps:
                    exportParameters.CloudFileName += ".xps";
                    break;
                case LlExportTarget.Mhtml:
                    exportParameters.CloudFileName += ".mhtml";
                    break;
                case LlExportTarget.Text:
                    exportParameters.CloudFileName += ".txt";
                    break;
                case LlExportTarget.Pptx:
                    exportParameters.CloudFileName += ".pptx";
                    break;
                default:
                    exportParameters.CloudFileName += ".zip";
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportSaveAsZip, "1");
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipFile, exportParameters.CloudFileName);
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipPath, Path.GetDirectoryName(exportConfiguration.Path));
                    break;
            }

            ll.Export(exportConfiguration);
            FileStream stream = System.IO.File.Open(Path.Combine(Path.GetDirectoryName(exportConfiguration.Path), exportParameters.CloudFileName), FileMode.Open);
            Upload(ll, new MicrosoftOneDriveUploadParameter()
            {
                UploadStream = stream,
                CloudFileName = exportParameters.CloudFileName,
                CloudPath = exportParameters.CloudPath,
                ApplicationId = exportParameters.ApplicationId
            });
        }
    }

    /// <summary>
    /// Adapts an existing token-acquisition delegate to an <see cref="TokenCredential"/> so it can be handed to
    /// <see cref="GraphServiceClient"/>, mirroring the authentication approach used by the Microsoft Graph provider.
    /// The original token acquisition logic (MSAL / refresh token) is preserved unchanged; this only replaces the
    /// removed DelegateAuthenticationProvider plumbing from the classic Graph SDK.
    /// </summary>
    internal sealed class DelegateTokenCredential : TokenCredential
    {
        private readonly Func<CancellationToken, Task<string>> _getToken;

        internal DelegateTokenCredential(Func<CancellationToken, Task<string>> getToken)
        {
            _getToken = getToken;
        }

        public override async ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            string token = await _getToken(cancellationToken).ConfigureAwait(false);
            return new AccessToken(token, DateTimeOffset.UtcNow.AddMinutes(5));
        }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            return GetTokenAsync(requestContext, cancellationToken).AsTask().GetAwaiter().GetResult();
        }
    }

    internal class SDKHelper
    {
        // Get an authenticated Microsoft Graph Service client.
        internal static GraphServiceClient GetAuthenticatedClient(string refreshToken, string clientID, string clientSecret, string scope, string redirectUri)
        {
            var credential = new DelegateTokenCredential(
                _ => Task.FromResult(GetAuthToken(refreshToken, clientID, clientSecret, scope, redirectUri)));

            GraphServiceClient graphClient = new GraphServiceClient(credential);
            return graphClient;
        }

        internal static string GetAuthToken(string refreshToken, string clientID, string clientSecret, string scope, string redirectUri)
        {
            string url = @"https://login.microsoftonline.com/common/oauth2/v2.0/token";
            //Ticket #57341 for postdata we need a instance of FormUrlEncodedContent for HttpClient
            var parameter = new Dictionary<string, string>
            {
                { "grant_type", "refresh_token" },
                { "client_id", clientID.Trim() },
                { "client_secret", clientSecret.Trim() },
                { "refresh_token", refreshToken.Trim() },
                { "scope", "user.read files.readwrite" },
                { "redirect_uri", redirectUri },
            };
            FormUrlEncodedContent content = new FormUrlEncodedContent(parameter);

            string result = GetPostResult(url, content, "application/x-www-form-urlencoded");
            JObject o = JObject.Parse(result);
            return o["access_token"].ToString();
        }

        internal static string GetPostResult(string url, FormUrlEncodedContent content, string contentType)
        {
            HttpClient client = new HttpClient();
            HttpResponseMessage response = client.PostAsync(url, content).Result;
            return response.Content.ReadAsStringAsync().Result;
        }

    }

    internal class AuthenticationHelper
    {
        internal static string[] Scopes = { "User.Read", "Files.ReadWrite" };

        internal static string TokenForUser = null;
        internal static DateTimeOffset Expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        internal static GraphServiceClient GetAuthenticatedClient(string clientId)
        {
            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    var credential = new DelegateTokenCredential(_ => GetTokenForUserAsync(clientId));
                    graphClient = new GraphServiceClient(credential, Scopes);
                    return graphClient;
                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }


        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        internal static async Task<string> GetTokenForUserAsync(string clientId)
        {
            AuthenticationResult authResult;
            IPublicClientApplication IdentityClientApp = PublicClientApplicationBuilder.Create(clientId).Build();
            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilent(Scopes, (await IdentityClientApp.GetAccountsAsync()).First()).ExecuteAsync();
                TokenForUser = authResult.AccessToken;
            }

            catch (Exception)
            {
                if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
                    authResult = await IdentityClientApp.AcquireTokenInteractive(Scopes).ExecuteAsync();
                    TokenForUser = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;
                }
            }

            return TokenForUser;
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        internal async static void SignOut(string clientId)
        {
            IPublicClientApplication IdentityClientApp = PublicClientApplicationBuilder.Create(clientId).Build();
            foreach (var user in await IdentityClientApp.GetAccountsAsync())
            {
                await IdentityClientApp.RemoveAsync(user);
            }

            graphClient = null;
            TokenForUser = null;

        }

    }

}
