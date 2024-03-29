﻿using combit.Reporting;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

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
                var user = await graphClient.Me.Request().GetAsync();
                var uploadedItem = await graphClient
                                             .Drive
                                             .Root
                                             .ItemWithPath($"{uploadParameters.CloudPath}/{uploadParameters.CloudFileName}")
                                             .Content
                                             .Request()
                                             .PutAsync<DriveItem>(uploadParameters.UploadStream);
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
            // Create the upload session. The access token is no longer required as you have session established for the upload.  
            // POST /v1.0/drive/root:/UploadLargeFile.bmp:/microsoft.graph.createUploadSession
            var uploadSession = await graphClient.Me.Drive.Root
                .ItemWithPath(fileName)
                .CreateUploadSession()
                .Request()
                .PostAsync();

            int maxSliceSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.
            var fileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

            _ = await fileUploadTask.UploadAsync();
        }

        /// <summary>
        /// Check credentials to access Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="credentials">requied parameters for Microsoft OneDrive OAuth 2.0 authentication.</param>
        public static bool CheckCredentials(MicrosoftOneDriveCredentials credentials)
        {
            // Initialize the GraphServiceClient.
            GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient(credentials.RefreshToken, credentials.ApplicationId, credentials.ApplicationSecret, credentials.Scope, credentials.RedirectUri);
            if (graphClient != null)
            {
                var user = graphClient.Me.Request().GetAsync();
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
            FileStream stream = System.IO.File.Open(string.Concat(Path.GetDirectoryName(exportConfiguration.Path), "\\", exportParameters.CloudFileName), FileMode.Open);
            Upload(ll, new MicrosoftOneDriveUploadParameter()
            {
                UploadStream = stream,
                CloudFileName = exportParameters.CloudFileName,
                CloudPath = exportParameters.CloudPath,
                ApplicationId = exportParameters.ApplicationId
            });
        }
    }

    internal class SDKHelper
    {

        // Get an authenticated Microsoft Graph Service client.
        internal static GraphServiceClient GetAuthenticatedClient(string refreshToken, string clientID, string clientSecret, string scope, string redirectUri)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        string accessToken = GetAuthToken(refreshToken, clientID, clientSecret, scope, redirectUri); //await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                        //requestMessage.Headers.Add("SampleID", "aspnet-snippets-sample");
                        return Task.FromResult(0);
                    }));
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
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        //static string clientId = App.Current.Resources["ida:ClientID"].ToString();
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
                    graphClient = new GraphServiceClient(
                        "https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                var token = await GetTokenForUserAsync(clientId);
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                            }));
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
