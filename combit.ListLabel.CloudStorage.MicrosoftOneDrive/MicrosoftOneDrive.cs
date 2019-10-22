using combit.ListLabel25;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace combit.ListLabel25.CloudStorage
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
        public static async void Upload(this ListLabel25.ListLabel ll, MicrosoftOneDriveUploadParameter uploadParameters)
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
        public static void UploadSilently(this ListLabel25.ListLabel ll, MicrosoftOneDriveUploadParameter uploadParameters, MicrosoftOneDriveCredentials credentials)
        {
            // Initialize the GraphServiceClient.
            GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient(credentials.RefreshToken, uploadParameters.ApplicationId, credentials.ApplicationSecret, credentials.Scope, credentials.RedirectUri);

            // Add the file.
            UploadLargeFile(graphClient, uploadParameters.UploadStream, string.Concat(uploadParameters.CloudPath, "/", uploadParameters.CloudFileName));
        }

        // Uploads a large file to the current user's root directory.
        private static void UploadLargeFile(GraphServiceClient graphClient, Stream fileStream, string fileName)
        {
            // Create the upload session. The access token is no longer required as you have session established for the upload.  
            // POST /v1.0/drive/root:/UploadLargeFile.bmp:/microsoft.graph.createUploadSession
            Microsoft.Graph.UploadSession uploadSession = graphClient.Me.Drive.Root.ItemWithPath(fileName).CreateUploadSession().Request().PostAsync().Result;

            int maxChunkSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.
            ChunkedUploadProvider provider = new ChunkedUploadProvider(uploadSession, graphClient, fileStream, maxChunkSize);

            // Set up the chunk request necessities.
            IEnumerable<UploadChunkRequest> chunkRequests = provider.GetUploadChunkRequests();
            byte[] readBuffer = new byte[maxChunkSize];
            List<Exception> trackedExceptions = new List<Exception>();
            DriveItem uploadedFile = null;

            // Upload the chunks.
            foreach (var request in chunkRequests)
            {
                // Do your updates here: update progress bar, etc.
                // ...
                // Send chunk request
                UploadChunkResult result = provider.GetChunkRequestResponseAsync(request, readBuffer, trackedExceptions).Result;

                if (result.UploadSucceeded)
                {
                    uploadedFile = result.ItemResponse;
                }


                // Check that upload succeeded.
                if (uploadedFile == null)
                {
                    throw new System.IO.IOException("Upload failed.");
                }
            }
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
        public static void Export(this ListLabel25.ListLabel ll, ExportConfiguration exportConfiguration, MicrosoftOneDriveExportParameter exportParameters)
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
            string data = @"grant_type=refresh_token&client_id=" + clientID.Trim() + "&client_secret=" + clientSecret.Trim() + "&refresh_token=" + refreshToken.Trim() + "&scope=user.read files.readwrite&redirect_uri=" + redirectUri;
            string result = GetPostResult(url, data, "application/x-www-form-urlencoded");
            JObject o = JObject.Parse(result);
            return o["access_token"].ToString();
        }

        internal static string GetPostResult(string url, string data, string contentType)
        {
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(new Uri(url));
            req.Method = "POST";
            req.ContentType = contentType;

            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] dataBytes = encoding.GetBytes(data);
            req.ContentLength = dataBytes.Length;
            using (Stream stream = req.GetRequestStream())
            {
                stream.Write(dataBytes, 0, dataBytes.Length);
            }

            HttpWebResponse response = (HttpWebResponse)req.GetResponse();
            Stream str = response.GetResponseStream();
            using (StreamReader reader = new StreamReader(str))
            {
                return reader.ReadToEnd();
            }
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
            PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId);

            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes, IdentityClientApp.Users.First());
                TokenForUser = authResult.AccessToken;
            }

            catch (Exception)
            {
                if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
                    authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);

                    TokenForUser = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;
                }
            }

            return TokenForUser;
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        internal static void SignOut(string clientId)
        {
            PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId);
            foreach (var user in IdentityClientApp.Users)
            {
                IdentityClientApp.Remove(user);
            }
            graphClient = null;
            TokenForUser = null;

        }

    }

}
