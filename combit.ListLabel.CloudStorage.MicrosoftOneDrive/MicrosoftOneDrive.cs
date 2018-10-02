using combit.ListLabel24;
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

namespace combit.ListLabel24.CloudStorage
{
    public static class MicrosoftOneDrive
    {
        /// <summary>
        /// Uploads given content to a file in the Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadStream">content to upload</param>
        /// <param name="cloudFileName">destination file name in Microsoft OneDrive</param>
        /// <param name="cloudPath">destination path in Microsoft OneDrive root</param>
        /// <param name="applicationId">application Id of your Microsoft App</param>
        public static async void Upload(this ListLabel24.ListLabel ll, FileStream uploadStream, string cloudFileName, string cloudPath, string applicationId)
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient(applicationId);

            if (graphClient != null)
            {
                var user = await graphClient.Me.Request().GetAsync();
                UploadLargeFile(graphClient, uploadStream, string.Concat(cloudPath, "/", cloudFileName));
            }
        }

        /// <summary>
        /// Uploads given content to a file in the Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadStream">content to upload</param>
        /// <param name="cloudFileName">destination file name in Microsoft OneDrive</param>
        /// <param name="cloudPath">destination path in Microsoft OneDrive root</param>
        /// <param name="applicationId">application Id of Microsoft app</param>
        /// <param name="applicationSecret">application secret of your Microsoft App</param>
        /// <param name=""></param>
        /// <param name=""></param>
        /// <param name=""></param>
        public static void UploadSilently(this ListLabel24.ListLabel ll, FileStream uploadStream, string cloudFileName, string cloudPath, string applicationId, string applicationSecret, string redirectUri, string scope, string refreshToken)
        {
            // Initialize the GraphServiceClient.
            GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient(refreshToken, applicationId, applicationSecret, scope, redirectUri);

            // Add the file.
            UploadLargeFile(graphClient, uploadStream, string.Concat(cloudPath, "/", cloudFileName));
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

        public static bool CheckCredentials(string applicationId, string applicationSecret, string redirectUri, string scope, string refreshToken)
        {
            // Initialize the GraphServiceClient.
            GraphServiceClient graphClient = SDKHelper.GetAuthenticatedClient(refreshToken, applicationId, applicationSecret, scope, redirectUri);
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
        /// <param name="cloudFileName">destination file name in Microsoft OneDrive</param>
        /// <param name="cloudPath">destination path in Microsoft OneDrive root</param>
        /// <param name="applicationId">application Id of your Microsoft App</param>
        public static void Export(this ListLabel24.ListLabel ll, ExportConfiguration exportConfiguration, string cloudFileName, string cloudPath, string applicationId)
        {
            ll.AutoShowSelectFile = false;
            ll.AutoShowPrintOptions = false;
            ll.AutoDestination = LlPrintMode.Export;
            ll.AutoProjectType = LlProject.List;
            ll.AutoBoxType = LlBoxType.None;
            exportConfiguration.ExportOptions.Add("Export.Quiet", "1");
            string mimeType = string.Empty;
            switch (exportConfiguration.ExportTarget)
            {
                case LlExportTarget.Pdf:
                    cloudFileName += ".pdf";
                    break;
                case LlExportTarget.Rtf:
                    cloudFileName += ".rtf";
                    break;
                case LlExportTarget.Xls:
                    cloudFileName += ".xls";
                    break;
                case LlExportTarget.Xlsx:
                    cloudFileName += ".xlsx";
                    break;
                case LlExportTarget.Docx:
                    cloudFileName += ".docx";
                    break;
                case LlExportTarget.Xps:
                    cloudFileName += ".xps";
                    break;
                case LlExportTarget.Mhtml:
                    cloudFileName += ".mhtml";
                    break;
                case LlExportTarget.Text:
                    cloudFileName += ".txt";
                    break;
                case LlExportTarget.Pptx:
                    cloudFileName += ".pptx";
                    break;
                default:
                    cloudFileName += ".zip";
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportSaveAsZip, "1");
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipFile, cloudFileName);
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipPath, Path.GetDirectoryName(exportConfiguration.Path));
                    break;
            }

            ll.Export(exportConfiguration);
            FileStream stream = System.IO.File.Open(string.Concat(Path.GetDirectoryName(exportConfiguration.Path), "\\", cloudFileName), FileMode.Open);
            Upload(ll, stream, cloudFileName, cloudPath, applicationId);            
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
