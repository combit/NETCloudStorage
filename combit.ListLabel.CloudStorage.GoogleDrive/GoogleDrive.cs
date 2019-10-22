using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Upload;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static Google.Apis.Drive.v3.FilesResource;

namespace combit.ListLabel25.CloudStorage
{
    public class GoogleDriveExportParameter
    {
        /// <summary>
        /// Destination file name in Google Drive.
        /// </summary>
        public string CloudFileName { get; set; }

        /// <summary>
        /// Destination path in Google Drive root.
        /// </summary>
        public string CloudPath { get; set; }

        /// <summary>
        /// Application name of your Google App.
        /// </summary>
        public string ApplicationName { get; set; }
    }

    public class GoogleDriveUploadParameter : GoogleDriveExportParameter
    {
        /// <summary>
        /// Content to upload.
        /// </summary>
        public FileStream UploadStream { get; set; }

        /// <summary>
        /// MIME-Type of the file.
        /// </summary>
        public string MimeType { get; set; }
    }

    public class GoogleDriveCredentials
    {
        /// <summary>
        /// Application name of your Google App.
        /// </summary>
        public string ApplicationName { get; set; }

        /// <summary>
        /// Application secret of your Google App.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Application secret of your Google App.
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// The current valid refresh token.
        /// </summary>
        public string RefreshToken { get; set; }
    }

    public static class GoogleDrive
    {
        static string[] Scopes = { DriveService.Scope.Drive };

        /// <summary>
        /// Uploads given content to a file in the Google Drive Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadParameters">requied parameters for Google Drive OAuth 2.0 upload.</param>
        /// <param name="clientSecretFilePath">The full path of client secret file.</param>
        public static void Upload(this ListLabel25.ListLabel ll, GoogleDriveUploadParameter uploadParameters, string clientSecretFilePath)
        {
            UserCredential credential;

            using (var stream =
                new FileStream(clientSecretFilePath, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user2",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                Upload2GoogleDrive(credential, uploadParameters.UploadStream, uploadParameters.CloudFileName, uploadParameters.CloudPath, uploadParameters.ApplicationName, uploadParameters.MimeType);

            }
        }

        /// <summary>
        /// Uploads given content to a file in the Google Drive Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadParameters">requied parameters for Google Drive OAuth 2.0 upload silently.</param>
        /// <param name="credentials">requied parameters for Google Drive OAuth 2.0 authentication.</param>
        public static void UploadSilently(this ListLabel25.ListLabel ll, GoogleDriveUploadParameter uploadParameters, GoogleDriveCredentials credentials)
        {
            var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
            {
                ClientSecrets = new ClientSecrets
                {
                    ClientId = credentials.ClientId,
                    ClientSecret = credentials.ClientSecret
                },
                Scopes = Scopes
            });

            var accessToken = ListLabel25.DataProviders.GoogleDataProviderHelper.GetAuthToken(credentials.RefreshToken, credentials.ClientId, credentials.ClientSecret);

            var token = new TokenResponse
            {
                AccessToken = accessToken,
                RefreshToken = credentials.RefreshToken
            };

            var credential = new UserCredential(flow, Environment.UserName, token);

            Upload2GoogleDrive(credential, uploadParameters.UploadStream, uploadParameters.CloudFileName, uploadParameters.CloudPath, uploadParameters.ApplicationName, uploadParameters.MimeType);

        }

        private static void Upload2GoogleDrive(UserCredential credential, FileStream uploadStream, string fileName, string path, string applicationName, string mimeType)
        {
            // Create the service using the client credentials.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            });

            string parentId = "root";

            if (path != string.Empty)
            {
                string[] folders = path.Split('/');
                foreach (var folder in folders)
                {
                    var request = service.Files.List();
                    request.Q = string.Format("'{0}' IN parents and name='{1}' and trashed=false and mimeType='application/vnd.google-apps.folder'", parentId, folder);
                    request.Spaces = "drive";
                    request.Fields = "nextPageToken, files(id, name)";
                    var result = request.Execute();
                    if (result.Files.Count == 0)
                    {
                        var fileMetadata = new Google.Apis.Drive.v3.Data.File();
                        fileMetadata.Name = folder;
                        fileMetadata.MimeType = "application/vnd.google-apps.folder";
                        fileMetadata.Parents = new List<string>() { parentId };
                        var createFolderRequest = service.Files.Create(fileMetadata);
                        createFolderRequest.Fields = "id";
                        var file = createFolderRequest.Execute();
                        parentId = file.Id;
                    }
                    else
                    {
                        parentId = result.Files[0].Id;
                    }
                }
            }

            // Get the media upload request object.
            CreateMediaUpload insertRequest = service.Files.Create(
                new Google.Apis.Drive.v3.Data.File
                {
                    Name = fileName,
                    Parents = new List<string>() { parentId }
                },
                uploadStream,
                mimeType);

            // Add handlers which will be notified on progress changes and upload completion.
            // Notification of progress changed will be invoked when the upload was started,
            // on each upload chunk, and on success or failure.
            insertRequest.ProgressChanged += Upload_ProgressChanged;
            insertRequest.ResponseReceived += Upload_ResponseReceived;

            insertRequest.Upload();
        }

        /// <summary>
        /// Check credentials to access Google Drive Cloud Storage.
        /// </summary>
        /// <param name="credentials">requied parameters for Google Drive OAuth 2.0 authentication.</param>
        public static bool CheckCredentials(GoogleDriveCredentials credentials)
        {
            var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
            {
                ClientSecrets = new ClientSecrets
                {
                    ClientId = credentials.ClientId,
                    ClientSecret = credentials.ClientSecret
                },
                Scopes = Scopes
            });

            var accessToken = ListLabel25.DataProviders.GoogleDataProviderHelper.GetAuthToken(credentials.RefreshToken, credentials.ClientId, credentials.ClientSecret);

            var token = new TokenResponse
            {
                AccessToken = accessToken,
                RefreshToken = credentials.RefreshToken
            };

            var credential = new UserCredential(flow, Environment.UserName, token);

            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = credentials.ApplicationName
            });
            var aboutRequest = service.About.Get();
            aboutRequest.Fields = "user";
            var about = aboutRequest.Execute();
            return about.User.DisplayName != string.Empty;

        }

        static void Upload_ProgressChanged(IUploadProgress progress)
        {
            Console.WriteLine(progress.Status + " " + progress.BytesSent);
        }

        static void Upload_ResponseReceived(Google.Apis.Drive.v3.Data.File file)
        {
            Console.WriteLine(file.Name + " was uploaded successfully");
        }

        /// <summary>
        /// Export a report using current instance of ListLabel and upload it directly to the Google Drive Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="exportConfiguration">required export configuration for native ListLabel Export method</param>
        /// <param name="exportParameters">requied parameters for Google Drive OAuth 2.0 upload.</param>
        /// <param name="clientSecretFilePath">The full path of client secret file.</param>
        public static void Export(this ListLabel25.ListLabel ll, ExportConfiguration exportConfiguration, GoogleDriveExportParameter exportParameters, string clientSecretFilePath)
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
                    exportParameters.CloudFileName += ".pdf";
                    mimeType = "application/pdf";
                    break;
                case LlExportTarget.Rtf:
                    exportParameters.CloudFileName += ".rtf";
                    mimeType = "application/rtf";
                    break;
                case LlExportTarget.Xls:
                    exportParameters.CloudFileName += ".xls";
                    mimeType = "application/vnd.ms-excel";
                    break;
                case LlExportTarget.Xlsx:
                    exportParameters.CloudFileName += ".xlsx";
                    mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    break;
                case LlExportTarget.Docx:
                    exportParameters.CloudFileName += ".docx";
                    mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    break;
                case LlExportTarget.Xps:
                    exportParameters.CloudFileName += ".xps";
                    mimeType = "application/vnd.ms-xpsdocument";
                    break;
                case LlExportTarget.Mhtml:
                    exportParameters.CloudFileName += ".mhtml";
                    mimeType = "message/rfc822";
                    break;
                case LlExportTarget.Text:
                    exportParameters.CloudFileName += ".txt";
                    mimeType = "text/plain";
                    break;
                case LlExportTarget.Pptx:
                    exportParameters.CloudFileName += ".pptx";
                    mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                    break;
                default:
                    exportParameters.CloudFileName += ".zip";
                    mimeType = "application/zip";
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportSaveAsZip, "1");
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipFile, exportParameters.CloudFileName);
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipPath, Path.GetDirectoryName(exportConfiguration.Path));
                    break;
            }

            ll.Export(exportConfiguration);
            FileStream stream = System.IO.File.Open(string.Concat(Path.GetDirectoryName(exportConfiguration.Path), "\\", exportParameters.CloudFileName), FileMode.Open);
            Upload(ll, new GoogleDriveUploadParameter()
            {
                UploadStream = stream,
                CloudFileName = exportParameters.CloudFileName,
                CloudPath = exportParameters.CloudPath,
                ApplicationName = exportParameters.ApplicationName,
                MimeType = mimeType
            }, clientSecretFilePath);
        }
    }
}
