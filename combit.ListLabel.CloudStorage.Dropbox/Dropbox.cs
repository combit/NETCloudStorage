﻿using combit.ListLabel24;
using Dropbox.Api;
using Dropbox.Api.Files;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace combit.ListLabel24.CloudStorage
{
    public static class Dropbox
    {
        /// <summary>
        /// Uploads given content to a file in the Dropbox Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadStream">content to upload</param>
        /// <param name="cloudFileName">destination file name in Dropbox</param>
        /// <param name="cloudPath">destination path in Dropbox root</param>
        /// <param name="appKey">App Key of your Dropbox App</param>
        public static void Upload(this ListLabel24.ListLabel ll, FileStream uploadStream, string cloudFileName, string cloudPath, string appKey)
        {   using (var client = new DropboxClient(GetAccessToken(appKey).Result))
            {
                if (cloudPath[0] != '/')
                {
                    cloudPath = string.Concat("/", cloudPath);
                }
                CreateFolder(client, cloudPath);
                Upload(client, cloudPath, cloudFileName, uploadStream);
            }
        }

        /// <summary>
        /// Uploads given content to a file in the Dropbox Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadStream">content to upload</param>
        /// <param name="cloudFileName">destination file name in Dropbox</param>
        /// <param name="cloudPath">destination path in Dropbox root</param>
        /// <param name="accesstoken">access token for Dropbox App</param>
        public static void UploadSilently(this ListLabel24.ListLabel ll, FileStream uploadStream, string cloudFileName, string cloudPath, string accesstoken)
        {
            using (var client = new DropboxClient(accesstoken))
            {
                if (cloudPath[0] != '/')
                {
                    cloudPath = string.Concat("/", cloudPath);
                }
                CreateFolder(client, cloudPath);
                Upload(client, cloudPath, cloudFileName, uploadStream);
            }
        }

        /// <summary>
        /// Uploads given content to a file in the Dropbox Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="uploadStream">content to upload</param>
        /// <param name="cloudFileName">destination file name in Dropbox</param>
        /// <param name="cloudPath">destination path in Dropbox root</param>
        /// <param name="accesstoken">access token for Dropbox App</param>
        public static bool CheckCredentials(string accesstoken)
        {
            using (var client = new DropboxClient(accesstoken))
            {
                var user = client.Users.GetCurrentAccountAsync().Result;
                return user.Name.DisplayName != string.Empty;                
            }            
        }

        /// <summary>
        /// Export a report using current instance of ListLabel and upload it directly to the Dropbox Cloud Storage.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="exportConfiguration">required export configuration for native ListLabel Export method</param>
        /// <param name="cloudFileName">destination file name in Dropbox</param>
        /// <param name="cloudPath">destination path in Dropbox root</param>
        /// <param name="appKey">App Key of your Dropbox App</param>
        public static void Export(this ListLabel24.ListLabel ll, ExportConfiguration exportConfiguration, string cloudFileName, string cloudPath, string appKey)
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
            using (FileStream stream = System.IO.File.Open(string.Concat(Path.GetDirectoryName(exportConfiguration.Path), "\\", cloudFileName), FileMode.Open))
            {
                Upload(ll, stream, cloudFileName, cloudPath, appKey);
            }
        }

        /// <summary>
        /// Gets the Dropbox access token.
        /// <para>
        /// This fetches the access token from the applications settings, if it is not found there
        /// (or if the user chooses to reset the settings) then the UI in <see cref="LoginForm"/> is
        /// displayed to authorize the user.
        /// </para>
        /// </summary>
        /// <returns>A valid access token or null.</returns>
        private static async Task<string> GetAccessToken(string appKey)
        {
            var accessToken = string.Empty;

            var completion = new TaskCompletionSource<Tuple<string, string>>();

            var thread = new Thread(() =>
            {
                try
                {
                    var app = new Application();
                    var login = new LoginForm(appKey);
                    app.Run(login);
                    if (login.Result)
                    {
                        completion.TrySetResult(Tuple.Create(login.AccessToken, login.Uid));
                    }
                    else
                    {
                        completion.TrySetCanceled();
                    }
                }
                catch (Exception e)
                {
                    completion.TrySetException(e);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            var result = await completion.Task;

            accessToken = result.Item1;
            var uid = result.Item2;            

            return accessToken;
        }

        /// <summary>
        /// Creates the specified folder.
        /// </summary>
        /// <remarks>This demonstrates calling an rpc style api in the Files namespace.</remarks>
        /// <param name="path">The path of the folder to create.</param>
        /// <param name="client">The Dropbox client.</param>
        /// <returns>The result from the ListFolderAsync call.</returns>
        private static FolderMetadata CreateFolder(DropboxClient client, string path)
        {
            FolderMetadata folder = null;
            try
            {
                folder = client.Files.GetMetadataAsync(path).Result as FolderMetadata;
            }
            catch (Exception ex)
            {
                if (ex.InnerException is ApiException<GetMetadataError> && ((ex.InnerException as ApiException<GetMetadataError>).ErrorResponse as GetMetadataError.Path).Value is LookupError.NotFound)
                {
                    var folderArg = new CreateFolderArg(path);
                    folder = client.Files.CreateFolderV2Async(folderArg).Result.Metadata;
                }
                else
                {
                    throw;
                }
            }
            return folder;
        }

        /// <summary>
        /// Uploads given content to a file in Dropbox.
        /// </summary>
        /// <param name="client">The Dropbox client.</param>
        /// <param name="folder">The folder to upload the file.</param>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="fileContent">The file content.</param>
        /// <returns></returns>
        private static void Upload(DropboxClient client, string folder, string fileName, FileStream stream)
        {
            using (stream)
            {
                client.Files.UploadAsync(folder + "/" + fileName, WriteMode.Overwrite.Instance, body: stream).Wait();
            }
        }

        /// <summary>
        /// Uploads a big file in chunk. The is very helpful for uploading large file in slow network condition
        /// and also enable capability to track upload progerss.
        /// </summary>
        /// <param name="client">The Dropbox client.</param>
        /// <param name="folder">The folder to upload the file.</param>
        /// <param name="fileName">The name of the file.</param>
        /// <returns></returns>
        private static async Task ChunkUpload(DropboxClient client, string folder, string fileName, FileStream stream)
        {
            // Chunk size is 128KB.
            const int chunkSize = 128 * 1024;

            using (stream)
            {
                int numChunks = (int)Math.Ceiling((double)stream.Length / chunkSize);

                byte[] buffer = new byte[chunkSize];
                string sessionId = null;

                for (var idx = 0; idx < numChunks; idx++)
                {
                    var byteRead = stream.Read(buffer, 0, chunkSize);

                    using (MemoryStream memStream = new MemoryStream(buffer, 0, byteRead))
                    {
                        if (idx == 0)
                        {
                            var result = await client.Files.UploadSessionStartAsync(body: memStream);
                            sessionId = result.SessionId;
                        }

                        else
                        {
                            UploadSessionCursor cursor = new UploadSessionCursor(sessionId, (ulong)(chunkSize * idx));

                            if (idx == numChunks - 1)
                            {
                                await client.Files.UploadSessionFinishAsync(cursor, new CommitInfo(folder + "/" + fileName), memStream);
                            }

                            else
                            {
                                await client.Files.UploadSessionAppendV2Async(cursor, body: memStream);
                            }
                        }
                    }
                }
            }
        }
    }
}