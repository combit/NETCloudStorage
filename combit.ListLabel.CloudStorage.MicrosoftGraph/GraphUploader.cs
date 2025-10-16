using Azure.Identity;
using combit.Reporting;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;

namespace combit.ListLabel31.CloudStorage.MicrosoftGraph
{
    /// <summary>
    /// Contains various credentials used to authenticate with Microsoft Entra ID.
    /// </summary>
    public class MicrosoftCredentials
    {
        /// <summary>
        /// Your organizations tenant ID
        /// </summary>
        public string TenantId { get; set; }
        /// <summary>
        /// Application Id of your Microsoft Entra App
        /// </summary>
        public string ApplicationId { get; set; }

        /// <summary>
        /// Redirect uri of your Microsoft Entra App
        /// </summary>
        public string RedirectUri { get; set; }

        /// <summary>
        /// Required access scope
        /// </summary>
        public string Scope { get; set; }

        /// <summary>
        /// Users Microsoft accounts email adress
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// Users Microsoft account password
        /// </summary>
        public string Password { get; set; }
    }

    /// <summary>
    /// Base class for exporting files from LL, do not use this.
    /// </summary>
    public abstract class MicrosoftGraphExportParameters
    {
        /// <summary>
        /// Name of the file being uploaded
        /// </summary>
        public string CloudFileName { get; set; }

        /// <summary>
        /// Destination path in MicrosoftOneDrive and MicrosoftSharePoint. 
        /// Do not include your Drive name here, as it will generate another folder with your drives name.
        /// </summary>
        public string CloudPath { get; set; }
    }

    /// <summary>
    /// Parameters used to directly export Files from LL to MicrosoftSharePoint, using the Graph client.
    /// </summary>
    public class MicrosoftSharePointExportParameters : MicrosoftGraphExportParameters
    {
        /// <summary>
        /// ID of the target MicrosoftSharePoint drive
        /// </summary>
        public string DriveId { get; set; }
    }

    /// <summary>
    /// Parameters used to upload a file to MicrosoftSharePoint using the Graph client.
    /// </summary>
    public class MicrosoftSharePointUploadParameters : MicrosoftSharePointExportParameters
    {
        /// <summary>
        /// Content to upload
        /// </summary>
        public Stream UploadStream { get; set; }
    }

    /// <summary>
    /// Parameters used to directly export Files from LL to MicrosoftOneDrive, using the Graph client.
    /// </summary>
    public class MicrosoftOneDriveExportParameters : MicrosoftGraphExportParameters
    {
    }

    /// <summary>
    /// Parameters used to upload a file to MicrosoftOneDrive using the Graph client.
    /// </summary>
    public class MicrosoftOneDriveUploadParameters : MicrosoftOneDriveExportParameters
    {
        /// <summary>
        /// Content to upload
        /// </summary>
        public Stream UploadStream { get; set; }
    }
    internal class GraphUploader
    {
        static string[] _scopes = { "User.Read", "Files.ReadWrite" };

        /// <summary>
        /// Uploads a given file to SharePoint or OneDrive, depending on the given parameters.
        /// </summary>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <param name="sharePointUploadParameters">Parameters used to upload a file to Microsoft SharePoint</param>
        /// <param name="oneDriveUploadParameters">Parameters used to upload a file to Microsoft OneDrive</param>
        /// <returns></returns>
        internal async Task Upload(MicrosoftCredentials credentials, MicrosoftSharePointUploadParameters sharePointUploadParameters = null, MicrosoftOneDriveUploadParameters oneDriveUploadParameters = null)
        {
            var client = await GetAuthenticatedClient(credentials);
            RequestInformation requestInfo = null;

            if (client == null)
            {
                throw new InvalidOperationException("Failed to obtain an authenticated client.");
            }

            if (sharePointUploadParameters == null && oneDriveUploadParameters == null)
            {
                throw new InvalidOperationException("Missing required upload parameter instance.");
            }

            if (oneDriveUploadParameters != null)
            {
                var drive = await client.Me.Drive.GetAsync();
                requestInfo = client
                                .Drives[drive.Id]
                                .Root
                                .ItemWithPath(Path.Combine(oneDriveUploadParameters.CloudPath, oneDriveUploadParameters.CloudFileName))
                                .Content
                                .ToPutRequestInformation(oneDriveUploadParameters.UploadStream);
            }

            if (sharePointUploadParameters != null)
            {
                if (sharePointUploadParameters.CloudPath == null)
                {
                    sharePointUploadParameters.CloudPath = "";
                }
                requestInfo = client
                                .Drives[sharePointUploadParameters.DriveId]
                                .Root
                                .ItemWithPath(Path.Combine(sharePointUploadParameters.CloudPath, sharePointUploadParameters.CloudFileName))
                                .Content
                                .ToPutRequestInformation(sharePointUploadParameters.UploadStream);
            }

            var response = await client
                .RequestAdapter
                .SendAsync<DriveItem>(requestInfo, DriveItem.CreateFromDiscriminatorValue, cancellationToken: default);
        }

        /// <summary>
        /// Uploads a large file to SharePoint or OneDrive, depending on the given parameters.
        /// </summary>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <param name="sharePointUploadParameters">Parameters used to upload a file to Microsoft SharePoint</param>
        /// <param name="oneDriveUploadParameters">Parameters used to upload a file to Microsoft OneDrive</param>
        /// <returns></returns>
        internal async Task UploadLargeFile(MicrosoftCredentials credentials, MicrosoftSharePointUploadParameters sharePointUploadParameters = null, MicrosoftOneDriveUploadParameters oneDriveUploadParameters = null)
        {
            var client = await GetAuthenticatedClient(credentials);

            if (client == null)
            {
                throw new InvalidOperationException("Failed to obtain an authenticated client.");
            }

            if (sharePointUploadParameters == null && oneDriveUploadParameters == null)
            {
                throw new InvalidOperationException("Missing required upload parameter instance.");
            }

            UploadSession uploadSession = null;
            Stream stream = null;

            var uploadRequestBody = new DriveUpload.CreateUploadSessionPostRequestBody
            {
                Item = new DriveItemUploadableProperties
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" },
                    },
                },
            };

            if (oneDriveUploadParameters != null)
            {
                stream = oneDriveUploadParameters.UploadStream;
                var drive = await client.Me.Drive.GetAsync();
                uploadSession = await client
                         .Drives[drive.Id]
                         .Root
                         .ItemWithPath(Path.Combine(oneDriveUploadParameters.CloudPath, oneDriveUploadParameters.CloudFileName))
                         .CreateUploadSession
                         .PostAsync(uploadRequestBody);
            }

            if (sharePointUploadParameters != null)
            {
                stream = sharePointUploadParameters.UploadStream;
                if (sharePointUploadParameters.CloudPath == null)
                {
                    sharePointUploadParameters.CloudPath = String.Empty;
                }
                uploadSession = await client
                .Drives[sharePointUploadParameters.DriveId]
                .Root
                .ItemWithPath(Path.Combine(sharePointUploadParameters.CloudPath, sharePointUploadParameters.CloudFileName))
                .CreateUploadSession
                .PostAsync(uploadRequestBody);
            }


            int maxSliceSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.
            var fileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, stream, maxSliceSize);

            _ = await fileUploadTask.UploadAsync();
        }

        /// <summary>
        /// Gets an authenticated Microsoft Graph Service Client.
        /// </summary>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <returns>An authenticated instance of GraphServiceClient</returns>
        private static async Task<GraphServiceClient> GetAuthenticatedClient(MicrosoftCredentials credentials)
        {
            var clientCredentials = await GetCredentialInteractive(credentials);

            var graphClient = new GraphServiceClient(clientCredentials, _scopes);

            return graphClient;
        }

        /// <summary>
        /// Opens an instance of the Systems default browser to authenticate interactively.
        /// </summary>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <returns>An InteractiveBrowserCredential to use as TokenCredential</returns>
        private static async Task<InteractiveBrowserCredential> GetCredentialInteractive(MicrosoftCredentials credentials)
        {
            await Task.CompletedTask; // removes warning for lacking an await operator

            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = credentials.TenantId,
                ClientId = credentials.ApplicationId,
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                RedirectUri = new Uri(credentials.RedirectUri)
            };

            return new InteractiveBrowserCredential(options);
        }

        /// <summary>
        /// Exports the contents of the provided <see cref="ListLabel"/> instance to a file stream using the specified export configuration and Microsoft Graph export parameters.
        /// </summary>
        /// <param name="ll">The <see cref="ListLabel"/> instance that is used to configure and execute the export operation.</param>
        /// <param name="exportConfiguration">The export configuration that specifies the export target, path, and additional export options.</param>
        /// <param name="exportParameters">The Microsoft Graph export parameters containing details such as the cloud file name which is appended with the appropriate file extension.</param>
        /// <returns>
        /// A <see cref="FileStream"/> representing the exported file.
        /// </returns>
        internal static FileStream ExportToStream(ListLabel ll, ExportConfiguration exportConfiguration, MicrosoftGraphExportParameters exportParameters)
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
            return stream;
        }
    }
}
