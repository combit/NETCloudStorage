using combit.Reporting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace combit.ListLabel31.CloudStorage.MicrosoftGraph
{
    public static class MicrosoftSharePoint
    {
        /// <summary>
        /// Uploads given content to a file in the Microsoft SharePoint Cloud Storage using the GraphUploader.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="creds">Required credentials for authenticating with Entra ID</param>
        /// <param name="uploadParams">Parameters used to upload a file to MicrosoftSharePoint</param>
        /// <returns></returns>
        public static async Task Upload(this ListLabel ll, MicrosoftCredentials creds, MicrosoftSharePointUploadParameters uploadParams)
        {
            GraphUploader uploader = new GraphUploader();
            await uploader.Upload(creds, sharePointUploadParameters: uploadParams);
        }

        /// <summary>
        /// Uploads a large file to the Microsoft SharePoint Cloud Storage using the GraphUploader.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="creds">Required credentials for authenticating with Entra ID</param>
        /// <param name="uploadParams">Parameters used to upload a file to MicrosoftSharePoint</param>
        /// <returns></returns>
        public static async Task UploadSilently(this ListLabel ll, MicrosoftCredentials creds, MicrosoftSharePointUploadParameters uploadParams)
        {
            GraphUploader uploader = new GraphUploader();
            await uploader.UploadLargeFile(creds, sharePointUploadParameters: uploadParams);
        }

        /// <summary>
        /// Export a report using current instance of ListLabel and upload it directly to the Microsoft SharePoint Cloud Storage.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="exportConfiguration">Required export configuration for native ListLabel Export method</param>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <param name="exportParameters">Parameters used to directly export Files from LL to MicrosoftSharePoint</param>
        public static void Export(this ListLabel ll, ExportConfiguration exportConfiguration, MicrosoftCredentials credentials, MicrosoftSharePointExportParameters exportParameters)
        {
            FileStream stream = GraphUploader.ExportToStream(ll, exportConfiguration, exportParameters);
            Upload(ll, credentials, new MicrosoftSharePointUploadParameters()
            {
                UploadStream = stream,
                CloudFileName = exportParameters.CloudFileName,
                CloudPath = exportParameters.CloudPath,
                DriveId = exportParameters.DriveId
            }).Wait();
        }
    }
}
