using combit.Reporting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace combit.ListLabel31.CloudStorage.MicrosoftGraph
{
    public static class MicrosoftOneDrive
    {
        /// <summary>
        /// Uploads given content to a file in the Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <param name="uploadParameters">Parameters used to upload a file to MicrosoftOneDrive</param>
        /// <returns></returns>
        public static async Task Upload(this ListLabel ll, MicrosoftCredentials credentials, MicrosoftOneDriveUploadParameters uploadParameters)
        {
            GraphUploader uploader = new GraphUploader();
            await uploader.Upload(credentials, oneDriveUploadParameters: uploadParameters);
        }

        /// <summary>
        /// Uploads a large file to the Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <param name="uploadParameters">Parameters used to upload a file to MicrosoftOneDrive</param>
        /// <returns></returns>
        public static async Task UploadSilently(this ListLabel ll, MicrosoftCredentials credentials, MicrosoftOneDriveUploadParameters uploadParameters)
        {
            GraphUploader uploader = new GraphUploader();
            await uploader.UploadLargeFile(credentials, oneDriveUploadParameters: uploadParameters);
        }

        /// <summary>
        /// Export a report using current instance of ListLabel and upload it directly to the Microsoft OneDrive Cloud Storage.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="exportConfiguration">Required export configuration for native ListLabel Export method</param>
        /// <param name="credentials">Required credentials for authenticating with Entra ID</param>
        /// <param name="exportParameters">Parameters used to directly export Files from LL to MicrosoftOneDrive</param>
        public static void Export(this ListLabel ll, ExportConfiguration exportConfiguration, MicrosoftCredentials credentials, MicrosoftOneDriveExportParameters exportParameters)
        {
            FileStream stream = GraphUploader.ExportToStream(ll, exportConfiguration, exportParameters);
            Upload(ll, credentials, new MicrosoftOneDriveUploadParameters()
            {
                UploadStream = stream,
                CloudFileName = exportParameters.CloudFileName,
                CloudPath = exportParameters.CloudPath,
            }).Wait();
        }
    }
}
