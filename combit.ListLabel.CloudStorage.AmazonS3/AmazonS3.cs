using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Amazon;
using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using System.IO;
using System.Diagnostics;
using System.Threading;
using combit.Reporting.Dom;
using System.Net;

namespace combit.Reporting.CloudStorage
{
    public class S3ClientCredentials
    {
        /// <summary>
        /// Current users AWS access key
        /// </summary>
        public string AccessKey { get; set; }

        /// <summary>
        /// Current users AWS secret key
        /// </summary>
        public string SecretKey { get; set; }

        /// <summary>
        /// Current users Amazon S3 connection config
        /// </summary>
        public AmazonS3Config S3Config { get; set; }   
    }

    public class S3ExportParameters
    {
        /// <summary>
        /// Destination S3 bucket name
        /// </summary>
        public string DestinationBucketName { get; set; }

        /// <summary>
        /// Destination file name in S3
        /// </summary>
        public string CloudFileName { get; set; }

        /// <summary>      
        /// <para>Setting DisablePayloadSigning to <see langword="true"/> disables the SigV4 payload signing data integrity check for the PutObject request.</para>
        /// <para>Some Amazon S3 compatible implementation will require this option to be <see langword="true"/> since they may not yet support Streaming SigV4.</para>
        /// </summary>
        public bool DisablePayloadSigning { get; set; }
    }

    public class S3UploadParameters : S3ExportParameters
    {
        /// <summary>
        /// Content to upload
        /// </summary>
        public Stream UploadStream { get; set; }
    }

    public static class S3
    {
        /// <summary>
        /// Uploads given content to a given S3 bucket silently.
        /// To successfully upload to an S3 bucket, it is required that you have the s3:PutObject permission on the bucket.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="clientCreds">Required credentials to initialize the AmazonS3Client</param>
        /// <param name="uploadParameters">Required parameters for S3Client to upload silently</param>
        public static async Task UploadSilently(this ListLabel ll, S3ClientCredentials clientCreds, S3UploadParameters uploadParameters)
        {
            BasicAWSCredentials creds = new BasicAWSCredentials(clientCreds.AccessKey, clientCreds.SecretKey);
            using (AmazonS3Client client = new AmazonS3Client(creds, clientCreds.S3Config))
            {
                var request = new PutObjectRequest
                {
                    BucketName = uploadParameters.DestinationBucketName,
                    InputStream = uploadParameters.UploadStream,
                    DisablePayloadSigning = uploadParameters.DisablePayloadSigning,
                    Key = uploadParameters.CloudFileName
                };
                await client.PutObjectAsync(request);
            }
        }

        /// <summary>
        /// Sends a GetBucketLocation request to the given <paramref name="bucketName"/> to check server connection and credentials. 
        /// </summary>
        /// <param name="clientCreds">Currently entered user credentials</param>
        /// <param name="bucketName">The bucket to test the credentials against</param>
        /// <returns>True if connection to the server is successful; otherwise, false</returns>
        public static async Task<bool> CheckCredentials(S3ClientCredentials clientCreds, string bucketName)
        {
            BasicAWSCredentials creds = new BasicAWSCredentials(clientCreds.AccessKey, clientCreds.SecretKey);
            using (AmazonS3Client client = new AmazonS3Client(creds, clientCreds.S3Config))
            {
                try
                {
                    await client.GetBucketLocationAsync(bucketName);
                    return true;
                }
                catch (AmazonS3Exception e)
                {
                    return e.StatusCode != HttpStatusCode.Forbidden;
                }
            }
        }

        /// <summary>
        /// Export a report using current instance of ListLabel and upload it directly to an S3 bucket.
        /// To successfully upload to an S3 bucket, it is required that you have the s3:PutObject permission on a bucket.
        /// </summary>
        /// <param name="ll">Current instance of List & Label</param>
        /// <param name="exportConfiguration">Required export configuration for native ListLabel Export method</param>
        /// <param name="clientCreds">Required credentials to initialize the AmazonS3Client</param>
        /// <param name="exportParameters">Required parameters for uploading to S3</param>
        public static async Task ExportAndUpload(this ListLabel ll, ExportConfiguration exportConfiguration, S3ClientCredentials clientCreds, S3ExportParameters exportParameters)
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
            FileStream stream = File.Open(Path.Combine(Path.GetDirectoryName(exportConfiguration.Path), exportParameters.CloudFileName), FileMode.Open);
            await UploadSilently(ll, clientCreds, new S3UploadParameters()
            {
                DestinationBucketName = exportParameters.DestinationBucketName,
                UploadStream = stream,
                DisablePayloadSigning = exportParameters.DisablePayloadSigning,
                CloudFileName = exportParameters.CloudFileName
            });
        }
    }
}
