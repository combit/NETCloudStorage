using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using DecaTec.WebDav;

namespace combit.Reporting.CloudStorage
{
    public class WebDavBasicParameters
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public string ServerUrl { get; set; }
    }

    public class WebDavUploadParameters : WebDavBasicParameters
    {
        public FileStream FileStream { get; set; }
        public string DestinationFileName { get; set; }
        public FolderListItem FolderListItem { get; set; }
    }

    public class FolderListItem
    {
        public FolderListItem(string folderName)
        {
            FolderName = folderName;
        }
        public string FolderName { get; set; }
    }

    public static class WebDAV
    {
        /// <summary>
        /// Upload a file  directly to a WebDAV-Storage of your choice.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="webDavUploadParameters">required parameters for WebDAV upload.</param>
        public static async Task Upload(this Reporting.ListLabel ll, WebDavUploadParameters webDavUploadParameters)
        {

            //Create and configure NetworkCredential, and the WebDavSession.
            NetworkCredential credential = new NetworkCredential(webDavUploadParameters.Username, webDavUploadParameters.Password);
            WebDavSession session = new WebDavSession(webDavUploadParameters.ServerUrl, credential, true);
            //Create a nulled WebDavSessionItem, because otherwise there will be a compiler error.
            WebDavSessionItem destinationDirectoryItem = null;
            //Get again all folders from the Server. This is needed since the Destination-Folder can only be defined by a valid WebDavSessionItem, so we need to fetch all Items...
            IList<WebDavSessionItem> contentList = await session.ListAsync("/");

            //Select correct destinationDirectoryItem from contents
            try
            {
                destinationDirectoryItem = (from content in contentList where content.Name == webDavUploadParameters.FolderListItem.FolderName select content).Single();
            }
            catch (Exception)
            {
               //Do nothing. We check later if the Item got set successfully. If not, an exception will be thrown anyways.
            }
            //Check if getting the right item from the RemoteServer failed, and/or assigning its value to the destinationDirectoryItem. So, if it should be still null, throw an Exception.
            if (destinationDirectoryItem == null)
            {
                throw new DirectoryNotFoundException("Destination folder not found on Server, please try again.");
            }

            //Now hand the destinationDirectoryItem, the String containing the destination Filename - which simply is the SourceFilename, aswell as the FileStream containing the sourceFile over to the UploadFileAsync method of our beforehand configured WebDavSession.
            await session.UploadFileAsync(destinationDirectoryItem, webDavUploadParameters.DestinationFileName, webDavUploadParameters.FileStream);
        }

        /// <summary>
        /// Upload a file  directly to a WebDAV-Storage of your choice.
        /// Returns <see langword="true"/> if a connection was successfully established, <see langword="false"/> if not.
        /// </summary>
        /// <param name="webDavBasicParameters">required parameters for basic connection to a WebDAV Server.</param>
        public static async Task<bool> ConnectionTest(WebDavBasicParameters webDavBasicParameters)
        {

            //Create and configure NetworkCredential, and the WebDavSession.
            NetworkCredential credential = new NetworkCredential(webDavBasicParameters.Username, webDavBasicParameters.Password);
            WebDavSession session = new WebDavSession(webDavBasicParameters.ServerUrl, credential, false);
            //Get all Folders from Server.
            IList<WebDavSessionItem> contentList = await session.ListAsync("/");

            //Check if the List contains any items. If so, the Connection to the Server was successful. If not, return false, indicating failed connection.
            if (contentList.ToArray().Length < 1)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Export a report using current instance of ListLabel and upload it directly to a WebDAV-Storage of your choice.
        /// </summary>
        /// <param name="ll">current instance of List & Label</param>
        /// <param name="exportConfiguration">required export configuration for native ListLabel Export method</param>
        /// <param name="webDavUploadParameters">required parameters for WebDAV upload.</param>
        public static async Task Export(this Reporting.ListLabel ll, ExportConfiguration exportConfiguration, WebDavUploadParameters webDavUploadParameters)
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
                    webDavUploadParameters.DestinationFileName += ".pdf";
                    break;
                case LlExportTarget.Rtf:
                    webDavUploadParameters.DestinationFileName += ".rtf";
                    break;
                case LlExportTarget.Xls:
                    webDavUploadParameters.DestinationFileName += ".xls";
                    break;
                case LlExportTarget.Xlsx:
                    webDavUploadParameters.DestinationFileName += ".xlsx";
                    break;
                case LlExportTarget.Docx:
                    webDavUploadParameters.DestinationFileName += ".docx";
                    break;
                case LlExportTarget.Xps:
                    webDavUploadParameters.DestinationFileName += ".xps";
                    break;
                case LlExportTarget.Mhtml:
                    webDavUploadParameters.DestinationFileName += ".mhtml";
                    break;
                case LlExportTarget.Text:
                    webDavUploadParameters.DestinationFileName += ".txt";
                    break;
                case LlExportTarget.Pptx:
                    webDavUploadParameters.DestinationFileName += ".pptx";
                    break;
                default:
                    webDavUploadParameters.DestinationFileName += ".zip";
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportSaveAsZip, "1");
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipFile, webDavUploadParameters.DestinationFileName);
                    exportConfiguration.ExportOptions.Add(LlExportOption.ExportZipPath, Path.GetDirectoryName(exportConfiguration.Path));
                    break;
            }

            ll.Export(exportConfiguration);
            //Push the FileStream to the webDavUploadParameters
            webDavUploadParameters.FileStream = System.IO.File.Open(string.Concat(Path.GetDirectoryName(exportConfiguration.Path), "\\", webDavUploadParameters.DestinationFileName), FileMode.Open);
            await Upload(ll, webDavUploadParameters);
        }

    }
}