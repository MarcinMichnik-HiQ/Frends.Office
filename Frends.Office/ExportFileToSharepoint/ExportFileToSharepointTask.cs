using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Linq;
using System.ComponentModel.DataAnnotations;

namespace Frends.Office
{
    /// <summary>
    /// Input data class for exporting files to Sharepoint.
    /// </summary>
    public class ExportFileInput
    {
        /// <summary>
        /// Full path of the target file to be written, e.g. c:\FileName.xlsx
        /// </summary>
        [DefaultValue("@\"c:\\temp\\file.xlsx\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string SourceFilePath { get; set; }

        /// <summary>
        /// Target folder path on Sharepoint.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string TargetFolderPath { get; set; }
    }

    /// <summary>
    /// Input data class for authenticating to Sharepoint.
    /// </summary>
    public class SharepointAuthentication
    {
        /// <summary>
        /// Azure Active Directory Registered APP Client ID.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string ClientID { get; set; }

        /// <summary>
        /// Azure Active Directory Registered APP Client Secret.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string ClientSecret { get; set; }

        /// <summary>
        /// Azure Active Directory tenant ID.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string TenantID { get; set; }

        /// <summary>
        /// Sharepoint Site ID - retrievable from Microsoft API once the site is created.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string SiteID { get; set; }
    }

    /// <summary>
    /// Office task for sending files to sharepoint.
    /// </summary>
    public class ExportFileToSharepointTask
    {
        /// <summary>
        /// Allows you to send files to sharepoint. Repository: https://github.com/MarcinMichnik-HiQ/Frends.Office
        /// </summary>
        /// <param name="fileExportInput"></param>
        /// <param name="authentication"></param>
        /// <returns>Returns JToken.</returns>
        public static async Task<JToken> ExportFileToSharepoint([PropertyTab] ExportFileInput fileExportInput,
            [PropertyTab] SharepointAuthentication authentication)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(authentication.ClientID)
                .WithTenantId(authentication.TenantID)
                .WithClientSecret(authentication.ClientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            // Create a new instance of GraphServiceClient with the authentication provider.
            IGraphServiceClient graphClient = new GraphServiceClient(authProvider);
            string fileLength;
            string url = "";

            // Get fileName from the sourceFilePath
            string[] sourcePathSplit = fileExportInput.SourceFilePath.Split('\\');
            string fileName = sourcePathSplit.Last();

            Drive drive = await graphClient.Sites[authentication.SiteID].Drive
                .Request()
                .GetAsync();

            try
            {
                using (FileStream fileStream = System.IO.File.OpenRead(fileExportInput.SourceFilePath))
                {
                    fileLength = fileStream.Length.ToString();
                    try
                    {
                        // Use properties to specify the conflict behavior
                        // in this case, replace
                        DriveItemUploadableProperties uploadProps = new DriveItemUploadableProperties
                        {
                            ODataType = null,
                            AdditionalData = new Dictionary<string, object>
                        {
                            { "@microsoft.graph.conflictBehavior", "replace" }
                        }
                        };

                        // Create the upload session
                        // itemPath does not need to be a path to an existing item
                        UploadSession uploadSession = await graphClient
                            .Sites[authentication.SiteID]
                            .Drives[drive.Id]
                            .Root
                            .ItemWithPath(fileExportInput.TargetFolderPath + fileName)
                            .CreateUploadSession(uploadProps)
                            .Request()
                            .PostAsync();

                        // Max slice size must be a multiple of 320 KiB
                        int maxSliceSize = 320 * 2048;
                        LargeFileUploadTask<DriveItem> fileUploadTask =
                            new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

                        // Create a callback that is invoked after each slice is uploaded
                        IProgress<long> progress = new Progress<long>();

                        url = uploadSession.UploadUrl;

                        try
                        {
                            // Upload the file
                            UploadResult<DriveItem> uploadResult = await fileUploadTask.UploadAsync(progress);
                        }
                        catch (ServiceException ex)
                        {
                            await fileUploadTask.DeleteSessionAsync();
                            throw new Exception("Unable to send file.", ex);
                        }
                    }
                    catch (ServiceException ex) {
                        throw new Exception("Unable to establish connection to Sharepoint.", ex);
                    }
                }
            }
            catch (ServiceException ex) {
                throw new Exception("Unable to open file.", ex);
            }

            JToken taskResponse = JToken.Parse("{}");
            taskResponse["FileSize"] = fileLength;
            taskResponse["Path"] = fileExportInput.SourceFilePath;
            taskResponse["FileName"] = fileName;
            taskResponse["TargetFolderName"] = fileExportInput.TargetFolderPath;
            taskResponse["ClientID"] = authentication.ClientID;
            taskResponse["TenantID"] = authentication.TenantID;
            taskResponse["SiteID"] = authentication.SiteID;
            taskResponse["DriveID"] = drive.Id;
            taskResponse["UploadUrl"] = url;

            return taskResponse;
        }
    }
}

