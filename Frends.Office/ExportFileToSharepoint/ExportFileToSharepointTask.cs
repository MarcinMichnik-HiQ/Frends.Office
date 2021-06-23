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
    /// Input class for exporting files to Sharepoint.
    /// </summary>
    public class ExportFileToSharepointInput : IExportFileToSharepointInput
    {
        /// <summary>
        /// Constructor used foor deriving fileName from sourceFilePath.
        /// </summary>
        public ExportFileToSharepointInput(string sourceFilePath, string clientID,
            string clientSecret, string tenantID, string siteID, string driveID, string targetFolderPath)
        {
            // Get fileName from the sourceFilePath
            string[] sourcePathSplit = sourceFilePath.Split('\\');
            fileName = sourcePathSplit.Last();

            this.sourceFilePath = sourceFilePath;
            this.clientID = clientID;
            this.clientSecret = clientSecret;
            this.tenantID = tenantID;
            this.siteID = siteID;
            this.driveID = driveID;
            this.targetFolderPath = targetFolderPath;
        }

        /// <summary>
        /// File name derived from sourceFilePath.
        /// </summary>
        public string fileName { get; }

        /// <summary>
        /// Full path of the target file to be written, e.g. c:\FileName.xlsx
        /// </summary>
        [DefaultValue(@"c:\temp\file.xlsx")]
        [DisplayFormat(DataFormatString = "Text")]
        public string sourceFilePath { get; set; }

        /// <summary>
        /// Azure Active Directory Registered APP Client ID
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string clientID { get; set; }

        /// <summary>
        /// Azure Active Directory Registered APP Client Secret
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string clientSecret { get; set; }

        /// <summary>
        /// Azure Active Directory tenant ID
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string tenantID { get; set; }

        /// <summary>
        /// Azure Active Directory Site ID
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string siteID { get; set; }

        /// <summary>
        /// Azure Active Directory Drive ID
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string driveID { get; set; }

        /// <summary>
        /// Target folder name on Sharepoint.
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string targetFolderPath { get; set; }
    }

    /// <summary>
    /// Office task for sending excel to sharepoint.
    /// </summary>
    public class ExportFileToSharepointTask
    {
        /// <summary>
        /// Allows you to send files to sharepoint. Repository: https://github.com/MarcinMichnik-HiQ/Frends.Office
        /// </summary>
        /// <param name="input"></param>
        /// <returns>Returns JToken.</returns>
        public static async Task<JToken> ExportFileToSharepoint(ExportFileToSharepointInput input)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(input.clientID)
                .WithTenantId(input.tenantID)
                .WithClientSecret(input.clientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            // Create a new instance of GraphServiceClient with the authentication provider.
            IGraphServiceClient graphClient = new GraphServiceClient(authProvider);
            string fileLength;
            string url = "";

            try
            {
                using (FileStream fileStream = System.IO.File.OpenRead(input.sourceFilePath))
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
                            .Sites[input.siteID]
                            .Drives[input.driveID]
                            .Root
                            .ItemWithPath(input.targetFolderPath + input.fileName)
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
            taskResponse["Path"] = input.sourceFilePath.ToString();
            taskResponse["FileName"] = input.fileName.ToString();
            taskResponse["TargetFolderName"] = input.targetFolderPath.ToString();
            taskResponse["ClientID"] = input.clientID;
            taskResponse["TenantID"] = input.tenantID.ToString();
            taskResponse["SiteID"] = input.siteID.ToString();
            taskResponse["DriveID"] = input.driveID.ToString();
            taskResponse["UploadUrl"] = url;

            return taskResponse;
        }
    }
}

