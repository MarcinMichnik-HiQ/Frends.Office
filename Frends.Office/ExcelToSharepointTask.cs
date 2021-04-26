using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using System.Threading.Tasks;


namespace Frends.Office
{
    /// <summary>
    /// Input for excel to sharepoint.
    /// </summary>
    /// 
    public class InputExcelSharepoint
    {
        /// <summary>
        /// Full path of the target file to be written, e.g. FileName.xlsx
        /// </summary>
        [DefaultValue(@"c:\temp\file.xlsx")]
        public string path { get; set; }

        /// <summary>
        /// The name of the file, e.g. FileName.xlsx
        /// </summary>
        [DefaultValue(@"file.xlsx")]
        public string fileName { get; set; }

        /// <summary>
        /// Azure AD Registered APP Client ID
        /// </summary>
        [DefaultValue("")]
        public string clientID { get; set; }

        /// <summary>
        /// Azure AD tenant ID
        /// </summary>
        [DefaultValue("")]
        public string tenantID { get; set; }

        /// <summary>
        /// Azure AD Registered APP Client Secret
        /// </summary>
        [DefaultValue("")]
        public string clientSecret { get; set; }

        /// <summary>
        /// Azure AD Site ID
        /// </summary>
        [DefaultValue("")]
        public string siteID { get; set; }

        /// <summary>
        /// Azure AD Drive ID
        /// </summary>
        [DefaultValue("")]
        public string driveID { get; set; }

        /// <summary>
        /// Target folder path
        /// </summary>
        [DefaultValue("")]
        public string targetFolderName { get; set; }

    }

    /// <summary>
    /// Office task for sending excel to sharepoint.
    /// </summary>
    /// 
    public class ExcelToSharepointTask
    {
       
        /// <summary>
        /// Allows you to send excel files to sharepoint. https://github.com/MarcinMichnik-HiQ/Frends.Office
        /// </summary>
        /// <param name="inputExcelSharepoint"></param>
        /// <returns>Returns true if the file was written to correctly Otherwise throws an exception</returns>
        public static async Task<bool> ExcelToSharepoint(InputExcelSharepoint inputExcelSharepoint)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(inputExcelSharepoint.clientID)
                .WithTenantId(inputExcelSharepoint.tenantID)
                .WithClientSecret(inputExcelSharepoint.clientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            // Create a new instance of GraphServiceClient with the authentication provider.
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            using (var fileStream = System.IO.File.OpenRead(inputExcelSharepoint.path))
            {
                // Use properties to specify the conflict behavior
                // in this case, replace
                var uploadProps = new DriveItemUploadableProperties
                {
                    ODataType = null,
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" }
                    }
                };

                // Create the upload session
                // itemPath does not need to be a path to an existing item
                var uploadSession = await graphClient
                    .Sites[inputExcelSharepoint.siteID]
                    .Drives[inputExcelSharepoint.driveID]
                    .Root
                    .ItemWithPath(inputExcelSharepoint.targetFolderName + inputExcelSharepoint.fileName)
                    .CreateUploadSession(uploadProps)
                    .Request()
                    .PostAsync();

                // Max slice size must be a multiple of 320 KiB
                int maxSliceSize = 320 * 1024;
                var fileUploadTask =
                    new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Console.WriteLine($"Uploaded {prog} bytes of {fileStream.Length} bytes");
                });

                try
                {
                    // Upload the file
                    var uploadResult = await fileUploadTask.UploadAsync(progress);

                }
                catch (ServiceException ex)
                {
                    throw new Exception("Unable to send file.", ex);
                }
                return true;
            }
        }
    }
}
