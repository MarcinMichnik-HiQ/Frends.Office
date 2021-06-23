namespace Frends.Office
{
    public interface IExportFileToSharepointInput
    {
        string clientID { get; set; }
        string clientSecret { get; set; }
        string driveID { get; set; }
        string fileName { get; }
        string siteID { get; set; }
        string sourceFilePath { get; set; }
        string targetFolderPath { get; set; }
        string tenantID { get; set; }
    }
}