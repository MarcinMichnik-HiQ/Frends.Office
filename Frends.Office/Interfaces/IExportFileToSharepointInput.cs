namespace Frends.Office
{
    public interface IExportFileToSharepointInput
    {
        string ClientID { get; set; }
        string ClientSecret { get; set; }
        string DriveID { get; set; }
        string SiteID { get; set; }
        string SourceFilePath { get; set; }
        string TargetFolderPath { get; set; }
        string TenantID { get; set; }
    }
}