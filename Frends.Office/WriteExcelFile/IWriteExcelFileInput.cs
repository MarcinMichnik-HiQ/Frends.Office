using System.Data;

namespace Frends.Office
{
    public interface IWriteExcelFileInput
    {
        string csv { get; set; }
        char cellDelimiter { get; set; }
        string lineDelimiter { get; set; }
        string path { get; set; }

        DataTable CsvToDataTable();
    }
}