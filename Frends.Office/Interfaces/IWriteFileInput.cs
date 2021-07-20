using System.Data;

namespace Frends.Office
{
    public interface IWriteFileInput
    {
        string StringInput { get; set; }
        string TargetPath { get; set; }
        string LineDelimiter { get; set; }
    }
}