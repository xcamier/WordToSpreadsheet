namespace WordToSpreadsheet.Filtering;

public interface IFileIO
{
    bool Exists(string fileName);
    string ReadFileContent(string fileName);
    bool WriteFileContent(string content, string fileName);
}
