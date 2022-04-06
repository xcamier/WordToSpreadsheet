namespace WordToSpreadsheet.Filtering;

public class FileIO : IFileIO
{
    public bool Exists(string fileName)
    {
        return File.Exists(fileName);
    }

    public string ReadFileContent(string fileName)
    {
        if (File.Exists(fileName))
            return File.ReadAllText(fileName);
        else 
            return string.Empty;
    }

    public bool WriteFileContent(string content, string fileName)
    {
        try
        {
            File.WriteAllText(fileName, content);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Could not create the filters file: {ex.Message}");
            return false;
        }
    }
}
