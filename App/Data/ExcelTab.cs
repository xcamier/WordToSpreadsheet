namespace WordToSpreadsheet.Data;

public class ExcelTab
{
    public string TabName { get; set; } = string.Empty;
    public List<WordTable> WordTables { get; set; } = new List<WordTable>();
}
