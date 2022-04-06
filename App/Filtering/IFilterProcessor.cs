namespace WordToSpreadsheet.Filtering;

public interface IFilterProcessor
{
    WordTable FilterTable(WordTable table);
}
