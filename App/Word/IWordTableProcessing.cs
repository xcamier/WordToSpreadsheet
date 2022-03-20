namespace WordToSpreadsheet.Word;

public interface IWordTableProcessing
{
    List<ExcelTab> PrepareExcelTabs(IEnumerable<WordTable> tables, bool shallMergeTablesOfSameKind, bool oneKindPerTab);
}
