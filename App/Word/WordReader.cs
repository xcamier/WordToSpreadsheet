using NPOI.XWPF.UserModel;

namespace WordToSpreadsheet.Word;

public class WordReader
{
    WordTableProcessing? _wtp = null;

    public WordReader(WordTableProcessing wtp)
    {
        _wtp = wtp;
    }

    public List<ExcelTab> GetTabs(string fileName, bool mergeTablesOfSameKind, bool oneKindPerTable)
    {
        List<WordTable> allTables = new List<WordTable>();

        XWPFDocument doc = OpenFile(fileName);
        if (doc != null)
        {
            IList<XWPFTable> tables =  doc.Tables;
            foreach (XWPFTable table in tables)
            {
                WordTable newTable = ReadTableContent(table);
                if (newTable.Rows.Count > 0)
                    allTables.Add(newTable);
            }
        }

        return _wtp!.PrepareExcelTabs(allTables, mergeTablesOfSameKind, oneKindPerTable);
    }

    private WordTable ReadTableContent(XWPFTable table)
    {
        WordTable newTable = new WordTable();

        foreach (XWPFTableRow row in table.Rows)
        {
            string[] newRow = BuildRow(row);
            newTable.Rows.Add(newRow);
        }

        return newTable;
    }

    private string[] BuildRow(XWPFTableRow row)
    {
        List<XWPFTableCell> cells = row.GetTableCells();
        string[] newRow = new string[cells.Count];

        for (int idx = 0; idx < cells.Count; idx ++)
        {
            XWPFTableCell cell = row.GetCell(idx);
            string value = cell.GetText();
            newRow[idx] =  value;
        }

        return newRow;
    }

    private XWPFDocument OpenFile(string fileName)
    {
        if (!File.Exists(fileName))
            throw new ArgumentException($"{fileName} not found.");

        try
        {
            using FileStream sr = File.Open(fileName, FileMode.Open);
            return new XWPFDocument(sr);
        }
        catch (Exception ex)
        {
            throw new Exception($"Unable to open the word file: {ex.Message}");
        }
    }
}
