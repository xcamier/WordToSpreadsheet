using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace WordToSpreadsheet.Excel;

public class ExcelWriter
{
    private IFilterProcessor _filterProcessor;

    private int _lastRowIdx = 0;

    public ExcelWriter(IFilterProcessor filterProcessor)
    {
        _filterProcessor = filterProcessor;
    }

    public void CreateExcel(List<ExcelTab> excelTabs, string fileName)
    {
        IWorkbook workbook = new XSSFWorkbook();
        foreach (ExcelTab tab in excelTabs)
        {
            _lastRowIdx = 0;
            AddTabContent(workbook, tab);
        }
    
        FileStream stream = SaveFile(workbook, fileName);
    }

    private void AddTabContent(IWorkbook workbook, ExcelTab tab)
    {
        ISheet sheet = workbook.CreateSheet(tab.TabName);
        
        //Creation of the cell style for the headers
        ICellStyle colorStyle = workbook.CreateCellStyle();
        colorStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
        colorStyle.FillPattern = FillPattern.SolidForeground;

        foreach (WordTable table in tab.WordTables)
        {
            AddTableContent(sheet, table, colorStyle);
        }
    }


    private void AddTableContent(ISheet sheet, WordTable table, ICellStyle style)
    {    
        WordTable tableToAdd = _filterProcessor.FilterTable(table);

        for (int rowIdx = 0; rowIdx < tableToAdd.Rows.Count; rowIdx++)
        {
            int realRowIdx = rowIdx + _lastRowIdx;
            IRow row = sheet.CreateRow(realRowIdx);
            for (int idxCol = 0; idxCol < tableToAdd.Rows[rowIdx].Count(); idxCol++)
            {
                ICell cell = row.CreateCell(idxCol);
                cell.SetCellValue(tableToAdd.Rows[rowIdx][idxCol]);

                if (realRowIdx == _lastRowIdx)
                    cell.CellStyle = style;
            }   
        }  

        _lastRowIdx += tableToAdd.Rows.Count;
    }

    private static FileStream SaveFile(IWorkbook workbook, string fileName)
    {
        FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
        workbook.Write(stream);
        return stream;
    }
}
