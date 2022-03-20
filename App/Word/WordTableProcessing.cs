namespace WordToSpreadsheet.Word;

public class WordTableProcessing : IWordTableProcessing
{
    public List<ExcelTab> PrepareExcelTabs(IEnumerable<WordTable> tables, bool shallMergeTablesOfSameKind, bool oneKindPerTab)
    {
        IEnumerable<WordTable> wordTablesToConvertToExcelTab = shallMergeTablesOfSameKind ? RecomposeTables(tables) : tables;

        if (oneKindPerTab)
        {
            if (shallMergeTablesOfSameKind)
            {
                return Get_Tabs_When_Content_Per_Tab_And_Tables_Grouped(wordTablesToConvertToExcelTab);
            }
            else
            {
                return Get_Tabs_When_Content_Per_Tab_And_Tables_Not_Grouped(wordTablesToConvertToExcelTab);
            }
        }
        else
        {
            if (wordTablesToConvertToExcelTab.Count() == 0)
                return new List<ExcelTab>();
            else
                return Get_Tab_When_All_Content_In_Same_Tab_And_Tables_Not_Grouped(wordTablesToConvertToExcelTab);
        }
    }

    private static List<ExcelTab> Get_Tabs_When_Content_Per_Tab_And_Tables_Not_Grouped(IEnumerable<WordTable> wordTablesToConvertToExcelTab)
    {
        List<ExcelTab> tabs = new List<ExcelTab>();
        foreach (WordTable table in wordTablesToConvertToExcelTab)
        {
            ExcelTab? tabWhereToAddTable = Find_Tab_Containing_Same_Kind_Of_Table(tabs, table);
            if (tabWhereToAddTable == null)
            {
                CreateNewTabWithTable(tabs, table);
            }
            else
            {
                tabWhereToAddTable.WordTables.Add(table);
            }
        }

        return tabs;
    }

    private static ExcelTab? Find_Tab_Containing_Same_Kind_Of_Table(List<ExcelTab> tabs, WordTable table)
    {
        foreach (ExcelTab tab in tabs)
        {
            foreach (WordTable tableFromTab in tab.WordTables)
            {
                bool isSameKind = AreHeadersTheSame(tableFromTab, table);
                if (isSameKind)
                    return tab;
            }
        }

        return null;
    }

    private static List<ExcelTab> Get_Tabs_When_Content_Per_Tab_And_Tables_Grouped(IEnumerable<WordTable> wordTablesToConvertToExcelTab)
    {
        List<ExcelTab> tabs = new List<ExcelTab>();
        foreach (WordTable table in wordTablesToConvertToExcelTab)
        {
            CreateNewTabWithTable(tabs, table);
        }

        return tabs;
    }

    private static void CreateNewTabWithTable(List<ExcelTab> tabs, WordTable table)
    {
        int tabIndex = tabs.Count + 1;
        ExcelTab newTab = new ExcelTab()
        {
            TabName = $"Tab_{tabIndex}",
            WordTables = new List<WordTable>()
                    {
                        table
                    }
        };
        tabs.Add(newTab); 
    }

    private static List<ExcelTab> Get_Tab_When_All_Content_In_Same_Tab_And_Tables_Not_Grouped(IEnumerable<WordTable> wordTablesToConvertToExcelTab)
    {
        return new List<ExcelTab>()
                {
                    new ExcelTab()
                    {
                        TabName = "Tab_1",
                        WordTables = wordTablesToConvertToExcelTab.ToList()
                    }
                };
    }

    private IEnumerable<WordTable> RecomposeTables(IEnumerable<WordTable> tables)
    {
        List<WordTable> recomposedTables = new List<WordTable>();
        foreach (WordTable table in tables)
        {
            WordTable? tableWithSameHeaders = FindTableWithSameHeaders(table, recomposedTables);
            WordTable tableCopy = table.Copy<WordTable>();
            if (tableWithSameHeaders == null)
            {
                recomposedTables.Add(tableCopy);
            }
            else
            {
                //Could not do it with Skip
                List<string[]> rowsToKeep = tableCopy.Rows.ToList();
                rowsToKeep.RemoveAt(0);
                tableWithSameHeaders.Rows.AddRange(rowsToKeep);
            }
        }

        return recomposedTables;
    }

    private WordTable? FindTableWithSameHeaders(WordTable table, List<WordTable> recomposedTables)
    {
        //Considering the first row is the header of the table
        //if the intersection of the first row of the reference and the first row of the table have the same
        //number of elements => all the elements are common => same headers
        return recomposedTables.FirstOrDefault(rct => AreHeadersTheSame(rct, table));
    }

    private static bool AreHeadersTheSame(WordTable table1, WordTable table2)
    {
        if (table1.Rows[0].Count() != table2.Rows[0].Count())
            return false;

        return table1.Rows[0].Intersect(table2.Rows[0]).Count() == table2.Rows[0].Count();
    }
}
