namespace UnitTests;
 
public class WordTableProcessingUnitTests
{
    [Fact]
    public void No_Word_Table_No_Recomposition_Only_One_Tab()
    {
        //Arrange
        List<WordTable> tables = new List<WordTable>();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, false, false);

        //Assert
        Assert.Empty(excelTabs);
    }

    [Fact]
    public void No_Word_Table_With_Recomposition_Only_One_Tab()
    {
        //Arrange
        List<WordTable> tables = new List<WordTable>();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, true, false);

        //Assert
        Assert.Empty(excelTabs);
    }

    [Fact]
    public void No_Word_Table_No_Recomposition_One_Tab_Per_Kind()
    {
        //Arrange
        List<WordTable> tables = new List<WordTable>();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, false, true);

        //Assert
        Assert.Empty(excelTabs);
    }

    [Fact]
    public void No_Word_Table_With_Recomposition_One_Tab_Per_Kind()
    {
        //Arrange
        List<WordTable> tables = new List<WordTable>();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, true, true);

        //Assert
        Assert.Empty(excelTabs);
    }

    [Fact]
    public void Word_Table_Only_Headers_No_Recomposition_Only_One_Tab()
    {
        //Arrange
        List<WordTable> tables = GetOnlyHeadersDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, false, false);

        //Assert
        Assert.Single(excelTabs);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal(tables.Count, excelTabs[0].WordTables.Count);
    }

    [Fact]
    public void Word_Table_Only_Headers_With_Recomposition_Only_One_Tab()
    {
        //Arrange
        List<WordTable> tables = GetOnlyHeadersDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, true, false);

        //Assert
        Assert.Single(excelTabs);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal(4, excelTabs[0].WordTables.Count);
    }

    [Fact]
    public void Word_Table_Only_Headers_No_Recomposition_One_Tab_per_Kind()
    {
        //Arrange
        List<WordTable> tables = GetOnlyHeadersDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, false, true);

        //Assert
        Assert.Equal(4, excelTabs.Count);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal("Tab_2", excelTabs[1].TabName);
        Assert.Equal("Tab_3", excelTabs[2].TabName);
        Assert.Equal("Tab_4", excelTabs[3].TabName);
        Assert.Equal(2, excelTabs[0].WordTables.Count);
        Assert.Equal(3, excelTabs[1].WordTables.Count);
        Assert.Single(excelTabs[2].WordTables);
        Assert.Single(excelTabs[3].WordTables);
    }



    [Fact]
    public void Word_Table_Only_Headers_With_Recomposition_One_Tab_per_Kind()
    {
        //Arrange
        List<WordTable> tables = GetOnlyHeadersDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, true, true);

        //Assert
        Assert.Equal(4, excelTabs.Count);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal("Tab_2", excelTabs[1].TabName);
        Assert.Equal("Tab_3", excelTabs[2].TabName);
        Assert.Equal("Tab_4", excelTabs[3].TabName);
        Assert.Single(excelTabs[0].WordTables);
        Assert.Single(excelTabs[1].WordTables);
        Assert.Single(excelTabs[2].WordTables);
        Assert.Single(excelTabs[3].WordTables);
    }

    [Fact]
    public void Word_Table_No_Recomposition_Only_One_Tab()
    {
        //Arrange
        List<WordTable> tables = GetFullDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, false, false);

        //Assert
        Assert.Single(excelTabs);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal(tables.Count, excelTabs[0].WordTables.Count);
        
        //The data rows + the headers
        int total = tables.SelectMany(tbl => tbl.Rows).Count();
        int sec = excelTabs[0].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(total, sec);
    }

    [Fact]
    public void Word_Table_With_Recomposition_Only_One_Tab()
    {
        //Arrange
        List<WordTable> tables = GetFullDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, true, false);

        //Assert
        Assert.Single(excelTabs);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal(4, excelTabs[0].WordTables.Count);

        //The data rows + the headers
        int total = tables.SelectMany(tbl => tbl.Rows).Count();
        //The regrouped data rows + duplicated headers (3 headers were removed  of 3)
        int sec = excelTabs[0].WordTables.SelectMany(tbl => tbl.Rows).Count() + 3;
        Assert.Equal(total, sec);
    }

    [Fact]
    public void Word_Table_No_Recomposition_One_Tab_Per_Kind()
    {
        //Arrange
        List<WordTable> tables = GetFullDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, false, true);

        //Assert
        Assert.Equal(4, excelTabs.Count);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal("Tab_2", excelTabs[1].TabName);
        Assert.Equal("Tab_3", excelTabs[2].TabName);
        Assert.Equal("Tab_4", excelTabs[3].TabName);
        Assert.Equal(2, excelTabs[0].WordTables.Count);
        Assert.Equal(3, excelTabs[1].WordTables.Count);
        Assert.Single(excelTabs[2].WordTables);
        Assert.Single(excelTabs[3].WordTables);
        
        int sec = excelTabs[0].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(5, sec);

        sec = excelTabs[1].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(9, sec);

        sec = excelTabs[2].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(4, sec);

        sec = excelTabs[3].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(2, sec);
    }

    [Fact]
    public void Word_Table_With_Recomposition_One_Tab_Per_Kind()
    {
        //Arrange
        List<WordTable> tables = GetFullDataSet();
        IWordTableProcessing wtp = new WordTableProcessing();

        //Act
        List<ExcelTab> excelTabs = wtp.PrepareExcelTabs(tables, true, true);

        //Assert
        Assert.Equal(4, excelTabs.Count);
        Assert.Equal("Tab_1", excelTabs[0].TabName);
        Assert.Equal("Tab_2", excelTabs[1].TabName);
        Assert.Equal("Tab_3", excelTabs[2].TabName);
        Assert.Equal("Tab_4", excelTabs[3].TabName);
        Assert.Single(excelTabs[0].WordTables);
        Assert.Single(excelTabs[1].WordTables);
        Assert.Single(excelTabs[2].WordTables);
        Assert.Single(excelTabs[3].WordTables);

        //The data rows + the headers
        int sec = excelTabs[0].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(4, sec);

        sec = excelTabs[1].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(7, sec);

        sec = excelTabs[2].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(4, sec);

        sec = excelTabs[3].WordTables.SelectMany(tbl => tbl.Rows).Count();
        Assert.Equal(2, sec);
    }

    private List<WordTable> GetOnlyHeadersDataSet()
    {
        string[] table1Row0 = new string[] {"A", "B", "C", "D" };
        string[] table2Row0 = new string[] {"AA", "BB", "CC", "DD" };
        string[] table3Row0 = new string[] {"A", "B", "C", "D" };
        string[] table4Row0 = new string[] {"AA", "BB", "CC", "DD" };
        string[] table5Row0 = new string[] {"AA", "BB", "CC", "DD" };
        string[] table6Row0 = new string[] {"E", "F", "G", "H" };
        string[] table7Row0 = new string[] {"A", "B" };

        return new List<WordTable>() 
        {
            CreateTableFromRow(table1Row0),
            CreateTableFromRow(table2Row0),
            CreateTableFromRow(table3Row0),
            CreateTableFromRow(table4Row0),
            CreateTableFromRow(table5Row0),
            CreateTableFromRow(table6Row0),
            CreateTableFromRow(table7Row0)
        };
    }

    private WordTable CreateTableFromRow(string[] row)
    {
        WordTable table = new WordTable();
        table.Rows.Add(row);

        return table;
    }

    private List<WordTable> GetFullDataSet()
    {
        string[] variousData1 = new string[] {"Papa", "Maman", "Soeur", "Frere" };
        string[] variousData2 = new string[] {"Tonton", "Tata", "Parrain", "Marraine" };
        string[] variousData3 = new string[] {"Renault", "Peugeot", "Alpine", "Citroen" };
        string[] variousData4 = new string[] {"Ferrari", "Porsche", "Mazeratti", "Aston Martin" };
        string[] variousData5 = new string[] {"Apple", "HP", "Dell", "Lenovo" };
        string[] variousData6 = new string[] {"Google", "Xiaomi", "Honor", "H" };
        string[] variousData7 = new string[] {"Ikea", "Conforama", "But", "Monsieur Meubles" };
        string[] variousData8 = new string[] {"Nike", "Puma", "Adidas", "Reebok" };
        string[] variousData9 = new string[] {"Alonso", "Ricciardo", "Hamilton", "Gasly" };
        string[] variousData10 = new string[] {"Sonos", "Bose", "B&W", "Warfedale" };
        string[] variousData11 = new string[] {"Evian", "Vittel", "Saint Amand", "Badoit" };
        string[] variousData12 = new string[] {"Coca-Cola", "Orangina", "Schweppes", "Pulco" };
        string[] variousData13 = new string[] {"You", "Me" };

        List<WordTable> data = GetOnlyHeadersDataSet();
        data[0].Rows.Add(variousData1);
        data[0].Rows.Add(variousData2);
        data[1].Rows.Add(variousData3);
        data[2].Rows.Add(variousData4);
        data[3].Rows.Add(variousData5);
        data[3].Rows.Add(variousData6);
        data[3].Rows.Add(variousData7);
        data[4].Rows.Add(variousData8);
        data[4].Rows.Add(variousData9);
        data[5].Rows.Add(variousData10);
        data[5].Rows.Add(variousData11);
        data[5].Rows.Add(variousData12);
        data[6].Rows.Add(variousData13);

        //*** Without merging ***
        //"A", "B", "C", "D"
        //"Papa", "Maman", "Soeur", "Frere"
        //"Tonton", "Tata", "Parrain", "Marraine"

        //"AA", "BB", "CC", "DD" 
        //"Renault", "Peugeot", "Alpine", "Citroen" 

        //"A", "B", "C", "D" 
        //"Ferrari", "Porsche", "Mazeratti", "Aston Martin" 

        //"AA", "BB", "CC", "DD"
        //"Apple", "HP", "Dell", "Lenovo" 
        //"Google", "Xiaomi", "Honor", "H" 
        //"Ikea", "Conforama", "But", "Monsieur Meubles" 

        //"AA", "BB", "CC", "DD"
        //"Nike", "Puma", "Adidas", "Reebok" 
        //"Alonso", "Ricciardo", "Hamilton", "Gasly" 
        
        //"E", "F", "G", "H" 
        //"Sonos", "Bose", "B&W", "Warfedale" 
        //"Evian", "Vittel", "Saint Amand", "Badoit" 
        //"Coca-Cola", "Orangina", "Schweppes", "Pulco" 

        //"A", "B"
        //"You", "Me" 


        // *** With merging ***
        //"A", "B", "C", "D"
        //"Papa", "Maman", "Soeur", "Frere"
        //"Tonton", "Tata", "Parrain", "Marraine"
        //"Ferrari", "Porsche", "Mazeratti", "Aston Martin" 
        
        //"AA", "BB", "CC", "DD" 
        //"Renault", "Peugeot", "Alpine", "Citroen" 
        //"Apple", "HP", "Dell", "Lenovo" 
        //"Google", "Xiaomi", "Honor", "H" 
        //"Ikea", "Conforama", "But", "Monsieur Meubles" 
        //"Nike", "Puma", "Adidas", "Reebok" 
        //"Alonso", "Ricciardo", "Hamilton", "Gasly" 

        //"E", "F", "G", "H" 
        //"Sonos", "Bose", "B&W", "Warfedale" 
        //"Evian", "Vittel", "Saint Amand", "Badoit" 
        //"Coca-Cola", "Orangina", "Schweppes", "Pulco" 

        //"A", "B"
        //"You", "Me" 


        return data;
    }
}