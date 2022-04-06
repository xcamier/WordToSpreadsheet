using FluentAssertions;
using Moq;

namespace UnitTests
{
    public class FilterProcessorUnitTests
    {
        [Fact]
        public void No_filter()
        {
            //Arrange
            WordTable table = GetWordTable();

            var filterReader = new Mock<IFilterReader>();
            filterReader.Setup(x => x.ReadFilters()).Returns(new FilterOptions());

            IFilterProcessor fp = new FilterProcessor(filterReader.Object);
            
            //Act
            WordTable filteredTable = fp.FilterTable(table);

            //Assert
            table.Should().BeEquivalentTo(filteredTable);
        }

        [Fact]
        public void One_Filter_Keep_UnkownColumn()
        {
            //Arrange
            WordTable table = GetWordTable();

            var filterReader = new Mock<IFilterReader>();
            FilterOptions opt = new FilterOptions();
            FilterOption filter = new FilterOption()
            {
                FilterType = FilterType.KeepWhenContains,
                ColumnName = "UnknownCol",
                FilterValue = "0"

            };
            opt.Filters.Add(filter);
            filterReader.Setup(x => x.ReadFilters()).Returns(opt);

            IFilterProcessor fp = new FilterProcessor(filterReader.Object);

            //Act
            WordTable filteredTable = fp.FilterTable(table);

            //Assert
            table.Should().BeEquivalentTo(filteredTable);
        }    

        [Fact]
        public void One_Filter_KownColumn_Keep_Contains_But_Nothing_To_Filter_Found()
        {
            //Arrange
            WordTable table = GetWordTable();

            var filterReader = new Mock<IFilterReader>();
            FilterOptions opt = new FilterOptions();
            FilterOption filter = new FilterOption()
            {
                FilterType = FilterType.KeepWhenContains,
                ColumnName = "Reference",
                FilterValue = "00"

            };
            opt.Filters.Add(filter);
            filterReader.Setup(x => x.ReadFilters()).Returns(opt);

            IFilterProcessor fp = new FilterProcessor(filterReader.Object);
            
            //Act
            WordTable filteredTable = fp.FilterTable(table);

            //Assert
            table.Should().BeEquivalentTo(filteredTable);
        } 

       [Fact]
        public void One_Filter_KownColumn_Keep_Contains_One_Found()
        {
            //Arrange
            WordTable table = GetWordTable();

            var filterReader = new Mock<IFilterReader>();
            FilterOptions opt = new FilterOptions();
            FilterOption filter = new FilterOption()
            {
                FilterType = FilterType.KeepWhenContains,
                ColumnName = "Reference",
                FilterValue = "WA"

            };
            opt.Filters.Add(filter);
            filterReader.Setup(x => x.ReadFilters()).Returns(opt);

            IFilterProcessor fp = new FilterProcessor(filterReader.Object);
            
            //Act
            WordTable filteredTable = fp.FilterTable(table);

            //Assert
            WordTable targetTable = GetWordTable_Minus_Item_Id_1();
            targetTable.Should().BeEquivalentTo(filteredTable);
        } 

        [Fact]
        public void Two_Filters_Two_KownColumns_Keep_One_Equals_One_Contains_Two_Found()
        {
            //Arrange
            WordTable table = GetWordTable();

            var filterReader = new Mock<IFilterReader>();
            FilterOptions opt = new FilterOptions();
            FilterOption filter1 = new FilterOption()
            {
                FilterType = FilterType.KeepWhenContains,
                ColumnName = "Reference",
                FilterValue = "WA"

            };
            FilterOption filter2 = new FilterOption()
            {
                FilterType = FilterType.KeepWhenEquals,
                ColumnName = "Identifier",
                FilterValue = "2"

            };
            opt.Filters.Add(filter1);
            opt.Filters.Add(filter2);
            filterReader.Setup(x => x.ReadFilters()).Returns(opt);

            IFilterProcessor fp = new FilterProcessor(filterReader.Object);
            
            //Act
            WordTable filteredTable = fp.FilterTable(table);

            //Assert
            WordTable targetTable = GetWordTable_Minus_Items_Id_0_And_1();
            targetTable.Should().BeEquivalentTo(filteredTable);
        } 

        [Fact]
        public void Two_Filters_Two_KownColumns_Exclude_One_Equals_One_Contains_Two_Found()
        {
            //Arrange
            WordTable table = GetWordTable();

            var filterReader = new Mock<IFilterReader>();
            FilterOptions opt = new FilterOptions();
            FilterOption filter1 = new FilterOption()
            {
                FilterType = FilterType.ExcludeWhenContains,
                ColumnName = "Reference",
                FilterValue = "GPR"

            };
            FilterOption filter2 = new FilterOption()
            {
                FilterType = FilterType.ExcludeWhenEquals,
                ColumnName = "Identifier",
                FilterValue = "0"

            };
            opt.Filters.Add(filter1);
            opt.Filters.Add(filter2);
            filterReader.Setup(x => x.ReadFilters()).Returns(opt);

            IFilterProcessor fp = new FilterProcessor(filterReader.Object);
            
            //Act
            WordTable filteredTable = fp.FilterTable(table);

            //Assert
            WordTable targetTable = GetWordTable_Minus_Items_Id_0_And_1();
            targetTable.Should().BeEquivalentTo(filteredTable);
        }


        private WordTable GetWordTable()
        {
            WordTable table = new WordTable();
            string[] headerRow = new string[] { "Identifier", "Description", "Reference" };
            string[] row1 = new string[] { "0", "Dad", "AWA-000" };
            string[] row2 = new string[] { "1", "Mom", "GPR-001" };
            string[] row3 = new string[] { "2", "Audrey", "TWA-002" };

            table.Rows.Add(headerRow);
            table.Rows.Add(row1);
            table.Rows.Add(row2);
            table.Rows.Add(row3);

            return table;
        }

        private WordTable GetWordTable_Minus_Item_Id_1()
        {
            WordTable table = GetWordTable();
            table.Rows.RemoveAt(2);

            //{ "Identifier", "Description", "Reference" };
            //{ "0", "Dad", "AWA-000" };
            //{ "2", "Audrey", "TWA-002" };

            return table;
        }

        private WordTable GetWordTable_Minus_Items_Id_0_And_1()
        {
            WordTable table = GetWordTable();
            table.Rows.RemoveAt(2);
            table.Rows.RemoveAt(1); 

            //{ "2", "Audrey", "TWA-002" };

            return table;
        }
    }
}