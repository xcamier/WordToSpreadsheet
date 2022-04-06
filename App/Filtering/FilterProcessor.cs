namespace WordToSpreadsheet.Filtering;

public class FilterProcessor: IFilterProcessor
{
    IFilterReader _filterReader;

    public FilterProcessor(IFilterReader filterReader)
    {
        _filterReader = filterReader;
    }

    public WordTable FilterTable(WordTable table)
    {
        //Copy the table to remove the rows without altering the source
        WordTable wtCopy = table.Copy();

        FilterOptions opt =  _filterReader.ReadFilters();
        string[] headers = wtCopy.Rows[0];
        foreach (FilterOption filter in opt.Filters)
        {
            //find the column index in order to apply the filter
            int colIdxToFilter = -1;
            for (int colIdx = 0; colIdx < headers.Count(); colIdx++)
            {
                if (headers[colIdx] == filter.ColumnName)
                {
                    colIdxToFilter = colIdx;
                    break;
                }
            }

            // Tests the value of the columns against the filter
            if (colIdxToFilter > -1)
            {
                for (int rowIdx = wtCopy.Rows.Count -1; rowIdx > 0; rowIdx--)   //>0 to not test the header
                {

                    if ((filter.FilterType == FilterType.KeepWhenContains && !wtCopy.Rows[rowIdx][colIdxToFilter].Contains(filter.FilterValue))
                        || (filter.FilterType == FilterType.KeepWhenEquals && wtCopy.Rows[rowIdx][colIdxToFilter] != filter.FilterValue)
                        || (filter.FilterType == FilterType.ExcludeWhenContains && wtCopy.Rows[rowIdx][colIdxToFilter].Contains(filter.FilterValue))
                        || filter.FilterType == FilterType.ExcludeWhenEquals && wtCopy.Rows[rowIdx][colIdxToFilter] == filter.FilterValue)
                    {
                        wtCopy.Rows.RemoveAt(rowIdx);
                    }
                }
            }
        }

        return wtCopy;
    }
}
