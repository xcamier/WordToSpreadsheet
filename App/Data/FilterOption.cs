namespace WordToSpreadsheet.Data;

public enum FilterType : int
{
    KeepWhenContains = 1,   
    KeepWhenEquals = 2,     
    ExcludeWhenContains = 3,   
    ExcludeWhenEquals = 4       
}

public class FilterOption
{
    public string ColumnName { get; set; } = string.Empty;
    public FilterType FilterType { get; set; }
    public string FilterValue { get; set; } = string.Empty;
}
