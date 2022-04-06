using System.Text.Json;

namespace WordToSpreadsheet.Filtering;

public class FilterReader : IFilterReader
{
    private IFileIO? _fileIO = null;
    private string _fileName = string.Empty;

    public FilterReader(IFileIO fileIO, string fileName)
    {
        _fileIO = fileIO;
        _fileName = fileName;
    }

    public FilterOptions ReadFilters()
    {
        string jsonString = string.Empty;
        if (!_fileIO!.Exists(_fileName))
            jsonString = CreateFiltersFile(_fileName);
        else
            jsonString = DeserializeFile(_fileName);

        if (!string.IsNullOrEmpty(jsonString))
            return JsonSerializer.Deserialize<FilterOptions>(jsonString)!;

        throw new Exception("Unable to create or read the json file containing the filters");
    }

    private static string SerializeEmptyFilters()
    {
        FilterOptions opts = new FilterOptions();
        FilterOption example = new FilterOption()
        {
            ColumnName = "testCol",
            FilterType = FilterType.ExcludeWhenContains,
            FilterValue = "testValue"
        };
        opts.Filters.Add(example);
        var serializationOptions = new JsonSerializerOptions { WriteIndented = true };
        return JsonSerializer.Serialize(opts, serializationOptions);
    }

    private string CreateFiltersFile(string fileName)
    {
        string jsonString = SerializeEmptyFilters();

        _fileIO!.WriteFileContent(jsonString, fileName);

        return jsonString;
    }

    private string DeserializeFile(string fileName)
    {
        return _fileIO!.ReadFileContent(fileName);
    }
}
