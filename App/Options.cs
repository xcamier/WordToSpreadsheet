namespace WordToSpreadsheet;

public class Options
{
  [Option('i', "input", Required = true, HelpText = "Word file to be processed.")]
  public string InputFile { get; set; } = string.Empty;

  [Option('o', "output", Required = true, HelpText = "Excel Spreadsheet to generate.")]
  public string OutputFile { get; set; } = string.Empty;

  [Option('m', "merge", Required = false, Default = false, HelpText = "Merge the tables of same nature into one")]
  public bool Merge { get; set; } 

  [Option('t', "tab", Required = false, Default = false, HelpText = "One type of table per tab")]
  public bool OnePerTab { get; set; } 

  [Option('f', "filters", Required = false, HelpText = "JSon file containing the filters rules")]
  public string FiltersFile { get; set; } = string.Empty;

}
