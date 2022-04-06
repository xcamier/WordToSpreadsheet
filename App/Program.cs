using WordToSpreadsheet;
using WordToSpreadsheet.Excel;
using WordToSpreadsheet.Word;


CommandLine.Parser.Default.ParseArguments<Options>(args)
    .WithParsed(RunOptions)
    .WithNotParsed(HandleParseError);


static void RunOptions(Options opts)
{
  try
  {
    Console.WriteLine("Reading the filtering options...");
    IFileIO fileIO = new FileIO();
    string fileName = string.IsNullOrEmpty(opts.FiltersFile) ? "filters.json" : opts.FiltersFile; 
    IFilterReader filterReader = new FilterReader(fileIO, fileName);
    IFilterProcessor filterProcessor = new FilterProcessor(filterReader);

    Console.WriteLine("Getting the tables from Word Document...");
    WordTableProcessing wtp = new WordTableProcessing();
    WordReader wr = new WordReader(wtp);
    List<ExcelTab> wordTables = wr.GetTabs(opts.InputFile, opts.Merge, opts.OnePerTab);

    Console.WriteLine("Building the Excel spreadsheet...");
    ExcelWriter ew = new ExcelWriter(filterProcessor);
    ew.CreateExcel(wordTables, opts.OutputFile);

    Console.WriteLine("Done !");
  }
  catch (ArgumentException argEx)
  {
    Console.WriteLine(argEx.Message);
  }
}

static void HandleParseError(IEnumerable<Error> errs)
{
  //handle errors
}