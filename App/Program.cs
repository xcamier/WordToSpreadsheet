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
    Console.WriteLine("Getting the tables from Word Document...");
    WordTableProcessing wtp = new WordTableProcessing();
    WordReader wr = new WordReader(wtp);
    List<ExcelTab> wordTables = wr.GetTabs(opts.InputFile, opts.Merge, opts.OnePerTab);

    Console.WriteLine("Building the Excel spreadsheet...");
    ExcelWriter ew = new ExcelWriter();
    ew.CreateExcel(wordTables, opts.Outputfile);
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