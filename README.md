# WordToSpreadsheet

## What is WordToSpreadsheet ?
WordToSpreadsheet is a very basic tool that copies tables from a Word file to an Excel Spreadsheet. It provides few options to finetune how the data have to be organized in the Excel spreadsheet.

## Use case
### Why
WordToSpreadsheet was originally implemented to facilitate the comparison and the review of data between documents. Let's imagine you have some data in a word document and those data are structured in Word tables. To compare those data with data coming from other documents, you may want to use the powerfull lookup features of Excel. 

### Example
Let's assume you have a document containing the following tables:

| Brand      | Model |
|------------|-------|
| Renault    | Clio  |
| Peugeot    | 208   |
| Volkswagen | Polo  |

| Brand    | Model   |
|----------|---------|
| Mercedes | Class A |
| BMW      | Serie 1 |

| Model sold | Value €  | Benefit % |
|------------|----------|-----------|
| iPhone     | 15000    | 25        |
| Pixel 6    | 17000    | 12        |
| Galaxy S22 | 19000    | 17        |

#### Usage
```
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx"
```
The tool creates an Excel spreadsheet containing:
- 1 tab concatenating the 3 tables above


```
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx" -t
```
The tool creates an Excel spreadsheet containing:
- 1 tab containing the first and the second of the tables above
- 1 tab containing the third of the tables above


```
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx" -r
```
The tool:
- Merges the first and the second of the tables above as they have the same header
- Creates 1 tab containing merged table and the third of the tables above


```
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx" -r
```
The tool:
- Merges the first and the second of the tables above as they have the same header
- Creates 1 tab containing merged table 
- Creates 1 tab containing the third of the tables above


#### From Windows Binaries 
You have the same options but run the executable instead. For example:
```
WordToSpreadsheet.exe -i "WordSource.docx" -o "ExcelTarget.xlsx" -t -m
```

## Credits
My friend Rakesh for having inpired me the idea.

### Dependencies
- .net Core 6
- Command Line Parser: https://github.com/commandlineparser/commandline
- NPOI: https://github.com/nissl-lab/npoi