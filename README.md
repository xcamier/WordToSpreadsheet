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

```bash
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx"
```


The tool creates an Excel spreadsheet containing:



**Tab_1**

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




```bash
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx" -t
```


The tool creates an Excel spreadsheet containing:



**Tab_1**
| Brand      | Model |
|------------|-------|
| Renault    | Clio  |
| Peugeot    | 208   |
| Volkswagen | Polo  |

| Brand    | Model   |
|----------|---------|
| Mercedes | Class A |
| BMW      | Serie 1 |



**Tab_2**

| Model sold | Value €  | Benefit % |
|------------|----------|-----------|
| iPhone     | 15000    | 25        |
| Pixel 6    | 17000    | 12        |
| Galaxy S22 | 19000    | 17        |




```bash
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx" -m
```


The tool creates an Excel spreadsheet containing:



**Tab_1**

| Brand      | Model   |
|------------|---------|
| Renault    | Clio    |
| Peugeot    | 208     |
| Volkswagen | Polo    |
| Mercedes   | Class A |
| BMW        | Serie 1 |

| Model sold | Value €  | Benefit % |
|------------|----------|-----------|
| iPhone     | 15000    | 25        |
| Pixel 6    | 17000    | 12        |
| Galaxy S22 | 19000    | 17        |




```bash
dotnet run -i "WordSource.docx" -o "ExcelTarget.xlsx" -m -t
```


The tool creates an Excel spreadsheet containing:



**Tab_1**

| Brand      | Model   |
|------------|---------|
| Renault    | Clio    |
| Peugeot    | 208     |
| Volkswagen | Polo    |
| Mercedes   | Class A |
| BMW        | Serie 1 |



**Tab_2**

| Model sold | Value €  | Benefit % |
|------------|----------|-----------|
| iPhone     | 15000    | 25        |
| Pixel 6    | 17000    | 12        |
| Galaxy S22 | 19000    | 17        |



##### Filtering options

When running the app for the first time, a file called filters.json is created. That JSon file allows you to configure filters.

###### File structure
{
  "Filters": [
    {
      "ColumnName": "TestCol1",
      "FilterType": 1,
      "FilterValue": "testValue1"
    },
    {
      "ColumnName": "TesCol2",
      "FilterType": 3,
      "FilterValue": "testValue2"
    }
  ]
}

- ColumnName: column header that will be considered for the filter

- FilterValue: the value to filter

- FilterType: mode of filtering

  - 1: Keep the line when the text in column *ColumnName* contains *FilterValue*
  - 2: Keep the line when the text in column *ColumnName* equals *FilterValue* (case sensitive)
  - 3: Exclude the line when the text in column *ColumnName* contains *FilterValue*
  - 4: Exclude the line when the text in column *ColumnName* equals *FilterValue* (case sensitive)

  


#### From Windows Binaries 
You have the same options but run the executable instead. For example:
```bash
WordToSpreadsheet.exe -i "WordSource.docx" -o "ExcelTarget.xlsx" -t -m
```



## Credits

My friend Rakesh for having inpired me the idea.



### Dependencies

- .net Core 6
- Command Line Parser: https://github.com/commandlineparser/commandline
- NPOI: https://github.com/nissl-lab/npoi
- FluentAssersions: https://github.com/fluentassertions/fluentassertions
- Moq: https://github.com/moq/moq4