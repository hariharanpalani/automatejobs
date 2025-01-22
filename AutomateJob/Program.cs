using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Data.Analysis;
using Microsoft.ML;
// Path to your Excel file

//-i=Ids.xlsx -s=50mb.xlsx -e=Name,Email,Phone,Company,"Job Title"

if (args.Length == 0)
{
    throw new ArgumentException("No arguments passed");
}

string inputFile = "";
string sourceFile = "";
List<string> outputColumns = new List<string>();

foreach (var arg in args)
{
    var items = arg.Split('=');
    switch (items[0].ToUpper())
    {
        case "--INPUT":
        case "-I":
            inputFile = items[1];
            break;
        case "-SOURCE":
        case "-S":
            sourceFile = items[1];
            break;
        case "--EXTRACT":
        case "-E":
            outputColumns = items[1].Trim().Split(',').ToList();
            break;
    }
}

Console.WriteLine("input file: {0}", inputFile);
Console.WriteLine("Source file: {0}", sourceFile);
Console.WriteLine("Columns to extract: {0}", string.Join(",", outputColumns));

var workbook = new XLWorkbook(sourceFile);
var worksheet = workbook.Worksheets.First();
var dataFrame = DataFrameFromWorksheet(worksheet);

var inputbook = new XLWorkbook(inputFile);
var firstSheet = inputbook.Worksheets.First();
var columns = firstSheet.ColumnsUsed();
var fields = new Dictionary<string, List<string>>();
foreach (var column in columns)
{
    var header = column.Cell(1).GetValue<string>();
    var data = column.CellsUsed().Skip(1).Select(c => c.GetValue<string>()).ToList();
    if (data.Count > 0)
    {
        fields.Add(header, data);
    }
}

var conditions = new List<PrimitiveDataFrameColumn<bool>>();
foreach (var field in fields.Keys)
{
    var index = 0;
    var conditionMask = dataFrame[field].ElementwiseEquals(fields[field][0]); ;
    foreach (var value in fields[field])
    {
        if (index != 0)
        {
            conditionMask = (PrimitiveDataFrameColumn<bool>)conditionMask.Or(dataFrame[field].ElementwiseEquals(value));
        }
        index++;
    }
    conditions.Add(conditionMask);
}

if (conditions.Count > 0)
{
    var combinedMask = conditions[0];
    var index = 0;
    foreach (var condition in conditions)
    {
        if (index != 0)
        {
            combinedMask = (PrimitiveDataFrameColumn<bool>)combinedMask.And(condition);
        }
        index++;
    }

    dataFrame = dataFrame.Filter(combinedMask);
}

var output = dataFrame.ToDataFrame(-1, outputColumns.ToArray()).ToTable();
// Display the DataFrame
var resultBook = new XLWorkbook();
var ws = resultBook.Worksheets.Add(output, "output");
ws.Table(0).ShowAutoFilter = false;
ws.Table(0).Theme = XLTableTheme.None;
ws.Columns().AdjustToContents();
resultBook.SaveAs($"{sourceFile.Split('.')[0]}_results.xlsx");
Console.WriteLine("Filter extracted successfully.");



DataFrame DataFrameFromWorksheet(IXLWorksheet worksheet)
{
    var rows = worksheet.RowsUsed();
    Console.WriteLine("Number of Rows: " + rows.Count());
    var columns = worksheet.ColumnsUsed();
    Console.WriteLine("Number of Columns: " + columns.Count());
    var dataFrameColumns = new DataFrameColumn[columns.Count()];

    int colIndex = 0;
    foreach (var column in columns)
    {
        var header = column.Cell(1).GetValue<string>();
        /* var data = column.Cells($"2:{rows.Count() - 1}").Select(cell =>
         {
             string? content = null;
             cell.TryGetValue<string>(out content);
             return content;
         }).ToList();*/
        var data = column.Cells($"2:{rows.Count()}").Select(cell => cell.IsEmpty() ? null : cell.GetValue<string>()).ToList();
        //data = PadData(data, rows.Count() - 1);
        var dataColumn = new StringDataFrameColumn(header, data);
        //dataColumn = PadColumn(dataColumn, rows.Count());
        dataFrameColumns[colIndex++] = dataColumn;
    }

    return new DataFrame(dataFrameColumns);
}
