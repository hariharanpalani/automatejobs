using ClosedXML.Excel;
using Microsoft.Data.Analysis;
using Microsoft.ML;
class Program1
{
     void Main(string[] args)
    {
        // Path to your Excel file
        string filePath = "50mb.xlsx";

        // Load Excel file using ClosedXML
        var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheets.First();
        var dataFrame = DataFrameFromWorksheet(worksheet);
        string[] allowedNames = { "Able Seamen", "Account Manager" };
        string[] allowedCities = { "Tremblay, Crona and Hagenes", "Christiansen, Shields and Ernser", "Abbott PLC" };

        var nameMask = dataFrame["Job Title"].ElementwiseEquals(allowedNames[0]);
        for (int i = 1; i < allowedNames.Length; i++)
        {
            nameMask = (PrimitiveDataFrameColumn<bool>)nameMask.Or(dataFrame["Job Title"].ElementwiseEquals(allowedNames[i]));
        }

        var cityMask = dataFrame["Company"].ElementwiseEquals(allowedCities[0]);
        for (int i = 1; i < allowedCities.Length; i++)
        {
            cityMask = (PrimitiveDataFrameColumn<bool>)cityMask.Or(dataFrame["Company"].ElementwiseEquals(allowedCities[i]));
        }
        // dataFrame = dataFrame.Filter(dataFrame["Id"].v);

        var combinedMask = (PrimitiveDataFrameColumn<bool>)nameMask.And(cityMask);
        // Filter the DataFrame
        var filteredDf = dataFrame.Filter(combinedMask);

        var output = filteredDf.ToDataFrame("Name", "Company", "Job Title", "Address", "Email").ToTable();
        // Display the DataFrame
        var resultBook = new XLWorkbook();
        resultBook.Worksheets.Add(output, "output");
        resultBook.SaveAs("results.xlsx");
        Console.WriteLine(dataFrame);
    }

    static DataFrame DataFrameFromWorksheet(IXLWorksheet worksheet)
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
            var data = column.CellsUsed().Skip(1).Select(cell => cell.GetValue<string>()).ToList();
            data = PadData(data, rows.Count() - 1);
            var dataColumn = new StringDataFrameColumn(header, data);
            //dataColumn = PadColumn(dataColumn, rows.Count());
            dataFrameColumns[colIndex++] = dataColumn;
        }

        return new DataFrame(dataFrameColumns);
    }

    static List<string> PadData(List<string> data, int maxLength)
    {
        if (data.Count < maxLength)
        {
            var startIndex = data.Count + 1;
            while (startIndex <= maxLength)
            {
                data.Add("");
                startIndex++;
            }
        }
        return data;
    }
}
