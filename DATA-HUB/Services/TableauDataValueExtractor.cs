using OfficeOpenXml;

public class TableauDataValueExtractor
{
    public Dictionary<string, List<decimal>> ExtractNumericValues(string excelPath, List<string> columnNames)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var columnValues = new Dictionary<string, List<decimal>>();

        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Get header row to find column indices
            var headerRow = worksheet.Cells[1, 1, 1, colCount];
            var columnIndices = new Dictionary<string, int>();

            // Map column names to their indices
            for (int col = 1; col <= colCount; col++)
            {
                var headerValue = worksheet.Cells[1, col].Text;
                if (columnNames.Contains($"[{headerValue}]"))
                {
                    columnIndices[headerValue] = col;
                    columnValues[$"[{headerValue}]"] = new List<decimal>();
                }
            }

            // Extract values for each numeric column
            for (int row = 2; row <= rowCount; row++)
            {
                foreach (var colPair in columnIndices)
                {
                    var cellValue = worksheet.Cells[row, colPair.Value].Value;
                    if (decimal.TryParse(cellValue?.ToString(), out decimal numValue))
                    {
                        columnValues[$"[{colPair.Key}]"].Add(numValue);
                    }
                }
            }
        }

        return columnValues;
    }
}
