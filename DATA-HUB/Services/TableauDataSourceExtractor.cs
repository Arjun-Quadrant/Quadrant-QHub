using System;
using System.Xml.Linq;
using System.Data.SqlClient;
using OfficeOpenXml;

public class TableauDataSourceExtractor
{
    public DataSourceInfo ExtractDataSourceInfo(string twbFilePath)
    {
        XDocument doc = XDocument.Load(twbFilePath);
        var namedConnection = doc.Descendants("named-connection").FirstOrDefault();
        
        if (namedConnection != null)
        {
            var excelConnection = namedConnection.Descendants("connection")
                .FirstOrDefault(c => c.Attribute("class")?.Value == "excel-direct");
                
            return new DataSourceInfo
            {
                Type = "Excel",
                // File must be accessed at a local destination in order for ExcelPackage to work
                FilePath = @"C:\Users\ArjunNarendra(Quadra\sampleData.xlsx",
                ConnectionDetails = GetConnectionDetails(doc)
            };
        }
        
        return new DataSourceInfo { Type = "Unknown" };
    }

    private string GetConnectionDetails(XDocument doc)
    {
        var relation = doc.Descendants("relation").FirstOrDefault();
        return relation?.Attribute("table")?.Value;
    }

    public Dictionary<string, List<object>> ExtractData(DataSourceInfo dataSource, List<string> columnNames)
    {
        if (dataSource == null || string.IsNullOrEmpty(dataSource.FilePath))
        {
            Console.WriteLine("Data source information is missing");
            return new Dictionary<string, List<object>>();
        }

        return dataSource.Type switch
        {
            "Excel" => ExtractFromExcel(dataSource.FilePath, columnNames),
            "SQL" => ExtractFromSql(dataSource.ConnectionString, columnNames),
            _ => new Dictionary<string, List<object>>()
        };
    }

    private Dictionary<string, List<object>> ExtractFromExcel(string filePath, List<string> columnNames)
    {
        var data = new Dictionary<string, List<object>>();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        try
        {
            string fullPath = Path.GetFullPath(filePath);
            Console.WriteLine($"Attempting to read Excel file: {fullPath}");

            using var package = new ExcelPackage(fullPath);

            if (package.Workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Excel file contains no worksheets");
                return data;
            }

            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            
            if (worksheet?.Dimension == null)
            {
                Console.WriteLine("Worksheet is empty");
                return data;
            }

            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Initialize columns
            foreach (var col in columnNames)
            {
                data[col] = new List<object>();
            }

            var headerIndices = MapHeaderToColumn(colCount, columnNames, worksheet);

            // Extract data
            for (int row = 2; row <= rowCount; row++)
            {
                foreach (var header in headerIndices)
                {
                    var value = worksheet.Cells[row, header.Value].Value;
                    data[header.Key].Add(value);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting Excel data: {ex.Message}");
        }

        return data;
    }

    public Dictionary<string, List<decimal>> ExtractNumericValues(string excelPath, List<string> columnNames)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var data = new Dictionary<string, List<decimal>>();

        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Initialize columns
            foreach (var col in columnNames)
            {
                data[col] = new List<decimal>();
            }

            var headerIndices = MapHeaderToColumn(colCount, columnNames, worksheet);

            // Extract data
            for (int row = 2; row <= rowCount; row++)
            {
                foreach (var header in headerIndices)
                {
                    var cellValue = worksheet.Cells[row, header.Value].Value;
                    if (decimal.TryParse(cellValue?.ToString(), out decimal numValue))
                    {
                        data[$"{header.Key}"].Add(numValue);
                    }
                }
            }
        }

        return data;
    }

    private Dictionary<string, int> MapHeaderToColumn(int colCount, List<String> columnNames, ExcelWorksheet worksheet) {
        var headerIndices = new Dictionary<string, int>();
        for (int col = 1; col <= colCount; col++)
        {
            var cellValue = worksheet.Cells[1, col].Text;
            var header = $"[{cellValue}]";
            if (columnNames.Contains(header))
            {
                headerIndices[header] = col;
            }
        }
        return headerIndices;
    }

    private Dictionary<string, List<object>> ExtractFromSql(string connectionString, List<string> columnNames)
    {
        var data = new Dictionary<string, List<object>>();
        
        foreach (var col in columnNames)
        {
            data[col] = new List<object>();
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();

            var cleanColumnNames = columnNames.Select(c => c.Trim('[', ']'));
            var query = $"SELECT {string.Join(", ", cleanColumnNames)} FROM YourTableName";

            using var command = new SqlCommand(query, connection);
            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                foreach (var col in columnNames)
                {
                    var cleanName = col.Trim('[', ']');
                    data[col].Add(reader[cleanName]);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting SQL data: {ex.Message}");
        }

        return data;
    }
}

public class DataSourceInfo
{
    public string Type { get; set; }
    public string FilePath { get; set; }
    public string ConnectionString { get; set; }
    public string ConnectionDetails { get; set; }
    public string Server { get; set; }
    public string Database { get; set; }
}
