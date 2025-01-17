using System.Text.Json;
using System.Xml.Linq;
using OfficeOpenXml; // Install EPPlus NuGet package

public class TableauConnectionInfoExtractor
{
    public class DataSourceInfo
    {
        public string DataSourceType { get; set; }
        public string ConnectionInfo { get; set; }
        public List<TableInfo> Tables { get; set; } = new();
    }

    public class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new();
    }

    public DataSourceInfo ExtractDataSourceInfo(string filePath)
    {
        var xmlDoc = XDocument.Load(filePath);
        var dataSource = xmlDoc.Descendants("datasource").FirstOrDefault();

        if (dataSource == null)
            throw new InvalidOperationException("No datasource found in the Tableau file.");

        // Extract the type of datasource being used
        var dataSourceInfo = dataSource.Descendants("connection")
            .FirstOrDefault().Descendants("named-connections").FirstOrDefault()
            .Descendants("connection").FirstOrDefault();
            
        var dataSourceType = dataSourceInfo.Attribute("class")?.Value ?? "Unknown";

        // Extract the datasource connection info
        string dataSourceConnection = dataSourceInfo.Attribute("filename")?.Value ?? "Unknown";

        var tableColumnMapping = new Dictionary<string, List<string>>();
        var columns = xmlDoc.Descendants("cols").FirstOrDefault().Descendants("map");
        foreach (var col in columns) {
            var info = col.Attribute("value").Value;
            var tableAndColumnNames = info.Split(".");
            var tableName = tableAndColumnNames[0];
            var columnName = tableAndColumnNames[1];
            tableName = tableName.Trim('[', ']');
            columnName = columnName.Trim('[', ']');
            if (tableColumnMapping.ContainsKey(tableName)) {
                    tableColumnMapping[tableName].Add(columnName);
            } else {
                    tableColumnMapping[tableName] = new List<string> {columnName};
            }
        }

        List<TableInfo> allTables = new List<TableInfo>();
        foreach (var table in tableColumnMapping) {
            allTables.Add(new TableInfo {
                TableName = table.Key,
                Columns = table.Value
            });
        }

        return new DataSourceInfo
        {
            DataSourceType = dataSourceType,
            ConnectionInfo = dataSourceConnection,
            Tables = allTables
        };
    }

    public void SaveToJSONAndExcel(DataSourceInfo dataSourceInfo, string jsonFilePath, string excelFilePath)
    {
        // Save to JSON
        string json = JsonSerializer.Serialize(dataSourceInfo, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonFilePath, json);

        // Save to Excel
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("DataSourceInfo");

        worksheet.Cells[1, 1].Value = "Data Source(s)";
        worksheet.Cells[2, 1].Value = dataSourceInfo.DataSourceType;

        worksheet.Cells[1, 2].Value = "Connection Info";
        worksheet.Cells[2, 2].Value = dataSourceInfo.ConnectionInfo;

        worksheet.Cells[1, 3].Value = "Data Table(s)";
        worksheet.Cells[1, 4].Value = "Column(s)";
        int row = 2;
        foreach (var table in dataSourceInfo.Tables) {
            worksheet.Cells[row, 3].Value = table.TableName;
            foreach (var column in table.Columns) {
                string currentValue = worksheet.Cells[row, 4].Text;
                string updatedValue = string.IsNullOrEmpty(currentValue) ? column : currentValue + ", " + column;
                worksheet.Cells[row, 4].Value = updatedValue;
            }
            row++;
        }

        // Make header cells bold
        int columnCount = worksheet.Dimension.Columns;
        for (int i = 1; i <= columnCount; i++) {
            worksheet.Cells[1, i].Style.Font.Bold = true;
        }

        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

        // Merge cells
        worksheet.Cells[2, 1, 1 + dataSourceInfo.Tables.Count, 1].Merge = true;
        worksheet.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        worksheet.Cells[2, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        worksheet.Cells[2, 2, 1 + dataSourceInfo.Tables.Count, 2].Merge = true;
        worksheet.Cells[2, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        worksheet.Cells[2, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        package.SaveAs(excelFilePath);
    }
}