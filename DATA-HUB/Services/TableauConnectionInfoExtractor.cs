using System.Text.Json;
using System.Xml.Linq;
using OfficeOpenXml;

public class TableauConnectionInfoExtractor
{
    public class DatasourceInfo {
        public List<ConnectionInfo> connections { get; set; } = new();
    }
    public class ConnectionInfo
    {
        public string ConnectionType { get; set; }
        public string ConnectionString { get; set; }
        public List<TableInfo> Tables { get; set; } = new();
    }

    public class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new();
    }

    public DatasourceInfo ExtractDataSourceInfo(string filePath) {
        DatasourceInfo datasourceInfo = new DatasourceInfo();
        var xmlDoc = LoadXmlDocument(filePath);
        var dataSource = GetDataSourceElement(xmlDoc);
        var namedConnectionInfo = GetNamedConnectionElements(dataSource);

        // Iterate through each connection to extract details
        foreach (var connection in namedConnectionInfo) {
            var connectionInfo = GetConnectionElement(connection);
            var connectionType = GetConnectionType(connectionInfo);
            string connectionString = BuildConnectionString(connectionType, connection, connectionInfo);
            var tableColumnMapping = ExtractTableColumnMapping(xmlDoc);
            var allTables = BuildTableInfoList(tableColumnMapping);
            datasourceInfo.connections.Add(new ConnectionInfo {
                ConnectionType = connectionType,
                ConnectionString = connectionString,
                Tables = allTables
            });
        }
        return datasourceInfo;
    }

    private XDocument LoadXmlDocument(string filePath) {
        return XDocument.Load(filePath);
    }

    private XElement GetDataSourceElement(XDocument xmlDoc) {
        var dataSource = xmlDoc.Descendants("datasource").FirstOrDefault();
        if (dataSource == null) {
            throw new InvalidOperationException("No datasource found in the Tableau file.");
        }
        return dataSource;
    }

    private IEnumerable<XElement> GetNamedConnectionElements(XElement dataSource) {
        return dataSource.Descendants("connection")
                     .FirstOrDefault()?
                     .Descendants("named-connection")
                     ?? Enumerable.Empty<XElement>();
    }

    private XElement GetConnectionElement(XElement connection) {
        return connection.Descendants("connection").FirstOrDefault();
    }

private string GetConnectionType(XElement connectionInfo) {
    return connectionInfo.Attribute("class")?.Value ?? "Unknown";
}

private string BuildConnectionString(string connectionType, XElement connection, XElement connectionInfo) {
    // Build the connection string based on the connection type
    if (connectionType == "sqlserver") {
        // For SQL Server connections, extract specific connection details
        return $"sqlserver; Server: {connection.Attribute("caption").Value}; " +
               $"Database: {connectionInfo.Attribute("dbname").Value}; " +
               $"Authentication: {connectionInfo.Attribute("authentication").Value}; " +
               $"Require SSL: {(connectionInfo.Attribute("sslmode").Value == "require" ? "Yes" : "No")}";
    }
    // For other connection types, use the "filename" attribute
    return connectionInfo.Attribute("filename")?.Value ?? "Unknown";
}

private Dictionary<string, List<string>> ExtractTableColumnMapping(XDocument xmlDoc) {
    var tableColumnMapping = new Dictionary<string, List<string>>();

    // Locate all column mappings in the "cols" section of the XML
    var columns = xmlDoc.Descendants("cols").FirstOrDefault()?.Descendants("map") ?? Enumerable.Empty<XElement>();

    foreach (var col in columns) {
        var info = col.Attribute("value")?.Value;
        if (info == null) continue;

        // Split the value to separate table and column names
        var tableAndColumnNames = info.Split(".");
        var tableName = tableAndColumnNames[0].Trim('[', ']');
        var columnName = tableAndColumnNames[1].Trim('[', ']');

        // Add the column to the corresponding table in the dictionary
        if (!tableColumnMapping.ContainsKey(tableName)) {
            tableColumnMapping[tableName] = new List<string>();
        }
        tableColumnMapping[tableName].Add(columnName);
    }

    return tableColumnMapping;
}

private List<TableInfo> BuildTableInfoList(Dictionary<string, List<string>> tableColumnMapping) {
    return tableColumnMapping.Select(table => new TableInfo {
        TableName = table.Key,
        Columns = table.Value
    }).ToList();
}

    public void SaveConnectionInfoToJSONAndExcel(DatasourceInfo dataSourceInfo, string jsonFilePath, string excelFilePath) {
        SaveToJSON(dataSourceInfo, jsonFilePath);
        SaveToExcel(dataSourceInfo, excelFilePath);
    }

    private void SaveToJSON(DatasourceInfo dataSourceInfo, string jsonFilePath) {
        string json = JsonSerializer.Serialize(dataSourceInfo, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonFilePath, json);
    }

    private void SaveToExcel(DatasourceInfo dataSourceInfo, string excelFilePath) {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("DataSourceInfo");
        AddExcelHeaders(worksheet);
        PopulateExcelData(worksheet, dataSourceInfo);
        package.SaveAs(excelFilePath);
    }

    private void AddExcelHeaders(ExcelWorksheet worksheet) {
        worksheet.Cells[1, 1].Value = "Data Source(s)";
        worksheet.Cells[1, 2].Value = "Connection Info";
        worksheet.Cells[1, 3].Value = "Data Table(s)";
        worksheet.Cells[1, 4].Value = "Column(s)";
        worksheet.Row(1).Style.Font.Bold = true;
    }

    private void PopulateExcelData(ExcelWorksheet worksheet, DatasourceInfo dataSourceInfo) {
        int connectionRow = 2;
        int tableRow = 2;

        foreach (var connection in dataSourceInfo.connections) {
            worksheet.Cells[connectionRow, 1].Value = connection.ConnectionType;
            worksheet.Cells[connectionRow, 2].Value = connection.ConnectionString;

            foreach (var table in connection.Tables) {
                worksheet.Cells[tableRow, 3].Value = table.TableName;
                worksheet.Cells[tableRow, 4].Value = string.Join(", ", table.Columns);
                tableRow++;
            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            if (connection.Tables.Count > 0) {
                MergeConnectionCells(worksheet, connectionRow, tableRow, connection.Tables.Count);
                connectionRow = tableRow;
            } else {
                worksheet.Cells[connectionRow, 3].Value = "No tables found";
                worksheet.Cells[connectionRow, 4].Value = "No columns found";
                connectionRow++;
                tableRow++;
            }
        }
    }

    private void MergeConnectionCells(ExcelWorksheet worksheet, int connectionRow, int tableRow, int tableCount) {
        worksheet.Cells[connectionRow, 1, connectionRow + tableCount - 1, 1].Merge = true;
        worksheet.Cells[connectionRow, 2, connectionRow + tableCount - 1, 2].Merge = true;

        worksheet.Cells[connectionRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        worksheet.Cells[connectionRow, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        worksheet.Cells[connectionRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        worksheet.Cells[connectionRow, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
    }
}