using System.Text.Json;
using System.Xml.Linq;
using OfficeOpenXml; // Install EPPlus NuGet package

public class TableauInfoExtractor
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

    public class VisualizationInfo {
        public string WorksheetName { get; set; }
        public string VisualizationTitle { get; set; }
        public string VisualizationType { get; set; }
    }

    public DatasourceInfo ExtractDataSourceInfo(string filePath) {
        // Initialize the DatasourceInfo object to store the extracted data
        DatasourceInfo datasourceInfo = new DatasourceInfo();

        // Load the XML document from the specified file path
        var xmlDoc = LoadXmlDocument(filePath);

        // Locate the first "datasource" element in the XML document
        var dataSource = GetDataSourceElement(xmlDoc);

        // Extract all "named-connection" elements from the datasource
        var namedConnectionInfo = GetNamedConnectionElements(dataSource);

        // Iterate through each connection to extract details
        foreach (var connection in namedConnectionInfo) {
            // Get the details of the connection element
            var connectionInfo = GetConnectionElement(connection);

            // Extract the "class" attribute to determine the type of connection
            var connectionType = GetConnectionType(connectionInfo);

            // Build the connection string based on the connection type
            string connectionString = BuildConnectionString(connectionType, connection, connectionInfo);

            // Extract the mapping of tables and columns from the XML
            var tableColumnMapping = ExtractTableColumnMapping(xmlDoc);

            // Convert the table-column mapping into a list of TableInfo objects
            var allTables = BuildTableInfoList(tableColumnMapping);

            // Add the connection information to the DatasourceInfo object
            datasourceInfo.connections.Add(new ConnectionInfo {
                ConnectionType = connectionType,
                ConnectionString = connectionString,
                Tables = allTables
            });
        }

        // Return the populated DatasourceInfo object
        return datasourceInfo;
    }

    private XDocument LoadXmlDocument(string filePath) {
        // Load the XML document from the file path
        return XDocument.Load(filePath);
    }

    private XElement GetDataSourceElement(XDocument xmlDoc) {
        // Locate the first "datasource" element in the XML document
        var dataSource = xmlDoc.Descendants("datasource").FirstOrDefault();
        if (dataSource == null) {
            throw new InvalidOperationException("No datasource found in the Tableau file.");
        }
        return dataSource;
    }

    private IEnumerable<XElement> GetNamedConnectionElements(XElement dataSource) {
        // Extract all "named-connection" elements under the "connection" section of the datasource
        return dataSource.Descendants("connection")
                     .FirstOrDefault()?
                     .Descendants("named-connection")
                     ?? Enumerable.Empty<XElement>();
}

    private XElement GetConnectionElement(XElement connection) {
        // Locate the first "connection" element inside a "named-connection" element
        return connection.Descendants("connection").FirstOrDefault();
    }

private string GetConnectionType(XElement connectionInfo) {
    // Extract the "class" attribute to determine the type of connection
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
    // Initialize a dictionary to store the mapping between tables and their columns
    var tableColumnMapping = new Dictionary<string, List<string>>();

    // Locate all column mappings in the "cols" section of the XML
    var columns = xmlDoc.Descendants("cols").FirstOrDefault()?.Descendants("map") ?? Enumerable.Empty<XElement>();

    // Iterate through each column mapping
    foreach (var col in columns) {
        // Extract the "value" attribute, which contains table and column information
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
    // Convert the table-column mapping dictionary into a list of TableInfo objects
    return tableColumnMapping.Select(table => new TableInfo {
        TableName = table.Key,
        Columns = table.Value
    }).ToList();
}
    // public VisualizationInfo ExtractVisualizationInfo(string filePath) {

    // }

    public void SaveConnectionInfoToJSONAndExcel(DatasourceInfo dataSourceInfo, string jsonFilePath, string excelFilePath) {
        // Save the DataSourceInfo object to a JSON file
        SaveToJSON(dataSourceInfo, jsonFilePath);

        // Save the DataSourceInfo object to an Excel file
        SaveToExcel(dataSourceInfo, excelFilePath);
    }

    private void SaveToJSON(DatasourceInfo dataSourceInfo, string jsonFilePath) {
        // Serialize the DatasourceInfo object to a JSON string with indented formatting
        string json = JsonSerializer.Serialize(dataSourceInfo, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON content to the specified file
        File.WriteAllText(jsonFilePath, json);
    }

    private void SaveToExcel(DatasourceInfo dataSourceInfo, string excelFilePath) {
        using var package = new ExcelPackage();

        // Create an Excel worksheet for the datasource information
        var worksheet = package.Workbook.Worksheets.Add("DataSourceInfo");

        // Add headers to the worksheet
        AddExcelHeaders(worksheet);

        // Populate the worksheet with connection and table details
        PopulateExcelData(worksheet, dataSourceInfo);

        // Save the Excel package to the specified file path
        package.SaveAs(excelFilePath);
    }

    private void AddExcelHeaders(ExcelWorksheet worksheet) {
        // Add headers to the first row
        worksheet.Cells[1, 1].Value = "Data Source(s)"; // Header for data source type
        worksheet.Cells[1, 2].Value = "Connection Info"; // Header for connection information
        worksheet.Cells[1, 3].Value = "Data Table(s)"; // Header for table names
        worksheet.Cells[1, 4].Value = "Column(s)"; // Header for column names

        // Make the header cells bold for better visibility
        worksheet.Row(1).Style.Font.Bold = true;
    }

    private void PopulateExcelData(ExcelWorksheet worksheet, DatasourceInfo dataSourceInfo) {
        int connectionRow = 2; // Start populating connections from the second row
        int tableRow = 2; // Start populating tables from the second row

        foreach (var connection in dataSourceInfo.connections) {
            // Add connection type and connection string
            worksheet.Cells[connectionRow, 1].Value = connection.ConnectionType;
            worksheet.Cells[connectionRow, 2].Value = connection.ConnectionString;

            // Add table and column details for the connection
            foreach (var table in connection.Tables) {
                worksheet.Cells[tableRow, 3].Value = table.TableName; // Table name
                worksheet.Cells[tableRow, 4].Value = string.Join(", ", table.Columns); // Columns as a comma-separated list
                tableRow++;
            }

            // Merge cells for connections with multiple tables
            if (connection.Tables.Count > 0) {
                MergeConnectionCells(worksheet, connectionRow, tableRow, connection.Tables.Count);
                connectionRow = tableRow;
            } else {
                // Handle connections with no tables
                worksheet.Cells[connectionRow, 3].Value = "No tables found";
                worksheet.Cells[connectionRow, 4].Value = "No columns found";
                connectionRow++;
                tableRow++;
            }
        }

        // Auto-fit columns to adjust content
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
    }

    private void MergeConnectionCells(ExcelWorksheet worksheet, int connectionRow, int tableRow, int tableCount) {
        // Merge cells for "Data Source(s)" and "Connection Info" columns
        worksheet.Cells[connectionRow, 1, connectionRow + tableCount - 1, 1].Merge = true;
        worksheet.Cells[connectionRow, 2, connectionRow + tableCount - 1, 2].Merge = true;

        // Center align the merged cells
        worksheet.Cells[connectionRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        worksheet.Cells[connectionRow, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        worksheet.Cells[connectionRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        worksheet.Cells[connectionRow, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
    }
    
    // public void SaveVisualizationInfoToJSONAndExcel(VisualizationInfo visualizationInfo, string jsonFilePath, string excelFilePath) {
    // }
}