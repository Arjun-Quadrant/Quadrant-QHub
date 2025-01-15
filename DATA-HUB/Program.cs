using System.Net.WebSockets;
using Newtonsoft.Json;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args) {

        string inputXmlPath = "data/tableauPOC.twb";
        string outputDirectory = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis";
        Directory.CreateDirectory(outputDirectory);

        // Step 1: Column Information
        var columnExtractor = new TableauColumnExtractor();
        var columns = columnExtractor.ExtractDataSourceColumns(inputXmlPath);
        SaveToJson(columns, Path.Combine(outputDirectory, "columns_info.json"));
        SaveToExcel(columns, Path.Combine(outputDirectory, "columns_info.xlsx"));
        Console.WriteLine("Column information extracted");

        // Step 2: Visualization Mapping
        var vizMapper = new TableauVisualizationMapper();
        var visualizationMapping = vizMapper.MapWorksheetToDataSourceColumns(inputXmlPath);
        SaveToJson(visualizationMapping, Path.Combine(outputDirectory, "visualization_mapping.json"));
        SaveToExcel(visualizationMapping, Path.Combine(outputDirectory, "visualization_mapping.xlsx"));
        Console.WriteLine("Visualization mapping completed");

        // Step 3: Map used columns to the data they contain
        var dataSourceExtractor = new TableauDataSourceExtractor();
        var dataSourceInfo = dataSourceExtractor.ExtractDataSourceInfo(inputXmlPath);

        if (dataSourceInfo != null && !string.IsNullOrEmpty(dataSourceInfo.FilePath)) {
            var columnNames = visualizationMapping.Values
                .SelectMany(v => v.UsedColumns)
                .Distinct()
                .ToList();
            var extractedData = dataSourceExtractor.ExtractData(dataSourceInfo, columnNames);
            SaveToJson(extractedData, Path.Combine(outputDirectory, "extracted_data.json"));
            SaveToExcel(extractedData, Path.Combine(outputDirectory, "extracted_data.xlsx"));
            Console.WriteLine("Column to data mapping completed.");
        } else {
            Console.WriteLine("Data source information could not be extracted.");
        }

        // Step 4: Column Usage
        var columnUsage = columnExtractor.ExtractColumnUsage(inputXmlPath);
        SaveToJson(columnUsage, Path.Combine(outputDirectory, "column_usage.json"));
        SaveToExcel(columnUsage, Path.Combine(outputDirectory, "column_usage.xlsx"));
        Console.WriteLine("Column usage analysis completed");

        var dataExtractor = new TableauDataSourceExtractor();
        // These are hardcoded in. Is there a better way to do this?
        var numericColumns = new List<string> { "[Sales]", "[Discounts]", "[Profit]", "[Units Sold]",
            "[Manufacturing Price]", "[Gross Sales]" };
        // Step 5: Map numeric columns to the data they contain
        if (dataSourceInfo != null) {
            var columnValues = dataExtractor.ExtractNumericValues(dataSourceInfo.FilePath, numericColumns);
            SaveToJson(columnValues, Path.Combine(outputDirectory, "column_values.json"));
            SaveToExcel(columnValues, Path.Combine(outputDirectory, "column_values.xlsx"));
        }
        Console.WriteLine("Numeric column to data mapping completed.");
        Console.WriteLine($"\nAll analyses complete. Results saved in {outputDirectory}");
    }

    private static void SaveToJson<T>(T data, string filePath)
    {
        var jsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore
        };

        string jsonOutput = JsonConvert.SerializeObject(data, jsonSettings);
        File.WriteAllText(filePath, jsonOutput);
    }

    private static void SaveToExcel<T>(IEnumerable<T> data, string filePath) {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        try {
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");
            worksheet.Cells["A1"].LoadFromCollection(data, true);
            package.SaveAs(filePath);
        } catch (Exception ex) {
            Console.WriteLine($"Error extracting Excel data: {ex.Message}");
        }
    }
}