using TableauConverter.Services;
using Newtonsoft.Json;

class Program
{
    static void Main(string[] args)
    {
        string inputXmlPath = "data/tableauPOC.twb";
        string outputDirectory = "tableau_analysis";
        Directory.CreateDirectory(outputDirectory);






        // Column Information



        var columnExtractor = new TableauColumnExtractor();
        var columns = columnExtractor.ExtractDataSourceColumns(inputXmlPath);
        SaveToJson(columns, Path.Combine(outputDirectory, "columns_info.json"));
        Console.WriteLine("Column information extracted");

        // Visualization Mapping
        var vizMapper = new TableauVisualizationMapper();
        var visualizationMapping = vizMapper.MapVisualizationsToDataSources(inputXmlPath);
        SaveToJson(visualizationMapping, Path.Combine(outputDirectory, "visualization_mapping.json"));
        Console.WriteLine("Visualization mapping completed");


        // Get the column names from your visualization mapping
                    var dataSourceExtractor = new TableauDataSourceExtractor();
            var dataSourceInfo = dataSourceExtractor.ExtractDataSourceInfo(inputXmlPath);

            if (dataSourceInfo != null && !string.IsNullOrEmpty(dataSourceInfo.FilePath))
            {
                var columnNames = visualizationMapping.Values
                    .SelectMany(v => v.UsedColumns)
                    .Distinct()
                    .ToList();

                var extractedData = dataSourceExtractor.ExtractData(dataSourceInfo, columnNames);
                SaveToJson(extractedData, Path.Combine(outputDirectory, "extracted_data.json"));
            }
            else
            {
                Console.WriteLine("Data source information could not be extracted.");
            }
        // Column Usage
        var columnUsage = columnExtractor.ExtractColumnUsage(inputXmlPath);
        SaveToJson(columnUsage, Path.Combine(outputDirectory, "column_usage.json"));
        Console.WriteLine("Column usage analysis completed");
        var dataExtractor = new TableauDataValueExtractor();
        var numericColumns = new List<string> { "[Sales]", "[Discounts]", "[Profit]", "[Units Sold]",
            "[Manufacturing Price]", "[Gross Sales]" };

        var columnValues = dataExtractor.ExtractNumericValues("path/to/Sample data.xlsx", numericColumns);
        SaveToJson(columnValues, Path.Combine(outputDirectory, "column_values.json"));

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
}


