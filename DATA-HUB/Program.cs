using System.Diagnostics;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

class Program
{
    public static async Task Main(string[] args) {
        // Define file paths for the output files
        string connectionInfoJsonFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\connection info.json"; 
        string connectionInfoExcelFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\connection info.xlsx"; 
        string visualizationInfoJsonFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\visualization info.json"; 
        string visualizationInfoExcelFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\visualization info.xlsx"; 

        // Get information about the workbook
        string workbookfilePath = @"Data/World Wide Importers Analysis.twb"; // Path to the Tableau workbook file (adjust as needed). This is needed for XML extraction.
        string pattern = @"Data/(.*)\.twb";
        Regex regex = new Regex(pattern);
        var match = regex.Match(workbookfilePath);
        string workbookName = match.Groups[1].Value; // The corresponding workbook name in Tableau Cloud. This is needed for API calls.
        string packagedWorkbookFilePath = @"Data/Superstore Analysis.twbx";

        // Project to store workbooks
        string projectName = "demo";

        TableauConnectionInfoExtractor connectionInfoExtractor = new TableauConnectionInfoExtractor();
        // TableauVisualizationInfoExtractor visualizationInfoExtractor = new TableauVisualizationInfoExtractor();

        try
        {
            // Step 1: Extract data source information from the specified Tableau workbook
            var dataSourceInfo = connectionInfoExtractor.ExtractDataSourceInfo(workbookfilePath);
            connectionInfoExtractor.SaveConnectionInfoToJSONAndExcel(dataSourceInfo, connectionInfoJsonFilePath, connectionInfoExcelFilePath);
            Console.WriteLine("Connection data extracted and saved successfully.");

            // Step 2: Extract visualization metadata
            // visualizationInfoExtractor.SetUp(packagedWorkbookFilePath, workbookName);
            // visualizationInfoExtractor.ExtractVisualizationInfo();
            // visualizationInfoExtractor.SaveVisualizationInfoToJSONAndExcel();
            Console.WriteLine("Visualization metadata extracted and saved successfully.");
        }
        catch (Exception ex) {
            Console.WriteLine(ex);
        }
    }
}