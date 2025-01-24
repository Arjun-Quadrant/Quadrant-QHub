using System.Text.RegularExpressions;

class Program
{
    public static async Task Main(string[] args) {
        // Define file paths for the output files
        string connectionInfoJsonFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\connection info.json"; 
        string connectionInfoExcelFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\connection info.xlsx"; 
        string visualizationInfoJsonFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\visualization info.json"; 
        string visualizationInfoExcelFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\visualization info.xlsx"; 

        // Get information about the workbook
        string filePath = @"Data/Netflix Titles.twb"; // Path to the Tableau workbook file (adjust as needed). This is needed for XML extraction.
        string pattern = @"Data/(.*)\.twb";
        Regex regex = new Regex(pattern);
        var match = regex.Match(filePath);
        string workbookName = match.Groups[1].Value; // The corresponding workbook name in Tableau Cloud. This is needed for API calls.

        TableauConnectionInfoExtractor connectionInfoExtractor = new TableauConnectionInfoExtractor();
        TableauVisualizationInfoExtractor visualizationInfoExtractor = new TableauVisualizationInfoExtractor();

        try
        {
            // Extract data source information from the specified Tableau workbook
            var dataSourceInfo = connectionInfoExtractor.ExtractDataSourceInfo(filePath);
            
            // Save the extracted data source information to JSON and Excel files
            connectionInfoExtractor.SaveConnectionInfoToJSONAndExcel(dataSourceInfo, connectionInfoJsonFilePath, connectionInfoExcelFilePath);
            Console.WriteLine("Connection data extracted and saved successfully.");

            // Extract visualization metadata
            var visualizationInfo = await visualizationInfoExtractor.ExtractVisualizationInfo(filePath, workbookName);
            visualizationInfoExtractor.SaveVisualizationInfoToJSONAndExcel(visualizationInfo, visualizationInfoJsonFilePath, visualizationInfoExcelFilePath);
            Console.WriteLine("Visualization metadata extracted and saved successfully.");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during extraction or saving
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}