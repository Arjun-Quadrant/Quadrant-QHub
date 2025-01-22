class Program
{
    static void Main(string[] args) {
        // Define file paths for the Tableau workbook and output files
        string filePath = @"Data/multiple datasources.twb"; // Path to the Tableau workbook file (adjust as needed)
        string connectionInfoJsonFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\connection info.json"; // Path for saving connection info in JSON format
        string connectionInfoExcelFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\connection info.xlsx"; // Path for saving connection info in Excel format
        string visualizationInfoJsonFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\visualization info.json"; // Path for saving visualization metadata in JSON format
        string visualizationInfoExcelFilePath = @"C:\Users\ArjunNarendra(Quadra\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis\visualization info.xlsx"; // Path for saving visualization metadata in Excel format

        // Instantiate the TableauInfoExtractor object to extract and save information
        TableauConnectionInfoExtractor connectionInfoExtractor = new TableauConnectionInfoExtractor();

        TableauVisualizationInfoExtractor visualizationInfoExtractor = new TableauVisualizationInfoExtractor();

        try
        {
            // Extract data source information from the specified Tableau workbook
            var dataSourceInfo = connectionInfoExtractor.ExtractDataSourceInfo(filePath);
            
            // Save the extracted data source information to JSON and Excel files
            connectionInfoExtractor.SaveConnectionInfoToJSONAndExcel(dataSourceInfo, connectionInfoJsonFilePath, connectionInfoExcelFilePath);
            Console.WriteLine("Connection data extracted and saved successfully.");

            var visualizationInfo = visualizationInfoExtractor.ExtractVisualizationInfo(filePath);
            visualizationInfoExtractor.SaveVisualizationInfoToJSONAndExcel(visualizationInfo, visualizationInfoJsonFilePath, visualizationInfoExcelFilePath);
            Console.WriteLine("Visualization metadata extracted and saved successfully.");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during extraction or saving
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}