using System.Xml.Linq;

public class TableauVisualizationMapper
{
    public Dictionary<string, VisualizationSource> MapWorksheetToDataSourceColumns(string twbFilePath)
    {
        var mapping = new Dictionary<string, VisualizationSource>();
        XDocument doc = XDocument.Load(twbFilePath);
        
        // Get all worksheets and their data source references
        var worksheets = doc.Descendants("worksheet");
        foreach (var worksheet in worksheets)
        {
            var worksheetName = worksheet.Attribute("name")?.Value;
            var datasourceRef = worksheet.Descendants("datasource-dependencies")
                .Select(d => d.Attribute("datasource")?.Value)
                .FirstOrDefault();
                
            mapping[worksheetName] = new VisualizationSource
            {
                WorksheetName = worksheetName,
                DataSourceId = datasourceRef,
                UsedColumns = GetUsedColumns(worksheet)
            };
        }
        return mapping;
    }

    private List<string> GetUsedColumns(XElement worksheet)
    {
        return worksheet.Descendants("column")
            .Select(c => c.Attribute("name")?.Value)
            .Where(name => !string.IsNullOrEmpty(name))
            .ToList();
    }
}

public class VisualizationSource
{
    public string WorksheetName { get; set; }
    public string DataSourceId { get; set; }
    public List<string> UsedColumns { get; set; }
}