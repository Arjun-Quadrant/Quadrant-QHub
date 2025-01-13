using System.Xml.Linq;

public class TableauColumnExtractor
{
    public Dictionary<string, string> ExtractDataSourceColumns(string twbFilePath)
{
    var columns = new Dictionary<string, string>();
    XDocument doc = XDocument.Load(twbFilePath);
    
    var datasource = doc.Descendants("datasource").First();
    var columnList = datasource.Descendants("columns");
    var columnElements = columnList.Descendants("column");
    
    foreach (var column in columnElements) {
        var name = column.Attribute("name")?.Value;
        var datatype = column.Attribute("datatype")?.Value;
        if (name != null && datatype != null) {
            columns.TryAdd(name, datatype);
        }
    }
    return columns;
}

public class DataSourceColumn
{
    public string Name { get; set; }
    public string DataType { get; set; }

}

     public List<ColumnUsage> GetColumnDetailsForSheet(string twbFilePath, string sheetName, string columnName)
    {
        var columnUsage = ExtractColumnUsage(twbFilePath);
        
        return columnUsage.Where(c => 
            c.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase) &&
            c.ColumnName.Contains(columnName, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }
    public List<ColumnUsage> ExtractColumnUsage(string twbFilePath)
    {
        var columnUsage = new List<ColumnUsage>();
        XDocument doc = XDocument.Load(twbFilePath);
        
        // Get columns from worksheets
        var worksheets = doc.Descendants("worksheet");
        foreach (var worksheet in worksheets)
        {
            var sheetName = worksheet.Attribute("name")?.Value;
            
            ExtractFromEncodings(worksheet, sheetName, columnUsage);
            ExtractFromColumns(worksheet, sheetName, columnUsage);
            ExtractFromCalculations(worksheet, sheetName, columnUsage);
        }
        
        return columnUsage;
    }

    private void ExtractFromEncodings(XElement worksheet, string sheetName, List<ColumnUsage> usage)
    {
        var encodings = worksheet.Descendants("encoding");
        foreach (var encoding in encodings)
        {
            var field = encoding.Attribute("field")?.Value;
            var role = encoding.Attribute("role")?.Value;
            var type = encoding.Attribute("type")?.Value;
            
            if (!string.IsNullOrEmpty(field))
            {
                usage.Add(new ColumnUsage
                {
                    SheetName = sheetName,
                    ColumnName = field,
                    Usage = "Visual Encoding",
                    Role = role,
                    Type = type
                });
            }
        }
    }

    private void ExtractFromColumns(XElement worksheet, string sheetName, List<ColumnUsage> usage)
    {
        var columns = worksheet.Descendants("column");
        foreach (var column in columns)
        {
            var name = column.Attribute("name")?.Value;
            var role = column.Attribute("role")?.Value;
            
            if (!string.IsNullOrEmpty(name))
            {
                usage.Add(new ColumnUsage
                {
                    SheetName = sheetName,
                    ColumnName = name,
                    Usage = "Column Definition",
                    Role = role
                });
            }
        }
    }

    private void ExtractFromCalculations(XElement worksheet, string sheetName, List<ColumnUsage> usage)
    {
        var calculations = worksheet.Descendants("calculation");
        foreach (var calc in calculations)
        {
            var formula = calc.Value;
            var name = calc.Attribute("name")?.Value;
            
            usage.Add(new ColumnUsage
            {
                SheetName = sheetName,
                ColumnName = name,
                Usage = "Calculation",
                Formula = formula
            });
        }
    }
}


public class ColumnUsage
{
    public string SheetName { get; set; }
    public string ColumnName { get; set; }
    public string Usage { get; set; }
    public string Role { get; set; }
    public string Type { get; set; }
    public string Formula { get; set; }
}
