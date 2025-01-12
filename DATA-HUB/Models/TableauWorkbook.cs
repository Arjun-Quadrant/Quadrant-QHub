using System.Xml.Serialization;

namespace TableauConverter.Models
{
    public class TableauWorkbook
    {
        public string WorkbookName { get; set; }
        public List<Sheet> Sheets { get; set; } = new();
        public List<Dashboard> Dashboards { get; set; } = new();
        public List<Datasource> Datasources { get; set; } = new();
        public List<CustomScript> CustomScripts { get; set; } = new();
    }

    public class Sheet
    {
        public string Name { get; set; }
        public List<SheetColumn> Columns { get; set; } = new();
        public List<Filter> Filters { get; set; } = new();
        public List<string> Dependencies { get; set; } = new();
    }

    public class SheetColumn
    {
        public string Name { get; set; }
        public string Caption { get; set; }
        public string AggregationType { get; set; }
        public string DataType { get; set; }
        public List<string> Calculations { get; set; } = new();
        public List<string> References { get; set; } = new();
        public string Role { get; set; }
    }

    public class Filter
    {
        public string Field { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }

    public class Dashboard
    {
        public string Name { get; set; }
        public Size Size { get; set; }
        public List<View> Views { get; set; } = new();
        public List<string> Parameters { get; set; } = new();
    }

    public class Size
    {
        public string Width { get; set; }
        public string Height { get; set; }
    }

    public class View
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public List<SheetColumn> UsedColumns { get; set; } = new();
        public string WorksheetName { get; set; }
    }

    public class Datasource
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public ConnectionDetails ConnectionDetails { get; set; }
        public List<SheetColumn> Columns { get; set; } = new();
    }

    public class ConnectionDetails
    {
        public string Server { get; set; }
        public string Database { get; set; }
        public string ConnectionType { get; set; }
    }

    public class CustomScript
    {
        public string Name { get; set; }
        public string Content { get; set; }
        public List<string> ReferencedFields { get; set; } = new();
    }
}
