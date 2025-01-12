using System.Xml.Linq;
using Newtonsoft.Json;
using TableauConverter.Models;

namespace TableauConverter.Services
{
    public class TableauXmlConverter
    {
        public TableauWorkbook ParseTableauXml(string xmlFilePath)
        {
            XDocument doc = XDocument.Load(xmlFilePath);
            var root = doc.Root;

            var workbook = new TableauWorkbook
            {
                WorkbookName = root.Attribute("name")?.Value ?? string.Empty
            };

            // Parse Datasources first to get column information
            ParseDatasources(root, workbook);

            // Parse Sheets with enhanced column tracking
            ParseSheets(root, workbook);

            // Parse Dashboards with column usage
            ParseDashboards(root, workbook);

            // Parse Custom Scripts with field references
            ParseCustomScripts(root, workbook);

            return workbook;
        }

        private void ParseDatasources(XElement root, TableauWorkbook workbook)
        {
            var datasources = root.Descendants("datasource");
            foreach (var datasource in datasources)
            {
                var ds = new Datasource
                {
                    Name = datasource.Attribute("name")?.Value ?? string.Empty,
                    Type = datasource.Attribute("type")?.Value ?? string.Empty,
                    ConnectionDetails = ParseConnectionDetails(datasource),
                    Columns = ExtractDatasourceColumns(datasource)
                };
                workbook.Datasources.Add(ds);
            }
        }

        private List<SheetColumn> ExtractDatasourceColumns(XElement datasource)
        {
            var columns = new List<SheetColumn>();
            var metadata = datasource.Descendants("column");
            
            foreach (var col in metadata)
            {
                columns.Add(new SheetColumn
                {
                    Name = col.Attribute("name")?.Value ?? string.Empty,
                    DataType = col.Attribute("datatype")?.Value ?? string.Empty,
                    Role = col.Attribute("role")?.Value ?? string.Empty,
                    Caption = col.Attribute("caption")?.Value ?? string.Empty
                });
            }
            
            return columns;
        }

        private ConnectionDetails ParseConnectionDetails(XElement datasource)
        {
            var connection = datasource.Element("connection");
            return new ConnectionDetails
            {
                Server = connection?.Attribute("server")?.Value ?? string.Empty,
                Database = connection?.Attribute("database")?.Value ?? string.Empty,
                ConnectionType = connection?.Attribute("class")?.Value ?? string.Empty
            };
        }

        private void ParseSheets(XElement root, TableauWorkbook workbook)
        {
            var worksheets = root.Descendants("worksheet");
            foreach (var worksheet in worksheets)
            {
                var sheet = new Sheet
                {
                    Name = worksheet.Attribute("name")?.Value ?? string.Empty,
                    Columns = ExtractColumns(worksheet),
                    Filters = ExtractFilters(worksheet),
                    Dependencies = ExtractDependencies(worksheet)
                };
                workbook.Sheets.Add(sheet);
            }
        }

        private List<SheetColumn> ExtractColumns(XElement worksheet)
        {
            var columns = new List<SheetColumn>();
            
            // Extract regular columns
            var regularColumns = worksheet.Descendants("column");
            foreach (var column in regularColumns)
            {
                var sheetColumn = new SheetColumn
                {
                    Name = column.Attribute("name")?.Value ?? string.Empty,
                    Caption = column.Attribute("caption")?.Value ?? string.Empty,
                    DataType = column.Attribute("datatype")?.Value ?? string.Empty,
                    AggregationType = column.Attribute("aggregation")?.Value ?? string.Empty,
                    Role = column.Attribute("role")?.Value ?? string.Empty
                };

                // Get column references
                var references = column.Descendants("reference")
                    .Select(r => r.Attribute("field")?.Value)
                    .Where(r => r != null)
                    .ToList();
                sheetColumn.References.AddRange(references);

                columns.Add(sheetColumn);
            }

            // Extract calculated fields
            var calculations = worksheet.Descendants("calculation");
            foreach (var calc in calculations)
            {
                var sheetColumn = new SheetColumn
                {
                    Name = calc.Attribute("name")?.Value ?? string.Empty,
                    Caption = calc.Attribute("caption")?.Value ?? string.Empty,
                    DataType = "Calculated",
                    Calculations = new List<string> { calc.Value }
                };
                columns.Add(sheetColumn);
            }

            return columns;
        }

        private List<Filter> ExtractFilters(XElement worksheet)
        {
            return worksheet.Descendants("filter")
                .Select(f => new Filter
                {
                    Field = f.Attribute("field")?.Value ?? string.Empty,
                    Type = f.Attribute("type")?.Value ?? string.Empty,
                    Value = f.Value
                })
                .ToList();
        }

        private List<string> ExtractDependencies(XElement worksheet)
        {
            return worksheet.Descendants("dependency")
                .Select(d => d.Attribute("name")?.Value)
                .Where(d => d != null)
                .ToList();
        }

        private void ParseDashboards(XElement root, TableauWorkbook workbook)
        {
            var dashboards = root.Descendants("dashboard");
            foreach (var dashboard in dashboards)
            {
                var dashboardData = new Dashboard
                {
                    Name = dashboard.Attribute("name")?.Value ?? string.Empty,
                    Size = new Size
                    {
                        Width = dashboard.Attribute("maxwidth")?.Value ?? string.Empty,
                        Height = dashboard.Attribute("maxheight")?.Value ?? string.Empty
                    },
                    Parameters = ExtractParameters(dashboard)
                };

                // Parse dashboard views with column usage
                foreach (var zone in dashboard.Descendants("zone"))
                {
                    var worksheetName = zone.Attribute("worksheet")?.Value;
                    if (!string.IsNullOrEmpty(worksheetName))
                    {
                        var worksheet = root.Descendants("worksheet")
                            .FirstOrDefault(w => w.Attribute("name")?.Value == worksheetName);

                        var view = new View
                        {
                            Name = zone.Attribute("name")?.Value ?? worksheetName,
                            Type = zone.Attribute("type")?.Value ?? "worksheet",
                            WorksheetName = worksheetName,
                            UsedColumns = worksheet != null ? ExtractColumns(worksheet) : new List<SheetColumn>()
                        };
                        dashboardData.Views.Add(view);
                    }
                }

                workbook.Dashboards.Add(dashboardData);
            }
        }

        private List<string> ExtractParameters(XElement dashboard)
        {
            return dashboard.Descendants("parameter")
                .Select(p => p.Attribute("name")?.Value)
                .Where(p => p != null)
                .ToList();
        }

        private void ParseCustomScripts(XElement root, TableauWorkbook workbook)
        {
            var scripts = root.Descendants("script");
            foreach (var script in scripts)
            {
                workbook.CustomScripts.Add(new CustomScript
                {
                    Name = script.Attribute("name")?.Value ?? string.Empty,
                    Content = script.Value,
                    ReferencedFields = ExtractScriptReferences(script)
                });
            }
        }

        private List<string> ExtractScriptReferences(XElement script)
        {
            // Extract field references from script content using regex or parsing
            // This is a simplified version
            return script.Value
                .Split(new[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
        }

        public void ConvertToJson(string xmlFilePath, string outputJsonPath)
        {
            var workbook = ParseTableauXml(xmlFilePath);
            var jsonString = JsonConvert.SerializeObject(workbook, Formatting.Indented);
            File.WriteAllText(outputJsonPath, jsonString);
        }
    }
}
