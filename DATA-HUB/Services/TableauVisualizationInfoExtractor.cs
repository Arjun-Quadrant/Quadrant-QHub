using System.Text.Json;
using System.Xml.Linq;
using OfficeOpenXml;

public class TableauVisualizationInfoExtractor
{
    public class VisualizationInfo {
        public string WorksheetName { get; set; }
        public string VisualizationTitle { get; set; }
        public string VisualizationType { get; set; }
    }

    public List<VisualizationInfo> ExtractVisualizationInfo(string filePath) {
        List<VisualizationInfo> visualizationInfoList = new List<VisualizationInfo>();
        var xmlDoc = LoadXmlDocument(filePath);
        var worksheets = GetWorksheetElements(xmlDoc);
        foreach (var worksheet in worksheets) {
            var worksheetName = GetWorksheetName(worksheet);
            var visualizationTitle = GetVisualizationTitle(worksheet);
            var visualizationType = GetVisualizationType(worksheet);
            visualizationInfoList.Add(new VisualizationInfo {
                WorksheetName = worksheetName,
                VisualizationTitle = visualizationTitle,
                VisualizationType = visualizationType
            });
        }
        return visualizationInfoList;
    }

    private XDocument LoadXmlDocument(string filePath) {
        return XDocument.Load(filePath);
    }

    private IEnumerable<XElement> GetWorksheetElements(XDocument xmlDoc) {
        return xmlDoc.Descendants("worksheet");
    }

    private string GetWorksheetName(XElement worksheet) {
        return worksheet.Attribute("name").Value;
    }
    
    private string GetVisualizationTitle(XElement worksheet) {
        var title = worksheet.Descendants("run").FirstOrDefault();
        if (title == null) {
            return "No Title";
        }
        return worksheet.Descendants("run").FirstOrDefault().Value;
    }

    private string GetVisualizationType(XElement worksheet) {
        var vizExists = worksheet.Descendants("datasource").FirstOrDefault();
        Console.WriteLine(vizExists);
        if (vizExists == null) {
            return "No Visualization";
        } else {
            return worksheet.Descendants("mark").FirstOrDefault().Attribute("class").Value;
        }
    }

    public void SaveVisualizationInfoToJSONAndExcel(List<VisualizationInfo> visualizationInfoList, string jsonFilePath, string excelFilePath) {
        SaveToJSON(visualizationInfoList, jsonFilePath);
        SaveToExcel(visualizationInfoList, excelFilePath);
    }

    private void SaveToExcel(List<VisualizationInfo> visualizationInfoList, string excelFilePath) {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("VisualizationInfo");
        AddExcelHeaders(worksheet);
        PopulateExcelData(worksheet, visualizationInfoList);
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        package.SaveAs(excelFilePath);
    }

    private void AddExcelHeaders(ExcelWorksheet worksheet) {
        worksheet.Cells[1, 1].Value = "Worksheet Name";
        worksheet.Cells[1, 2].Value = "Visualization Title";
        worksheet.Cells[1, 3].Value = "Visualization Type";
        worksheet.Row(1).Style.Font.Bold = true;
    }

    private void PopulateExcelData(ExcelWorksheet worksheet, List<VisualizationInfo> visualizationInfoList) {
        for (int i = 0; i < visualizationInfoList.Count; i++) {
            worksheet.Cells[i + 2, 1].Value = visualizationInfoList[i].WorksheetName;
            worksheet.Cells[i + 2, 2].Value = visualizationInfoList[i].VisualizationTitle;
            worksheet.Cells[i + 2, 3].Value = visualizationInfoList[i].VisualizationType;
        }
    }

    private void SaveToJSON(List<VisualizationInfo> visualizationInfoList, string jsonFilePath) {
        string json = JsonSerializer.Serialize(visualizationInfoList, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonFilePath, json);
    }
}