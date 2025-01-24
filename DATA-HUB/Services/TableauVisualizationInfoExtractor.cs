using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

public class TableauVisualizationInfoExtractor
{
    private const string TableauCloudUrl = "https://10ay.online.tableau.com";
    private const string ApiVersion = "3.24"; 
    private const string TokenName = "PLACEHOLDER"; // Generate from Tableau account
    private const string TokenSecret = "PLACEHOLDER"; // Generate from Tableau account
    private const string SiteId = "tableauuser6-41ea9b971e";

    public async Task<List<VisualizationInfo>> ExtractVisualizationInfo(string filePath, string workbookName) {
        // Set up
        List<VisualizationInfo> visualizationInfoList = new List<VisualizationInfo>();
        var xmlDoc = LoadXmlDocument(filePath);
        var worksheets = GetWorksheetElements(xmlDoc);
        var authToken = await Authenticate();
        var url = $"{TableauCloudUrl}/api/metadata/graphql";
        using var client = new HttpClient();

        foreach (var worksheet in worksheets) {
            var worksheetName = GetWorksheetName(worksheet);
            var visualizationTitle = GetVisualizationTitle(worksheet);
            var visualizationType = GetVisualizationType(worksheet);
            string tablesUsed;
            string columnsUsed;
            if (visualizationType == "No Visualization") {
                tablesUsed = "No tables used";
                columnsUsed = "No columns used";
            } else {
                var metadata = await QueryMetadataApi(authToken, workbookName, worksheetName, url, client);
                JObject jsonMetadata = JObject.Parse(metadata);
                JArray tablesAsJson = (JArray)jsonMetadata["data"]["workbooks"][0]["sheets"][0]["upstreamTables"];
                JArray columnsAsJson = (JArray)jsonMetadata["data"]["workbooks"][0]["sheets"][0]["upstreamColumns"];
                tablesUsed = string.Join(", ", tablesAsJson.Select(t => t["name"]));
                columnsUsed = string.Join(", ", columnsAsJson.Select(c => c["name"]));
            }
            visualizationInfoList.Add(new VisualizationInfo {
                WorksheetName = worksheetName,
                VisualizationTitle = visualizationTitle,
                VisualizationType = visualizationType,
                TablesUsed = tablesUsed,
                ColumnsUsed = columnsUsed
            });
        }
        return visualizationInfoList;
    }

    public async Task<string> QueryMetadataApi(string authToken, string workbookName, string worksheetName, string url, HttpClient client) {
        var query = $@"
        {{
            workbooks (filter: {{ name: ""{workbookName}"" }}) {{
                sheets (filter: {{name: ""{worksheetName}"" }}) {{
                    upstreamTables {{
                        name
                    }}
                    upstreamColumns {{
                        name
                    }}
                }}
            }}
        }}";
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

        var payload = new { query };
        var response = await client.PostAsync(url,
            new StringContent(JObject.FromObject(payload).ToString(), Encoding.UTF8, "application/json"));
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }

    public void SaveVisualizationInfoToJSONAndExcel(List<VisualizationInfo> visualizationInfoList, string jsonFilePath, string excelFilePath) {
        SaveToJSON(visualizationInfoList, jsonFilePath);
        SaveToExcel(visualizationInfoList, excelFilePath);
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
        if (vizExists == null) {
            return "No Visualization";
        } else {
            return worksheet.Descendants("mark").FirstOrDefault().Attribute("class").Value;
        }
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
        worksheet.Cells[1, 4].Value = "Tables Used";
        worksheet.Cells[1, 5].Value = "Columns Used";
        worksheet.Row(1).Style.Font.Bold = true;
    }

    private void PopulateExcelData(ExcelWorksheet worksheet, List<VisualizationInfo> visualizationInfoList) {
        for (int i = 0; i < visualizationInfoList.Count; i++) {
            worksheet.Cells[i + 2, 1].Value = visualizationInfoList[i].WorksheetName;
            worksheet.Cells[i + 2, 2].Value = visualizationInfoList[i].VisualizationTitle;
            worksheet.Cells[i + 2, 3].Value = visualizationInfoList[i].VisualizationType;
            worksheet.Cells[i + 2, 4].Value = string.Join(", ", visualizationInfoList[i].TablesUsed);
            worksheet.Cells[i + 2, 5].Value = string.Join(", ", visualizationInfoList[i].ColumnsUsed);
        }
    }

    private void SaveToJSON(List<VisualizationInfo> visualizationInfoList, string jsonFilePath) {
        string json = JsonSerializer.Serialize(visualizationInfoList, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonFilePath, json);
    }

    private async Task<string> Authenticate() {
        using var client = new HttpClient();
        var url = $"{TableauCloudUrl}/api/{ApiVersion}/auth/signin";
        var payload = new
        {
            credentials = new
            {
                personalAccessTokenName = TokenName,
                personalAccessTokenSecret = TokenSecret,
                site = new { contentUrl = SiteId }
            }
        };

        var request = new HttpRequestMessage(HttpMethod.Post, url) {
            Content = new StringContent(JObject.FromObject(payload).ToString(), Encoding.UTF8, "application/json")
        };
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var response = await client.SendAsync(request);
        response.EnsureSuccessStatusCode();

        var json = JObject.Parse(await response.Content.ReadAsStringAsync());
        var authToken = json["credentials"]["token"].ToString();

        return authToken;
    }

    public class VisualizationInfo {
        public string WorksheetName { get; set; }
        public string VisualizationTitle { get; set; }
        public string VisualizationType { get; set; }
        public string TablesUsed { get; set; }
        public string ColumnsUsed { get; set; }
    }
}
