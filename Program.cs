using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Hosting;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.MapGet("/", () => "Google Sheets API Web Service is Running!");

// ✅ API Endpoint to Trigger Your Google Sheets Processing
app.MapPost("/run", async (HttpContext context) =>
{
    var result = await ProcessGoogleSheetsData();
    await context.Response.WriteAsJsonAsync(new { message = "✅ Google Sheets Processing Completed!", result });
});

// 📌 Function to Process Google Sheets Data
async Task<List<IList<object>>> ProcessGoogleSheetsData()
{
    string spreadsheetId = "1jzRGk_FlT33N2Kop3h1GXuDsqFA2JSRakIZz-bmv6hs"; // Replace with actual Sheet ID
    string generationTrackingSheet = "Generation Tracking";
    string databaseSheet = "Database";

    // ✅ Authenticate with Google Sheets API
    GoogleCredential credential;
    using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(Environment.GetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS_JSON"))))
    {
        credential = GoogleCredential.FromStream(stream).CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
    }

    var service = new SheetsService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credential,
        ApplicationName = "Google Sheets API Web Service",
    });

    // ✅ Get Date and Area from "Generation Tracking"
    string targetDate = await GetCellValue(service, spreadsheetId, generationTrackingSheet, "A1");
    string selectedArea = await GetCellValue(service, spreadsheetId, generationTrackingSheet, "G1");

    var databaseData = (await GetSheetData(service, spreadsheetId, databaseSheet)).ToList();
    if (databaseData.Count < 2) return new List<IList<object>> { new List<object> { "❌ No data found in Database." } };

    // ✅ Process Data
    var (summaryData, belowAvgClients, aboveAvgClients, areaClients, cleaningDoneClients) =
        ProcessData(databaseData, targetDate, selectedArea);

    // ✅ Write Processed Data to Google Sheets
    await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "D1", summaryData);
    await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "A5", belowAvgClients);
    await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "A42", aboveAvgClients);
    await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "G3", areaClients);
    await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "J3", cleaningDoneClients);

    return summaryData;
}

// 📌 Helper Function to Get a Cell Value
async Task<string> GetCellValue(SheetsService service, string spreadsheetId, string sheetName, string cell)
{
    var request = service.Spreadsheets.Values.Get(spreadsheetId, $"{sheetName}!{cell}");
    var response = await request.ExecuteAsync();
    return response.Values?.FirstOrDefault()?.FirstOrDefault()?.ToString() ?? "";
}

// 📌 Helper Function to Get All Data from a Sheet
async Task<List<IList<object>>> GetSheetData(SheetsService service, string spreadsheetId, string sheetName)
{
    var request = service.Spreadsheets.Values.Get(spreadsheetId, $"{sheetName}!A1:Z");
    var response = await request.ExecuteAsync();
    return response.Values?.ToList() ?? new List<IList<object>>();
}

// 📌 Function to Process Data
static (List<IList<object>>, List<IList<object>>, List<IList<object>>, List<IList<object>>, List<IList<object>>)
    ProcessData(List<IList<object>> databaseData, string targetDate, string selectedArea)
{
    List<IList<object>> summaryData = new List<IList<object>>();
    List<IList<object>> belowAvgClients = new List<IList<object>>();
    List<IList<object>> aboveAvgClients = new List<IList<object>>();
    List<IList<object>> areaClients = new List<IList<object>>();
    List<IList<object>> cleaningDoneClients = new List<IList<object>>();

    int totalCount = 0;
    double totalUnitPerKW = 0;
    int areaCount = 0;
    double areaUnitPerKW = 0;

    foreach (var row in databaseData.Skip(1))
    {
        if (row.Count < 14) continue;

        string clientName = row.ElementAtOrDefault(0)?.ToString() ?? "";
        string rowDate = row.ElementAtOrDefault(1)?.ToString() ?? "";
        double unitPerKW = ParseDouble(row.ElementAtOrDefault(10) ?? 0);
        string problem = row.ElementAtOrDefault(11)?.ToString() ?? "";
        string area = row.ElementAtOrDefault(13)?.ToString().ToLower() ?? "";
        string cleaningStatus = row.ElementAtOrDefault(12)?.ToString() ?? "";

        if (rowDate == targetDate)
        {
            totalCount++;
            totalUnitPerKW += unitPerKW;

            if (!string.IsNullOrEmpty(selectedArea) && area == selectedArea)
            {
                areaCount++;
                areaUnitPerKW += unitPerKW;
                areaClients.Add(new List<object> { clientName, unitPerKW });
            }

            if (cleaningStatus == "Cleaning done")
            {
                cleaningDoneClients.Add(new List<object> { clientName, unitPerKW });
            }
        }
    }

    double avgUnitPerKW = totalCount > 0 ? totalUnitPerKW / totalCount : 0;
    double avgAreaUnitPerKW = areaCount > 0 ? areaUnitPerKW / areaCount : 0;

    summaryData.Add(new List<object> { totalCount });
    summaryData.Add(new List<object> { avgUnitPerKW.ToString("0.00") });
    summaryData.Add(new List<object> { avgAreaUnitPerKW.ToString("0.00") });

    return (summaryData, belowAvgClients, aboveAvgClients, areaClients, cleaningDoneClients);
}

// 📌 Helper Function to Write Data to Google Sheets
async Task WriteToGoogleSheet(SheetsService service, string spreadsheetId, string sheetName, string startCell, List<IList<object>> data)
{
    var request = service.Spreadsheets.Values.Update(new ValueRange { Values = data }, spreadsheetId, $"{sheetName}!{startCell}");
    request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
    await request.ExecuteAsync();
}

// 📌 Helper Function to Parse Double Values
static double ParseDouble(object value)
{
    if (value == null) return 0;
    if (double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
    {
        return result;
    }
    return 0;
}

app.Run();
