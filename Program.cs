using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("🔄 Connecting to Google Sheets...");

        string[] Scopes = { SheetsService.Scope.Spreadsheets };
        string ApplicationName = "Google Sheets API Test";
        string spreadsheetId = "1jzRGk_FlT33N2Kop3h1GXuDsqFA2JSRakIZz-bmv6hs"; // 🔹 Replace with actual Google Sheet ID
        string databaseSheet = "Database";
        string generationTrackingSheet = "Generation Tracking";

        GoogleCredential credential;
        using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(Environment.GetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS_JSON"))))

        {
            credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
        }

        var service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });

        // ✅ Step 1: Get selected Date from A1 in "Generation Tracking"
        string targetDate = await GetCellValue(service, spreadsheetId, generationTrackingSheet, "A1");
        if (string.IsNullOrEmpty(targetDate))
        {
            Console.WriteLine("❌ No date found in A1 of Generation Tracking.");
            return;
        }
        Console.WriteLine($"📅 Target Date: {targetDate}");

        // ✅ Step 2: Get selected Area from G1
        string selectedArea = await GetCellValue(service, spreadsheetId, generationTrackingSheet, "G1");

        // ✅ Step 3: Read all data from "Database"
        var databaseData = (await GetSheetData(service, spreadsheetId, databaseSheet)).ToList();
        if (databaseData.Count < 2)
        {
            Console.WriteLine("❌ No data found in Database.");
            return;
        }

        // ✅ Step 4: Process Data
        var (summaryData, belowAvgClients, aboveAvgClients, areaClients, cleaningDoneClients) = 
            ProcessGenerationTrackingData(databaseData, targetDate, selectedArea);

        // ✅ Step 5: Write Processed Data to "Generation Tracking"
        await WriteMultipleRanges(service, spreadsheetId, generationTrackingSheet, 
            "D1", new List<IList<object>> { summaryData[0] }, // Total Count
            "D2", new List<IList<object>> { summaryData[1] }, // Average Unit/kW
            "H1", new List<IList<object>> { summaryData[2] }); // Area-Specific Average

        await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "A5", belowAvgClients);
        await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "A42", aboveAvgClients);
        await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "G3", areaClients);
        await WriteToGoogleSheet(service, spreadsheetId, generationTrackingSheet, "J3", cleaningDoneClients);

        Console.WriteLine("✅ Data successfully updated in Generation Tracking!");
    }

    // 📌 Function to Get a Single Cell Value
    static async Task<string> GetCellValue(SheetsService service, string spreadsheetId, string sheetName, string cell)
    {
        string range = $"{sheetName}!{cell}";
        var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
        ValueRange response = await request.ExecuteAsync();
        return response.Values != null && response.Values.Count > 0 && response.Values[0].Count > 0 
    ? response.Values[0][0]?.ToString() ?? "" 
    : "";
    }

    // 📌 Function to Read All Data from a Sheet
    static async Task<List<IList<object>>> GetSheetData(SheetsService service, string spreadsheetId, string sheetName)
    {
        string range = $"{sheetName}!A1:Z"; // Adjust the range if necessary
        var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
        ValueRange response = await request.ExecuteAsync();
        return response.Values?.ToList() ?? new List<IList<object>>();
    }

    // 📌 Function to Process Generation Tracking Data
    static (List<IList<object>>, List<IList<object>>, List<IList<object>>, List<IList<object>>, List<IList<object>>) 
        ProcessGenerationTrackingData(List<IList<object>> databaseData, string targetDate, string selectedArea)
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

    // 📌 Function to Write Data to a Sheet
    static async Task WriteToGoogleSheet(SheetsService service, string spreadsheetId, string sheetName, string startCell, List<IList<object>> data)
    {
        string range = $"{sheetName}!{startCell}";
        var valueRange = new ValueRange { Values = data };

        var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        await updateRequest.ExecuteAsync();
    }

    // 📌 Function to Write Multiple Ranges in One API Call
    static async Task WriteMultipleRanges(SheetsService service, string spreadsheetId, string sheetName, 
        string range1, IList<IList<object>> data1, 
        string range2, IList<IList<object>> data2, 
        string range3, IList<IList<object>> data3)
    {
        var updateData = new List<ValueRange>
        {
            new ValueRange { Range = $"{sheetName}!{range1}", Values = data1 },
            new ValueRange { Range = $"{sheetName}!{range2}", Values = data2 },
            new ValueRange { Range = $"{sheetName}!{range3}", Values = data3 }
        };

        var requestBody = new BatchUpdateValuesRequest { ValueInputOption = "USER_ENTERED", Data = updateData };
        var request = service.Spreadsheets.Values.BatchUpdate(requestBody, spreadsheetId);
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
}
