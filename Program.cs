using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Dalek
{
    class Program
    {
        private static SheetsService service;

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = typeof(Program).Namespace;

        static void saySomethingAlready(string said, string spreadsheetId, string range)
        {
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS
            var oblist = new List<object>() { said };
            valueRange.Values = new List<IList<object>> { oblist };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
        }

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream(strings.PathJsonCredentials, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                //      string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                // automatically when the authorization flow completes for the first time.
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(strings.PathOAuthSessionToken, true)).Result;
                Console.WriteLine(strings.Program_Main_Credential_Saved_To_PathOAuthSessionToken);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "11fkJcoCfZPt3wUVnON9Y_jNg76EUsGwWm3Flz5QB6xU";
            String range = $"{ApplicationName}!A2:E2";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            var lastResponse=strings.Program_Main_We_will_explain_later;
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Dalek, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                    lastResponse = row[0].ToString();
                }
            }
            else
            {
                Console.WriteLine(strings.Program_Main_No_data_found_);
            }


            // Dalek speaks on the first sheet, new sheet, and newly added spreadsheet.
            var numSheets = 0;
            Console.WriteLine(strings.Program_Main_Dalek_speaks);
            Google.Apis.Sheets.v4.Data.Spreadsheet spread;
            try
            {
                spread = service.Spreadsheets.Get(spreadsheetId).Execute();
                var s = new Sheet();
                s.Properties = new SheetProperties();
                numSheets = spread.Sheets.Count + 1;
                s.Properties.Title =
                    $"{ApplicationName} {numSheets}  {DateTime.Now.Month}/{DateTime.Now.Day} {DateTime.Now.ToString("HH:mm")}";
                spread.Sheets.Add(s);

                // Add new Sheet
                var addSheetRequest = new AddSheetRequest();
                addSheetRequest.Properties = new SheetProperties();
                addSheetRequest.Properties.Title = s.Properties.Title;
                BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateSpreadsheetRequest.Requests = new List<Request>();
                batchUpdateSpreadsheetRequest.Requests.Add(new Request { AddSheet = addSheetRequest });
                var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
                batchUpdateRequest.Execute();

            // Dalek speaks on the first sheet
                saySomethingAlready(lastResponse, spreadsheetId, "!F5");
            // Dalek speaks on the new sheet, and newly added spreadsheet.
                saySomethingAlready(lastResponse, spreadsheetId, $"{s.Properties.Title}!F5");
            }
            catch (Exception ex)
            {
                Console.WriteLine(strings.Program_Main_This_is_obvious_);
                throw ex;
            }

            var myNewSheet = new Google.Apis.Sheets.v4.Data.Spreadsheet();
            myNewSheet.Properties = new SpreadsheetProperties();
            myNewSheet.Properties.Title = $"{ApplicationName} speaks to {credential.UserId} at  {DateTime.Now.ToString("HH:mm")}";
            var sheet = new Sheet();
            sheet.Properties = new SheetProperties();
            sheet.Properties.Title = $"{ApplicationName} speaks to {credential.UserId}";
            myNewSheet.Sheets = new List<Sheet>() { sheet };
            var newSheet = service.Spreadsheets.Create(myNewSheet).Execute();
            // Dalek speaks on the newly added spreadsheet.
            saySomethingAlready(lastResponse, newSheet.SpreadsheetId, "Dalek speaks to user!F5");

            /*
            var reqs = new BatchUpdateSpreadsheetRequest();
            reqs.Requests = new List<Request>();
            string[] colNames = new[] { "timestamp", "videoid", "videoname", "firstname", "lastname", "email" };

            // Create starting coordinate where data would be written to

            GridCoordinate gc = new GridCoordinate();
            gc.ColumnIndex = 0;
            gc.RowIndex = 0;
            gc.SheetId = newSheet.GetHashCode();

            Request rq = new Request();
            rq.UpdateCells = new UpdateCellsRequest();
            rq.UpdateCells.Start = gc;
            rq.UpdateCells.Fields = "*"; // needed by API, throws error if null

            // Assigning data to cells
            RowData rd = new RowData();
            List<CellData> lcd = new List<CellData>();
            foreach (String column in colNames)
            {
                ExtendedValue ev = new ExtendedValue();
                ev.StringValue = column;

                CellData cd = new CellData();
                cd.UserEnteredValue = ev;
                lcd.Add(cd);
            }
            rd.Values = lcd;

            // Put cell data into a row
            List<RowData> lrd = new List<RowData>();
            lrd.Add(rd);
            rq.UpdateCells.Rows = lrd;

            // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
            reqs.Requests.Add(rq);
            var r = service.Spreadsheets.BatchUpdate(reqs, newSheet.SpreadsheetId).Execute(); // Replace Spreadsheet.SpreadsheetId with your recently created spreadsheet ID
            */

            Console.Read();
        }
    }
}
// [END Horner's Dalek Introduction]
