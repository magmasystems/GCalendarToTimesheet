using System;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GCalendarToTimesheet
{
    public class GoogleSheetsDriver
    {
        private SheetsService Service { get; }
        
        public GoogleSheetsDriver(UserCredential credential, string applicationName)
        {
            this.Service = this.GetGoogleSheetsService(credential, applicationName);
        }

        private SheetsService GetGoogleSheetsService(UserCredential credential, string applicationName)
        {
            // Create Google Sheets API service.
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
        }

        public ValueRange GetSheet(string spreadsheetId, string range)
        {
            // const string range = "Mosaik!A4:C37";
            var request = this.Service.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = request.Execute();
            return response;
        }
        
        public Spreadsheet GetSheets(string spreadsheetId)
        {
            var request = this.Service.Spreadsheets.Get(spreadsheetId);
            var response = request.Execute();
            return response;
        }

        public int GetLastRowInSheet(string spreadsheetId, Sheet sheet)
        {
            var response = this.GetSheet(spreadsheetId, $"{sheet.Properties.Title}!A1:D5000");
            var idxLastRow = response.Values.Count;
            return idxLastRow;
        }
        
        public int AppendToSheet(string spreadsheetId, Sheet sheet, string range, List<IList<object>> data)
        {
            try
            {
                var body = new ValueRange { MajorDimension = "ROWS", Values = data };
                var request = this.Service.Spreadsheets.Values.Append(body, spreadsheetId, range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                
                var response = request.Execute();
                return response.Updates.UpdatedRows.GetValueOrDefault();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}