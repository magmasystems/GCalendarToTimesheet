using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;

// ReSharper disable InconsistentNaming
// ReSharper disable StringLiteralTypo

namespace GCalendarToTimesheet
{
    internal class Program
    {
        #region Variables
        // https://console.cloud.google.com/apis/credentials/consent/edit?project=gcalendartotimesheet
        
        // If modifying these scopes, delete your previously saved credentials at ~/.credentials/calendar-dotnet-quickstart.json
        private readonly string[] Scopes = { CalendarService.Scope.CalendarReadonly, SheetsService.Scope.Spreadsheets };
        private const string ApplicationName = "GCalendarToTimesheet";
        private string CaaSClientsSpreadsheetId { get; set; }
        
        private DateTime? StartDate { get; set; }
        private DateTime? EndDate { get; set; }
        private bool Verbose { get; set; }
        private bool NoSheets { get; set; }
        private bool CompressDailyMultipleTasks { get; set; }
        private IEnumerable<string> ClientFilters { get; set; }

        private GoogleCalendarDriver GoogleCalendarDriver { get; set; }
        private GoogleSheetsDriver GoogleSheetsDriver { get; set; }
        private Events CalendarEvents { get; set; }

        private readonly HashSet<string> ExistingClients = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private IDictionary<string, Sheet> ClientToSheetMap { get; set; }
        private Dictionary<string, List<GCalEventInfo>> ClientToEventsMap { get; set; }
        #endregion

        #region Main
        private static void Main(string[] args)
        {
            var program = new Program();

            program.ParseCommandLine(args);
            program.Initialize();

            program.ClientToEventsMap = program.ProcessCalendarEvents(program.CalendarEvents);
            program.DumpTimesheetEvents(program.ClientToEventsMap);

            if (!program.NoSheets)
            {
                program.WriteClientTimeToGoogleSheets(program.ClientToEventsMap);
            }

            Console.WriteLine("Press ENTER to quit");
            Console.ReadLine();
        }
        #endregion

        #region Command Line
        private void ParseCommandLine(IReadOnlyList<string> args)
        {
            if (args.Count <= 0)
                return;
            
            for (var i = 0; i < args.Count; i++)
            {
                switch (args[i].ToLower())
                {
                    case "-start":
                        this.StartDate = DateTime.Parse(args[++i]);
                        break;

                    case "-end":
                        this.EndDate = DateTime.Parse(args[++i]).AddDays(1); // use midnight of the next day
                        break;

                    case "-verbose":
                        this.Verbose = true;
                        break;

                    case "-clients":
                        this.ClientFilters = args[++i].Split(',').Select(s => s.ToLower());
                        break;
                    
                    case "-nosheet":
                        this.NoSheets = true;
                        break;
                    
                    case "-compress":
                        this.CompressDailyMultipleTasks = true;
                        break;
                }
            }
        }
        #endregion

        #region Initialization
        private void Initialize()
        {
            var config = this.InitConfiguration();
            this.CaaSClientsSpreadsheetId = config.GetSection("Application:timesheetId").Value;

            // Get the list of existing clients
            var clientList = config.GetSection("Application:clients").AsEnumerable();
            foreach (var kvp in clientList.Where(x => x.Value != null))
                this.ExistingClients.Add(kvp.Value);
            
            var apiDriver = new GoogleApiDriver();
            var credentials = apiDriver.Authorize(Scopes);
            
            this.GoogleCalendarDriver = new GoogleCalendarDriver(credentials, ApplicationName);
            this.GoogleSheetsDriver = new GoogleSheetsDriver(credentials, ApplicationName);

            this.ClientToSheetMap = this.CreateSheetsMap();
            this.CalendarEvents = this.GoogleCalendarDriver.GetCalendarEvents(this.StartDate, this.EndDate);
        }

        private IConfiguration InitConfiguration()
        {
            var config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();
            return config;
        }
        #endregion
        
        #region Calendar Events
        private Dictionary<string, List<GCalEventInfo>> ProcessCalendarEvents(Events events)
        {
            if (events.Items == null || events.Items.Count == 0)
                return null;
            
            var clientMap = new Dictionary<string, List<GCalEventInfo>>();
            foreach (var client in ExistingClients)
            {
                clientMap[client] = new List<GCalEventInfo>();
            }

            foreach (var eventItem in events.Items)
            {
                if (eventItem.End.DateTime == null || eventItem.Start.DateTime == null)
                    continue;

                var start = eventItem.Start.DateTime.Value;
                var summary = eventItem.Summary.ToLower();

                if (summary.IndexOf("canceled", StringComparison.Ordinal) >= 0 ||
                    summary.IndexOf("cancelled", StringComparison.Ordinal) >= 0)
                    continue;

                var clientName = summary.Split(' ', ':', '-', '.', '<', '>', ',').FirstOrDefault(s => ExistingClients.Contains(s));
                if (clientName == null)
                {
                    var emailComponents = eventItem.Organizer.Email.Split('.', '@');
                    clientName = emailComponents[^2].ToLower();
                }

                if (!clientMap.ContainsKey(clientName))
                {
                    Console.WriteLine($"Cannot find client {clientName} for event [{eventItem.Summary}]");
                }
                else if (this.ClientFilters != null && !this.ClientFilters.Contains(clientName))
                {
                    // If we have a filter of client names and this client is not in the list, then don't do anything
                    // ReSharper disable once RedundantJumpStatement
                    continue;
                }
                else if (eventItem.DidIDecline())
                {
                    // If I (magmasystems@gmail.com) declined the event, then ignore it.
                    // ReSharper disable once RedundantJumpStatement
                    continue;
                }
                else
                {
                    clientMap[clientName].Add(new GCalEventInfo
                    {
                        Date = start,
                        Duration = (eventItem.End.DateTime - eventItem.Start.DateTime).GetValueOrDefault(),
                        Summary = summary,
                        NativeEvent = eventItem
                    });
                }
            }

            return clientMap;
        }

        private void DumpTimesheetEvents(Dictionary<string, List<GCalEventInfo>> clientMap)
        {
            foreach (var (key, value) in clientMap)
            {
                if (value.Count == 0) // No events for this client?
                    continue;

                // Calculate how much time we spent this week for the client
                var totalDuration = new TimeSpan();
                foreach (var e in value)
                {
                    totalDuration += e.Duration;
                    Console.WriteLine($"Client: {key}, {e}");
                    if (this.Verbose)
                    {
                        Console.WriteLine(e.ToStringEx());
                        Console.WriteLine();
                    }
                }

                Console.WriteLine($"Total: {totalDuration}\n");
            }
        }
        #endregion
        
        #region Sheets Support
        private Dictionary<string, Sheet> CreateSheetsMap()
        {
            var clientToSheetMap = new Dictionary<string, Sheet>(StringComparer.OrdinalIgnoreCase);
            var sheets = this.GoogleSheetsDriver.GetSheets(CaaSClientsSpreadsheetId);

            foreach (var sheet in sheets.Sheets)
            {
                if (this.ExistingClients.Contains(sheet.Properties.Title))
                {
                    clientToSheetMap[sheet.Properties.Title] = sheet;
                }
            }

            return clientToSheetMap;
        }
        
        private void WriteClientTimeToGoogleSheets(Dictionary<string, List<GCalEventInfo>> clientMap)
        {
            foreach (var (clientName, calEvents) in clientMap)
            {
                if (calEvents.Count == 0) // No events for this client?
                    continue;

                // Make sure that there is a worksheet for this client
                if (!this.ClientToSheetMap.TryGetValue(clientName, out var sheet))
                    continue;

                // Create a nested list of the new event data that we want to append to the sheet
                var data = new List<IList<object>>();  // RowData<ColumnData>
                var lastDateAdded = DateTime.MinValue;
                foreach (var e in calEvents)
                {
                    // If we are coalescing multiple events for one date into a single entry, see if we are dealing with the same data as the last one added
                    if (this.CompressDailyMultipleTasks && e.Date.Month == lastDateAdded.Month && e.Date.Day == lastDateAdded.Day)
                    {
                        // Get the last item and add this event's duration and append the summary.
                        var lastItem = data[^1];
                        lastItem[1] = (double) lastItem[1] + e.Duration.TotalMinutes / 60.0;
                        lastItem[3] = (string) lastItem[3] + ", " + e.Summary;
                    }
                    else
                    {
                        // A new event. Add it to the column data for this row.
                        data.Add(new List<object>
                        {
                            $"{e.Date.Month}/{e.Date.Day:00}", 
                            e.Duration.TotalMinutes / 60.0,
                            null, 
                            e.Summary
                        });
                    }

                    lastDateAdded = e.Date;
                }
                
                // We need to find the last row of the sheet, leave a blank line after the last entry, and create a custom range.
                const int number_of_row_to_leave_blank = 2;
                var idxLastRow = this.GoogleSheetsDriver.GetLastRowInSheet(CaaSClientsSpreadsheetId, sheet);
                var range = $"{sheet.Properties.Title}!A{idxLastRow + number_of_row_to_leave_blank}:D{idxLastRow + number_of_row_to_leave_blank + data.Count}";
                
                // Append the calendar event data
                var rowsInserted = this.GoogleSheetsDriver.AppendToSheet(CaaSClientsSpreadsheetId, sheet, range, data);
                if (rowsInserted != calEvents.Count)
                {
                }
            }
        }
        #endregion
    }
}