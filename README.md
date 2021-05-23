# GCalendarToTimesheet

This is an app that is only really useful to Marc Adler and the way he accumulates client hours in his Google Sheets timesheet.

It will go through the week's Google Calendar, looking for the name of a CaaS client within the summary of the invitation. Then it will create a map of all of the different clients, the hours that Marc put into each meeting every day, and then it will insert those entries in the corresponding client tab inthe Google timesheet.

This repo contains an app which is really of no great importance to anyone other that CTO as a Service. But it illustrates the integration you can have between Google Calendar and Google Sheets.

If the `start` or `end` date options are not specified, the default date range is from the beginning of the week (Sunday at midnight) until today.

## The Appsettings File

Although it is not checked into the Github repo, the `appsettings.json` file is used to get a list of clients and the id of the Google Sheet that is the timesheet.

`appsettings.json`

```json
{
    "Logging": {
        "LogLevel": {
            "Default": "Debug",
            "System": "Information",
            "Microsoft": "Information"
        }
    },
    "Application": {
        "clients": [
            "client1",
            "client2",
            "client3"
        ],
        "timesheetId": "XXXXXXXXXXXXXYYYYYYYYYYYYYYZZZZZZZZ"
    }
}
```

## Command-line Options

| Name | Description |
| ---- | ----------- |
| -compress | If multiple events occur on the same day, combines the multiple entries into a single entry |
| -verbose | Debugging info |
| -nosheet | Do not write the new events to the Google timesheet |
| -clients [list] | Only processes clients which are in the comma-delimitd list |
| -start [date] | The starting date of the events |
| -end [date] | The ending date of the events |
