using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;

namespace GCalendarToTimesheet
{
    public class GoogleCalendarDriver
    {
        private CalendarService Service { get; }
        
        public GoogleCalendarDriver(UserCredential credential, string applicationName)
        {
            this.Service = this.GetGCalendarService(credential, applicationName);
        }

        #region Interface to Google Calendar
        public CalendarService GetGCalendarService(UserCredential credential, string applicationName)
        {
            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
            return service;
        }
        
        public Events GetCalendarEvents(DateTime? startDate, DateTime? endDate)
        {
            // Define parameters of request.
            var request = this.Service.Events.List("primary");
            request.TimeMin = startDate != null && startDate != DateTime.MinValue ? startDate : DateTime.Now.StartOfWeek(DayOfWeek.Sunday);
            request.TimeMax = endDate != null && endDate != DateTime.MinValue ? endDate : DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 200;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            var events = request.Execute();
            return events;
        }
        #endregion
    }
}