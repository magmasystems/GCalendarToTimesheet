using System;
using System.Linq;
using Google.Apis.Calendar.v3.Data;
using Newtonsoft.Json;

namespace GCalendarToTimesheet
{
    public class GCalEventInfo
    {
        public DateTime Date { get; set; }
        public TimeSpan Duration { get; set; }
        public string Summary { get; set; }
        public Event NativeEvent { get; set; }

        public GCalEventInfo()
        {
        }
        
        public GCalEventInfo(DateTime date, TimeSpan duration, string summary, Event nativeEvent) : this()
        {
            this.Date = date;
            this.Duration = duration;
            this.Summary = summary;
            this.NativeEvent = nativeEvent;
        }

        public bool DidIDecline()
        {
            var record = this.MyAttendeeRecord();
            return record is { ResponseStatus: "declined" };
        }

        private EventAttendee MyAttendeeRecord()
        {
            var myRecord = NativeEvent?.Attendees?.Where(a => a.Email == "magmasystems@gmail.com").FirstOrDefault();
            return myRecord ?? null;
        }
        
        public override string ToString()
        {
            return $"Desc: {this.Summary}, Date: {this.Date}, Duration: {this.Duration}";
        }
        
        public string ToStringEx()
        {
            return JsonConvert.SerializeObject(this.NativeEvent, Formatting.Indented);
        }
    }

    public static class EventUtilities
    {
        public static bool DidIDecline(this Event e)
        {
            var record = e.MyAttendeeRecord();
            return record is { ResponseStatus: "declined" };
        }

        private static EventAttendee MyAttendeeRecord(this Event e)
        {
            var myRecord = e?.Attendees?.Where(a => a.Email == "magmasystems@gmail.com").FirstOrDefault();
            return myRecord ?? null;
        }
    }
}