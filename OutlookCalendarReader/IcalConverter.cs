using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Ical.Net;
using Ical.Net.CalendarComponents;

namespace OutlookCalendarReader
{
    internal class IcalConverter
    {
        public List<CalendarEvent> ConvertIcalToEvents(string iCalFileName)
        {
            var iCalText = File.ReadAllText(iCalFileName);
            var calendar = Calendar.Load(iCalText);
            
            var movedRecurringEvents = calendar.Events
                .Where(e => e.RecurrenceId is not null && !Equals(e.RecurrenceId, e.Start))
                .Select(e =>
                {
                    if (string.IsNullOrWhiteSpace(e.Summary))
                    {
                        e.Summary = "unknown moved recurring event";
                    }
                    return e;
                });

            return calendar.Events
                .Where(e => !string.IsNullOrWhiteSpace(e.Summary))
                .Where(e => e.Start is not null && e.End is not null)
                //.Where(e => e.Start.AsSystemLocal < DateTime.Today + TimeSpan.FromDays(365))  // exclude events which start more than 365 days from today
                //.Where(e => e.End.AsSystemLocal > DateTime.Today - TimeSpan.FromDays(365))    // exclude events which ended more than 365 days ago
                .Concat(movedRecurringEvents)
                .ToList();
        }
    }
}
