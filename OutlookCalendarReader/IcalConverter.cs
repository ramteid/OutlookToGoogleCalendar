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
            return calendar.Events
                .Where(e => !string.IsNullOrWhiteSpace(e.Summary))
                .Where(e => e.Start is not null && e.End is not null)
                //.Where(e => e.Start.AsSystemLocal < DateTime.Today + TimeSpan.FromDays(365))
                //.Where(e => e.End.AsSystemLocal > DateTime.Today - TimeSpan.FromDays(365))
                .ToList();
        }
    }
}
