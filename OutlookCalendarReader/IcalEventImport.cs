using System;
using System.Collections.Generic;
using System.Linq;
using Ical.Net.CalendarComponents;

namespace OutlookCalendarReader;

internal sealed class IcalEventImport
{
    public IcalEventImport(CalendarEvent calendarEvent, IEnumerable<DateTime>? exceptionDates = null)
    {
        CalendarEvent = calendarEvent;
        ExceptionDates = (exceptionDates ?? Enumerable.Empty<DateTime>()).ToList();
    }

    public CalendarEvent CalendarEvent { get; }

    public List<DateTime> ExceptionDates { get; }
}