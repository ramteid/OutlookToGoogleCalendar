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
        public List<IcalEventImport> ConvertIcalToEvents(string iCalFileName)
        {
            var iCalText = File.ReadAllText(iCalFileName);
            var calendar = Calendar.Load(iCalText);

            var masterEvents = calendar.Events
                .Where(e => e.RecurrenceId is null)
                .Where(e => e.Start is not null && e.End is not null)
                .ToList();

            var importedMasters = masterEvents
                .Where(e => !string.IsNullOrWhiteSpace(e.Summary))
                .Where(IncludeEvent)
                //.Where(e => e.Start.AsSystemLocal < DateTime.Today + TimeSpan.FromDays(365))  // exclude events which start more than 365 days from today
                //.Where(e => e.End.AsSystemLocal > DateTime.Today - TimeSpan.FromDays(365))    // exclude events which ended more than 365 days ago
                .Select(e => new IcalEventImport(
                    e,
                    e.ExceptionDates.SelectMany(periods => periods).Select(period => period.StartTime.AsUtc)))
                .ToList();

            var seriesMastersByUid = masterEvents
                .Where(e => !string.IsNullOrWhiteSpace(e.Uid))
                .ToLookup(e => e.Uid!);

            var importedOverrides = calendar.Events
                .Where(e => e.RecurrenceId is not null)
                .Where(e => e.Start is not null && e.End is not null)
                .Select(e => TryConvertRecurringOverride(e, seriesMastersByUid, importedMasters))
                .Where(e => e is not null)
                .Cast<IcalEventImport>();

            return importedMasters
                .Concat(importedOverrides)
                .ToList();
        }

        private IcalEventImport? TryConvertRecurringOverride(
            CalendarEvent recurringEvent,
            ILookup<string, CalendarEvent> seriesMastersByUid,
            IEnumerable<IcalEventImport> importedMasters)
        {
            var master = ResolveSeriesMaster(recurringEvent, seriesMastersByUid);
            if (master is null)
            {
                return string.IsNullOrWhiteSpace(recurringEvent.Summary) || !IncludeEvent(recurringEvent)
                    ? null
                    : new IcalEventImport(recurringEvent);
            }

            ApplyMasterDefaults(recurringEvent, master);

            if (string.IsNullOrWhiteSpace(recurringEvent.Summary) || !IncludeEvent(recurringEvent))
            {
                return null;
            }

            if (!ShouldImportRecurringOverride(master, recurringEvent))
            {
                return null;
            }

            var masterImport = importedMasters.FirstOrDefault(item => ReferenceEquals(item.CalendarEvent, master));
            if (masterImport is not null)
            {
                AddSyntheticException(masterImport.ExceptionDates, recurringEvent.RecurrenceId!.AsUtc);
            }

            return new IcalEventImport(recurringEvent);
        }

        private static CalendarEvent? ResolveSeriesMaster(
            CalendarEvent recurringEvent,
            ILookup<string, CalendarEvent> seriesMastersByUid)
        {
            if (string.IsNullOrWhiteSpace(recurringEvent.Uid))
            {
                return null;
            }

            var candidates = seriesMastersByUid[recurringEvent.Uid!].ToList();
            if (candidates.Count == 0)
            {
                return null;
            }

            if (candidates.Count == 1)
            {
                return candidates[0];
            }

            var recurrenceStart = recurringEvent.RecurrenceId?.AsUtc ?? recurringEvent.Start!.AsUtc;

            return candidates
                .OrderByDescending(e => e.Start!.AsUtc <= recurrenceStart)
                .ThenBy(e => Math.Abs((e.Start!.AsUtc - recurrenceStart).Ticks))
                .First();
        }

        private static void ApplyMasterDefaults(CalendarEvent recurringEvent, CalendarEvent master)
        {
            if (string.IsNullOrWhiteSpace(recurringEvent.Summary))
            {
                recurringEvent.Summary = master.Summary;
            }

            if (string.IsNullOrWhiteSpace(recurringEvent.Description))
            {
                recurringEvent.Description = master.Description;
            }

            if (string.IsNullOrWhiteSpace(recurringEvent.Location))
            {
                recurringEvent.Location = master.Location;
            }

            recurringEvent.Organizer ??= master.Organizer;
        }

        private static bool ShouldImportRecurringOverride(CalendarEvent master, CalendarEvent recurringEvent)
        {
            var originalStart = recurringEvent.RecurrenceId!.AsUtc;
            var originalEnd = originalStart + (master.End!.AsUtc - master.Start!.AsUtc);

            return recurringEvent.Start!.AsUtc != originalStart
                || recurringEvent.End!.AsUtc != originalEnd
                || !string.Equals(recurringEvent.Summary, master.Summary, StringComparison.Ordinal);
        }

        private static void AddSyntheticException(List<DateTime> exceptionDates, DateTime recurrenceStart)
        {
            if (exceptionDates.Contains(recurrenceStart))
            {
                return;
            }

            exceptionDates.Add(recurrenceStart);
        }

        private bool IncludeEvent(CalendarEvent e)
        {
            return e.Summary switch
            {
                // Add any event title here that should be excluded from syncing. Matching is case-sensitive and looks for an exact match of the entire summary.
                "ELTERNZEIT" => false,
                "ABWESEND" => false,
                "Thursdays Blocker" => false,
                "PRIVAT" => false,
                "URLAUB" => false,
                _ => true
            };
        }
    }
}
