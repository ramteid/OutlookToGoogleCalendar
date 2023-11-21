using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Ical.Net.CalendarComponents;
using System.Reflection;
using System.Text.Json;
using Ical.Net.DataTypes;

namespace OutlookCalendarReader
{
    internal class GoogleCalendar
    {
        private readonly string _calendarId;
        private readonly CalendarService _service;
        private string _calendarTimeZone;

        public GoogleCalendar()
        {
            var basePath = Path.GetDirectoryName(Assembly.GetEntryAssembly()?.Location);

            var serviceAccountKeyFilePath = Path.Join(basePath, "service-account-key-file.json");
            if (!File.Exists(serviceAccountKeyFilePath))
            {
                throw new InvalidOperationException($"Could not find '{serviceAccountKeyFilePath}'");
            }

            var calendarInfoFile = Path.Join(basePath, "google-calendar-info.json");
            if (!File.Exists(calendarInfoFile))
            {
                throw new InvalidOperationException($"Could not find '{calendarInfoFile}'");
            }

            var calendarInfo = JsonSerializer.Deserialize<GoogleCalendarInfo>(File.ReadAllText(calendarInfoFile));
            if (string.IsNullOrWhiteSpace(calendarInfo?.GoogleCalendarId))
            {
                throw new InvalidOperationException($"Error reading Google Calendar Id from '{calendarInfoFile}'");
            }

            _calendarId = calendarInfo.GoogleCalendarId;

            _service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GoogleCredential.FromFile(serviceAccountKeyFilePath).CreateScoped(CalendarService.Scope.Calendar),
                ApplicationName = "Outlook Calendar Importer"
            });
        }

        internal async Task<(Event ConvertedEvent, IEnumerable<Period> ExceptionDates)> ConvertIcalToConvertedEvent(CalendarEvent iCalEvent)
        {
            // Specifying attendees in Google calendar would send out invitations on creation. Also some business account would be required.
            // Also sometimes there is no E-Mail in the Outlook export specified. Thus only append the attendees to the description.
            var attendees = iCalEvent.Attendees.Select(a =>
            {
                var name = string.IsNullOrWhiteSpace(a.CommonName) ? "Unknown" : a.CommonName;
                var email = string.IsNullOrWhiteSpace(a.Value.ToString())
                    ? ""
                    : a.Value.ToString().Replace($"{a.Value.Scheme}:", "");
                return $"{name} <{email}>";
            });

            var description = iCalEvent.Attendees.Any() 
                ? (iCalEvent.Description ?? "") + "\n\nAttendees:\n" + string.Join("\n", attendees)
                : iCalEvent.Description ?? "";

            var convertedEvent = new Event
            {
                Id = NormalizeUid(iCalEvent),
                Created = iCalEvent.Created?.AsSystemLocal,
                Summary = iCalEvent.Summary,
                Description = description,
                Location = iCalEvent.Location ?? "",
                Organizer = new Event.OrganizerData { DisplayName = iCalEvent.Organizer?.CommonName ?? "" },
                Recurrence = iCalEvent.RecurrenceRules.Select(r => "RRULE:" + r).ToList(),
                Sequence = iCalEvent.Sequence,
                Start = new EventDateTime
                {
                    DateTime = iCalEvent.Start.AsUtc,
                    TimeZone = await GetCalendarTimeZone()
                },
                End = new EventDateTime
                {
                    DateTime = iCalEvent.End.AsUtc,
                    TimeZone = await GetCalendarTimeZone()
                },
                RecurringEventId = iCalEvent.RecurrenceId?.ToString()
            };

            return (ConvertedEvent: convertedEvent, ExceptionDates: iCalEvent.ExceptionDates.SelectMany(e => e));
        }

        private async Task<string> GetCalendarTimeZone()
        {
            if (string.IsNullOrWhiteSpace(_calendarTimeZone))
            {
                var timeZone = await GetCalendarTimeZone(_calendarId);
                _calendarTimeZone = timeZone;
            }
            return _calendarTimeZone;
        }

        internal async Task<IList<Event>> GetExistingEvents()
        {
            var list = new List<Event>();
            string? nextPageToken = null;

            do
            {
                var listRequest = _service.Events.List(_calendarId);
                if (nextPageToken is not null)
                {
                    listRequest.PageToken = nextPageToken;
                }
                //listRequest.TimeMax = DateTime.Today + TimeSpan.FromDays(365);  // exclude events which start more than 365 days from today
                //listRequest.TimeMin = DateTime.Today - TimeSpan.FromDays(365);    // exclude events which ended more than 365 days ago
                var listResponse = await listRequest.ExecuteAsync();
                list.AddRange(listResponse.Items);
                nextPageToken = listResponse.NextPageToken;
            }
            while (nextPageToken != null);

            return list;
        }

        private async Task<List<Event>> GetInstances(Event convertedEvent)
        {
            try
            {
                var list = new List<Event>();
                string? nextPageToken = null;

                do
                {
                    var request = _service.Events.Instances(_calendarId, convertedEvent.Id);
                    if (nextPageToken is not null)
                    {
                        request.PageToken = nextPageToken;
                    }
                    var response = await request.ExecuteAsync();
                    list.AddRange(response.Items);
                    nextPageToken = response.NextPageToken;
                }
                while (nextPageToken != null);

                return list;
            }
            catch (Exception e)
            {
                Logger.Log($"Error retrieving instances: {e}");
                return new List<Event>();
            }
        }

        internal async Task InsertEvent(Event calendarEvent)
        {
            var request = _service.Events.Insert(calendarEvent, _calendarId);
            try
            {
                var response = await request.ExecuteAsync();
                Logger.Log("Inserted event " + response.Id);
            }
            catch (Exception e)
            {
                Logger.Log($"Error inserting event {calendarEvent.Id}: {e}");
            }
        }

        internal async Task DeleteEvent(string eventId)
        {
            var request = _service.Events.Delete(_calendarId, eventId);
            var response = "";
            try
            {
                response = await request.ExecuteAsync();
            }
            catch (Exception e)
            {
                Logger.Log($"Error deleting event {eventId}. Response: {response}, Error: {e}");
                return;
            }
            Logger.Log($"Deleted event {eventId}");
        }

        internal async Task UpdateEvent(Event calendarEvent, string eventId)
        {
            var request = _service.Events.Update(calendarEvent, _calendarId, eventId);
            Event response;
            try
            {
                response = await request.ExecuteAsync();
            }
            catch (Exception e)
            {
                Logger.Log($"Error updating event {calendarEvent.Id}: {e}");
                return;
            }
            Logger.Log($"Updated event {response.Id}");
        }

        internal async Task DeleteExceptions(Event googleEvent, IEnumerable<Period> exceptionDates)
        {
            var instances = await GetInstances(googleEvent);
            var exceptionsToDelete = instances
                .Where(instance => instance.Status != "cancelled")
                .Where(instance => instance.Start.DateTime.HasValue)
                .Where(instance => exceptionDates.Any(e => e.StartTime.AsUtc == instance.Start.DateTime!.Value.ToUniversalTime()))
                .ToList();

            if (exceptionsToDelete.Count == 0)
            {
                // All exceptions for this event have already been cancelled
                return;
            }

            Logger.Log($"Cancelling {exceptionsToDelete.Count} exceptions for event {googleEvent.Id}");

            foreach (var exception in exceptionsToDelete)
            {
                // Retrieve instance again as the ETAG could have changed if the base event was modified in another iteration
                var refreshedInstances = await GetInstances(googleEvent);
                var instance = refreshedInstances.SingleOrDefault(i => i.Id == exception.Id);
                if (instance is null)
                {
                    Logger.Log($"Error retrieving instance again: {exception.Id}");
                }

                instance!.Status = "cancelled";
                var cancelRequest = _service.Events.Update(instance, _calendarId, instance.Id);
                try
                {
                    var result = await cancelRequest.ExecuteAsync();
                }
                catch (Exception e)
                {
                    Logger.Log($"Error cancelling instance: {e}");
                }
            }
        }

        private async Task<string> GetCalendarTimeZone(string calendarId)
        {
            var request = _service.Calendars.Get(calendarId);
            var response = await request.ExecuteAsync();
            return response.TimeZone;
        }

        /// <summary>
        /// Outlook event UIDs are not standard-conform, not consistent and possibly duplicated, so transform it into a standard-conform hash
        /// </summary>
        private static string NormalizeUid(CalendarEvent iCalEvent)
        {
            var id = iCalEvent.Uid + iCalEvent.Summary + iCalEvent.Start.AsUtc + iCalEvent.End.AsUtc + "3";

            using var sha256Hash = SHA256.Create();

            // Convert the input string to a byte array and compute the hash.
            var data = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(id));

            // Create a new Stringbuilder to collect the bytes and create a string.
            var sBuilder = new StringBuilder();

            // Loop through each byte of the hashed data and format each one as a hexadecimal string.
            foreach (var t in data)
            {
                sBuilder.Append(t.ToString("x2"));
            }

            // Return the hexadecimal string.
            var hash = sBuilder.ToString();
            return hash[..Math.Min(hash.Length, 1024)];
        }
    }
}
