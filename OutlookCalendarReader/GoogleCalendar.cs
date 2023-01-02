﻿using System;
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

namespace OutlookCalendarReader
{
    internal class GoogleCalendar
    {
        private readonly string _calendarId;
        private readonly CalendarService _service;

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

        internal async Task<Event> ConvertIcalToGoogleEvent(CalendarEvent iCalEvent)
        {
            var calendarTimeZone = await GetCalendarTimeZone(_calendarId);

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

            return new Event
            {
                Id = NormalizeUid(iCalEvent),
                Created = iCalEvent.Created.AsSystemLocal,
                Summary = iCalEvent.Summary,
                Description = description,
                Location = iCalEvent.Location ?? "",
                Organizer = new Event.OrganizerData { DisplayName = iCalEvent.Organizer?.CommonName ?? "" },
                Recurrence = iCalEvent.RecurrenceRules.Select(r => "RRULE:" + r).ToList(),
                Sequence = iCalEvent.Sequence,
                Start = new EventDateTime
                {
                    DateTime = iCalEvent.Start.AsUtc,
                    TimeZone = calendarTimeZone
                },
                End = new EventDateTime
                {
                    DateTime = iCalEvent.End.AsUtc,
                    TimeZone = calendarTimeZone
                }
            };
        }

        internal async Task<IList<Event>> GetExistingEvents()
        {
            var listRequest = _service.Events.List(_calendarId);
            //listRequest.TimeMax = DateTime.Today + TimeSpan.FromDays(365);  // exclude events which start more than 365 days from today
            //listRequest.TimeMin = DateTime.Today - TimeSpan.FromDays(365);    // exclude events which ended more than 5 days ago
            var listResponse = await listRequest.ExecuteAsync();
            return listResponse.Items;
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
            Logger.Log("Updated event " + response.Id);
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
            var id = iCalEvent.Uid + iCalEvent.Created + iCalEvent.Summary + iCalEvent.Start.AsUtc + iCalEvent.End.AsUtc + "a";

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
