using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using dotenv.net;
using Google.Apis.Calendar.v3.Data;

namespace OutlookCalendarReader;

internal class Program
{
    static async Task Main(string[] _)
    {
        DotEnv.Load();

        try
        {
            await ExportOutlookCalendarToGoogle();
        }
        catch (HttpRequestException)
        {
            Logger.Log("HTTP connection failed");
        }
        catch (Exception e)
        {
            Logger.Log($"Exception terminated program: {e}");
        }
    }

    private static async Task ExportOutlookCalendarToGoogle()
    {
        Logger.Log("Started");

        var iCalFileName = Path.GetTempFileName();
        var outlookExporter = new OutlookExporter();
        outlookExporter.SaveCalendarToDisk(iCalFileName);

        var converter = new IcalConverter();
        var iCalImports = converter.ConvertIcalToEvents(iCalFileName);
        File.Delete(iCalFileName);
        
        var googleCalendar = new GoogleCalendar();
        var existingGoogleEvents = await googleCalendar.GetExistingEvents();
        
        var eventsFromOutlookTasks = iCalImports
            .Select(async e => new
            {
                Import = e,
                ConvertedEvent = await googleCalendar.ConvertIcalToConvertedEvent(e.CalendarEvent)
            });
        var eventsFromOutlook = await Task.WhenAll(eventsFromOutlookTasks);

        var alreadyImportedEvents = eventsFromOutlook
            .Where(outlookEvent => existingGoogleEvents.Any(e => e.Id == outlookEvent.ConvertedEvent.Id))
            .ToList();

        var eventsDeletedInOutlook = existingGoogleEvents
            .Where(e => e.Status != "cancelled")  // ignore cancelled instances of recurring events
            .ExceptBy(eventsFromOutlook.Select(e => e.ConvertedEvent.Id), e => e.Id)
            .ToList();

        var newEvents = eventsFromOutlook
            .Select(e => e.ConvertedEvent)
            .ExceptBy(alreadyImportedEvents.Select(e => e.ConvertedEvent).Select(x => x.Id), x => x.Id)
            .ToList();

        var eventsToUpdate = existingGoogleEvents
            .Where(e => e.Status != "cancelled")
            .Except(eventsDeletedInOutlook)
            .Where(existing =>
            {
                var matchingOutlookEvent = eventsFromOutlook
                    .FirstOrDefault(o => o.ConvertedEvent.Id == existing.Id);
                return matchingOutlookEvent is not null && WasUpdated(existing, matchingOutlookEvent.ConvertedEvent);
            })
            .ToList();

        var eventsWithExceptions = eventsFromOutlook
            .Where(e => e.Import.ExceptionDates.Any())
            .ToList();
        
        if (eventsDeletedInOutlook.Count > 0)
        {
            Logger.Log($"Events to delete: {eventsDeletedInOutlook.Count}");
        }
        foreach (var deletedEvent in eventsDeletedInOutlook)
        {
            await googleCalendar.DeleteEvent(deletedEvent.Id);
        }

        if (eventsToUpdate.Count > 0)
        {
            Logger.Log($"Events to update: {eventsToUpdate.Count}");
        }
        foreach (var eventToUpdate in eventsToUpdate)
        {
            // Duplicates are possible here for unknown reasons
            var updatedOutlookEvent = eventsFromOutlook.First(o => o.ConvertedEvent.Id == eventToUpdate.Id);
            updatedOutlookEvent.ConvertedEvent.Sequence = eventToUpdate.Sequence;

            await googleCalendar.UpdateEvent(updatedOutlookEvent.ConvertedEvent, eventToUpdate.Id);
        }

        if (newEvents.Count > 0)
        {
            Logger.Log($"Events to insert: {newEvents.Count}");
        }
        foreach (var newEvent in newEvents)
        {
            await googleCalendar.InsertEvent(newEvent);
        }
        
        foreach (var eventWithExceptions in eventsWithExceptions)
        {
            await googleCalendar.DeleteExceptions(eventWithExceptions.ConvertedEvent, eventWithExceptions.Import.ExceptionDates);
        }
    }
    
    private static bool WasUpdated(Event event1, Event event2)
    {
        var desc1 = event1.Description?.Substring(0, Math.Min(event1.Description.Length, 8000)) ?? "";
        var desc2 = event2.Description?.Substring(0, Math.Min(event2.Description.Length, 8000)) ?? "";
        var wasUpdated= event1.Summary != event2.Summary
                || desc1 != desc2
                || event1.Start.DateTime!.Value != event2.Start.DateTime!.Value
                || event1.End.DateTime!.Value != event2.End.DateTime!.Value;
        return wasUpdated;
    }
}