using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Google.Apis.Calendar.v3.Data;

namespace OutlookCalendarReader;

internal class Program
{
    static async Task Main(string[] _)
    {
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
        var iCalFileName = Path.GetTempFileName();
        var outlookExporter = new OutlookExporter();
        outlookExporter.SaveCalendarToDisk(iCalFileName);

        var converter = new IcalConverter();
        var iCalEvents = converter.ConvertIcalToEvents(iCalFileName);
        File.Delete(iCalFileName);

        var googleCalendar = new GoogleCalendar();

        var existingGoogleEvents = await googleCalendar.GetExistingEvents();

        var eventsFromOutlookTasks = iCalEvents
            .Select(async e => await googleCalendar.ConvertIcalToGoogleEvent(e));
        var eventsFromOutlook = await Task.WhenAll(eventsFromOutlookTasks);

        var alreadyImportedEvents = eventsFromOutlook
            .Where(outlookEvent => existingGoogleEvents
                .Any(googleEvent => googleEvent.Id == outlookEvent.Id))
            .ToList();

        var eventsDeletedInOutlook = existingGoogleEvents
            .ExceptBy(eventsFromOutlook.Select(e => e.Id), e => e.Id)
            .ToList();

        var newEvents = eventsFromOutlook
            .Except(alreadyImportedEvents)
            .ToList();

        var eventsToUpdate = existingGoogleEvents
            .Except(eventsDeletedInOutlook)
            .Where(e => WasUpdated(e, eventsFromOutlook.Single(o => o.Id == e.Id)))
            .ToList();
        
        Logger.Log("Started");

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
            var updatedOutlookEvent = eventsFromOutlook.Single(o => o.Id == eventToUpdate.Id);
            await googleCalendar.UpdateEvent(updatedOutlookEvent, eventToUpdate.Id);
        }
        
        if (newEvents.Count > 0)
        {
            Logger.Log($"Events to insert: {newEvents.Count}");
        }
        foreach (var newEvent in newEvents)
        {
            await googleCalendar.InsertEvent(newEvent);
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