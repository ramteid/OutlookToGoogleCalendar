using System;

namespace OutlookCalendarReader;

internal class OutlookExporter
{
    public void SaveCalendarToDisk(string calendarFileName)
    {
        if (string.IsNullOrEmpty(calendarFileName))
        {
            throw new Exception("calendarFileName must contain a value.");
        }

        var outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        var calendar = outlookApplication.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar) as Microsoft.Office.Interop.Outlook.Folder;
        var exporter = calendar!.GetCalendarExporter();

        // Set the properties for the export
        exporter.CalendarDetail = Microsoft.Office.Interop.Outlook.OlCalendarDetail.olFullDetails;
        exporter.IncludeAttachments = true;
        exporter.IncludePrivateDetails = true;
        exporter.RestrictToWorkingHours = false;
        exporter.IncludeWholeCalendar = true;
        exporter.StartDate = DateTime.Today;
        exporter.EndDate = DateTime.Today + TimeSpan.FromDays(365);

        // Save the calendar to disk
        exporter.SaveAsICal(calendarFileName);
    }
}