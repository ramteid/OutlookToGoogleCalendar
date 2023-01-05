# Outlook To Google Calendar Reader/Exporter/Importer
Exports Outlook calendar items from your lokal computer and imports it to Google Calendar. <br>
Does not communicate with Office365 cloud services. The export of Outlook calendar items runs purely locally. <br>

Tested with Outlook 16.0. Based on .NET6 C#.


## Details
The software basically does:
- Export Outlook main calendar items to an `.ical` file
- The `.ical` file is parsed to `Ical.Net` items as an intermediate format
- The ical items are converted into Google Calendar Event format
- The events are synced _one-way_ into the Google Calendar:
  - Google Calendar events that don't exist in Outlook are deleted in Google Calendar
  - Events that exist in both calendars and were modified in Outlook are updated in Google Calendar
  - Outlook events that don't yet exist in Google Calendar are inserted in Google Calendar

Changes in Google Calendar are not monitored nor synced back to Outlook.

## HowTo
- Create a *SEPARATE and EMPTY* Google Calendar *(!)* as **ALL EXISTING EVENTS OF THE TARGETED GOOGLE CALENDAR WILL BE DELETED !!!**
- In your Google calendar settings, find the Calendar Id, which has this format: `xxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com`
- Open `google-calendar-info.json` and paste the Google calendar Id into the appropriate field
- Build the project
- Create a Google Calendar service account
- Download the service account key file in `.json` format
- Copy the `json` file to the directory of the executable
- Rename the `json` file to `service-account-key-file.json`
- Open the `json` file and copy the value of the field `client_email`. This is the user name of the service account in email format.
- Go to the settings of your Google calendar. Add the service account user as a share (use the `client_email`) and assign write permissions.
- With both `.json` files present and edited in the same directory as your executable, run `OutlookCalendarReader.exe`.
- Check the console output and the logfile

Hint: As you'll probably want your calendar events synced regularly, you can use the `Windows Task Scheduler` to invoke the executable regularly (e. g. every 30 min)

## Known issues
- If the sharing options are disabled in your Google calendar and if you're using Google Workspace, check the Workspace settings for Calendar and the sharing settings for secondary calendars.
- If you're getting `404` / `Not Found` errors, check if the Google Calendar Id is really for the correct calendar and that your service account user has full permissions.
