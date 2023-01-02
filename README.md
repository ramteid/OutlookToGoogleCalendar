# Outlook To Google Calendar Reader/Exporter/Importer
Exports Outlook calendar items from your lokal computer and imports it to Google Calendar. <br>
Does not communicate with Office365 cloud services. The export of Outlook calendar items runs purely locally. <br>

Tested with Outlook 16.0. Based on .NET6 C#.


## Details
The software basically does:
- Export Outlook main calendar items to an `.ical` file
- The `.ical` file is parsed to `Ical.Net` items as an intermediate format
- The ical items are converted into Google Calendar Event format
- The events are synced one-way into the Google Calendar

## HowTo
- In the Google calendar settings, find the Calendar Id, which has this format: `xxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com`
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
