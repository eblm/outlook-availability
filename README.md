# Outlook Availability Finder

A small Windows desktop tool (a single PowerShell script with a Windows Forms interface) that reads a calendar from **classic Outlook** and lists your **free time slots** over a period you choose, ready to paste into an email.

No installation, no admin rights, no external dependencies: it talks to the Outlook you already have installed.

## Features

- Pick **which calendar** to analyze: your default calendar, calendars of other mailboxes or data files, calendars shown in the Outlook Calendar pane, or someone else's shared calendar (by email address, read permission required).
- Choose a **date range** and a **daily time window** (e.g. 09:00-17:00).
- **Exclude days of the week** (Saturday and Sunday are excluded by default).
- Optional **lunch break**: when enabled, you set a start and an end time, and that period counts as "not available" every day.
- **Minimum slot duration** to discard slots that are too short (default 30 minutes).
- Options for how events are treated:
  - all-day events block the whole day (on by default);
  - events marked as "Free" are ignored (on by default).
- Recurring events are expanded, so every occurrence is taken into account.
- Cancelled meetings and meetings you declined are ignored.
- Results are shown three ways: a table, an editable plain-text summary (to paste into an email) and a CSV export.

Example of the text output:

```
Mon 12/10/2026: 10:30-14:00, 15:00-17:00
Tue 13/10/2026: 09:00-17:00
Fri 16/10/2026: 09:00-12:00
```

## Requirements

- Windows with **Windows PowerShell 5.1** (included in Windows 10 and 11).
- **Classic Outlook** for Windows, installed and configured. The "new Outlook" and Outlook on the web do **not** support the COM interface this script uses.
- Outlook and PowerShell must run with the **same privilege level** (both normal, or both as administrator).

## Usage

1. Download `outlook-availability.ps1`.
2. If the file was downloaded from the internet, unblock it once:

   ```powershell
   Unblock-File .\outlook-availability.ps1
   ```

3. Start it:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\outlook-availability.ps1
   ```

   Or create a desktop shortcut with this target, to launch it like a regular application:

   ```
   powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\outlook-availability.ps1"
   ```

4. Select the calendar, set the criteria and click **Calculate availability**.
5. Use **Copy text** to copy the summary, or **Export CSV...** to save the slots.

The CSV export is semicolon-separated, UTF-8 with BOM (so Excel opens it correctly), with the columns `Date;From;To;Duration (min)`.

## How the calculation works

For each selected day that is not excluded, the script starts from your daily time window and subtracts the busy periods:

- Events from the chosen calendar that overlap the window (partially overlapping events are clipped to the window).
- The lunch break, if enabled.

What remains is split into free slots, and slots shorter than the minimum duration are dropped. Overlapping events are merged correctly.

An event counts as busy when its status is Busy, Tentative, Out of Office or Working Elsewhere. Events marked Free are ignored unless you untick the option.

## Privacy

Everything runs locally on your PC. The script does not send data anywhere and does not read event titles, descriptions or attendees: it only reads start and end times, the all-day flag, the busy status, and the meeting/response status (to skip cancelled or declined meetings).

## Troubleshooting

- **"Unable to connect to Outlook"**: make sure you are using classic Outlook (if you see a *New Outlook* toggle at the top right, turn it off), that Outlook is configured, and that Outlook and PowerShell have the same privileges.
- **The calendar I need is not in the list**: click *Refresh list*. For a colleague's or shared mailbox calendar use *Shared calendar...* and enter their email address (you need read permission).
- **Unexpected number of events**: Outlook filters events using the short date format of your Windows regional settings. Check that the regional format is configured correctly.
- **Any other error**: the error window shows the details, and they are also written to `%TEMP%\outlook-availability.log`. The log can contain mailbox names, so check it before sharing it.

## Optional: build an .exe

You can wrap the script in an executable with the [PS2EXE](https://github.com/MScholtes/PS2EXE) module:

```powershell
Install-Module ps2exe -Scope CurrentUser
Invoke-ps2exe .\outlook-availability.ps1 .\outlook-availability.exe -noConsole -STA
```

Keep `-STA`, it is required for the Windows Forms interface and for Outlook. Executables built this way are often flagged by antivirus software and application-control policies; the shortcut approach above avoids this.

## Limitations

- Classic Outlook for Windows only.
- Only one calendar is analyzed at a time.
- Time zones are those handled by Outlook for the selected calendar; no manual conversion is done.
