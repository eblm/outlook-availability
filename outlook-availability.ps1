# ============================================================================
#  outlook-availability.ps1
#  Computes free time slots by reading a calendar directly from
#  CLASSIC Outlook for Windows (via COM). Windows Forms user interface.
#
#  Requirements: classic Outlook installed and configured (NOT "new Outlook").
#  Run:    powershell -ExecutionPolicy Bypass -File .\outlook-availability.ps1
#          (if the file was downloaded: Unblock-File .\outlook-availability.ps1)
#  The script and Outlook must run with the same privilege level
#  (both normal, or both as administrator).
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

# Prevents a global setting ($ErrorActionPreference = 'Stop' in the profile) from turning
# Outlook COM errors into unhandled exceptions inside window events
$ErrorActionPreference = 'Continue'

# Error log: %TEMP%\outlook-availability.log
$script:logPath = Join-Path $env:TEMP 'outlook-availability.log'

function Write-Log([string]$text) {
    try {
        Add-Content -Path $script:logPath -Encoding UTF8 -Value ('[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $text)
    } catch { }
}

# Full text of an error: type, message, inner exceptions, script line and the last errors recorded by PowerShell
function Get-FailureText($rec) {
    $sb = New-Object System.Text.StringBuilder
    try {
        $ex = $rec.Exception
        while ($null -ne $ex) {
            [void]$sb.AppendLine($ex.GetType().FullName + ': ' + $ex.Message)
            $ex = $ex.InnerException
        }
        if ($null -ne $rec.InvocationInfo) { [void]$sb.AppendLine($rec.InvocationInfo.PositionMessage) }
        if ($rec.ScriptStackTrace) { [void]$sb.AppendLine($rec.ScriptStackTrace) }
        # PipelineStoppedException hides the original error: look for it among the latest errors
        $n = 0
        foreach ($e in $global:Error) {
            if ($n -ge 4) { break }
            [void]$sb.AppendLine('[recent error] ' + $e.Exception.GetType().Name + ': ' + $e.Exception.Message)
            if ($null -ne $e.InvocationInfo) { [void]$sb.AppendLine('   ' + $e.InvocationInfo.PositionMessage.Replace("`r`n", ' ')) }
            $n++
        }
    } catch { }
    return $sb.ToString()
}

$script:ol          = $null
$script:ns          = $null
$script:calFolders  = New-Object System.Collections.Generic.List[object]
$script:slots       = @()
$script:inv = [System.Globalization.CultureInfo]::InvariantCulture

# ---------------------------------------------------------------------------
#  Outlook (COM)
# ---------------------------------------------------------------------------

function Connect-Outlook {
    if ($null -ne $script:ns) { return }
    try {
        $script:ol = New-Object -ComObject Outlook.Application
    } catch {
        throw ("Unable to connect to Outlook. This script only works with CLASSIC Outlook " +
               "(not with 'new Outlook'): if you see a 'New Outlook' toggle at the top right, " +
               "turn it off and try again. Also check that Outlook and PowerShell were not started with " +
               "different privileges (one as administrator and the other not).`r`n`r`nDetails: " + $_.Exception.Message)
    }
    $script:ns = $script:ol.GetNamespace('MAPI')
}

# Adds a calendar and its sub-calendars, if any
function Add-CalendarTree($folder, $list, $seen) {
    try {
        $key = [string]$folder.StoreID + '|' + [string]$folder.EntryID
        if ($seen.Add($key)) { $list.Add($folder) }
        foreach ($sub in $folder.Folders) {
            if ($sub.DefaultItemType -eq 1) { Add-CalendarTree $sub $list $seen }
        }
    } catch { }
}

# Lists the calendars: default one, those of every mailbox/data file and those shown in the Calendar pane
function Get-CalendarFolders {
    Connect-Outlook
    $list = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'

    try { Add-CalendarTree ($script:ns.GetDefaultFolder(9)) $list $seen }
    catch { Write-Log ('Default calendar not read: ' + $_.Exception.Message) }

    foreach ($root in $script:ns.Folders) {
        try {
            foreach ($f in $root.Folders) {
                if ($f.DefaultItemType -eq 1) { Add-CalendarTree $f $list $seen }
            }
        } catch { Write-Log ('Mailbox not read: ' + $_.Exception.Message) }
    }

    # Calendars shown in the navigation pane (including shared ones opened from other people)
    try {
        $exp = $script:ol.ActiveExplorer()
        if ($null -ne $exp) {
            $mod = $exp.NavigationPane.Modules.GetNavigationModule(1)   # 1 = olModuleCalendar
            foreach ($grp in $mod.NavigationGroups) {
                foreach ($nf in $grp.NavigationFolders) {
                    try { Add-CalendarTree $nf.Folder $list $seen } catch { }
                }
            }
        }
    } catch { Write-Log ('Navigation pane not read: ' + $_.Exception.Message) }

    return $list
}

# Reads the events of a calendar in the period [from, toExclusive), expanding recurrences
function Read-OutlookEvents($folder, [datetime]$from, [datetime]$toExclusive) {
    $list  = New-Object System.Collections.Generic.List[object]
    $items = $folder.Items
    $items.IncludeRecurrences = $true
    $items.Sort('[Start]')
    # Outlook wants dates in the short format of the Windows regional settings
    $filter = "[Start] < '" + $toExclusive.ToString('g') + "' AND [End] > '" + $from.ToString('g') + "'"
    $res = $items.Restrict($filter)
    $it  = $res.GetFirst()
    while ($null -ne $it) {
        try {
            $canceled = ($it.MeetingStatus -eq 5 -or $it.MeetingStatus -eq 7)
            $declined = ($it.ResponseStatus -eq 4)
            if (-not $canceled -and -not $declined) {
                # BusyStatus: 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere
                $list.Add([pscustomobject]@{
                    Start  = [datetime]$it.Start
                    End    = [datetime]$it.End
                    AllDay = [bool]$it.AllDayEvent
                    Free   = ($it.BusyStatus -eq 0)
                })
            }
        } catch { }
        $it = $res.GetNext()
    }
    return $list
}

# ---------------------------------------------------------------------------
#  Availability calculation
# ---------------------------------------------------------------------------

function Get-FreeSlots($events, [datetime]$from, [datetime]$to, [timespan]$dayStart, [timespan]$dayEnd,
                       $excludedDays, [timespan]$minSlot, [bool]$allDayBlocks, [bool]$ignoreFree,
                       [bool]$lunchOn, [timespan]$lunchStart, [timespan]$lunchEnd) {
    $result = New-Object System.Collections.Generic.List[object]
    for ($day = $from.Date; $day -le $to.Date; $day = $day.AddDays(1)) {
        if ($excludedDays -contains [int]$day.DayOfWeek) { continue }
        $ws = $day + $dayStart
        $we = $day + $dayEnd
        if ($we -le $ws) { continue }

        $busy = New-Object System.Collections.Generic.List[object]
        foreach ($e in $events) {
            if ($ignoreFree -and $e.Free) { continue }
            if ($e.AllDay -and -not $allDayBlocks) { continue }
            if ($e.End -le $e.Start) { continue }
            if ($e.End -le $ws -or $e.Start -ge $we) { continue }
            $s  = if ($e.Start -lt $ws) { $ws } else { $e.Start }
            $en = if ($e.End -gt $we) { $we } else { $e.End }
            $busy.Add([pscustomobject]@{ S = $s; E = $en })
        }

        # Optional lunch break: counts as "not available" every day
        if ($lunchOn) {
            $ls = $day + $lunchStart
            $le = $day + $lunchEnd
            if ($le -gt $ws -and $ls -lt $we) {
                $s  = if ($ls -lt $ws) { $ws } else { $ls }
                $en = if ($le -gt $we) { $we } else { $le }
                $busy.Add([pscustomobject]@{ S = $s; E = $en })
            }
        }

        $cursor = $ws
        foreach ($b in @($busy | Sort-Object S)) {
            if ($b.S -gt $cursor -and ($b.S - $cursor) -ge $minSlot) {
                $result.Add([pscustomobject]@{ Date = $day; From = $cursor; To = $b.S })
            }
            if ($b.E -gt $cursor) { $cursor = $b.E }
        }
        if ($cursor -lt $we -and ($we - $cursor) -ge $minSlot) {
            $result.Add([pscustomobject]@{ Date = $day; From = $cursor; To = $we })
        }
    }
    return $result
}

# ---------------------------------------------------------------------------
#  Formatting
# ---------------------------------------------------------------------------

function Format-Day([datetime]$d) {
    $n = $script:inv.DateTimeFormat.GetAbbreviatedDayName($d.DayOfWeek)
    return "$n " + $d.ToString('dd/MM/yyyy', $script:inv)
}

function Format-Time([datetime]$t) {
    return $t.ToString('HH:mm', $script:inv)
}

function Format-Duration([timespan]$t) {
    $h = [int][math]::Floor($t.TotalHours)
    $m = $t.Minutes
    if ($h -eq 0) { return "$m min" }
    if ($m -eq 0) { return "$h h" }
    return "$h h $m min"
}

function Build-Text($slots) {
    if (@($slots).Count -eq 0) { return 'No availability found in the selected period.' }
    $dash = [string][char]0x2013
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($g in ($slots | Group-Object { $_.Date.Date.Ticks })) {
        $parts = @($g.Group | ForEach-Object { (Format-Time $_.From) + $dash + (Format-Time $_.To) })
        $lines.Add((Format-Day $g.Group[0].Date) + ': ' + ($parts -join ', '))
    }
    return ($lines -join "`r`n")
}

function Show-Error([string]$title, [string]$message) {
    Write-Log ($title + ' - ' + $message)
    try {
        [void][System.Windows.Forms.MessageBox]::Show(
            ($message + "`r`n`r`n(Details saved in: " + $script:logPath + ')'), $title, 'OK', 'Error')
    } catch { }
}

# ---------------------------------------------------------------------------
#  User interface
# ---------------------------------------------------------------------------

function New-Label([string]$text, [int]$x, [int]$y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.AutoSize = $true
    return $l
}

function New-TimePicker([int]$x, [int]$y, [int]$hour, [int]$minute) {
    $p = New-Object System.Windows.Forms.DateTimePicker
    $p.Format = 'Custom'
    $p.CustomFormat = 'HH:mm'
    $p.ShowUpDown = $true
    $p.Location = New-Object System.Drawing.Point($x, $y)
    $p.Width = 80
    $p.Value = [datetime]::Today.AddHours($hour).AddMinutes($minute)
    return $p
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Availability from Outlook'
$form.Size = New-Object System.Drawing.Size(980, 800)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# --- top panel ---
$top = New-Object System.Windows.Forms.Panel
$top.Dock = 'Top'
$top.Height = 315

$cmbCal = New-Object System.Windows.Forms.ComboBox
$cmbCal.DropDownStyle = 'DropDownList'
$cmbCal.Location = New-Object System.Drawing.Point(100, 12)
$cmbCal.Width = 520

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh list'
$btnRefresh.Location = New-Object System.Drawing.Point(632, 10)
$btnRefresh.Size = New-Object System.Drawing.Size(120, 28)

$btnShared = New-Object System.Windows.Forms.Button
$btnShared.Text = 'Shared calendar...'
$btnShared.Location = New-Object System.Drawing.Point(762, 10)
$btnShared.Size = New-Object System.Drawing.Size(170, 28)

$dtFrom = New-Object System.Windows.Forms.DateTimePicker
$dtFrom.Format = 'Short'
$dtFrom.Location = New-Object System.Drawing.Point(120, 58)
$dtFrom.Width = 120
$dtFrom.Value = [datetime]::Today

$dtTo = New-Object System.Windows.Forms.DateTimePicker
$dtTo.Format = 'Short'
$dtTo.Location = New-Object System.Drawing.Point(360, 58)
$dtTo.Width = 120
$dtTo.Value = [datetime]::Today.AddDays(14)

$tpStart = New-TimePicker 120 92 9 0
$tpEnd   = New-TimePicker 360 92 17 0

$numMin = New-Object System.Windows.Forms.NumericUpDown
$numMin.Minimum = 5
$numMin.Maximum = 480
$numMin.Increment = 5
$numMin.Value = 30
$numMin.Location = New-Object System.Drawing.Point(165, 126)
$numMin.Width = 70

# --- optional lunch break ---
$chkLunch = New-Object System.Windows.Forms.CheckBox
$chkLunch.Text = 'Lunch break (not available)'
$chkLunch.AutoSize = $true
$chkLunch.Checked = $false
$chkLunch.Location = New-Object System.Drawing.Point(290, 127)

$lblLunchFrom = New-Label 'From:' 500 130
$tpLunchStart = New-TimePicker 540 126 12 0
$lblLunchTo   = New-Label 'To:' 635 130
$tpLunchEnd   = New-TimePicker 665 126 13 0
$tpLunchStart.Enabled = $false
$tpLunchEnd.Enabled   = $false

$chkLunch.Add_CheckedChanged({
    $tpLunchStart.Enabled = $chkLunch.Checked
    $tpLunchEnd.Enabled   = $chkLunch.Checked
})

$gbDays = New-Object System.Windows.Forms.GroupBox
$gbDays.Text = 'Days of the week to EXCLUDE'
$gbDays.Location = New-Object System.Drawing.Point(12, 160)
$gbDays.Size = New-Object System.Drawing.Size(930, 56)

$dayNames = @(
    @{ N = 'Monday';    D = [System.DayOfWeek]::Monday },
    @{ N = 'Tuesday';   D = [System.DayOfWeek]::Tuesday },
    @{ N = 'Wednesday'; D = [System.DayOfWeek]::Wednesday },
    @{ N = 'Thursday';  D = [System.DayOfWeek]::Thursday },
    @{ N = 'Friday';    D = [System.DayOfWeek]::Friday },
    @{ N = 'Saturday';  D = [System.DayOfWeek]::Saturday },
    @{ N = 'Sunday';    D = [System.DayOfWeek]::Sunday }
)
$dayChecks = @()
$i = 0
foreach ($d in $dayNames) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $d.N
    $cb.Tag = [int]$d.D
    $cb.AutoSize = $true
    $cb.Location = New-Object System.Drawing.Point((14 + $i * 128), 24)
    $cb.Checked = ($d.D -eq [System.DayOfWeek]::Saturday -or $d.D -eq [System.DayOfWeek]::Sunday)
    $gbDays.Controls.Add($cb)
    $dayChecks += $cb
    $i++
}

$chkAllDay = New-Object System.Windows.Forms.CheckBox
$chkAllDay.Text = 'All-day events block the whole day'
$chkAllDay.AutoSize = $true
$chkAllDay.Checked = $true
$chkAllDay.Location = New-Object System.Drawing.Point(14, 226)

$chkFree = New-Object System.Windows.Forms.CheckBox
$chkFree.Text = 'Ignore events marked "Free"'
$chkFree.AutoSize = $true
$chkFree.Checked = $true
$chkFree.Location = New-Object System.Drawing.Point(14, 252)

$btnCalc = New-Object System.Windows.Forms.Button
$btnCalc.Text = 'Calculate availability'
$btnCalc.Size = New-Object System.Drawing.Size(200, 34)
$btnCalc.Location = New-Object System.Drawing.Point(520, 236)
$btnCalc.Enabled = $false

$lblStatus = New-Label 'Connecting to Outlook...' 14 286

$top.Controls.AddRange(@(
    (New-Label 'Calendar:' 14 16), $cmbCal, $btnRefresh, $btnShared,
    (New-Label 'Start date:' 14 62), $dtFrom,
    (New-Label 'End date:' 270 62), $dtTo,
    (New-Label 'Start time:' 14 96), $tpStart,
    (New-Label 'End time:' 270 96), $tpEnd,
    (New-Label 'Minimum duration (min):' 14 130), $numMin,
    $chkLunch, $lblLunchFrom, $tpLunchStart, $lblLunchTo, $tpLunchEnd,
    $gbDays, $chkAllDay, $chkFree, $btnCalc, $lblStatus
))

# --- results ---
$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'

$lv = New-Object System.Windows.Forms.ListView
$lv.View = 'Details'
$lv.FullRowSelect = $true
$lv.GridLines = $true
$lv.Dock = 'Fill'
[void]$lv.Columns.Add('Day', 140)
[void]$lv.Columns.Add('From', 70)
[void]$lv.Columns.Add('To', 70)
[void]$lv.Columns.Add('Duration', 90)
$split.Panel1.Controls.Add($lv)

$tb = New-Object System.Windows.Forms.TextBox
$tb.Multiline = $true
$tb.ScrollBars = 'Vertical'
$tb.Dock = 'Fill'

$bottom = New-Object System.Windows.Forms.Panel
$bottom.Dock = 'Bottom'
$bottom.Height = 44

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = 'Copy text'
$btnCopy.Size = New-Object System.Drawing.Size(120, 28)
$btnCopy.Location = New-Object System.Drawing.Point(8, 8)
$btnCopy.Enabled = $false

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export CSV...'
$btnExport.Size = New-Object System.Drawing.Size(120, 28)
$btnExport.Location = New-Object System.Drawing.Point(138, 8)
$btnExport.Enabled = $false
$bottom.Controls.AddRange(@($btnCopy, $btnExport))

$lblText = New-Label 'Text (editable, ready to paste into an email):' 0 0
$lblText.Dock = 'Top'
$lblText.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)

$split.Panel2.Controls.Add($tb)
$split.Panel2.Controls.Add($bottom)
$split.Panel2.Controls.Add($lblText)

# Order matters: add the Fill control first, then the Top one
$form.Controls.Add($split)
$form.Controls.Add($top)

# --- button logic ---

function Update-CalendarList {
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblStatus.Text = 'Reading the calendar list from Outlook...'
    $form.Refresh()
    try {
        Write-Log 'Calendar list: start'
        $folders = @(Get-CalendarFolders)
        Write-Log ('Calendar list: found ' + $folders.Count)
        $script:calFolders = New-Object System.Collections.Generic.List[object]
        $cmbCal.Items.Clear()
        foreach ($f in $folders) {
            # An unreadable calendar (e.g. mailbox no longer reachable) is skipped, it does not block the rest
            try {
                $name = ([string]$f.FolderPath).TrimStart('\')
                if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$f.Name }
                [void]$cmbCal.Items.Add($name)
                $script:calFolders.Add($f)
            } catch {
                Write-Log ('Calendar skipped: ' + $_.Exception.Message)
            }
        }
        if ($cmbCal.Items.Count -gt 0) {
            $cmbCal.SelectedIndex = 0
            $btnCalc.Enabled = $true
            $lblStatus.Text = "$($cmbCal.Items.Count) calendars found. The first one is the default calendar."
        } else {
            $btnCalc.Enabled = $false
            $lblStatus.Text = 'No calendar found.'
        }
    } catch {
        $btnCalc.Enabled = $false
        $lblStatus.Text = 'Outlook is not reachable.'
        Show-Error 'Outlook is not reachable' (Get-FailureText $_)
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# The calendar list is loaded AFTER the window is shown, using a timer:
# this way an error is not lost inside the Shown event (which showed
# "The pipeline has been stopped" and hid the real cause)
$startTimer = New-Object System.Windows.Forms.Timer
$startTimer.Interval = 300
$startTimer.Add_Tick({
    $startTimer.Stop()
    try { Update-CalendarList } catch { Show-Error 'Unexpected error' (Get-FailureText $_) }
})

$form.Add_Shown({
    try { $split.SplitterDistance = [int]($split.Width * 0.45) } catch { Write-Log ('SplitterDistance: ' + $_.Exception.Message) }
    $startTimer.Start()
})

$btnRefresh.Add_Click({
    try { Update-CalendarList } catch { Show-Error 'Unexpected error' (Get-FailureText $_) }
})

$btnShared.Add_Click({
    $addr = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Email address (or name) of the person or shared mailbox whose calendar you want.`r`nYou need read permission on their calendar.",
        'Shared calendar', '')
    if ([string]::IsNullOrWhiteSpace($addr)) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Connect-Outlook
        $rcp = $script:ns.CreateRecipient($addr.Trim())
        [void]$rcp.Resolve()
        if (-not $rcp.Resolved) { throw "Could not find '$addr' in the address book." }
        $f = $script:ns.GetSharedDefaultFolder($rcp, 9)    # 9 = olFolderCalendar
        $script:calFolders.Add($f)
        $idx = $cmbCal.Items.Add(([string]$f.FolderPath).TrimStart('\'))
        $cmbCal.SelectedIndex = $idx
        $btnCalc.Enabled = $true
        $lblStatus.Text = 'Shared calendar added to the list.'
    } catch {
        Show-Error 'Shared calendar not available' ("Unable to open the calendar of '$addr'. " +
            "Check the address and that you have read permission.`r`n`r`nDetails: " + (Get-FailureText $_))
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnCalc.Add_Click({
    try {
        if ($cmbCal.SelectedIndex -lt 0) { Show-Error 'Calendar missing' 'Select a calendar.'; return }
        $from = $dtFrom.Value.Date
        $to   = $dtTo.Value.Date
        $ds   = $tpStart.Value.TimeOfDay
        $de   = $tpEnd.Value.TimeOfDay
        $lunchOn = $chkLunch.Checked
        $ls   = $tpLunchStart.Value.TimeOfDay
        $le   = $tpLunchEnd.Value.TimeOfDay
        if ($to -lt $from) { Show-Error 'Invalid dates' 'The end date is earlier than the start date.'; return }
        if ($de -le $ds)   { Show-Error 'Invalid times' 'The end time must be later than the start time.'; return }
        if ($lunchOn -and $le -le $ls) { Show-Error 'Invalid lunch break' 'The lunch break end time must be later than its start time.'; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lblStatus.Text = 'Reading events from Outlook...'
        $form.Refresh()

        $folder = $script:calFolders[$cmbCal.SelectedIndex]
        $events = @(Read-OutlookEvents $folder $from $to.AddDays(1))

        $excluded = @($dayChecks | Where-Object { $_.Checked } | ForEach-Object { [int]$_.Tag })
        $min = [timespan]::FromMinutes([double]$numMin.Value)

        $script:slots = @(Get-FreeSlots $events $from $to $ds $de $excluded $min `
                          $chkAllDay.Checked $chkFree.Checked $lunchOn $ls $le)

        $lv.BeginUpdate()
        $lv.Items.Clear()
        foreach ($s in $script:slots) {
            $item = New-Object System.Windows.Forms.ListViewItem((Format-Day $s.Date))
            [void]$item.SubItems.Add((Format-Time $s.From))
            [void]$item.SubItems.Add((Format-Time $s.To))
            [void]$item.SubItems.Add((Format-Duration ($s.To - $s.From)))
            [void]$lv.Items.Add($item)
        }
        $lv.EndUpdate()

        $tb.Text = Build-Text $script:slots
        $has = ($script:slots.Count -gt 0)
        $btnCopy.Enabled = $has
        $btnExport.Enabled = $has
        $lblStatus.Text = "$($events.Count) events read from '$($cmbCal.SelectedItem)'; $($script:slots.Count) free slots."
    } catch {
        $lblStatus.Text = 'Error.'
        Show-Error 'Error while reading or calculating' (Get-FailureText $_)
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnCopy.Add_Click({
    if ($tb.Text.Length -gt 0) { [System.Windows.Forms.Clipboard]::SetText($tb.Text) }
})

$btnExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = 'Export availability'
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = 'availability.csv'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    try {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("Date;From;To;Duration (min)`r`n")
        foreach ($s in $script:slots) {
            [void]$sb.Append($s.Date.ToString('dd/MM/yyyy', $script:inv) + ';' + (Format-Time $s.From) + ';' +
                             (Format-Time $s.To) + ';' + [int](($s.To - $s.From).TotalMinutes) + "`r`n")
        }
        [System.IO.File]::WriteAllText($dlg.FileName, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    } catch {
        Show-Error 'Unable to save the file' $_.Exception.Message
    }
})

[void]$form.ShowDialog()
