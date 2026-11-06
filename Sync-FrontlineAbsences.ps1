# https://www.aesoponline.com/webservices/dataextract.asmx?op=GetData
[CmdletBinding()]
param(
 [string]$FrontlineApiKey,
 [int]$ReportId,
 [datetime]$EndDate,
 [string]$TenantId,
 [PSCredential]$GraphCredential,  # UserName = Entra app client ID, Password = client secret
 [string]$CalendarEmail,  # Shared group availability calendar mailbox
 [string]$SqlServer,
 [PSCredential]$SqlCredential,
 [string]$SqlDatabase,
 [Alias('wi')][switch]$WhatIf
)

function Get-Absences {
 param(
  [string]$apiKey,
  [int]$reportId,
  [datetime]$startDate,
  [datetime]$endDate
 )

 if ($endDate -lt $startDate) { throw 'EndDate must not be earlier than StartDate.' }
 $QueryDate = if ($startDate -eq $endDate) { $startDate.ToString('yyyy-MM-dd') } else { "$($startDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" }
 Write-Host "Retrieving Frontline absences for $QueryDate..." -ForegroundColor Cyan
 $FrontlineEndpoint = 'https://absence.frontlineeducation.com/webservices/dataextract.asmx/GetData'

 # The report's Date filter takes M-d-yyyy values (inclusive range)
 $FilterXml = '<Filters mode="add"><Filter name="Date" value1="{0}" value2="{1}" /></Filters>' -f $StartDate.ToString('M-d-yyyy'), $EndDate.ToString('M-d-yyyy')
 $FormBody = @{
  TemplateId    = $ReportId
  Filters       = $FilterXml
  WebServiceKey = $FrontlineApiKey
  extract       = 'false'
 }

 try {
  $Response = Invoke-WebRequest -Uri $FrontlineEndpoint -Method Post -Body $FormBody -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
  $Result = $Response.Content
  $Absences = [xml]$Result
 }
 catch {
  Write-Error "Failed to fetch absences from Frontline: $_"
  exit
 }

 # Each <row> carries the absence fields as attributes.
 $RowNodes = @($Absences.SelectNodes('//row'))

 $AbsenceRows = foreach ($Node in $RowNodes) {
  $Row = [ordered]@{}
  foreach ($Attr in $Node.Attributes) { $Row[$Attr.Name] = $Attr.Value }
  foreach ($Child in $Node.ChildNodes) {
   if ($Child.NodeType -eq 'Element') { $Row[$Child.Name] = $Child.InnerText }
  }
  [pscustomobject]$Row
 }

 if (-not $AbsenceRows) {
  Write-Warning "No absence rows found. Raw response ($($Result.Length) chars):"
  $Result
  return
 }

 Write-Host "Found $(@($AbsenceRows).Count) absence record(s) for $QueryDate." -ForegroundColor Green
 # $AbsenceRows | Format-Table -AutoSize
 $AbsenceRows
}

function Get-CalendarEvents {
 param(
  [string]$email,
  [datetime]$startDate,
  [datetime]$endDate
 )

 $QueryDate = if ($startDate -eq $endDate) { $startDate.ToString('yyyy-MM-dd') } else { "$($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))" }
 Write-Host "Retrieving Outlook calendar events for $QueryDate..." -ForegroundColor Cyan

 try {
  $pacific = [TimeZoneInfo]::FindSystemTimeZoneById('Pacific Standard Time')
  $startDay = [datetime]::SpecifyKind($startDate.Date, [DateTimeKind]::Unspecified)
  $endDay = [datetime]::SpecifyKind($endDate.Date.AddDays(1), [DateTimeKind]::Unspecified)
  $rangeStart = [DateTimeOffset]::new($startDay, $pacific.GetUtcOffset($startDay)).ToString('o')
  $rangeEnd = [DateTimeOffset]::new($endDay, $pacific.GetUtcOffset($endDay)).ToString('o')
  $Events = Get-MgUserCalendarView -UserId $email -StartDateTime $rangeStart -EndDateTime $rangeEnd -Headers @{ Prefer = 'outlook.timezone="Pacific Standard Time"' } -All -ErrorAction Stop
 }
 catch {
  Write-Error "Failed to fetch events from Outlook calendar: $_"
  exit
 }

 Write-Host "Found $(@($Events | Where-Object { $null -ne $_ }).Count) event(s) for $QueryDate." -ForegroundColor Green
 $Events
}

function Select-Absences {
 param(
  [array]$calendarMembers,
  [array]$absences
 )

 $SelectedAbsences = foreach ($Member in $calendarMembers) {
  $lastName = ([string]$Member.NameLast).Trim()
  $firstName = ([string]$Member.NameFirst).Trim()
  $preferredName = ([string]$Member.NameFirstPreferred).Trim()
  if (-not $lastName -or (-not $firstName -and -not $preferredName)) {
   Write-Warning "Skipping calendar member with incomplete name (employee ID: $($Member.empid))."
   continue
  }

  $names = @()
  if ($firstName) { $names += "$lastName, $firstName" }
  if ($preferredName) { $names += "$lastName, $preferredName" }
  $MemberAbsences = $absences | Where-Object { ([string]$_.employee_full_name).Trim() -in $names }
  $MemberAbsences
 }

 Write-Host "Selected $(@($SelectedAbsences).Count) absence record(s) for calendar members." -ForegroundColor Green
 # $SelectedAbsences | Format-Table -AutoSize
 $SelectedAbsences | Group-Object -Property employee_full_name
}

function Add-MissingAbsenceEvents {
 param(
  [string]$Email,
  [array]$Absences,
  [array]$CalendarEvents,
  [switch]$WhatIf
 )

 $candidates = foreach ($absence in $Absences) {
  $name = ([string]$absence.employee_full_name).Trim()
  $reason = ([string]$absence.absence_reason).Trim()
  $absenceDate = [datetime]::MinValue
  if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($reason) -or
   -not [datetime]::TryParseExact([string]$absence.date, [string[]]@('M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd'), [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$absenceDate)) {
   throw 'An absence is missing a valid employee name, absence reason, or date.'
  }
  [pscustomobject]@{ Subject = "$name - $reason"; Date = $absenceDate.Date }
 }

 $subjects = @{}
 foreach ($candidate in $candidates) { $subjects[$candidate.Subject] = $true }
 $eventRanges = @{}
 foreach ($myEvent in $CalendarEvents) {
  $subject = ([string]$myEvent.Subject).Trim()
  if (-not $subjects.ContainsKey($subject)) { continue }
  $eventStart = [datetime]::MinValue
  $eventEnd = [datetime]::MinValue
  if ($myEvent.Start.TimeZone -ne 'Pacific Standard Time' -or
   $myEvent.End.TimeZone -ne 'Pacific Standard Time' -or
   -not [datetime]::TryParse([string]$myEvent.Start.DateTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$eventStart) -or
   -not [datetime]::TryParse([string]$myEvent.End.DateTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$eventEnd) -or
   $eventEnd -le $eventStart) {
   throw "Cannot determine the Pacific date range for calendar event '$($myEvent.Id)'."
  }
  if (-not $eventRanges.ContainsKey($subject)) { $eventRanges[$subject] = [System.Collections.Generic.List[object]]::new() }
  $eventRanges[$subject].Add([pscustomobject]@{ Start = $eventStart; End = $eventEnd })
 }

 $selected = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
 $pending = foreach ($candidate in $candidates) {
  $key = "$($candidate.Date.ToString('yyyy-MM-dd'))|$($candidate.Subject)"
  if (-not $selected.Add($key)) { continue }
  $match = $false
  foreach ($range in $eventRanges[$candidate.Subject]) {
   if ($range.Start -lt $candidate.Date.AddDays(1) -and $range.End -gt $candidate.Date) {
    $match = $true
    break
   }
  }
  if (-not $match) { $candidate }
 }

 foreach ($item in $pending) {
  $subject = $item.Subject
  $start = @{ DateTime = $item.Date.ToString('yyyy-MM-ddTHH:mm:ss'); TimeZone = 'Pacific Standard Time' }
  $end = @{ DateTime = $item.Date.AddDays(1).ToString('yyyy-MM-ddTHH:mm:ss'); TimeZone = 'Pacific Standard Time' }
  $msg = $MyInvocation.MyCommand.Name, $email, ($subject, $start.DateTime, $end.DateTime -join ' | ')
  Write-Host ('{0},{1},{2}' -f $msg) -ForegroundColor Yellow
  New-MgUserEvent -UserId $Email -Subject $subject -Start $start -End $end -IsAllDay -ShowAs 'free' -WhatIf:$WhatIf -ErrorAction Stop
 }
}
# =======================================================================
Import-Module CommonScriptFunctions
Import-Module dbatools
Import-Module Microsoft.Graph.Authentication -Cmdlet Connect-MgGraph, Disconnect-MgGraph
Import-Module Microsoft.Graph.Calendar -Cmdlet Get-MgUserCalendarView, New-MgUserEvent, Remove-MgUserEvent

# App-only auth: the app registration needs admin-consented Application permissions (User.Read.All, GroupMember.Read.All, Calendars.ReadWrite)
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $GraphCredential -NoWelcome

$sqlInstance = Connect-DbaInstance -SqlInstance $SqlServer -SqlCredential $SqlCredential -Database $SqlDatabase
$membersSql = Get-Content '.\sql\get-it-staff.sql' -Raw
$calendarMembers = Invoke-DbaQuery -SqlInstance $sqlInstance -Query $membersSql
# $calendarMembers

$startDate = [datetime]::Today
$endDate = (Get-Date).AddYears(1)

$absenceRows = Get-Absences -apiKey $FrontlineApiKey -reportId $ReportId -startDate $startDate -endDate $EndDate
$absenceRows | Select-Object -Property absence_reason -Unique
$calendarEvents = Get-CalendarEvents -email $CalendarEmail -startDate $startDate -endDate $EndDate

$events = Select-Absences -calendarMembers $calendarMembers -absences $absenceRows
Add-MissingAbsenceEvents -Email $CalendarEmail -Absences @($events | Where-Object { $null -ne $_ } | ForEach-Object { $_.Group }) -CalendarEvents $calendarEvents -WhatIf:$WhatIf
Disconnect-MgGraph | Out-Null