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

function Get-MemberMatchNames {
 param($Member)

 $lastName = ([string]$Member.NameLast).Trim()
 $firstName = ([string]$Member.NameFirst).Trim()
 $preferredName = ([string]$Member.NameFirstPreferred).Trim()
 if ($lastName) {
  if ($firstName) { "$lastName, $firstName" }
  if ($preferredName -and $preferredName -ne $firstName) { "$lastName, $preferredName" }
 }
}

function Select-Absences {
 param(
  [array]$calendarMembers,
  [array]$absences
 )

 $SelectedAbsences = foreach ($Member in $calendarMembers) {
  $names = @(Get-MemberMatchNames -Member $Member)
  if (-not $names.Count) {
   Write-Warning "Skipping calendar member with incomplete name (employee ID: $($Member.empid))."
   continue
  }

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
  [array]$CalendarMembers,
  [array]$CalendarEvents,
  [switch]$WhatIf
 )

 $categoryPath = Join-Path $PSScriptRoot 'lib\color_categories.csv'
 $categoryRows = @(Import-Csv -LiteralPath $categoryPath -ErrorAction Stop)
 $categoriesByReason = @{}
 foreach ($categoryRow in $categoryRows) {
  $categoryReason = ([string]$categoryRow.absence_reason -replace '\s+', ' ').Trim()
  $categoryName = ([string]$categoryRow.default_category).Trim()
  if (-not $categoryReason -or -not $categoryName) {
   throw "Invalid absence reason or category in '$categoryPath'."
  }
  if ($categoriesByReason.ContainsKey($categoryReason)) {
   throw "Duplicate absence reason '$categoryReason' in '$categoryPath'."
  }
  $categoriesByReason[$categoryReason] = $categoryName
 }
 if (-not $categoriesByReason.ContainsKey('Release')) {
  throw "The default Release category mapping is missing from '$categoryPath'."
 }

 $membersByName = @{}
 foreach ($member in $CalendarMembers) {
  foreach ($memberName in @(Get-MemberMatchNames -Member $member)) {
   if (-not $membersByName.ContainsKey($memberName)) { $membersByName[$memberName] = [System.Collections.Generic.List[object]]::new() }
   $membersByName[$memberName].Add($member)
  }
 }

 $candidates = foreach ($absence in $Absences) {
  $name = ([string]$absence.employee_full_name).Trim()
  $reason = ([string]$absence.absence_reason).Trim()
  $absenceDate = [datetime]::MinValue
  if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($reason) -or
   -not [datetime]::TryParseExact([string]$absence.date, [string[]]@('M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd'), [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$absenceDate)) {
   throw 'An absence is missing a valid employee name, absence reason, or date.'
  }
  if (-not $membersByName.ContainsKey($name) -or $membersByName[$name].Count -ne 1) {
   throw "Cannot uniquely match absence employee '$name' to a calendar member."
  }
  $member = $membersByName[$name][0]
  $firstName = ([string]$member.NameFirst).Trim()
  $preferredName = ([string]$member.NameFirstPreferred).Trim()
  $displayName = if ($preferredName) { $preferredName } else { $firstName }
  $lastInitial = ([string]$member.NameLastInitial).Trim()
  if (-not $displayName -or -not $lastInitial) {
   throw "Calendar member '$name' is missing a first name or NameLastInitial."
  }
  $categoryReason = ($reason -replace '\s+', ' ').Trim()
  if ($categoriesByReason.ContainsKey($categoryReason)) {
   $category = $categoriesByReason[$categoryReason]
  }
  else {
   $category = $categoriesByReason['Release']
   Write-Warning "No category mapping for absence reason '$reason'; using the Release category '$category'."
  }
  $memberFirstName = if ($firstName) { $firstName } else { $preferredName }
  [pscustomobject]@{ Subject = "$displayName $lastInitial - $reason"; Date = $absenceDate.Date; MemberName = "$(([string]$member.NameLast).Trim()), $memberFirstName"; Category = $category }
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

 $selected = @{}
 $pending = foreach ($candidate in $candidates) {
  $key = "$($candidate.Date.ToString('yyyy-MM-dd'))|$($candidate.Subject)"
  if ($selected.ContainsKey($key)) {
   if ($selected[$key] -ne $candidate.MemberName) {
    throw "Multiple calendar members have the same event subject and date: '$($candidate.Subject)' on $($candidate.Date.ToString('yyyy-MM-dd'))."
   }
   continue
  }
  $selected[$key] = $candidate.MemberName
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
  # New-MgUserEvent -UserId $Email -Subject $subject -Start $start -End $end -IsAllDay -ShowAs 'free' -Categories @($item.Category) -WhatIf:$WhatIf -ErrorAction Stop
 }
}
# =======================================================================
Import-Module CommonScriptFunctions
Show-BlockInfo Start
Import-Module dbatools
Import-Module Microsoft.Graph.Authentication -Cmdlet Connect-MgGraph, Disconnect-MgGraph
Import-Module Microsoft.Graph.Calendar -Cmdlet Get-MgUserCalendarView, New-MgUserEvent, Remove-MgUserEvent


Show-BlockInfo Main
# App-only auth: the app registration needs admin-consented Application permissions (User.Read.All, GroupMember.Read.All, Calendars.ReadWrite)


$sqlInstance = Connect-DbaInstance -SqlInstance $SqlServer -SqlCredential $SqlCredential -Database $SqlDatabase
$membersSql = Get-Content '.\sql\get-it-staff.sql' -Raw
$calendarMembers = Invoke-DbaQuery -SqlInstance $sqlInstance -Query $membersSql
# $calendarMembers

$startDate = [datetime]::Today
# $endDate = (Get-Date).AddYears(1)

$stopTime = Get-Date '4:00PM'
$delay = 14400 # 4 hours in seconds
if (!$WhatIf) { Write-Host "Running every $($delay /60 /60) hours until $($stopTime.ToString('h:mm tt'))..." -F Green }

do {
 $absenceRows = Get-Absences -apiKey $FrontlineApiKey -reportId $ReportId -startDate $startDate -endDate $EndDate
 # $absenceRows | Select-Object -Property absence_reason -Unique

 Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $GraphCredential -NoWelcome
 $calendarEvents = Get-CalendarEvents -email $CalendarEmail -startDate $startDate -endDate $EndDate

 $events = Select-Absences -calendarMembers $calendarMembers -absences $absenceRows
 Add-MissingAbsenceEvents -Email $CalendarEmail -Absences @($events | Where-Object { $null -ne $_ } | ForEach-Object { $_.Group }) -CalendarMembers $calendarMembers -CalendarEvents $calendarEvents -WhatIf:$WhatIf

 Disconnect-MgGraph | Out-Null

 if ($WhatIf) { break }
 Start-Sleep -Seconds $delay

} until ((Get-Date) -lt $stopTime)

Show-BlockInfo End