# https://www.aesoponline.com/webservices/dataextract.asmx?op=GetData
[CmdletBinding()]
param(
 [Parameter(Mandatory)][string]$FrontlineApiKey,
 [Parameter(Mandatory)][int]$ReportId,
 [Parameter(Mandatory)][datetime]$EndDate,
 [Parameter(Mandatory)][string]$TenantId,
 [Parameter(Mandatory)][PSCredential]$GraphCredential,  # UserName = Entra app client ID, Password = client secret
 [Parameter(Mandatory)][string]$CalendarEmail,  # Shared group availability calendar mailbox
 [Parameter(Mandatory)][string]$SqlServer,
 [Parameter(Mandatory)][PSCredential]$SqlCredential,
 [Parameter(Mandatory)][string]$SqlDatabase,
 [Alias('wi')][switch]$WhatIf
)
Import-Module CommonScriptFunctions
Import-Module dbatools
Import-Module Microsoft.Graph.Authentication -Cmdlet Connect-MgGraph, Disconnect-MgGraph -Force
Import-Module Microsoft.Graph.Calendar -Cmdlet Get-MgUserCalendarView, New-MgUserEvent -Force
# =======================================================================
# Constants
# =======================================================================
$PacificTimeZoneId = 'Pacific Standard Time'
$FrontlineEndpoint = 'https://absence.frontlineeducation.com/webservices/dataextract.asmx/GetData'
$FallbackReason = 'Release'   # absence reason whose category is used when a reason has no mapping
$PollSeconds = 3600          # 1 hour

# =======================================================================
# Generic helpers (pure)
# =======================================================================

# Pipeline -> hashtable. The only mutation in the script lives here, so callers stay declarative.
function ConvertTo-Lookup {
 param(
  [Parameter(ValueFromPipeline)]$InputObject,
  [Parameter(Mandatory)][scriptblock]$Key,
  [Parameter(Mandatory)][scriptblock]$Value,
  [switch]$Ordered
 )
 begin { $lookup = if ($Ordered) { [ordered]@{} } else { @{} } }
 process { $lookup[($InputObject | ForEach-Object $Key)] = $InputObject | ForEach-Object $Value }
 end { $lookup }
}

function ConvertTo-NormalizedText {
 param([string]$Text)
 ($Text -replace '\s+', ' ').Trim()
}

function Format-DateRange {
 param([datetime]$Start, [datetime]$End)
 $s = $Start.ToString('yyyy-MM-dd')
 $e = $End.ToString('yyyy-MM-dd')
 if ($s -eq $e) { $s } else { "$s to $e" }
}

# =======================================================================
# Pure transformations
# =======================================================================

function New-FrontlineFilterXml {
 param([datetime]$StartDate, [datetime]$EndDate)
 # The report's Date filter takes M-d-yyyy values (inclusive range)
 '<Filters mode="add"><Filter name="Date" value1="{0}" value2="{1}" /></Filters>' -f $StartDate.ToString('M-d-yyyy'), $EndDate.ToString('M-d-yyyy')
}

# Each <row> carries the absence fields as attributes and/or child elements.
function ConvertFrom-FrontlineXml {
 param([xml]$Xml)
 $Xml.SelectNodes('//row') | ForEach-Object {
  $node = $_
  $attributes = @($node.Attributes | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Value = $_.Value } })
  $children = @($node.ChildNodes |
    Where-Object { $_.NodeType -eq 'Element' } |
     ForEach-Object { [pscustomobject]@{ Name = $_.Name; Value = $_.InnerText } })

    [pscustomobject](($attributes + $children) | ConvertTo-Lookup -Key { $_.Name } -Value { $_.Value } -Ordered)
   }
  }

  function ConvertTo-PacificOffsetString {
   param([datetime]$Day)
   $tz = [TimeZoneInfo]::FindSystemTimeZoneById($PacificTimeZoneId)
   $unspecified = [datetime]::SpecifyKind($Day, [DateTimeKind]::Unspecified)
   [DateTimeOffset]::new($unspecified, $tz.GetUtcOffset($unspecified)).ToString('o')
  }

  function Get-PacificRange {
   param([datetime]$StartDate, [datetime]$EndDate)
   [pscustomobject]@{
    Start = ConvertTo-PacificOffsetString -Day $StartDate.Date
    End   = ConvertTo-PacificOffsetString -Day $EndDate.Date.AddDays(1)
   }
  }

  function Get-MemberMatchNames {
   param($Member)
   $last = ([string]$Member.NameLast).Trim()
   $first = ([string]$Member.NameFirst).Trim()
   $preferred = ([string]$Member.NameFirstPreferred).Trim()

   if (-not $last) { return }
   $firstNames = @($first) + @($preferred | Where-Object { $_ -and $_ -ne $first })
   $firstNames | Where-Object { $_ } | ForEach-Object { "$last, $_" }
  }

  # name -> member lookup (values are collections so duplicate names can be detected)
  function New-MemberIndex {
   param([array]$Members)

   $Members |
    Where-Object { -not @(Get-MemberMatchNames -Member $_).Count } |
     ForEach-Object { Write-Warning "Skipping calendar member with incomplete name (employee ID: $($_.empid))." }

 $pairs = @($Members | ForEach-Object {
   $member = $_
   Get-MemberMatchNames -Member $member | ForEach-Object { [pscustomobject]@{ Name = $_; Member = $member } }
  })

 if ($pairs.Count -eq 0) { return @{} }
 $pairs | Group-Object -Property Name -AsHashTable -AsString
}

function ConvertTo-AbsenceDate {
 param([string]$Text)
 $parsed = [datetime]::MinValue
 $formats = [string[]]@('M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd')
 if ([datetime]::TryParseExact($Text, $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
  $parsed.Date
 }
}

# "8:00 AM", "08:00", "13:30:00" -> time of day; $null when blank or unparseable
function ConvertTo-AbsenceTime {
 param([string]$Text)
 $parsed = [datetime]::MinValue
 if ([datetime]::TryParse($Text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
  $parsed.TimeOfDay
 }
}

# Absence row -> candidate calendar event (throws on data it cannot map safely)
function ConvertTo-AbsenceCandidate {
 param(
  [Parameter(Mandatory, ValueFromPipeline)]$Absence,
  [Parameter(Mandatory)][hashtable]$MemberIndex,
  [Parameter(Mandatory)][hashtable]$CategoryMap
 )
 process {
  $name = ([string]$Absence.employee_full_name).Trim()
  $reason = ([string]$Absence.absence_reason).Trim()
  $date = ConvertTo-AbsenceDate -Text ([string]$Absence.date)

  if (-not $name -or -not $reason -or -not $date) {
   throw 'An absence is missing a valid employee name, absence reason, or date.'
  }

  $found = if ($MemberIndex.ContainsKey($name)) { @($MemberIndex[$name]) } else { @() }
  if ($found.Count -ne 1) { throw "Cannot uniquely match absence employee '$name' to a calendar member." }
  $member = $found[0].Member

  $first = ([string]$member.NameFirst).Trim()
  $preferred = ([string]$member.NameFirstPreferred).Trim()
  $lastInitial = ([string]$member.NameLastInitial).Trim()
  $displayName = if ($preferred) { $preferred } else { $first }
  if (-not $displayName -or -not $lastInitial) {
   throw "Calendar member '$name' is missing a first name or NameLastInitial."
  }

  $reasonKey = ConvertTo-NormalizedText -Text $reason
  $reasonLabel = ($reasonKey -split ' ')[0]   # first word only, used in the event subject
  $category = if ($CategoryMap.ContainsKey($reasonKey)) { $CategoryMap[$reasonKey] }
  else {
   Write-Warning "No category mapping for absence reason '$reason'; using the $FallbackReason category '$($CategoryMap[$FallbackReason])'."
   $CategoryMap[$FallbackReason]
  }

  # Blank start and end times mean a full-day absence; otherwise both must be valid times on the same day.
  $startText = ([string]$Absence.employee_start_time).Trim()
  $endText = ([string]$Absence.employee_end_time).Trim()
  $isAllDay = -not $startText -and -not $endText
  if ($isAllDay) {
   $start = $date
   $end = $date.AddDays(1)
  }
  else {
   $startTime = ConvertTo-AbsenceTime -Text $startText
   $endTime = ConvertTo-AbsenceTime -Text $endText
   if ($null -eq $startTime -or $null -eq $endTime -or $endTime -le $startTime) {
    throw "Absence for '$name' on $($date.ToString('yyyy-MM-dd')) has an invalid time range ('$startText' to '$endText')."
   }
   $start = $date + $startTime
   $end = $date + $endTime
  }

  [pscustomobject]@{
   Subject    = "$displayName $lastInitial - $reasonLabel"
   Date       = $date
   Start      = $start
   End        = $end
   IsAllDay   = $isAllDay
   MemberName = "$(([string]$member.NameLast).Trim()), $(if ($first) { $first } else { $preferred })"
   Category   = $category
  }
 }
}

# One candidate per (start, end, subject); the same subject and time range for different people is ambiguous.
function Get-UniqueCandidates {
 param([array]$Candidates)
 $Candidates |
  Group-Object -Property { '{0:yyyy-MM-dd HH:mm}|{1:yyyy-MM-dd HH:mm}|{2}' -f $_.Start, $_.End, $_.Subject } |
   ForEach-Object {
    $first = $_.Group | Select-Object -First 1
    $owners = @($_.Group | Select-Object -ExpandProperty MemberName -Unique)
    if ($owners.Count -gt 1) {
     throw "Multiple calendar members have the same event subject and date: '$($first.Subject)' on $($first.Date.ToString('yyyy-MM-dd'))."
    }
    $first
   }
}

function ConvertTo-EventRange {
 param($CalendarEvent)
 $start = [datetime]::MinValue
 $end = [datetime]::MinValue
 $culture = [System.Globalization.CultureInfo]::InvariantCulture
 $none = [System.Globalization.DateTimeStyles]::None

 $valid = $CalendarEvent.Start.TimeZone -eq $PacificTimeZoneId -and
 $CalendarEvent.End.TimeZone -eq $PacificTimeZoneId -and
 [datetime]::TryParse([string]$CalendarEvent.Start.DateTime, $culture, $none, [ref]$start) -and
 [datetime]::TryParse([string]$CalendarEvent.End.DateTime, $culture, $none, [ref]$end) -and
 $end -gt $start

 if (-not $valid) { throw "Cannot determine the Pacific date range for calendar event '$($CalendarEvent.Id)'." }

 [pscustomobject]@{ Subject = ([string]$CalendarEvent.Subject).Trim(); Start = $start; End = $end }
}

function Test-EventExists {
 param($Candidate, [array]$Ranges)
 @($Ranges | Where-Object { $_.Start -eq $Candidate.Start -and $_.End -eq $Candidate.End }).Count -gt 0
}

# Candidates that no existing event (same subject, same start and end) already covers.
function Select-MissingCandidates {
 param([array]$Candidates, [array]$CalendarEvents)

 $subjects = @($Candidates | Select-Object -ExpandProperty Subject)
 $ranges = @($CalendarEvents |
   Where-Object { ([string]$_.Subject).Trim() -in $subjects } |
    ForEach-Object { ConvertTo-EventRange -CalendarEvent $_ })

 $Candidates | Where-Object {
  $candidate = $_
  -not (Test-EventExists -Candidate $candidate -Ranges @($ranges | Where-Object { $_.Subject -eq $candidate.Subject }))
 }
}

function Select-MemberAbsences {
 param([array]$Absences, [hashtable]$MemberIndex)
 $selected = @($Absences | Where-Object { $MemberIndex.ContainsKey(([string]$_.employee_full_name).Trim()) })
 Write-Host "Selected $($selected.Count) absence record(s) for calendar members." -ForegroundColor Green
 $selected
}

function ConvertTo-EventParams {
 param($Candidate)
 @{
  Subject    = $Candidate.Subject
  Start      = @{ DateTime = $Candidate.Start.ToString('yyyy-MM-ddTHH:mm:ss'); TimeZone = $PacificTimeZoneId }
  End        = @{ DateTime = $Candidate.End.ToString('yyyy-MM-ddTHH:mm:ss'); TimeZone = $PacificTimeZoneId }
  IsAllDay   = $Candidate.IsAllDay
  ShowAs     = 'free'
  Categories = @($Candidate.Category)
 }
}

# =======================================================================
# I/O edges (HTTP, files, Graph)
# =======================================================================

function Import-CategoryMap {
 param([string]$Path, [string]$RequiredReason)

 $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop | ForEach-Object {
   [pscustomobject]@{
    Reason   = ConvertTo-NormalizedText -Text $_.absence_reason
    Category = ([string]$_.default_category).Trim()
   }
  })

 if ($rows | Where-Object { -not $_.Reason -or -not $_.Category }) {
  throw "Invalid absence reason or category in '$Path'."
 }
 $duplicate = $rows | Group-Object -Property Reason | Where-Object Count -GT 1 | Select-Object -First 1
 if ($duplicate) { throw "Duplicate absence reason '$($duplicate.Name)' in '$Path'." }

 $map = $rows | ConvertTo-Lookup -Key { $_.Reason } -Value { $_.Category }
 if (-not $map.ContainsKey($RequiredReason)) {
  throw "The default $RequiredReason category mapping is missing from '$Path'."
 }
 $map
}

function Invoke-FrontlineReport {
 param([string]$ApiKey, [int]$ReportId, [datetime]$StartDate, [datetime]$EndDate)
 $body = @{
  TemplateId    = $ReportId
  Filters       = New-FrontlineFilterXml -StartDate $StartDate -EndDate $EndDate
  WebServiceKey = $ApiKey
  extract       = 'false'
 }
 try {
  (Invoke-WebRequest -Uri $FrontlineEndpoint -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing).Content
 }
 catch { throw "Failed to fetch absences from Frontline: $_" }
}

function Get-Absences {
 param([string]$ApiKey, [int]$ReportId, [datetime]$StartDate, [datetime]$EndDate)

 $label = Format-DateRange -Start $StartDate -End $EndDate
 Write-Host "Retrieving Frontline absences for $label..." -ForegroundColor Cyan

 $content = Invoke-FrontlineReport -ApiKey $ApiKey -ReportId $ReportId -StartDate $StartDate -EndDate $EndDate
 $rows = @(ConvertFrom-FrontlineXml -Xml ([xml]$content))

 if (-not $rows) {
  Write-Warning "No absence rows found (raw response was $($content.Length) chars; use -Verbose to see it)."
  Write-Verbose $content
  return
 }

 Write-Host "Found $($rows.Count) absence record(s) for $label." -ForegroundColor Green
 $rows
}

function Get-CalendarEvents {
 param([string]$Email, [datetime]$StartDate, [datetime]$EndDate)

 $label = Format-DateRange -Start $StartDate -End $EndDate
 Write-Host "Retrieving Outlook calendar events for $label..." -ForegroundColor Cyan

 $range = Get-PacificRange -StartDate $StartDate -EndDate $EndDate
 $calendarEvents = try {
  @(Get-MgUserCalendarView -UserId $Email -StartDateTime $range.Start -EndDateTime $range.End `
    -Headers @{ Prefer = "outlook.timezone=`"$PacificTimeZoneId`"" } -All -ErrorAction Stop |
    Where-Object { $null -ne $_ })
 }
 catch { throw "Failed to fetch events from Outlook calendar: $_" }

 Write-Host "Found $($calendarEvents.Count) event(s) for $label." -ForegroundColor Green
 $calendarEvents
}

function Add-AbsenceEvent {
 param([string]$Email, $Candidate, [switch]$WhatIf)
 $params = ConvertTo-EventParams -Candidate $Candidate
 Write-Host ('{0},{1} | {2} | {3}' -f $Email, $params.Subject, $params.Start.DateTime, $params.End.DateTime) -ForegroundColor Yellow
 New-MgUserEvent -UserId $Email @params -WhatIf:$WhatIf -ErrorAction Stop | Out-Null
}

# =======================================================================
# Orchestration: one sync pass = fetch -> transform -> create what's missing
# =======================================================================

function Invoke-SyncPass {
 param(
  [string]$ApiKey,
  [int]$ReportId,
  [string]$CalendarEmail,
  [datetime]$StartDate,
  [datetime]$EndDate,
  [hashtable]$MemberIndex,
  [hashtable]$CategoryMap,
  [switch]$WhatIf
 )

 $absences = @(Get-Absences -ApiKey $ApiKey -ReportId $ReportId -StartDate $StartDate -EndDate $EndDate)
 $calendarEvents = @(Get-CalendarEvents -Email $CalendarEmail -StartDate $StartDate -EndDate $EndDate)

 $candidates = @(
  Select-MemberAbsences -Absences $absences -MemberIndex $MemberIndex |
   ConvertTo-AbsenceCandidate -MemberIndex $MemberIndex -CategoryMap $CategoryMap
 )

 Select-MissingCandidates -Candidates @(Get-UniqueCandidates -Candidates $candidates) -CalendarEvents $calendarEvents |
  ForEach-Object { Add-AbsenceEvent -Email $CalendarEmail -Candidate $_ -WhatIf:$WhatIf }
}

# =======================================================================
# Main: the only place with startup/shutdown side effects
# =======================================================================

Show-BlockInfo Start

$startDate = [datetime]::Today
if ($EndDate -lt $startDate) { throw 'EndDate must not be earlier than StartDate.' }

# App-only auth: the app registration needs admin-consented Application permissions (User.Read.All, GroupMember.Read.All, Calendars.ReadWrite)
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $GraphCredential -NoWelcome

Show-BlockInfo Main

$categoryMap = Import-CategoryMap -Path (Join-Path $PSScriptRoot 'lib\color_categories.csv') -RequiredReason $FallbackReason

$sqlInstance = Connect-DbaInstance -SqlInstance $SqlServer -SqlCredential $SqlCredential -Database $SqlDatabase
$memberIndex = New-MemberIndex -Members @(
 Invoke-DbaQuery -SqlInstance $sqlInstance -Query (Get-Content (Join-Path $PSScriptRoot 'sql\get-it-staff.sql') -Raw)
)

$syncArgs = @{
 ApiKey        = $FrontlineApiKey
 ReportId      = $ReportId
 CalendarEmail = $CalendarEmail
 StartDate     = $startDate
 EndDate       = $EndDate
 MemberIndex   = $memberIndex
 CategoryMap   = $categoryMap
 WhatIf        = $WhatIf
}

$stopTime = Get-Date '4:00PM'
if (-not $WhatIf) {
 $pollHours = $PollSeconds / 3600
 $pollUnit = if ($pollHours -eq 1) { 'hour' } else { 'hours' }
 Write-Host "Running every $pollHours $pollUnit until $($stopTime.ToString('h:mm tt'))..." -ForegroundColor Green
}

while ($true) {
 # A failed pass is logged and retried next interval; with -WhatIf, fail fast so problems are visible.
 try { Invoke-SyncPass @syncArgs }
 catch {
  if ($WhatIf) { throw }
  Write-Warning "Sync pass failed: $_"
 }

 if ($WhatIf) { break }
 # Stop now rather than sleeping past the stop time
 if ((Get-Date).AddSeconds($PollSeconds) -ge $stopTime) { break }
 Start-Sleep -Seconds $PollSeconds
}

Disconnect-MgGraph | Out-Null
Show-BlockInfo End