# https://www.aesoponline.com/webservices/dataextract.asmx?op=GetData
[CmdletBinding()]
param(
 [string]$FrontlineApiKey,
 [datetime]$StartDate = (Get-Date).Date,
 [datetime]$EndDate = $StartDate,
 [string]$TenantId,
 [System.Management.Automation.PSCredential]$GraphCredential,  # UserName = Entra app client ID, Password = client secret
 [string]$CalendarEmail,  # Shared group availability calendar mailbox
 [string]$SqlServer,
 [System.Management.Automation.PSCredential]$SqlCredential,
 [string]$SqlDatabase
)
if ($EndDate -lt $StartDate) { throw 'EndDate must not be earlier than StartDate.' }
$QueryDate = if ($StartDate -eq $EndDate) { $StartDate.ToString('yyyy-MM-dd') } else { "$($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" }
Write-Host "Retrieving Frontline absences for $QueryDate..." -ForegroundColor Cyan
$FrontlineEndpoint = 'https://absence.frontlineeducation.com/webservices/dataextract.asmx/GetData'
$ReportId = 1423663   # Report Writer template ID of the absence export report

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

if ($SqlServer) {
 if (-not (Get-Module -ListAvailable -Name dbatools)) {
  throw 'The dbatools module is required for SQL access. Install it with: Install-Module dbatools'
 }
 Import-Module dbatools

 $SqlConnectParams = @{ SqlInstance = $SqlServer }
 if ($SqlDatabase) { $SqlConnectParams.Database = $SqlDatabase }
 if ($SqlCredential) { $SqlConnectParams.SqlCredential = $SqlCredential }

 try {
  $SqlConnection = Connect-DbaInstance @SqlConnectParams
  Write-Host "Connected to SQL Server $($SqlConnection.Name)." -ForegroundColor Green
 }
 catch {
  Write-Error "Failed to connect to SQL Server '$SqlServer': $_"
  exit
 }
}