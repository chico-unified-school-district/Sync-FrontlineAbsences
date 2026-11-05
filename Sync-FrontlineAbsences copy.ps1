# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
# Microsoft Entra ID (Azure AD) Credentials
$TenantId = 'your-azure-tenant-id'
$ClientId = 'your-azure-client-id'
$ClientSecret = 'your-azure-client-secret'
$TargetUserEmail = 'administrator@yourdistrict.org'  # Target Outlook account

# Frontline API Configuration
$FrontlineApiUrl = 'https://api.frontlineeducation.com/v1/absences'  # Or your district endpoint
$FrontlineApiKey = 'your-frontline-api-key'

$TimeZone = 'Pacific Standard Time'  # Set to district timezone
$QueryDate = (Get-Date).ToString('yyyy-MM-dd')

# ---------------------------------------------------------------------------
# 1. AUTHENTICATE WITH MICROSOFT GRAPH API
# ---------------------------------------------------------------------------
Write-Host 'Authenticating with Microsoft Graph...' -ForegroundColor Cyan

$TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$TokenBody = @{
 grant_type    = 'client_credentials'
 client_id     = $ClientId
 client_secret = $ClientSecret
 scope         = 'https://graph.microsoft.com/.default'
}

try {
 $TokenResponse = Invoke-RestMethod -Uri $TokenUrl -Method Post -ContentType 'application/x-www-form-urlencoded' -Body $TokenBody
 $AccessToken = $TokenResponse.access_token
}
catch {
 Write-Error "Failed to authenticate with Microsoft Graph: $_"
 exit
}

# ---------------------------------------------------------------------------
# 2. FETCH ABSENCES FROM FRONTLINE API
# ---------------------------------------------------------------------------
Write-Host "Retrieving Frontline absences for $QueryDate..." -ForegroundColor Cyan

$FrontlineHeaders = @{
 'Authorization' = "Bearer $FrontlineApiKey"
 'Accept'        = 'application/json'
}

$FrontlineEndpoint = "$FrontlineApiUrl`?startDate=$QueryDate&endDate=$QueryDate"

try {
 $Absences = Invoke-RestMethod -Uri $FrontlineEndpoint -Headers $FrontlineHeaders -Method Get
}
catch {
 Write-Error "Failed to fetch absences from Frontline: $_"
 exit
}

# ---------------------------------------------------------------------------
# 3. POST EVENTS TO OUTLOOK CALENDAR
# ---------------------------------------------------------------------------
$GraphHeaders = @{
 'Authorization' = "Bearer $AccessToken"
 'Content-Type'  = 'application/json'
}
$GraphEventsUrl = "https://graph.microsoft.com/v1.0/users/$TargetUserEmail/events"

Write-Host "Syncing $($Absences.Count) record(s) to Outlook..." -ForegroundColor Cyan

foreach ($Absence in $Absences) {
 # Fallbacks for optional fields
 $Employee = if ($Absence.employeeName) { $Absence.employeeName } else { 'Staff Member' }
 $Reason = if ($Absence.reason) { $Absence.reason } else { 'Absence' }
 $ConfNum = if ($Absence.confirmationNumber) { $Absence.confirmationNumber } else { 'N/A' }
 $SubName = if ($Absence.substituteName) { $Absence.substituteName } else { 'Unfilled' }
 $StartTime = if ($Absence.startTime) { $Absence.startTime } else { '08:00:00' }
 $EndTime = if ($Absence.endTime) { $Absence.endTime } else { '15:00:00' }
 $Location = if ($Absence.locationName) { $Absence.locationName } else { 'School Site' }

 # Construct Microsoft Graph Event Payload
 $EventPayload = @{
  subject  = "Absence: $Employee - $Reason"
  body     = @{
   contentType = 'HTML'
   content     = @"
                <p><strong>Confirmation #:</strong> $ConfNum</p>
                <p><strong>Employee:</strong> $Employee</p>
                <p><strong>Reason:</strong> $Reason</p>
                <p><strong>Substitute Assigned:</strong> $SubName</p>
"@
  }
  start    = @{
   dateTime = "$($Absence.startDate)T$StartTime"
   timeZone = $TimeZone
  }
  end      = @{
   dateTime = "$($Absence.endDate)T$EndTime"
   timeZone = $TimeZone
  }
  location = @{
   displayName = $Location
  }
 } | ConvertTo-Json -Depth 4

 try {
  $Response = Invoke-RestMethod -Uri $GraphEventsUrl -Headers $GraphHeaders -Method Post -Body $EventPayload
  Write-Host "Successfully added event: Absence: $Employee - $Reason" -ForegroundColor Green
 }
 catch {
  Write-Warning "Failed to add event for $Employee: $_"
 }
}