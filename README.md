# Sync Frontline Absences

This PowerShell script retrieves absence-report rows from Frontline, selects employees returned by a SQL query, and creates missing all-day events in an Outlook mailbox calendar through Microsoft Graph.

## Prerequisites

- PowerShell with the `dbatools`, `Microsoft.Graph.Authentication`, and `Microsoft.Graph.Calendar` modules installed, plus the organization's `CommonScriptFunctions` module available to PowerShell.
- A Frontline web-service key and the Report Writer template ID for an absence export report. The report must support the `Date` filter and return XML `row` elements with `employee_full_name`, `absence_reason`, and `date` attributes.
- SQL Server access to the employee database and the `vwHREmploymentList` view. The script runs [sql/get-it-staff.sql](sql/get-it-staff.sql) to select active employees using site and job-category conditions; review those filters before a production run.
- A Microsoft Entra app registration with an admin-consented Microsoft Graph **application** permission for calendar access, such as `Calendars.ReadWrite`, and access to the target mailbox under any applicable Exchange access policy. `CalendarEmail` must identify a user or shared mailbox accessible via the Graph user calendar API; a Microsoft 365 group calendar uses a different API.
- The target mailbox and the machine running the script must support the `Pacific Standard Time` time-zone identifier.

The script calls `Show-BlockInfo` from `CommonScriptFunctions` at startup, before the main process, and at the end.

## Parameters

| Parameter         | Purpose                                                                                            |
| ----------------- | -------------------------------------------------------------------------------------------------- |
| `FrontlineApiKey` | Frontline web-service key.                                                                         |
| `ReportId`        | Frontline Report Writer template ID.                                                               |
| `EndDate`         | Inclusive end date for the Frontline and calendar queries.                                         |
| `TenantId`        | Microsoft Entra tenant ID.                                                                         |
| `GraphCredential` | `PSCredential` whose username is the app client ID and whose password is the client secret.        |
| `CalendarEmail`   | Address of the user or shared mailbox whose default calendar is used.                              |
| `SqlServer`       | SQL Server instance name.                                                                          |
| `SqlCredential`   | SQL login as a `PSCredential`.                                                                     |
| `SqlDatabase`     | Employee database name.                                                                            |
| `WhatIf`          | Preview event creation rather than creating events; alias `wi`.                                    |

All parameters except `WhatIf` are mandatory. Pass the required values explicitly or through a parameter hashtable. Do not store keys or passwords in the script or commit them to source control.

**Date range:** The script sets the start to the machine's local date at startup and uses the supplied `EndDate` as the inclusive end date. Supply `EndDate` explicitly; it has no usable default. The start date is fixed for the duration of the process.

## Run

Run from the project directory: the SQL query is loaded relative to the current working directory. Supply the app client ID as the username and the client secret as the password when prompted for the Graph credential.

```powershell
Set-Location 'C:\path\to\Sync-FrontlineAbsences'

$frontlineKey = Read-Host 'Frontline web-service key'
$graphCredential = Get-Credential -Message 'Enter the app client ID and client secret'
$sqlCredential = Get-Credential -Message 'Enter the SQL login'

$params = @{
    FrontlineApiKey = $frontlineKey
    ReportId        = 123456
    EndDate         = (Get-Date).Date.AddYears(1)
    TenantId        = 'your-tenant-id'
    GraphCredential = $graphCredential
    CalendarEmail   = 'calendar@example.org'
    SqlServer       = 'sql-server'
    SqlCredential   = $sqlCredential
    SqlDatabase     = 'employee-database'
}

.\Sync-FrontlineAbsences.ps1 @params -WhatIf
```

Replace the example values with your own. `Read-Host` displays the Frontline key as it is entered; use an approved secret-management method when running in a shared or automated environment. Review the preview before running the same command without `-WhatIf`.

`-WhatIf` applies to event creation only. The script still connects to Graph and SQL Server, retrieves Frontline data, and reads the calendar. Do not assume it makes the entire process offline or prevents all other side effects.

## Schedule and repeat runs

The script queries SQL once at startup, then fetches Frontline and calendar data and attempts synchronization in a loop. A `-WhatIf` run performs one iteration and exits without sleeping.

Without `-WhatIf`, the script sleeps for one hour after each iteration. **The current loop condition is inverted:** it exits after the sleep if the current time is *before* the 4:00 PM stop time, and continues if the time is at or after 4:00 PM. For example, a run whose first sleep ends before 4:00 PM exits after one iteration; a run reaching the check at or after 4:00 PM keeps repeating until stopped. The stop time is fixed at startup, so it does not reset the next day. Do not rely on the printed "until 4:00 PM" message as the actual stopping behavior.

The employee list and start date are not refreshed during the loop. Changes to SQL membership require restarting the process.

## Matching and event creation

1. The SQL query selects calendar members. A Frontline `employee_full_name` matches a member's `NameLast, NameFirst` or `NameLast, NameFirstPreferred`, ignoring case and trimming surrounding whitespace.
2. The event subject is `First LastInitial - absence_reason`. The preferred first name is used when present; otherwise, the first name is used.
3. Existing calendar events are compared by subject and overlapping Pacific calendar date. A match is skipped. Duplicate rows for the same member, subject, and date are also skipped. Ambiguous member matches or subject/date collisions between different members stop the run.
4. Missing events are created as all-day events in Pacific time with availability set to **Free**. The script creates missing events only; it does not update or delete existing events.

The comparison does not use a confirmation number. An existing event with an older subject format will not match the current subject, so the script may create another event. Name-based matching can also be ambiguous when employees share names.

## Categories

[lib/color_categories.csv](lib/color_categories.csv) maps `absence_reason` to `default_category`. The script reads and validates the mappings, normalizing whitespace when comparing reasons. If a reason is not mapped, it warns and uses the `Release` mapping, currently `Blue Category`; the `Release` mapping must be present.

**Category assignment is currently disabled:** the active `New-MgUserEvent` call does not include `-Categories`. A commented-out alternative includes it. If category assignment is enabled, verify that the named categories exist in the target mailbox and have the desired colors; the CSV's `color` and `graph_preset` columns are not applied by the script. Existing events are not updated.

## Troubleshooting

- **Graph 403:** Check the app's application permissions, admin consent, target mailbox, and any Exchange access restrictions.
- **No Frontline rows:** Verify the report ID, web-service key, report data, and its `Date` filter. The script logs a warning and the raw response when no rows are found.
- **Employees missing:** Check the SQL query's filters and compare the report's `Last, First` names with the SQL first and preferred names.
- **Category mapping error:** Check for empty or duplicate reason mappings and a missing `Release` row in the CSV.
- **Unexpected or duplicate events:** Compare the existing subject and date with the current naming scheme. Review a `-WhatIf` run before making changes.

Absence reasons and names are sensitive personnel information. Limit access to the calendar, logs, report output, and credentials according to your organization's policies.
