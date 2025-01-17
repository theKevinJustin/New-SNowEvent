# New-SNOWEvent
Use PowerShell script to take SCOM alerts to create ITSM ServiceNow (SNOW) events

```
New-SNOWEvent-TeamParameter-MM.ps1 v1.0.0.9
New-SNowEvent - Generic.ps1 v1.0.0.5
New-SNOWEvent.ps1 v1.0.0.0
```

Download [here](https://github.com/theKevinJustin/New-SNowEvent/blob/main/New-SNowEventGeneric.ps1)

### New-SNOWEvent files
Create ServiceNow events leveraging multiple SCOM alert fields, and updates SCOM alert Owner, TicketID, and ResolutionState with successful incident creation.

Blog [(https://kevinjustin.com/blog/2024/03/27/servicenow-event-integration/)](https://kevinjustin.com/blog/2024/03/27/servicenow-event-integration/)

Create SNOW Event

# Testing - depending on how you want to randomly choose an incident 
```
# Newer versions add SNowEnv parameter, leverage SCOM alert TicketID field, and CustomField1 leverages ServiceNow Event/Alert/Incident URL

# Lab example
$Alerts = get-scomalert -resolutionstate 0 | where { $_.Name -like "System Center*" }

# Gather Critical, New alerts
$Alerts = get-scomalert -ResolutionState 0 -severity 2

# Debug for warning alerts
$Alerts = get-scomalert -ResolutionState 0 -severity 1

# Debug
$Alerts[0] | fl ID,Name,Description,Severity,MonitoringObjectDisplayName

$AlertID = $Alerts[0].ID
$AlertName = $Alerts[0].Name
$TeamNameHere = "MECM"

# Run from PowerShell on SCOM MS (with successful pre-requisite verification)
.\New-SNOWEventGeneric.ps1 -AlertName $AlertName -AlertID $AlertID -Team $TeamNameHere

# Example 2
# Run from PowerShell on SCOM MS (with successful pre-requisite verification)
.\New-SNOWEventGeneric.ps1 -SNowEnv Prod -AlertName $AlertName -AlertID $AlertID -Team $TeamNameHere
```
