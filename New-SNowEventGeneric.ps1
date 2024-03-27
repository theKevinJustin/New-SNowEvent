<#
#=================================================================================
# Script test to setup SNOW events using SCOM alert data
#
#  Authors: Steven Brown, Kevin Justin, Joe Kelly
#  v1.0
#=================================================================================

# Global required parameters for ServiceNow (SNow) Event creation

# From SCOM, AlertName parameter is specified in the channel as $Data[Default='Not Present']/Context/DataItem/AlertName$
# AlertID can help link the alert to check for the alert to parse out $HostName, as it can reside in multiple alert properties
# AlertID translates to $Data/Context/DataItem/AlertId$
# Additional fields https://blog.tyang.org/2012/01/29/command-line-parameters-for-scom-command-notification-channel/

# Find and replace various variables before running

# Setup SNOW Event Name standard
Example SNOWAlertName
	$SNOWAlertName = "<Org> <Team> SCOM Test Event - $Alert"
Example SNOWAlertName
	$SNOWAlertName = "<ORG> <Team> SCOM Event - $AlertName"
Example SNOWAlertName
	$SNOWAlertName = "##CUSTOMER## SCOM Event - $AlertName"

# Replace these variables with valid values:
##CUSTOMER##
##TEAM##
##SERVICENOWURL##
##CallerID##

Example $CallerId = "##CallerID##"
Example $CallerID = "13ad1d814fb68c25038e92468c676063c"


# If required, add proxy URL for REST injection
##Proxy##


Hard coded variables for ServiceNow (SNow) for CallerId, URL, and if Proxy needed

Hard code URL, Proxy, and CallerID, ##CUSTOMER## into ServiceNow (SNow) caller_id field to create events
#===============================================

# Don't forget to replace the following variables!

SNOW DEV URL with Prod before going to production events
$ServiceNowURL="https://##ServiceNowURL##/api/now/table/em_event"

# Set AlertName for Testing
$SNOWAlertName = "##CUSTOMER## ##TEAM## SCOM Event - $AlertName"

Proxy
$Proxy = ##Proxy##

CallerID
$CallerID = "##CallerID##"


# AssignmentGroup & TicketID
$AssignmentGroup
$TicketID = "SNOW_event"

#>

# Hard code URL, Proxy, CallerID SNOW variables into script
#===============================================

# Don't forget to replace SNOW DEV URL with Prod before going to production events
$ServiceNowURL = ##ServiceNowURL##/api/now/table/em_event"
$Proxy = ##Proxy##
$CallerID = "##CallerID##"

# Assume module NOT loaded into current PowerShell profile
Import-Module -Name CredentialManager


Param (
     [Parameter(
         Mandatory=$true,
         ValueFromPipeline=$true,
         Position=0)]	 
     [ValidateNotNullorEmpty()]
     [String]$AlertName,
     [Parameter(
         Mandatory=$true,
         ValueFromPipeline=$true,
         Position=1)]
     [ValidateNotNullorEmpty()]
     [String]$AlertID,
     [Parameter(
		 Mandatory=$true,
         ValueFromPipeline=$true,
		 Position=5)]
		 [ValidateNotNullorEmpty()]
		 [String]$AssignmentGroup
)



#=================================================================================
# Starting Script section - All scripts get this
#=================================================================================
# Gather the start time of the script
$StartTime = Get-Date

# Set variable to be used in logging events
$whoami = whoami
 
# ScriptName should match the <scriptname.ps1> to log script details
#=================================================================================
# ScriptName
$ScriptName = "New-SNowEvent.ps1"
$EventID = "700"

# Create new object for MOMScript API, or SCOM alert properties
$momapi = New-Object -comObject MOM.ScriptAPI

# Begin logging script starting into event log
# write-host "Script is starting. `n Running as ($whoami)."
$momapi.LogScriptEvent($ScriptName,$EventID,0,"Script is starting. `n Running as ($whoami).")
#=================================================================================

# PropertyBag Script section - Monitoring scripts get this
#=================================================================================
# Load SCOM PropertyBag function
$bag = $momapi.CreatePropertyBag()

$date = get-date -uFormat "%Y-%m-%d"





<#
.Synopsis
   New-SNowEvent creates ServiceNow (SNow) events
.DESCRIPTION
   Create ServiceNow (SNow) events using New-SNowEvent
.EXAMPLE
   Example of how to use this cmdlet

   To provide input parameters for ALL Event pieces   
      New-SNowEvent -AlertName <> -AlertID <> -AssignmentGroup <> 
.EXAMPLE
   Example of New-SNowEvent required parameters
      New-SNowEvent -AlertName <> -AlertID <> -AssignmentGroup <>
.EXAMPLE
   Example of New-SNowEvent required parameters
      New-SNowEvent -AlertName "System Center Management Health Service Unloaded System Rule(s)" -AlertID 5e0f7d66-1aae-43d6-8002-73f7668dc889 -AssignmentGroup "System Admins"
.INPUTS
   Strings can be used for the following inputs:
	-AlertName 
	-AlertID
	-AssignmentGroup
.OUTPUTS
   Script leverages Add-Event function to create Operations Manager Event ID 700 events, as well as write-host elements to screen.
   NOTE: write-host elements largely disabled and intended for debug purposes.
   
   Example outputs running functions from PowerShell
   PS C:\Users\scomadmin> Get-SNowParameters
	https://##ServiceNowURL##/api/now/table/event
	
	PROD ServiceNow URL specified

	CredentialManager PoSH Module NOT Installed
	
	ServiceNow Credential NOT stored on server
	ServiceNow User NOT stored on server
	ServiceNow Password NOT stored on server
	
   	Additional Output example error - 
		Most likely when script run from non-DOD server, or server NOT on trusted network
	
	PS C:\Users\scomadmin> New-SNowEvent
	Error: The remote certificate is invalid according to the validation procedure.
.NOTES
   Validates required URL, ID/Password is stored on server.

.COMPONENT
   New-SNowEvent script used to create ServiceNow (SNow) events.
.ROLE
   Use New-SNowEvent in ITSM integration for Event Management
.FUNCTIONALITY
   Setup ServiceNow (SNoW) events, based on strategy 'intervention required' monitoring and alerting.
#>



function Add-Event
{
<#
.Synopsis
   Create Events in Operations Manager event log
.DESCRIPTION
   Setup MOMAPI Event Logging to Operations Manager event log

   Create $StartTime, $whoami, $ScriptName,$Event, $momapi
   Begin logging script runtime events using EventID 700.
	
   Log script runtime events using EventID 700.
.EXAMPLE
   Example of how to use this cmdlet:
	Log-Event <string>
.EXAMPLE
   Example using new line `n:
	Log-Event "Script is starting. `n Running as ($whoami)."
.INPUTS
   Input string of what you want added to Operations Manager event log
.OUTPUTS
   Creates EventID 700 events added to Operations Manager event log.
.NOTES
   Use newline `n, or carriage returns `r to format additional lines into Event.
.COMPONENT
   Leverage function to create events with debug or error conditions related to new ServiceNow (SNow) events.
.ROLE
   Used as event logging function
.FUNCTIONALITY
   Create events related to new ServiceNow (SNow) events.
#>

[CmdletBinding()]
Param (
     [Parameter(
         Mandatory=$true,
         ValueFromPipeline=$true,
         Position=0)]
     [ValidateNotNullorEmpty()]
     [String]$Message
)

$momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)

}



function Add-SCOMAlertFields
{
# AssignmentGroup specified & TicketID is freeform string field
$AssignmentGroup
$TicketID

if ( $Alerts.ResolutionState -ne 255 )
	{
	$TicketID = "SNOW_event"
	write-host "ServiceNow (SNow) Event created with $TicketID"
	Add-Event "ServiceNow (SNow) Event created with $TicketID"
	}
if ( $Alerts.ResolutionState -eq 255 )
	{
	$TicketID = "NO_SNOW_event"
	write-host "ServiceNow (SNow) Event created with $TicketID"
	Add-Event "ServiceNow (SNow) Event created with $TicketID"
	exit $0
	}


# Get-SCOM alert and update alert
#================================
Get-SCOMAlert -Name "$AlertName" -ResolutionState 0 | Set-SCOMAlert -ticketID $TicketID `
	-Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-Comment "ServiceNow SCOM event automation - Set Ticket, Owner, Resolution state in current alert"

write-host "Updated SCOM alert $TicketID for group $AssignmentGroup"
Add-Event "Updated SCOM alert $TicketID for group $AssignmentGroup"


<# Resolve alert?
#===========================
Get-SCOMAlert -Name "$AlertName" -ResolutionState 0 | Resolve-SCOMAlert -ticketID $TicketID `
	-Owner "$AssignmentGroup" `
	-Comment "Resolve ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert"

	Add-Event "Resolved SCOM alert $TicketID for group $AssignmentGroup"
#>

}



function Get-SNowParameters
{
<#
.Synopsis
   Get ServiceNow (SNow) Event parameters
.DESCRIPTION
   This function is used to gather and validate SNow parameters.

   Get-SNowParameters function validates multiple required parameters to create a populated ServiceNow Event.
	Parameters include ServiceNow URL (prod/test), ServiceNow user/pass (leveraging Credential Manager)

   ServiceNow specific fields that are required for RESTAPI Event creation include:
	CallerID, AssignmentGroup, AlertName, into the REST payload variable $EventData.  The $EventData array is tested, and then converted to JSON payload for invoke-RestMethod injection.

.EXAMPLE
   Example of how to use this cmdlet
   
   Get-SNowParameters

.EXAMPLE
   Another example of how to use this cmdlet
   
.INPUTS
   No Inputs
.OUTPUTS
   Function will output various validation messages for ServiceNow events
   
   TEST ServiceNow URL specified
   PROD ServiceNow URL specified
   NO ServiceNow URL specified
   
   CredentialManager PoSH Module NOT Installed
   
   ServiceNow Credential NOT stored on server
   ServiceNow User NOT stored on server
   ServiceNow Password NOT stored on server

   ServiceNow (SNow) Event NOT needed as SCOM alert is CLOSED
   
   Hostname = $HostName
   
   Invalid or null parameters passed for impact,Urgency,Priority in SNow Event
.NOTES
   Leverages parameter provided variables to create ServiceNow (SNow) Event
.COMPONENT
   Get-SNowParameters function is contained in the New-SNowEvent.ps1 script
.ROLE
   Get function of New-SNowEvent.ps1 script
.FUNCTIONALITY
   Function get's or validates required fields to create ServiceNow (SNow) Event
#>



<#
# Set up ServiceNow URL connection pieces
#===============================================
#>

#write-host $ServiceNowURL
Add-Event "ServiceNow (SNow) URL specified = $($ServiceNowURL)"


if ( ( $ServiceNowURL | where { $_ -like "*test*" } ) )
	{
	write-host "TEST ServiceNow URL specified"
	Add-Event "TEST ServiceNow URL specified"
	}

if ( ( $ServiceNowURL | where { $_ -notlike "*test*" } ) )
	{
	write-host "PROD ServiceNow URL specified"
	Add-Event "PROD ServiceNow URL specified"
	}

if ( $ServiceNowURL -eq $null )
	{
	write-host "NO ServiceNow URL specified"
	Add-Event "NO ServiceNow URL specified"
	exit $0
	}


<#
# Pre-req for CredentialManager powershell (posh) module
# Assume module NOT loaded into current PowerShell profile

Import-Module -Name CredentialManager
#>

# Verify Credential Manager snap in installed
$CredMgrModuleBase = Get-Module -Name CredentialManager

if ( $CredMgrModuleBase.ModuleBase -ne $Null )
	{
	write-host "CredentialManager PoSH Module Installed, ModuleBase = $CredMgrModuleBase.ModuleBase"
	Add-Event "ServiceNow Credential PowerShell module installed, ModuleBase = $CredMgrModuleBase.ModuleBase"
	}

if ( $CredMgrModuleBase.ModuleBase -eq $Null )
	{
	write-host "CredentialManager PoSH Module NOT Installed, ModuleBase = $CredMgrModuleBase.ModuleBase"
	Add-Event "ServiceNow Credential PowerShell module NOT installed, ModuleBase = $CredMgrModuleBase.ModuleBase"
	exit $0
	}


# Retrieve SNOW credential from Credential Manager
#===============================================
# Example
# $Credential = Get-StoredCredential -Target "SNOW_Account"
# $Credential = Get-StoredCredential -Target "ServiceNowCredential"
# ID, Password, and Caller_ID are provided by ##CUSTOMER## team

$Credential = Get-StoredCredential -Target "ServiceNowCredential"
$ServiceNowUser = $Credential.Username
$ServiceNowPassword = $Credential.GetNetworkCredential().Password


# Test Credential variables for User password are provided
#===============================================
if ( $ServiceNowUser -eq $null )
	{
	write-host "ServiceNow User NOT stored on server"
	Add-Event "ServiceNow User NOT stored on server"
	}
if ( $ServiceNowPassword -eq $null )
	{
	write-host "ServiceNow Password NOT stored on server"
	Add-Event "ServiceNow Password NOT stored on server"
	}


<#
# Alert may not have server listed as offending object with issue
# $AlertID parameter passed into script to then audit alert, and find where alert originated

# Other test scenarios
Gather Critical, New alerts

Example of new, critical alerts
$Alerts = get-scomalert -ResolutionState 0 -severity 2

Example of new, warning alerts
$Alerts = get-scomalert -ResolutionState 0 -severity 1

Example of alert with resolution LT 255
$Alerts = get-scomalert -Name $AlertName -ResolutionState (0..254)

Example of alert specified from input variables
$Alerts = get-scomalert -Name $AlertName -ResolutionState (0..254)

Example of alertID specified with where clause
$Alerts = Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq $AlertName ) -AND ( $_.ResolutionState -ne 255 ) }
#>

#Assuming No changes, inputs passed to SCOM channel for SNOW event creation
$Alerts = Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq $AlertName ) -AND ( $_.ResolutionState -ne 255 ) }


# Evaluate alert closed before SCOM channel SNOW script executed
#===============================================
if ( $Alerts.ResolutionState -eq 255 )
	{
	write-host "ServiceNow (SNow) Event NOT needed as SCOM alert is CLOSED"
	Add-Event "ServiceNow (SNow) Event NOT needed as SCOM alert is CLOSED"
	exit $0
	}


# Set Alert Parameters, ID, ManagementGroup, Category from SCOM alert to use in event fields
$AlertParameters = $Alerts[0].Parameters

# Used with manual testing $Alerts
$AlertID = $Alerts[0].ID.Guid

$AlertManagementGroup =  $Alerts[0].ManagementGroup.Name
$AlertCategory = $Alerts[0].Category


# Alert description
$AlertDescription = $Alerts[0].Description

# Display AlertDescription (Debug)
$AlertDescription

# Determine SCOM Alert Description excludex JSON special characters
#===============================================
$Description = $Alerts.Description

$Description = $Description -replace "^@", ""
$Description = $Description -replace "=", ":"
$Description = $Description -replace ";", ","
$Description = $Description -replace "`", "#"
$Description = $Description -replace "{", "*"
$Description = $Description -replace "}", "*"
$Description = $Description -replace "\n", "\\n"
$Description = $Description -replace "\r", "\\r"
$Description = $Description -replace "\\\\n", "\\n"

#write-host $Description
#Add-Event $Description


# Multiple locations for HostName in alerts based on class, path, and other variables
# Hostname
# Figure out hostname based on alert values
#===============================================

$MonitoringObjectPath = ($Alerts |select MonitoringObjectPath).MonitoringObjectPath
$MonitoringObjectDisplayName = ($Alerts[0] |select MonitoringObjectDisplayName).MonitoringObjectDisplayName	
$PrincipalName = ($Alerts |select PrincipalName).PrincipalName
$DisplayName = ($Alerts |select DisplayName).DisplayName
$PKICertPath = ($Alerts |select Path).Path

# Update tests with else for PowerShell script
if ( $MonitoringObjectPath -ne $null ) { $Hostname = $MonitoringObjectPath }
if ( $MonitoringObjectDisplayName -ne $null ) { $Hostname = $MonitoringObjectDisplayName }
if ( $PrincipalName -ne $null ) { $Hostname = $PrincipalName }
if ( $DisplayName -ne $null ) { $Hostname = $DisplayName }
if ( $PKICertPath -ne $null ) { $Hostname = $PKICertPath }

# Verify unique Hostname
if (( $Hostname | measure).Count -gt 1 )
	{
	$Hostname = $Hostname | sort -uniq
	write-host $Hostname
	}

$IP = Resolve-DNSName -Name $hostname -Type A
$ServerIP = ($IP.IPAddress)

# Remove FQDN, leaving servername
$ParseHost = $Hostname.Split(".")
$Hostname = $Parsehost[0]

<# 
Debug
$ServerIP
$MonitoringObjectPath
$MonitoringObjectDisplayName
$PrincipalName
$DisplayName
$PKICertPath

#>

# Debug showing hostname for Event
Add-Event "Hostname = $HostName"
write-host $Hostname


# Determine SCOM Alert Severity to ITSM tool impact
#===============================================
$Severity = $Alerts[0].Severity
 
If ( $Severity -eq "Warning" )
	{
	$EventSeverity = "Minor"
	}
If ( $Severity -eq "Critical" )
	{
	$EventSeverity = "Critical"
	}

# Debug Severity
write-host $Severity
Add-Event $Severity

}



function New-SNowEvent
{
<#
.Synopsis
   Create new ServiceNow (SNow) Event
.DESCRIPTION
   New-SNowEvent function will create SNOW events using passed alert data (SCOM source).

   New-SNowEvent function follows the Get-SNowEvent function to create a populated ServiceNow Event.

   ServiceNow specific fields that are required for RESTAPI Event creation include: 
	CallerID, AssignmentGroup, Business_Service, Category, SubCategory, AlertName, Priority, Impact, and Severity into the REST payload variable $EventData.  The $EventData array is tested, and then converted to JSON payload for invoke-RestMethod injection.

.EXAMPLE
   Example of how to use this cmdlet
   
   New-ServiceEvent   
.INPUTS
   No inputs for this function.
.OUTPUTS
   Function will output various validation messages for ServiceNow events
   
   ServiceNow Event payload `n `n $($EventData)
   
   Completed ServiceNow Event creation for ($date) `n $EventData"
   
   Attempting to create Event for $AlertName on $HostName...
   
   Event created successfully. Event Number: $($response.result.number)
   
   Failed to create Event. Error: $($response.result.number)
   
   Error: $errorMessage in REST Response
   
   Updated SCOM alert $TicketID for group $AssignmentGroup
   
   Completed ServiceNow Event creation for ($date) `n $EventData
   
   Script Completed. `n Script Runtime: ($ScriptTime) seconds.
.NOTES
   Generates events tracking pieces, and results from Event creation.
   Final piece is SCOM alert updated with Event #, Assignment Group.
.COMPONENT
   New-SNowEvent function belongs to New-SNowEvent.ps1
.ROLE
   The New-SNowEvent function creates NEW ServiceNow (SNow) events.
.FUNCTIONALITY
   Create new ServiceNow (SNow) event, and update SCOM alert with event number, AssignmentGroup
#>

#write-host "Create ServiceNow event for ($date)."
#Add-Event "Create ServiceNow event for ($date)."

<#
# Set up ServiceNow connection pieces from top level parameters
#===============================================

 	[String]$AlertName,
	[String]$AlertID,
	[String]$Impact,
	[String]$Urgency,
	[String]$Priority,
	[String]$AssignmentGroup,
	[String]$BusinessGroup,
	[String]$Category,
	[String]$SubCategory
#>


# Set up EventData variable with SCOM to SNow fields
#===============================================

# Test Credential variables for User password are provided
#===============================================
if ( $Credential -eq $null )
	{
	write-host "ServiceNow Credential NOT stored on server"
	Add-Event "ServiceNow Credential NOT stored on server"
	exit $0
	}


# Build SNow event payload
#=======================================================
#$info = @{Test="12345"} | ConvertTo-Json
$info = "{ SCOM Alert Parameters: $AlertParameters , SCOM Alert ID: $AlertID }" | ConvertTo-Json

#write-host "ServiceNow Event payload `n `n $($info)"
Add-Event "ServiceNow Event payload `n `n $($info)"

# SNOW provided event body example
#$body = @{source="SCOM";node="$Hostname";type="Alert";severity="$EventSeverity";Description="This is a SCOM alert -  $Description";resource="NIC2";metric_name="Adapter2";additional_info=$info;event_class="SCOM from $ServerIP";message_key="SCOM-$Hostname-Alert"} | ConvertTo-Json

$body = @{source="SCOM";node="$Hostname";type="Alert";severity="$EventSeverity";Description="{ $AlertManagementGroup - SCOM alert - $Description}";resource="$AlertCategory";metric_name="$AlertName";additional_info=$info;event_class="SCOM from $ServerIP";message_key="SCOM-$Hostname-Alert"} | ConvertTo-Json

#write-host "ServiceNow Event payload `n `n $($body)"
Add-Event "ServiceNow Event payload `n `n $($body)"

<#
# Debug verify Info and Body variables
$info
$body
#>

<#
#=======================================================
# Craft auth and headers for RESTAPI
#>

$base64Auth=[System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($ServiceNowUser):$($ServiceNowPassword)"))
$headers = @{"Content-Type"="application/json"; "Authorization"="Basic $base64Auth"}

#Include $proxy if specified

if ( $Proxy -eq $Null )
	{
	$result=Invoke-RestMethod -method "POST" -uri $ServiceNowURL -headers $headers -body $body;$result | ConvertTo-Json;
	}
if ( $Proxy -ne $Null )
	{
	$result=Invoke-RestMethod -method "POST" -uri $ServiceNowURL -headers $headers -body $body;$result -Proxy $Proxy | ConvertTo-Json;
	}


<#

# Debug of invoke 
$result | fl -property *

# Full object output debug
#$result

#>


<#
#===============================================================
# Gather Event values to update SCOM alert
#===============================================================

# Debug
# Additional scripting to pull event number,assignment_group fields into SCOM alert
# $TicketID = $Response.result.Number  # Unable to confirm value from array

# Set SCOM alert with SNOW Event number, and assignment group
# Post $Result > set TicketID, Owner
#===============================================================

#>
$ResolutionState = 249

Add-SCOMAlertFields

#============================================================
$Result = "GOOD"
$Message = "Completed ServiceNow Event creation for ($date)"

#Write-Host "Completed ServiceNow Event creation for ($date) `n $EventData"
Add-Event "Completed ServiceNow Event creation for ($date) `n $EventData"

<#
$bag.AddValue('Result',$Result)
$bag.AddValue('Count',$Test)
$bag.AddValue('Message',$Message)
$bag.AddValue('Summary',$DNSMessage)
#>

# Return all bags
$bag
#=================================================================================
# End MAIN script section
 
}

# Execute functions
Get-SNowParameters
New-SNowEvent


  
# End of script section
#=================================================================================
#Log an event for script ending and total execution time.
$EndTime = Get-Date
$ScriptTime = ($EndTime - $StartTime).TotalSeconds
Add-Event "Script Completed. `n Script Runtime: ($ScriptTime) seconds."
#=================================================================================
# End of script





