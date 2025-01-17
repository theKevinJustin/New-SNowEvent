<#
#=================================================================================
# Script test to setup SNOW events using SCOM alert data
#
#  Authors: Steven Brown, Kevin Justin, Joe Kelly
#=================================================================================

#=================================================================================
# Script change history
#
# v1.0.0.9	19 Dec 2024 - KWJ - Includes ServiceNow (SNow) maintenance mode test and TicketID update
# v1.0.0.8	 4 Dec 2024 - KWJ - Updated logic to use SCOM CustomField1 for ServiceNow (SNow) URL to allow users to click on incident and open in system browser
# v1.0.0.7  22 Nov 2024 - KWJ - Updated $result logic for URL handling and proper URL parsing
# v1.0.0.6   8 Nov 2024 - KWJ - Added logic to use SCOM CustomField1 for ServiceNow (SNow) URL to allow users to click on incident and open in system browser
# v1.0.0.5   9 Oct 2024 - KWJ - Added SnowEnv parameter for test/prod capability
# v1.0.0.4  19 Sep 2024 - KWJ - Updated ServiceNow ID/password encoding checks after ServiceNow password change
# v1.0.0.3  30 Jul 2024 - KWJ - Updated SCOM Alert Description and Parameters parsing logic
# v1.0.0.2  29 Jun 2024 - KWJ - Added SCOM Alert Description and Parameters parsing logic
# v1.0.0.0  22 Sep 2023 - KWJ - First version based off Steven Brown template

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
	$SNOWAlertName = "<Team> <ORG> SCOM Event - $AlertName"
Example SNOWAlertName
	$SNOWAlertName = "<Team> <ORG> SCOM $AlertName"
Example SNOWAlertName
	$SNOWAlertName = "##CUSTOMER## ##TEAM## SCOM Event - $AlertName"
Example SNOWAlertName
	$SNOWAlertName = "##TEAM## ##CUSTOMER##: SCOM - $AlertName"

# Replace these variables with valid values:
##CUSTOMER##
##TEAM##
##SERVICENOWURL##
##ServiceNowTestURL##
##ServiceNowProdURL##

# Change relevant ServiceNow SNow accounts for test and production
ServiceNowProdCredential
ServiceNowCredential


# ServiceNow URL examples
Example ##ServiceNowURL##
##ServiceNowURL##/api/now/table/em_event"

# If required, add proxy URL for REST injection
##Proxy##


Hard coded variables for ServiceNow (SNow) for CallerId, URL, and if Proxy needed

Hard code URL, Proxy, and CallerID, ##CUSTOMER## into ServiceNow (SNow) caller_id field to create events
#===============================================

# Don't forget to replace the following variables!

SNOW DEV URL with Prod before going to production events
$ServiceNowURL="https://##ServiceNowURL##/api/now/table/em_event"
# Test
$ServiceNowURL="https://##ServiceNowTestURL##/api/now/table/em_event"
# Prod
$ServiceNowURL="https://##ServiceNowProdURL##/api/now/table/em_event"
##ServiceNowURL##/api/now/table/em_event"

# Set AlertName for Testing
$SNOWAlertName = "##TEAM## ##CUSTOMER##: SCOM - $AlertName"

Proxy
$Proxy = ##PROXY##

# AssignmentGroup & TicketID
$AssignmentGroup
$TicketID = "SNOW_event"

#>


Param (
     [Parameter(
         Mandatory=$true,
         ValueFromPipeline=$true,
         Position=0)]	 
     [ValidateNotNullorEmpty()]
     [String]$SNowENV,
     [Parameter(
         Mandatory=$true,
         ValueFromPipeline=$true,
         Position=1)]
     [ValidateNotNullorEmpty()]
     [String]$AlertName,
     [Parameter(
		 Mandatory=$true,
         ValueFromPipeline=$true,
		 Position=2)]
		 [ValidateNotNullorEmpty()]
     [String]$AlertID,
     [Parameter(
	 	 Mandatory=$true,
         ValueFromPipeline=$true,
	 	 Position=3)]
		 [ValidateNotNullorEmpty()]
		 [String]$AssignmentGroup,
     [Parameter(
	 	 Mandatory=$true,
         ValueFromPipeline=$true,
	 	 Position=4)]
		 [ValidateNotNullorEmpty()]
		 [String]$Team
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

if ( $SNowENV -eq "Test" ) { $EventID = "711" }
if ( $SNowENV -eq "Prod" ) { $EventID = "712" }

# Create new object for MOMScript API, or SCOM alert properties
$momapi = New-Object -comObject MOM.ScriptAPI

# Begin logging script starting into event log
# write-host "Script is starting. `n Running as ($whoami)."
$momapi.LogScriptEvent($ScriptName,$EventID,0,"New-SNowEvent Script is starting. `n Running as ($whoami).")
#=================================================================================

# PropertyBag Script section - Monitoring scripts get this
#=================================================================================
# Load SCOM PropertyBag function
$bag = $momapi.CreatePropertyBag()

$date = get-date -uFormat "%Y-%m-%d"


# Hard code URL, Proxy, CallerID SNOW variables into script
#===============================================

# Don't forget to replace SNOW DEV URL with Prod before going to production events
# $global: myVariable
# Assume module NOT loaded into current PowerShell profile

Import-Module -Name CredentialManager
write-host ""
$GetCM = Get-Module CredentialManager
#$momapi.LogScriptEvent($ScriptName,$EventID,0,"CredentialManager module NOT loaded")
#write-host -foreground red "CredentialManager module NOT loaded"

write-host "$($GetCM)"

if ( $null -ne $GetCM )
	{	
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"CredentialManager module IS loaded")
	write-host ""
	write-host -foreground green "CredentialManager module IS loaded"
	write-host ""
	}

if ( $null -eq $GetCM )
	{
	try
		{
		$GetCM = Get-Module CredentialManager
		#Install-Module CredentialManager -force -Verbose -Scope CurrentUser
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Try statement - CredentialManager module NOT loaded")
		write-host -foreground red "Try statement = CredentialManager module NOT loaded)"
		}
	catch
		{
        $errorMessage = $_.Exception.Message

		$errormsg = $_.ToString()
		$exception = $_.Exception
		$stacktrace = $_.ScriptStackTrace
        $failingline = $_.InvocationInfo.Line
        $positionmsg = $_.InvocationInfo.PositionMessage
        $pscommandpath = $_.InvocationInfo.PSCommandPath
        $failinglinenumber = $_.InvocationInfo.ScriptLineNumber
		write-host -foreground red "Errormsg $errormsg `n Exception $exception `n Scriptname $scriptname `n Failinglinenumber $failinglinenumber `n Failingline $failingline `n PSCommandPath $pscommandpath `n Positionmsg $pscommandpath `n Stacktrace $stacktrace"
		write-host ""
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Errormsg $errormsg `n Exception $exception `n Scriptname $scriptname `n Failinglinenumber $failinglinenumber `n Failingline $failingline `n PSCommandPath $pscommandpath `n Positionmsg $pscommandpath `n Stacktrace $stacktrace")

        if ($_.Exception.InnerException)
        	{
			$errorMessage = $_.Exception.InnerException.Message
	        }

		# Test Incident returned
		if ( $GetCM -eq $null ) { write-host "Null catch response"}
		
		Write-Host "Error: REST Response Error Message: $errorMessage"
		Write-Host ""
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Error: REST Response Error Message: $errorMessage")
		}
	}


# Time to test Prod vs. Test parameter for single script to multiple ServiceNow (SNow) environments
if ( $SNowENV -eq "Test" )
	{
	$SNowHostName = "##ServiceNowTestURL##"
	# Test
	$ServiceNowURL = "https://$SNowHostName/api/now/table/em_event"
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNowURL = ($ServiceNowURL)")
	write-host "ServiceNowURL = $($ServiceNowURL)"
	}

if ( $SNowENV -eq "Prod" )
	{
	$SNowHostName = "##ServiceNowProdURL##"
	# Prod
	$ServiceNowURL = "https://$SNowHostName/api/now/table/em_event"
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNowURL = ($ServiceNowURL)")
	write-host "ServiceNowURL = $($ServiceNowURL)"
	}

$Proxy= $null # ##PROXY##
#$momapi.LogScriptEvent($ScriptName,$EventID,0," Proxy = ($Proxy)")
#write-host "Proxy = $($Proxy)"

<#
# Retrieve SNOW credential from Credential Manager
#===============================================
# Example
# $Credential = Get-StoredCredential -Target "SNOW_Account"
#
# ID, Password, and Caller_ID are provided by ServiceNow team
#
# Get-StoredCredential -Target "ServiceNowCredential"
# Example
# $Credential = Get-StoredCredential -Target "SNOW_Account"
# $Credential = Get-StoredCredential -Target "ServiceNowCredential"
# $Credential = Get-StoredCredential -Target "svc_rest_scom"
# ID, Password, and Caller_ID are provided by ##CUSTOMER## team
#>


# Credential check
if ( $SNowENV = "Prod" )
	{
	$Credential = Get-StoredCredential -Target "ServiceNowProdCredential" -Verbose
	}
if ( $SNowENV = "Test" )
	{
	$Credential = Get-StoredCredential -Target "ServiceNowCredential" -Verbose
	}
	
$momapi.LogScriptEvent($ScriptName,$EventID,0,"Stored ServiceNowCredential user = $($Credential.UserName)")
write-host "ServiceNowCredential user = $($Credential.UserName)"

# Test Credential variables for User password are provided
#===============================================

$ServiceNowUser = $Credential.Username
$ServiceNowPassword = $Credential.GetNetworkCredential().Password

write-host -f green "Stored Credential variable ServiceNowUser = $($ServiceNowUser)"
write-host -f green "Stored Credential variable ServiceNowPassword = $($ServiceNowPassword)"


if ( $Null -eq $ServiceNowUser )
	{
	write-host -f red "ServiceNow User NOT stored on server"
	Add-Event "ServiceNow User NOT stored on server"

	$EndTime = Get-Date
	$ScriptTime = ($EndTime - $StartTime).TotalSeconds
	write-host -f red "Script Completed. `n Script Runtime: ($ScriptTime) seconds."
	Add-Event "Script Completed. `n Script Runtime: ($ScriptTime) seconds."

	exit $0
	}
if ( $Null -eq $ServiceNowPassword )
	{
	write-host -f red "ServiceNow Password NOT stored on server"
	Add-Event "ServiceNow Password NOT stored on server"

	$EndTime = Get-Date
	$ScriptTime = ($EndTime - $StartTime).TotalSeconds
	write-host -f red "Script Completed. `n Script Runtime: ($ScriptTime) seconds."
	Add-Event "Script Completed. `n Script Runtime: ($ScriptTime) seconds."

	exit $0
	}


if ( $null -eq $Credential )
	{
	write-host -f red "ServiceNow Credential NOT stored on server"
	write-host ""
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNow Credential NOT stored on server")

	try
		{
		$Credential = Get-StoredCredential -Target "ServiceNowCredential"
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Try statement - Stored ServiceNowCredential")
		write-host "Try statement - ServiceNowCredential"

		$ServiceNowUser = $Credential.Username
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Try statement - ServiceNowUser = $($ServiceNowUser)")

		$ServiceNowPassword = $Credential.GetNetworkCredential().Password
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Try statement - ServiceNowPassword = $($ServiceNowPassword)")
		}
	catch
		{
        $errorMessage = $_.Exception.Message

		$errormsg = $_.ToString()
		$exception = $_.Exception
		$stacktrace = $_.ScriptStackTrace
		$failingline = $_.InvocationInfo.Line
		$positionmsg = $_.InvocationInfo.PositionMessage
		$pscommandpath = $_.InvocationInfo.PSCommandPath
		$failinglinenumber = $_.InvocationInfo.ScriptLineNumber
		write-host -foreground red "Errormsg $errormsg `n Exception $exception `n Scriptname $scriptname `n Failinglinenumber $failinglinenumber `n Failingline $failingline `n PSCommandPath $pscommandpath `n Positionmsg $pscommandpath `n Stacktrace $stacktrace"
		write-host ""
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Errormsg $errormsg `n Exception $exception `n Scriptname $scriptname `n Failinglinenumber $failinglinenumber `n Failingline $failingline `n PSCommandPath $pscommandpath `n Positionmsg $pscommandpath `n Stacktrace $stacktrace")

			
        if ($_.Exception.InnerException)
        	{
			$errorMessage = $_.Exception.InnerException.Message
	        }

		# Test Incident returned
		if ( $Credential -eq $null ) { write-host "Null catch response"}
		
		Write-Host "Error: REST result variable response Error Message: $errorMessage"
		Write-Host ""
		$momapi.LogScriptEvent($ScriptName,$EventID,0,"Error: REST result variable response Error Message: $errorMessage")
		}
	}



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
      New-SNowEvent -AlertName <> -AlertID <> -AssignmentGroup <> -Team <>
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







<#
# Credential Manager components complete, SCOM alert inputs passed to SCOM channel for SNOW event creation

# Write AlertName and AlertID variables

# NOTE * Uncomment as needed for commented (#) debug write-host lines as needed
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
$Alert = Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq $AlertName ) -AND ( $_.ResolutionState -ne 255 ) }

# Evaluate alert closed before SCOM channel SNOW script executed
#===============================================

# * Uncomment as needed for debug write-host lines
#>

#write-host -f green "SCOM Alert alertName = $AlertName, Alert ID = $AlertID"
#$momapi.LogScriptEvent($ScriptName,$EventID,0,"SCOM Alert alertName = $AlertName, Alert ID = $AlertID")

#Assuming No changes, inputs passed to SCOM channel for SNOW event creation
$Alert = Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq $AlertName ) } #-AND ( $_.ResolutionState -ne 255 ) }

#write-host -f green "SCOM Alert ready for parsing"

$momapi.LogScriptEvent($ScriptName,$EventID,0,"Parse out input variables. `n AlertName = $($AlertName) `n AlertID = $($AlertID) `n AssignmentGroup = $($AssignmentGroup) `n Team = $($Team)")

$AlertID = $AlertID.Replace('{','').Replace('}','')

$momapi.LogScriptEvent($ScriptName,$EventID,0,"Parsed out braces from AlertID with replace = $($AlertID)")

$momapi.LogScriptEvent($ScriptName,$EventID,0,"End Global section")
# write-host -f green "End Global section"
write-host ""

# End Global section



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

<#
#===============================================================
# Gather values to update SCOM alert fields after processing incident RESTAPI
#===============================================================

# Function created to verify IF SNOW Event created became an Alert, and INC (incident)
# If true, change resolution state of alert to acknowledged (249)
# If AssignmentGroup specified, change to 248 'Assigned to Engineering'

* Uncomment as needed for debug write-host lines

# write-host -f green "Begin Add-SCOMAlertFields function"
# add-event "Begin Add-SCOMAlertFields function"

# AssignmentGroup specified & TicketID is freeform string field
write-host -f green "Assignment Group = $AssignmentGroup"
write-host -f green "TicketID = $TicketID"
write-host -f green "SCOM Alert ResolutionState = $AlertResolutionState"


# Resolve alert?
#===========================
Get-SCOMAlert -Name "$AlertName" -ResolutionState 0 | Resolve-SCOMAlert -ticketID $TicketID `
	-Owner "$AssignmentGroup" `
	-Comment "Resolve ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert"

	Add-Event "Resolved SCOM alert $TicketID for group $AssignmentGroup"
#>


# Change resolution state of alert to acknowledged (249)
$AlertResolutionState = 249
$ResolutionState = 249

if ( $AlertResolutionState -ne 255 )
	{
	write-host ""
	write-host -f green "ServiceNow (SNow) Event created with TicketID = $TicketID"
	Add-Event "ServiceNow (SNow) Event created with TicketID = $TicketID"
	}
if ( $AlertResolutionState -eq 255 )
	{
	#$TicketID = "NO_SNOW_event"
	write-host ""
	write-host -f yellow "SCOM Alert in closed state (ResolutionState = 255) reference TicketID = $TicketID"
	Add-Event "SCOM Alert in closed state (ResolutionState = 255) reference TicketID = $TicketID"
	#write-host -f red "Exiting - ServiceNow (SNow) Event NOT created as SCOM alert closed with TicketID = $TicketID"
	#Add-Event "Exiting - ServiceNow (SNow) Incident NOT created as SCOM alert closed with TicketID = $TicketID"
	#exit $0

	# Update ticket, assignment group, resolution state -AND ( $_.ResolutionState -eq 0 ) } `
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-CustomField1 "NO_SNOW_event" `
	-Comment "ServiceNow SCOM alert automation - SCOM alert closed - Set Ticket, Owner, Resolution state in current alert. Updated SCOM alert $AlertName, Server = $Hosts, with TicketID $TicketID for group $AssignmentGroup"

	write-host ""
	write-host -f yellow "SCOM alert closed - Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = $TicketID, for group $AssignmentGroup"
	Add-Event "SCOM alert closed - Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"

	}




<#
#=======================================================
# Get-SCOM alert and update alert
# Did event created become an alert?
#================================
#>

# Event section
if ( $Null -ne $($EventSysID) )
	{
	$TicketID = $EventSysID
	# Debug
	# $TicketID
 
	# Update ticket, assignment group, resolution state -AND ( $_.ResolutionState -eq 0 ) } `
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-CustomField1 "$URL" `
	-Comment "ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert.  Updated SCOM alert $AlertName, Server = $Hosts, with TicketID $TicketID for group $AssignmentGroup"

	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = $TicketID, for group $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"
	}
 
if ( $null -eq $($EventSysID) )
	{
	$TicketID = "NO SNOW event created after 5 plus attempts over 30 seconds"
 
	# Update ticket, assignment group, resolution state -AND ( $_.ResolutionState -eq 0 ) } `
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-CustomField1 "NO_SNOW_event" `
	-Comment "ServiceNow SCOM alert automation - Set SCOM Alert Ticket, Owner, Resolution state in current alert.  Updated SCOM alert $AlertName, Server = $Hosts, with 'NO SNOW Alert' for TicketID $TicketID, for group $AssignmentGroup"

	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"

	}


# Alert section
if ( $Null -ne $($AlertCreated.result.number) )
	{
	$TicketID = $AlertCreated.result.number
	# Debug
	# $TicketID
 
	# Update ticket, assignment group, resolution state -AND ( $_.ResolutionState -eq 0 ) } `
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-CustomField1 "$AlertPublicSysURL" `
	-Comment "ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert. 	Updated SCOM alert $AlertName, Server = $Hosts, with TicketID $TicketID for group $AssignmentGroup"

	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = $TicketID, for group $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"
	}
 
if ( $null -eq $($AlertCreated.result.number) )
	{
	$TicketID = "NO SNOW alert created after 5 plus attempts over 30 seconds"
 
	# Update ticket, assignment group, resolution state -AND ( $_.ResolutionState -eq 0 ) } `
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-CustomField1 "NO_SNOW_alert_created" `
	-Comment "ServiceNow SCOM alert automation - Set SCOM Alert Ticket, Owner, Resolution state in current alert. Updated SCOM alert $AlertName, Server = $Hosts, with 'NO SNOW Alert' for TicketID $TicketID, for group $AssignmentGroup"

	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"

	}


# Incident section
if ( $null -ne $($IncidentCreated.result.number) )
	{
	$TicketID = $IncidentCreated.result.number

	# Update ticket, assignment group, resolution state
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$TEAM $AssignmentGroup"  -ResolutionState $ResolutionState `
	-CustomField1 "$IncidentPublicSysURL" `
	-Comment "ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert. Updated SCOM alert $AlertName, Server = $Hosts, for TicketID $TicketID, for group $AssignmentGroup"
	# ResolutionState 248 = Assigned to Engineering
	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 248 Assigned to Engineering, SNOW Incident $TicketID, for group $TEAM $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 248 Assigned to Engineering, SNOW Incident $TicketID, for group $TEAM $AssignmentGroup"
	}
if ( $null -eq $($IncidentCreated.result.number) )
	{
	$TicketID = "NO $SNowENV SNOW Incident created"
	write-host -f red "NO SNOW incident created after 20+ plus attempts over 3 minutes"
	Add-Event "NO SNOW Incident created after 20+ plus attempts over 3 minutes"

	if ( $AlertCreated.result.maintenance -eq "true" )
		{
		write-host "ServiceNow $SNowENV MaintenanceMode Enabled"
		$TicketID = "ServiceNow $SNowENV MaintenanceMode Enabled"
		Add-Event "ServiceNow MaintenanceMode Enabled, NO SNOW Incident created"
		}

	# Update ticket, assignment group, resolution state
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$TEAM $AssignmentGroup"  -ResolutionState $ResolutionState `
	-CustomField1 "NO_SNOW_incident_created" `
	-Comment "ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert.  Updated SCOM alert $AlertName, Server = $Hosts, with 'NO SNOW Incident' for TicketID $TicketID, for group $AssignmentGroup"
	# ResolutionState 248 = Assigned to Engineering
	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 248 Assigned to Engineering, TicketID = NO SNOW alert, for group $TEAM $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 248 Assigned to Engineering, SNOW Incident $TicketID, for group $TEAM $AssignmentGroup"
	}
#=======================================================

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

#write-host -f green $ServiceNowURL
#write-host -f green "ServiceNow (SNow) URL specified = $($ServiceNowURL)"
#Add-Event "ServiceNow (SNow) URL specified = $($ServiceNowURL)"

#>


if ( ( $ServiceNowURL | where { $_ -like "*test*" } ) )
	{
	write-host -f green "TEST ServiceNow URL specified"
	Add-Event "TEST ServiceNow URL specified"
	}

if ( ( $ServiceNowURL | where { $_ -notlike "*test*" } ) )
	{
	write-host -f green "PROD ServiceNow URL specified"
	Add-Event "PROD ServiceNow URL specified"
	}

if ( $Null -eq $ServiceNowURL )
	{
	write-host -f red "Exiting - NO ServiceNow URL specified"
	write-host ""
	Add-Event "Exiting - NO ServiceNow URL specified"

	#Log an event for script ending and total execution time.
	$EndTime = Get-Date
	$ScriptTime = ($EndTime - $StartTime).TotalSeconds
	write-host -f red "Script Completed. `n Script Runtime: ($ScriptTime) seconds."
	Add-Event "Script Completed. `n Script Runtime: ($ScriptTime) seconds."

	exit $0
	}


<#
# Pre-req for CredentialManager powershell (posh) module
# Assume module NOT loaded into current PowerShell profile

* Uncomment as needed for debug write-host lines

# Verify SNOW credential exists in Credential Manager
#===============================================
# Example
# $Credential = Get-StoredCredential -Target "SNOW_Account"
#
# ID, Password, and Caller_ID are provided by ServiceNow team

# From Global section
# $Credential = Get-StoredCredential -Target "ServiceNowCredential"
# $ServiceNowUser = $Credential.Username
# $ServiceNowPassword = $Credential.GetNetworkCredential().Password
#>

# Verify Credential Manager snap in installed
$CredMgrModuleBase = Get-Module -Name CredentialManager

if ( $Null -ne $CredMgrModuleBase.ModuleBase )
	{
	write-host -f yellow "CredentialManager PoSH Module Installed,`nModuleBase = $($CredMgrModuleBase.ModuleBase)"
	Add-Event "ServiceNow Credential PowerShell module installed,`nModuleBase = $($CredMgrModuleBase.ModuleBase)"

	if ( $null -ne $Credential )
		{
		Write-host ""
		Write-host -f green "CredentialManager PoSH Module Installed, Stored Credential variable exists"
		Write-host ""
		Add-Event "CredentialManager PoSH Module Installed,`nStored Credential variable exists"
		}
	#write-host -f green "CredentialManager PoSH Module Installed"
	#Add-Event "CredentialManager PoSH Module Installed"
	}
else
	{
	write-host -f red "CredentialManager PoSH Module NOT Installed"
	write-host ""
	Add-Event "ServiceNow Credential PowerShell module NOT installed"

	#Log an event for script ending and total execution time.
	$EndTime = Get-Date
	$ScriptTime = ($EndTime - $StartTime).TotalSeconds
	write-host -f red "Script Completed. `n Script Runtime: ($ScriptTime) seconds."
	Add-Event "Script Completed. `n Script Runtime: ($ScriptTime) seconds."

	#exit $0
	}

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

# write-host -f green "Create ServiceNow event for ($date)."
# Add-Event "Create ServiceNow event for ($date)."

<#
# Set up ServiceNow connection pieces from top level parameters
#===============================================

 	[String]$AlertName,
	[String]$AlertID,
	[String]$AssignmentGroup

# 
# Set up EventData variable with SCOM to SNow fields
#===============================================

#
# Multiple locations for HostName in alerts based on class, path, and other variables
# Hostname
# Figure out hostname based on alert values
#===============================================

#
# NOTE * Uncomment as needed for debug write-host lines
 
#Debug
#$EventInfo
 
write-host "MessageKey = $Hosts"
write-host "HostName = $HostName"
#>

Get-SNowParameters

# Verify that SNow Hostname/URL are correct
write-host ""
write-host -f green "New-SNowEvent function verifying SNowHostName = $($SNowHostName)"
Add-Event "New-SNowEvent function verifying SNowHostName = $($SNowHostName)"
#write-host -f green "New-SNowEvent function verifying ServiceNowURL = $($ServiceNowURL)"
#Add-Event "New-SNowEvent function verifying ServiceNowURL = $($ServiceNowURL)"


# Jump to validating alert fields are populated
$MonitoringObjectPath = ($Alert |select MonitoringObjectPath).MonitoringObjectPath
$MonitoringObjectDisplayName = ($Alert |select MonitoringObjectDisplayName).MonitoringObjectDisplayName	
$PrincipalName = ($Alert |select PrincipalName).PrincipalName
$DisplayName = ($Alert |select DisplayName).DisplayName
$PKICertPath = ($Alert |select Path).Path

# Update tests with else for PowerShell script
if ( $Null -ne $MonitoringObjectPath ) { $Hostname = $MonitoringObjectPath }
if ( $Null -ne $MonitoringObjectDisplayName ) { $Hostname = $MonitoringObjectDisplayName }
if ( $Null -ne $PrincipalName ) { $Hostname = $PrincipalName }
if ( $Null -ne $DisplayName ) { $Hostname = $DisplayName }
if ( $Null -ne $PKICertPath ) { $Hostname = $PKICertPath }

$CleanArray = $Hostname | sort -uniq | Where-Object { $_.Trim() -ne "" }
#$CleanArray

# Verify unique Hostname
if ( ( $CleanArray | measure).Count -ge 1 )
	{
	write-host ""
	write-host "Unique hostname array = $CleanArray"
	write-host ""

	# Management Servers Resource Pool
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "All Management Servers Resource Pool" ) )
		{
		$Pool = @( $ENV:ComputerName )
		$Slash = $Pool[0]
		$Server += $Pool[0]
		write-host "Server split for 'Resource Pool' hostname = $Slash"
		write-host ""
		}
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "\" ) )
		{
		$Split = @( $Hostname.split("\") )
		$Slash = $Split[0]
		$Server += $Split[0]
		write-host "Server split for '\' hostname = $Slash"
		write-host ""
		}
	# Cluster variables 
	$Cluster = @()
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "Cluster" ) )
		{
		$Cluster = @( $CleanArray | select-string -pattern "Cluster" -NotMatch | `
		select-string -pattern "Microsoft" -NotMatch | select-string -NOTMatch "ConfigMgr" `
		| select-string -pattern "Cert" -NotMatch )
		$Server += $Cluster
		write-host "Server split for 'Cluster' hostname = $Cluster"
		write-host ""
		}
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "Cluster Group" ) )
		{
		$ClusterGroup = $CleanArray | select-string -SimpleMatch "Cluster Group"
		$ClusterGroup = $ClusterGroup -split '[()]' # ; $ClusterGroup
		$Server += $ClusterGroup[1] | sort -uniq
		write-host "Server split for 'Cluster Group object' hostname = $($ClusterGroup[1])"
		write-host ""
		}
	if ( $null -ne ( $CleanArray | select-string -Pattern " (" -SimpleMatch ) )
		{
		$ClusterRG = $CleanArray | select-string -Pattern " (" -SimpleMatch
		$ClusterRG = $ClusterRG -split '[()]' # ; $ClusterR
		$Server += $ClusterRG[1] | sort -uniq
		write-host "Server split for 'Cluster Group object' hostname = $($ClusterRG[1])"
		write-host ""
		}
	# ConfigManager pack variables 
	$CfgMgr = @()
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "ConfigMgr" ) )
		{
		$CfgMgr = @( $CleanArray | select-string -SimpleMatch "ConfigMgr" )
		$CfgMgr = $CfgMgr -split '[-]'
		$Final = $CfgMgr[1] -split '[ ]'
		$Server += $Final[1]
		write-host "Server split for 'ConfigMgr' hostname = $($Final[1])"
		write-host ""
		}
	#Microsoft
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "Microsoft" ) )
		{
		$Microsoft = @( $CleanArray | select-string -pattern "Microsoft" -NotMatch `
		| select-string -pattern "Cert" -NotMatch | select-string -pattern "Cluster" -NotMatch `
		| select-string -pattern "ConfigMgr" -NotMatch )
		$Server += $Microsoft
		write-host "Server split for 'Microsoft*' hostname = $Microsoft"
		write-host ""
		}
	# PKI certificate variables 
	$PKICerts = @()
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "Cert CN" ) )
		{
		$PKICerts = @( $CleanArray | ForEach-Object { $_.Split('=')[1]; } )
		$Server += $PKICerts
		write-host "Server split for 'PKI certs' hostname = $PKICerts"
		write-host ""
		}
	# SCOM Agent issues
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "Microsoft.SystemCenter.AgentWatchersGroup" ) )
		{
		$AgentWatcher = @( $CleanArray | select-string -NOTMatch "Microsoft.SystemCenter.AgentWatchersGroup" `
		| select-string -pattern "Microsoft" -NotMatch | select-string -pattern "Cluster" -NotMatch )
		$Server += $AgentWatcher[0]
		#$Server = $Server | sort -uniq
		write-host "Server split for 'AgentWatchers' hostname = $AgentWatcher"
		write-host ""
		}
	#SQL
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "Microsoft.SQLServer" ) )
		{
		$SQLServer = @( $CleanArray | ForEach-Object { $_.Split(':')[1]; } )
		$Server += $SQLServer
		write-host "Server split for 'MSSQL object' hostname = $SQLServer"
		write-host ""
		}
	if ( $null -ne ( $CleanArray | select-string -SimpleMatch "SQL server" ) )
		{
		$SQLServer = $CleanArray | select-string -SimpleMatch "SQL server"
		$SQLServer = $name -split '[()]'  #; $SQLServer
		$Server += $SQLServer[1] | sort -uniq
		write-host "Server split for 'SQL Server object' hostname = $($SQLServer[1])"
		write-host ""
		}
	#Windows Server in fields
	if ( $null -eq $Server )
		{
		$FQDN = $Hostname.Split(".")
		# $FQDN[0]
		$Server += $FQDN[0]
		write-host "Server split for '.' hostname = $($FQDN[0])"
		write-host ""
		}
	[string]$Hosts = @( $Server | sort -uniq )
	$Hosts.Replace(" ", "")
	}

if ( $null -ne $Hosts )
	{
	$IP = Resolve-DNSName -Name $Hosts -Type A
	$ServerIP = ($IP.IPAddress)
	# Debug
	# write-host "Server IP = $($ServerIP)"
	}

$HostFinal = $Hosts
$HostFinal

# Combined event
write-host ""
write-host -f yellow "SCOM Alert - Hostname = $HostFinal, Server IP = $ServerIP, AlertName = $AlertName, Alert ID = $AlertID"
Add-Event "SCOM Alert - Hostname = $Hosts, Server IP = $ServerIP, AlertName = $AlertName, Alert ID = $AlertID"


# Get ResolutionState
$AlertResolutionState = $Alert.ResolutionState

<#
# Evaluate alert closed before SCOM channel SNOW script executed
#===============================================

if ( $AlertResolutionState -eq 255 )
	{
	write-host "ServiceNow (SNow) Event NOT needed as SCOM alert is CLOSED"
	write-host ""
	Add-Event "ServiceNow (SNow) Event NOT needed as SCOM alert is CLOSED"

	#Log an event for script ending and total execution time.
	$EndTime = Get-Date
	$ScriptTime = ($EndTime - $StartTime).TotalSeconds
	write-host -f red "Script Completed. `n Script Runtime: ($ScriptTime) seconds."
	Add-Event "Script Completed. `n Script Runtime: ($ScriptTime) seconds."

	exit $0
	}

# Optional $Description enrichment
$AlertParameters = $Alert.Parameters
$AlertID = $Alert.ID.Guid
$AlertManagementGroup = $Alert.ManagementGroup.Name
$AlertCategory = $Alert.Category

# Debug
$ServerIP
$MonitoringObjectDisplayName
$MonitoringObjectFullName
$MonitoringObjectPath
$PrincipalName
$DisplayName
$PKICertPath
$AlertParameters
$AlertID
$AlertManagementGroup
$AlertCategory

# * Uncomment as needed for debug write-host lines
#>

# Recommended SNOW Event $info with SCOM alert enrichment
$Severity = $Alert.Severity

<#
Set Alert Parameters, ID, ManagementGroup, Category from SCOM alert to use in event fields

Need to parse out invalid JSON characters.
Found customer alert examples with invalid characters, i.e. alerts with \, {}
Example parsing out invalid characters, as found $AlertParameters had alerts with \, {}, []

Parameters : {}, { servername.fqdn }, { Alert Text here }

# Setting up secondary logic when product group(s) used PowerShell specific variables were passed into the event description or parameters fields in SCOM alert

if ( $AlertFullName = "MSSQL on Windows: Monitoring error" )
	{
	# $Description = "Management Group: \"WRCCC\".Module: Microsoft.SQLServer.Windows.Module.Monitoring.Performance.ThreadCount.Version: 7.2.0.0..Error(s) was(were) occurred:.Message: .---------- Exception: ----------.Exception Type: System.TimeoutException.Message: Module execution was terminated due to timeout after 300.000 seconds.Source: Microsoft.SQLServer.Module4.Helper.Stack Trace: .   at Microsoft.SQLServer.Module.Helper.Base.ModuleBasePropertyHelper`1.\u003cGetOutputDataAsync\u003ed__16.MoveNext()..---------- Inner Exception: ----------.Exception Type: System.Threading.Tasks.TaskCanceledException.Message: A task was canceled..Source: mscorlib.Stack Trace: .   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Windows.Module.Monitoring.Performance.ThreadCount.\u003cGetPropertyBagAsync\u003ed__2.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetPropertyBagDataAsyncStatic\u003ed__5`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at Microsoft.SQLServer.Module.Helper.Base.ModuleBasePropertyHelper`1.HandleException(Exception exception, Boolean hideSqlExceptions).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetPropertyBagDataAsyncStatic\u003ed__5`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetModuleDataAsync\u003ed__3`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetModuleDataAsync\u003ed__0`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Core.Module.Helper.Base.ModuleBasePropertyHelperSql`1.\u003cGetModuleDataAsync\u003ed__9.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.ModuleBasePropertyHelper`1.\u003cGetOutputDataAsync\u003ed__16.MoveNext()...State:.The configuration properties are: .ManagementGroupName : WRCCC.Publisher : SQLMonitoringWindows.ConnectionString : BELVW054AAB7SS1.nae.ds.army.mil.InstanceEdition : Standard Edition.InstanceName : MSSQLSERVER.InstanceVersion : 15.0.4382.1.MachineName : BELVW054AAB7SS1.nae.ds.army.mil.MonitoringType : Local.NetbiosComputerName : BELVW054AAB7SS1.Login : .SqlExecTimeoutSeconds : 60.SqlTimeoutSeconds : 15.TimeoutSeconds : 300.Password : ********..Error(s):..---------- Exception: ----------.Exception Type: System.TimeoutException.Message: Module execution was terminated due to timeout after 300.000 seconds.Source: Microsoft.SQLServer.Module4.Helper.Stack Trace: .   at Microsoft.SQLServer.Module.Helper.Base.ModuleBasePropertyHelper`1.\u003cGetOutputDataAsync\u003ed__16.MoveNext()..---------- Inner Exception: ----------.Exception Type: System.Threading.Tasks.TaskCanceledException.Message: A task was canceled..Source: mscorlib.Stack Trace: .   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Windows.Module.Monitoring.Performance.ThreadCount.\u003cGetPropertyBagAsync\u003ed__2.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetPropertyBagDataAsyncStatic\u003ed__5`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at Microsoft.SQLServer.Module.Helper.Base.ModuleBasePropertyHelper`1.HandleException(Exception exception, Boolean hideSqlExceptions).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetPropertyBagDataAsyncStatic\u003ed__5`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetModuleDataAsync\u003ed__3`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.DataItemHelper.\u003cGetModuleDataAsync\u003ed__0`1.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Core.Module.Helper.Base.ModuleBasePropertyHelperSql`1.\u003cGetModuleDataAsync\u003ed__9.MoveNext().--- End of stack trace from previous location where exception was thrown ---.   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw().   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task).   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task).   at Microsoft.SQLServer.Module.Helper.Base.ModuleBasePropertyHelper`1.\u003cGetOutputDataAsync\u003ed__16.MoveNext().."
	$AlertParameters = "Management Group: $($AlertManagementGroup) Module: Microsoft.SQLServer.Windows.Module.Monitoring* workload caused stack trace.  See SCOM alert for details on what workflow failed, timed out, execution or termination failure after NOT completing for 300 seconds."
	}

if ( $AlertFullName = "MSSQL on Windows: Monitoring error" )
	{
	$AlertParameters = "component SMS_SITE_COMPONENT_MANAGER on computer $Hostname reported: Site Component Manager could not access site system $Hostname. The operating system reported error 2147942453: The network path was not found. `
	Possible cause: The site system is turned off, not connected to the network, or not functioning properly. Solution: Verify that the site system is turned on, connected to the network, and functioning properly."
	}

#>


<#
# Determine SCOM Alert Severity to ITSM tool impact
#===============================================
$Severity = $Alert.Severity
#>

If ( $Severity -eq "Warning" )
	{
	$EventSeverity = "Minor"
	}
If ( $Severity -eq "Error" )
	{
	$EventSeverity = "Minor"
	}
If ( $Severity -eq "Critical" )
	{
	$EventSeverity = "Critical"
	}
else
	{
	$EventSeverity = "Warning"
	}

# Closed alert example
If ( $AlertResolutionState -eq 255 )
	{
	$EventSeverity = "Clear"
	}

<#
# Debug Severity
# write-host -f green "Severity = $Severity"
# Add-Event $Severity
#
# NOTE * Uncomment as needed for debug write-host lines

# write-host -f green "SCOM Alert ID = $AlertID, Severity = $EventSeverity, ServerName = $Hosts, Server IP = $ServerIP, `
# Alert Parameters = $AlertParameters, SCOM management group = $AlertManagementGroup, Alert Category = $AlertCategory"
# Non-function write event
# $momapi.LogScriptEvent($ScriptName,$EventID,0,"SCOM Alert ID = $AlertID, Severity = $EventSeverity, Hostname = $Hosts, Server IP = $ServerIP, `
# Alert Parameters = $AlertParameters, SCOM management group = $AlertManagementGroup, Alert Category = $AlertCategory")
#>


$AlertParameters = $Alert.Parameters


<# 
# Filtering for Certificate, MECM and MSSQL related alerts with double quotes
#if ( ( $AlertName -like "Certificate lifespan alert" ) )
#	{
#	$AlertParameters = "{component SMS_DATABASE_NOTIFICATION_MONITOR on computer $PrincipalName reported:  The task configured to run today, but could not run within the scheduled time.}"
#	}
#>


if ( ( $AlertName -like "*Maintenance task failure alert" ) )
	{
	$AlertParameters = "{component SMS_DATABASE_NOTIFICATION_MONITOR on computer $PrincipalName reported:  The task configured to run today, but could not run within the scheduled time.}"
	}

if ( $AlertName -like "*Fail to access site system alert" )
	{
	$AlertParameters = "{SMS_SITE_COMPONENT_MANAGER, $HostName}"
	}

if ( $AlertName -like "*Failed to update Active Directory alert" )
	{
	$AlertParameters = "{component SMS_HIERARCHY_MANAGER on computer $Hostname reported: Configuration Manager cannot update the already existing object in Active Directory.}"
	}

if ( $AlertName -like "*File Dispatch Manager Not Connecting to Site Server" )
	{
	$AlertParameters = "{$PrincipalName, component SMS_MP_FILE_DISPATCH_MANAGER on computer $PrincipalName reported:  MP File Dispatch Manager running on the management point cannot connect to site server $PrincipalName. The operating system reported error 53: The network path was not found.}"
	}

if ( $AlertName -like "*Maintenance task failure alert" )
	{
	$AlertParameters = "{The site server failed to execute a maintenance task.`
Event Description: component SMS_DATABASE_NOTIFICATION_MONITOR on $hostname reported: The task 'Delete Obsolete Alerts' is configured to run today, but it could not run within the scheduled time.}"
	}

if ( $AlertName -like "*Management Point NOT available" )
	{
	$AlertParameters = "{SMS_MP_CONTROL_MANAGER, $HostName}"
	}

if ( $AlertName -like "*Sender connection failure alert" )
	{
	$AlertParameters = "{component SMS_LAN_SENDER on computer $Hostname reported:  The sender cannot connect to remote site over the LAN. The operating system reported error 53: `
The network path was not found....Possible cause: Remote site server might not be connected to the LAN.`
Solution: Verify that site server is connected to the LAN and functioning properly.`
Possible cause: Share 'SMS_SITE' might not be accessible.`
Solution: Verify that share 'SMS_SITE' is visible and that the site server machine account has the necessary permissions to access the share.`
Possible cause: Network load might be too high.`
Solution: Verify that the network is not saturated. Verify that you can move large amounts of data across the network.}"
	}

if ( $AlertName -like "*WSUS subscription alert" )
	{
	$AlertParameters = "{component SMS_WSUS_CONFIGURATION_MANAGER on computer $PrincipalName reported:  WSUS Configuration Manager failed to subscribe to update categories and classifications on WSUS Server.}"
	}

#MSSQL
if ( $AlertName -like "MSSQL on Windows: *" )
	{
	if ( $AlertName -eq "MSSQL on Windows: DB Engine is in unhealthy state" )
		{
		$AlertParameters = "{$Hostname, Instance is unavailable}"
		}

	if ( $AlertName -eq "MSSQL on Windows: Database is in offline/recovery pending/suspect/emergency state" )
		{
		$AlertParameters = "Database on SQL Server instance, computer $Hostname has one or more databases offline/recovery pending/suspect/emergency."
		}

	if ( $AlertName -eq "MSSQL on Windows: Discovery error" )
		{
		$AlertParameters = "{Management Group: $($AlertManagementGroup) Module: Microsoft.SQLServer.Windows.Module.Discovery.*`
Message: [Error number: 916] The server principal 'NT AUTHORITY\SYSTEM' is not able to access the database under the current security context.}"
		}

	if ( $AlertName -eq "MSSQL on Windows: Monitoring error" )
		{
		$AlertParameters = "{Management Group: $($AlertManagementGroup) Module: Microsoft.SQLServer.Windows.Module.Monitoring* workload caused stack trace.  `
See SCOM alert for details on what workflow failed, timed out, execution or termination failure after NOT completing for 300 seconds.}"
		}
	}

write-host ""
write-host "AlertParameters variable BEFORE JSON parsing = $AlertParameters"

$AlertParameter = $AlertParameters -replace "/", "-"
$AlertParameter = $AlertParameter -replace "{", ""
$AlertParameter = $AlertParameter -replace "}", ""
$AlertParameter = $AlertParameter -replace "{ }", "NotNull"
$AlertParameter = $AlertParameter -replace "^@", "amp"
$AlertParameter = $AlertParameter -replace "=", ":"
$AlertParameter = $AlertParameter -replace ";", ","
$AlertParameter = $AlertParameter -replace "`", "#"
$AlertParameter = $AlertParameter -replace "\\\\n", "...n"
$AlertParameter = $AlertParameter -replace "\\\n", "..n"
$AlertParameter = $AlertParameter -replace "\\n", ".n"
$AlertParameter = $AlertParameter -replace "\n", "."
$AlertParameter = $AlertParameter -replace "\r", ".."
$AlertParameter = $AlertParameter -replace "\\", "."
$AlertParameter = $AlertParameter -replace '"','.'
$AlertParameter = $AlertParameter -replace "'","."


write-host ""
write-host "AlertParameter AFTER JSON parsing = $AlertParameter"

# Used with manual testing $Alerts
#$AlertID = $Alert.ID.Guid

$AlertManagementGroup =  $Alert.ManagementGroup.Name
$AlertCategory = $Alert.Category

# Set SCOM Pack, Class, Object into Message_key
$AlertFullName = $Alert.MonitoringObjectFullName
#write-host "AlertFullName = $($AlertFullName)"
#write-host ""

<#
#
# NOTE * Uncomment as needed for debug write-host lines

#Debug
#$EventInfo

write-host "MessageKey = $Hosts"
write-host "HostName = $HostName"

# Create custom SNOW short_description for incident/event title or name
# Examples

Example SNOWAlertName
	$SNOWAlertName = "<Org> <Team> SCOM Test Event - $Alert"
Example SNOWAlertName
	$SNOWAlertName = "<Team> <ORG> SCOM Event - $AlertName"
Example SNOWAlertName
	$SNOWAlertName = "<Team> <ORG> SCOM $AlertName"
Example SNOWAlertName
	$SNOWAlertName = "##CUSTOMER## ##TEAM## SCOM Event - $AlertName"
Example SNOWAlertName
	$SNOWAlertName = "##TEAM## ##CUSTOMER##: SCOM - $AlertName"

# Customer example using passed -Team parameter
$SNOWAlertName = "$Team: SCOM - $AlertName"

# Examples
# M365 Messaging, SharePoint

#>

# Set AlertName for Testing
$SNOWAlertName = "$Team SCOM - $AlertName"


# Display AlertDescription (Debug)
#write-host "Alert Description = $($AlertDescription)

<# Determine SCOM Alert Description excludes JSON special characters
#===============================================
# NOTE * Uncomment as needed for debug write-host lines

# write-host -f green "Begin SCOM Alert Description JSON audit"

Change Alert.Description to new variable, then run through the meat grinder, to remove PowerShell and JSON characters
i.e. 
PKI Cert expiring alert, parameter has double quotes
Filtering for MECM and MSSQL related alerts with double quotes

#>

# Alert description
$AlertDescription = $Alert.Description

$Description = $Alert.Description


# AD Dell hardware failures create medium (non outage) incident
if ( $AlertName -like "DELL OMS*" )
	{
	$EventSeverity = "Warning"
	}


# PKI Cert expiring alert, parameter has double quotes
if ( ( $AlertName -like "Certificate Lifespan alert" ) )
	{
	$AlertParameter = $Description
	}


# Filtering for MECM and MSSQL related alerts with double quotes
if ( ( $AlertName -like "*Fail to access site system alert" ) )
	{
	$Description = "component SMS_SITE_COMPONENT_MANAGER on computer $Hostname reported: Site Component Manager could not access site system $Hostname. `
The operating system reported error 2147942453: The network path was not found. `
Possible cause: The site system is turned off, not connected to the network, or not functioning properly.`
Solution: Verify that the site system is turned on, connected to the network, and functioning properly."
	}

if ( $AlertName -like "*Failed to update Active Directory alert" )
	{
	$Description = "The site server failed to update objects in Active Directory.`
Event Description: On 7/22/2024 2:17:54 AM, component SMS_HIERARCHY_MANAGER on computer $Hostname reported:  Configuration Manager cannot update the already existing object in Active Directory."
	}

if ( $AlertName -like "*File Dispatch Manager Not Connecting to Site Server" )
	{
	$Description = "The file dispatch manager on $PrincipalName fails to connect to site server of another site. The operating system reported error 53: The network path was not found."
	}

if ( $AlertName -like "*Maintenance task failure alert" )
	{
	$Description = "The site server failed to execute a maintenance task.`
Event Description: component SMS_DATABASE_NOTIFICATION_MONITOR on computer $PrincipalName reported:  The task configured to run today, but could not run within the scheduled time."
	}

if ( $AlertName -like "*Sender connection failure alert" )
	{
	$Description = "component SMS_LAN_SENDER on computer $Hostname reported:  The sender cannot connect to remote site over the LAN. The operating system reported error 53: `
The network path was not found....Possible cause: Remote site server might not be connected to the LAN.`
Solution: Verify that site server is connected to the LAN and functioning properly.`
Possible cause: Share 'SMS_SITE' might not be accessible.`
Solution: Verify that share 'SMS_SITE' is visible and that the site server machine account has the necessary permissions to access the share.`
Possible cause: Network load might be too high.`
Solution: Verify that the network is not saturated. Verify that you can move large amounts of data across the network."
	}

if ( $AlertName -eq "MECM Central Site to Primary Site Global Data Sending Not Working" )
	{
	$EventSeverity = "Minor"
	}
if ( $AlertName -eq "MECM Primary Site to Central Site Global Data Receiving Not Working" )
	{
	$EventSeverity = "Minor"
	}
if ( $AlertName -eq "MECM Primary site to Secondary Site Global Data Receiving Not Working" )
	{
	$EventSeverity = "Minor"
	}



#MSSQL
if ( $AlertName -eq "MSSQL on Windows: SQL Server service stopped" )
	{
	$EventSeverity = "Critical"
	}
if ( $AlertName -like "MSSQL on Windows: *" )
	{
	if ( $AlertName -eq "MSSQL on Windows: Database is in offline/recovery pending/suspect/emergency state" )
		{
		$Description = "Database on SQL Server instance, computer $Hostname has one or more databases offline/recovery pending/suspect/emergency."
		}

	if ( $AlertName -eq "MSSQL on Windows: DB Engine is in unhealthy state" )
		{
		$Description = "The SQL Server instance on computer $Hostname is unhealthy and reports 'Instance is unavailable'."
		$EventSeverity = "Warning"
		}

	if ( $AlertName -eq "MSSQL on Windows: Discovery error" )
		{
		$Description = "Management Group: $($AlertManagementGroup), Module: Microsoft.SQLServer.Windows.Module.Discovery.* `
Message: [Error number: 916] The server principal 'NT AUTHORITY\SYSTEM' is not able to access the database under the current security context.}"
		}

	if ( $AlertName -eq "MSSQL on Windows: Monitoring error" )
		{
		$Description = "Management Group: $($AlertManagementGroup),  Module: Microsoft.SQLServer.Windows.Module.Monitoring* `
Message: Execution Timeout Expired.  The timeout period elapsed prior to completion of the operation or the server is not responding."
		}	  
	}


# Logical Disk Free Space is low
if ( $AlertName -eq "Logical Disk Free Space is low" )
	{
	$SNOWAlertName = "$Team SCOM - $MonitoringObjectDisplayName $AlertName"
	$EventSeverity = "Warning"
	}


# WSMT subscription (includes global for Logical Disk Free space
if ( $AlertName -eq "Failed to Connect to Computer" )
	{
	if ( ( $Team -eq "DBS" ) -OR ( $Team -eq "ADS" ) ) { $EventSeverity = "Warning" }
	}


# Adjust description for delimiters
$Description = $Description -replace "/", "."
$Description = $Description -replace "{", ""
$Description = $Description -replace "}", ""
$Description = $Description -replace "{ }", "NotNull"
$Description = $Description -replace "^@", "amp"
$Description = $Description -replace "=", ":"
$Description = $Description -replace ";", ","
$Description = $Description -replace "`", "#"
$Description = $Description -replace "\\\\n", "...n"
$Description = $Description -replace "\\\n", "..n"
$Description = $Description -replace "\\n", ".n"
$Description = $Description -replace "\n", "."
$Description = $Description -replace "\r", ".."
$Description = $Description -replace "\\", "."
$Description = $Description -replace '"','.'
$Description = $Description -replace "'","."


<#
# Completed JSON format checks

# write-host -f green "SCOM Alert Description formatted for JSON"
# write-host -f green $Description
# Add-Event $Description


# Check if any alert variables are null, and provide a value
# NOTE: ServiceNow does not like null!

# Testing shows that $info nulls result in 401 unauthorized SNOW REST submissions
#
$info = "{`"Alert Details`":`"$($Description)`",`"SCOM Alert Parameters`":`"$($AlertParameter)`",`"SCOM Alert ID`":`"$($AlertID)`",`"Hostname`":`"$($Server)`",`"MonitoringObjectPath`":`"$($MonitoringObjectPath)`",`"MonitoringObjectDisplayName`":`"$($MonitoringObjectDisplayName)`",`"PrincipalName`":`"$($PrincipalName)`",`"DisplayName`":`"$($DisplayName)`",`"PKICertPath`":`"$($PKICertPath)`",`"asset_org`":`"$($AssignmentGroup)`"}"
 
if ( ( $null -eq $PrincipalName ) -OR ( $null -eq $DisplayName ) -OR ( $null -eq $PKICertPath ) )
	{
	$info = "{`"Alert Details`":`"$($Description)`",`"SCOM Alert Parameters`":`"$($AlertParameter)`",`"SCOM Alert ID`":`"$($AlertID)`",`"Hostname`":`"$($Server)`",`"MonitoringObjectPath`":`"$($MonitoringObjectPath)`",`"MonitoringObjectDisplayName`":`"$($MonitoringObjectDisplayName)`",`"asset_org`":`"$($AssignmentGroup)`"}"
	}
if ( ( $null -ne $PKICertPath ) )
	{
	$info = "{`"Alert Details`":`"$($Description)`",`"SCOM Alert Parameters`":`"$($AlertParameter)`",`"SCOM Alert ID`":`"$($AlertID)`",`"Hostname`":`"$($Server)`",`"MonitoringObjectPath`":`"$($MonitoringObjectPath)`",`"MonitoringObjectDisplayName`":`"$($MonitoringObjectDisplayName)`",`"PKICertPath`":`"$($PKICertPath)`",`"asset_org`":`"$($AssignmentGroup)`"}"
	}

# SCOM alert parameters that typically change based on which author (PG) built pack

# Audit of SCOM Alert fields used in script:

# SCOM Alert pieces:
#
# AlertCategory
# Description,
# parameters,
# Hostname,
# MonitoringObjectDisplayName,
# MonitoringObjectFullName,
# MonitoringObjectPath,
# PrincipalName,
# DisplayName
# PKICertPath

#>


if ( $null -eq $AssignmentGroup ) { $AssignmentGroup = "$AssignmentGroup empty" }
if ( $null -eq $AlertCategory ) { $AlertParameter = "AlertParameter_empty" }
if ( $null -eq $AlertParameter ) { $AlertParameter = "AlertParameter_empty"	}
if ( $null -eq $Hostname ) { $Hostname = "Hostname_empty" }
if ( $null -eq $MonitoringObjectDisplayName ) { $MonitoringObjectDisplayName = "MonitoringObjectDisplayName_empty" }
if ( $MonitoringObjectDisplayName -like "*\*" )
	{ 
	$MonitoringObjectDisplayName = $MonitoringObjectDisplayName -replace "\", "."
	write-host "MonitoringObjectDisplayName Final = $MonitoringObjectDisplayName"
	}
if ( $null -eq $MonitoringObjectFullName ) { $MonitoringObjectFullName = "MonitoringObjectFullName_empty" }
if ( $null -eq $MonitoringObjectPath ) { $MonitoringObjectPath = "MonitoringObjectPath_empty" }
if ( $null -eq $PrincipalName ) { $PrincipalName = "PrincipalName_empty" }
if ( $null -eq $DisplayName ) { $DisplayName = "DisplayName_empty" }
if ( $null -eq $PKICertPath ) { $PKICertPath = "PKICertPath_empty" }



write-host ""
write-host -f yellow "SCOM Alert ID = $AlertID, Severity = $EventSeverity, ServerName = $Hosts, Server IP = $ServerIP, Alert Parameters = $AlertParameter, SCOM management group = $AlertManagementGroup, Alert Category = $AlertCategory"
Add-Event "SCOM Alert ID = $AlertID, Severity = $EventSeverity, ServerName = $HostFinal, Server IP = $ServerIP, Alert Parameters = $AlertParameter, SCOM management group = $AlertManagementGroup, Alert Category = $AlertCategory"


<#
# Build SNow event payload
#=======================================================
#
$info became hashtable causing JSON REST API insert failures.
TS helped with options.

Simple, converted twice, would cause lots of escaped characters with slashes (\)
$info = @{nic=”1234”} | ConvertTo-Json
$info = @{Test="12345"} | ConvertTo-Json
$info = "{ SCOM Alert Parameters: $AlertParameters , SCOM Alert ID: $AlertID }" | ConvertTo-Json

# Change from array to JSON formatted, requires comma's, squiggly braces
# $additionalContent = @ {"asset_org: $AssignmentGroup" } | ConvertTo-Json

# SNOW provided event body example
#$body = @{source="SCOM";node="$Hostname";type="Alert";severity="$EventSeverity";description="This is a SCOM alert -  $Description";resource="NIC2";metric_name="Adapter2";additional_info=$info;event_class="SCOM from $ServerIP";message_key="SCOM-$Hostname-Alert"} | ConvertTo-Json

# Array example that parses to JSON, but SNOW can't read as there is only one column
$info = @("Alert Details: $Description","SCOM Alert Parameters: $AlertParameters","SCOM Alert ID: $AlertID","Hostname: $Server","MonitoringObjectPath: $MonitoringObjectPath","MonitoringObjectDisplayName: $MonitoringObjectDisplayName","PrincipalName: $PrincipalName","DisplayName: $DisplayName","PKICertPath: $PKICertPath","asset_org: $AssignmentGroup" )

# Array example that also parses to JSON that SNOW can read, that causes System.Collections.Hashtable in PoSH
$info = @{ "Alert Details" = $Description;"SCOM Alert Parameters" = $AlertParameters;"SCOM Alert ID" = $AlertID;"Hostname" = $Server;"MonitoringObjectPath" = $MonitoringObjectPath;"MonitoringObjectDisplayName" = $MonitoringObjectDisplayName;"PrincipalName" = $PrincipalName;"DisplayName" = $DisplayName;"PKICertPath" = $PKICertPath;"asset_org" = $AssignmentGroup } #| ConvertTo-Json

SPECTRUM example
{"probable_cause": "DEVICE HAS STOPPED RESPONDING TO POLLSSYMPTOMS:Device has stopped responding to polls.PROBABLE CAUSES:1) Device Hardware Failure.2) Cable between this and upstream device broken.3) Power Failure.4) Incorrect Network Address.5) Device Firmware Failure.RECOMMENDED ACTIONS:1) Check power to device.2) Verify status lights on device.3) Verify reception of packets.4) Verify network address in device and SPECTRUM.5) Cycle power on device and recheck.6) If above fails, call repair.","asset_org": "RCC-C - Network Infrastructure","asset_tag": "PRX RCC-C","event_message": "Wed 03 Jul, 2024 - 15:05:42 - Device NOVOW054AAPXJH1 of type BluecoatProxySG has stopped responding to polls and/or external requests.  An alarm will be generated.   (event [0x00010d35])","network_address": "140.153.164.26","device_class": "BluecoatProxySG","device_type": "Reverse Proxy","device_id": "0x1f688587"}

SolarWinds Example
$Additional = "{'DNS Node Name':'$($args[0])','DNS Host Name':'$($args[1])','Alert Name':'$($args[2])','Host Status':'$($args[3])','Alert Description':'$($args[4])','vCenter':'$($args[5])','VCenter Description':'$($args[6])','Host Name':'$($args[7])'}"

#$Additional

#$Additional = "{"Alert Details"=$Description;"SCOM Alert Parameters"= $AlertParameter;"SCOM Alert ID"=$AlertID;"Hostname"=$Server;`
#"MonitoringObjectPath"=$MonitoringObjectPath;"MonitoringObjectDisplayName"=$MonitoringObjectDisplayName;"PrincipalName"=$PrincipalName;`
#"DisplayName"=$DisplayName;"PKICertPath"=$PKICertPath;"asset_org"=$AssignmentGroup}"

The JSON format allows brackets and other special characters inside of double quotes 
[ {"name":"[value%#<>","name":"Other^%#@!" ]. 
However, this requires any double quote (") to be removed or escaped (\").

Example
'{"scom-severity":"Medium","metric-value":"38","os_type":"Windows.Server.2008"}'


Example parsing out invalid characters, as found $AlertParameters had alerts with \, {}, []
$AlertParameter = $AlertParameters -replace "/", "."
$AlertParameter = $AlertParameters -replace "[", ""
$AlertParameter = $AlertParameters -replace "]", ""
$AlertParameter = $AlertParameters -replace "^@", ""
$AlertParameter = $AlertParameters -replace "=", ":"
$AlertParameter = $AlertParameters -replace ";", ","
$AlertParameter = $AlertParameters -replace "`", "#"
$AlertParameter = $AlertParameters -replace "{", " "
$AlertParameter = $AlertParameters -replace "}", " "
$AlertParameter = $AlertParameters -replace "\n", "\\n"
$AlertParameter = $AlertParameters -replace "\r", "\\r"
$AlertParameter = $AlertParameters -replace "\\\\n", "\\n"

Leverage $AlertParameter parsing out JSON characters ( variable as hash or array?)

write-host ""
write-host "ServiceNow Event info payload"
$info = "{'Alert Details':'$($Description)','SCOM Alert Parameters': $($AlertParameter),'SCOM Alert ID':$AlertID,'Hostname':$($Server),`
'MonitoringObjectPath':$($MonitoringObjectPath),'MonitoringObjectDisplayName':$($MonitoringObjectDisplayName),'PrincipalName':$($PrincipalName),`
'DisplayName':$($DisplayName),'PKICertPath':$($PKICertPath),'asset_org': $($AssignmentGroup)}"
#$info

write-host ""
write-host "ServiceNow Event info payload `n $($info)"
Add-Event "ServiceNow Event info payload `n $($info)"

write-host ""
write-host "============================================"
write-host "ServiceNow Event info JSON converted payload"
#$info | ConvertTo-Json

# Test Posh variable with data in hashtable
write-host ""
write-host "ServiceNow Event info payload"
$ArrayInfo = @{"Alert Details" = $($Description);"SCOM Alert Parameters" = $($AlertParameter);"SCOM Alert ID" = $AlertID;"Hostname" = $($Server);`
"MonitoringObjectPath" = $($MonitoringObjectPath);"MonitoringObjectDisplayName" = $($MonitoringObjectDisplayName);"PrincipalName" = $($PrincipalName);`
"DisplayName" = $($DisplayName);"PKICertPath" = $($PKICertPath);"asset_org" = $($AssignmentGroup)}

#$ArrayInfo

write-host ""
write-host "ServiceNow Event info ArrayInfo JSON converted payload"
#$ArrayInfo | ConvertTo-Json

# Test Posh variable without setting variable with $($variable) data in hashtable
$ArrayInfo2 = @{"Alert Details" = $Description;"SCOM Alert Parameters" = $AlertParameter;"SCOM Alert ID" = $AlertID;"Hostname" = $Server;`
"MonitoringObjectPath" = $MonitoringObjectPath;"MonitoringObjectDisplayName" = $MonitoringObjectDisplayName;"PrincipalName" = $PrincipalName;`
"DisplayName" = $DisplayName;"PKICertPath" = $PKICertPath;"asset_org" = $AssignmentGroup}

$ArrayInfo2

write-host ""
write-host "ServiceNow Event info ArrayInfo2 JSON converted payload"
$ArrayInfo2 | ConvertTo-Json


Recommended trying to code around the JSON - 
Need to use back ticks (`) and single quotes, no convert-toJSON

#$info = "{'Alert Details':'$($Description)','SCOM Alert Parameters':'$($AlertParameter)','SCOM Alert ID':'$($AlertID)','Hostname':'$($Server)','MonitoringObjectPath':'$($MonitoringObjectPath)','MonitoringObjectDisplayName':'$($MonitoringObjectDisplayName)','PrincipalName':'$($PrincipalName)','DisplayName':'$($DisplayName)','PKICertPath':'$($PKICertPath)','asset_org':'$($AssignmentGroup)'}"
#$info = "{'Alert Details':'$($Description)','SCOM Alert Parameters':'$($AlertParameter)','SCOM Alert ID':'$($AlertID)','Hostname':'$($Server)','MonitoringObjectPath':'$($MonitoringObjectPath)','MonitoringObjectDisplayName':'$($MonitoringObjectDisplayName)','asset_org':'$($AssignmentGroup)'}"
$info = "{`"Alert Details`":`"$($Description)`",`"SCOM Alert Parameters`":`"$($AlertParameter)`",`"SCOM Alert ID`":`"$($AlertID)`",`"Hostname`":`"$($Server)`",`"MonitoringObjectPath`":`"$($MonitoringObjectPath)`",`"MonitoringObjectDisplayName`":`"$($MonitoringObjectDisplayName)`",`"PrincipalName`":`"$($PrincipalName)`",`"DisplayName`":`"$($DisplayName)`",`"PKICertPath`":`"$($PKICertPath)`",`"asset_org`":`"$($AssignmentGroup)`"}"
$info = "{`"Alert Details`":`"$($Description)`",`"SCOM Alert Parameters`":`"$($AlertParameter)`",`"SCOM Alert ID`":`"$($AlertID)`",`"Hostname`":`"$($Server)`",`"MonitoringObjectPath`":`"$($MonitoringObjectPath)`",`"MonitoringObjectDisplayName`":`"$($MonitoringObjectDisplayName)`",`"PrincipalName`":`"$($PrincipalName)`",`"DisplayName`":`"$($DisplayName)`",`"PKICertPath`":`"$($PKICertPath)`",`"asset_org`":`"$($AssignmentGroup)`"}"

#write-host ""
#$info | ConvertTo-Json -Compress
#write-host ""
 
if ( ( $null -eq $PrincipalName ) -OR ( $null -eq $DisplayName ) -OR ( $null -eq $PKICertPath ) )
	{
	$info = "{'Alert Details':'$($Description)','SCOM Alert Parameters':'$($AlertParameter)','SCOM Alert ID':'$($AlertID)','Hostname':'$($Server)','MonitoringObjectPath':'$($MonitoringObjectPath)','MonitoringObjectDisplayName':'$($MonitoringObjectDisplayName)','asset_org':'$($AssignmentGroup)'}"
	}
if ( ( $null -ne $PKICertPath ) )
	{
	$info = "{'Alert Details':'$($Description)','SCOM Alert Parameters':'$($AlertParameter)','SCOM Alert ID':'$($AlertID)','Hostname':'$($Server)','MonitoringObjectPath':'$($MonitoringObjectPath)','MonitoringObjectDisplayName':'$($MonitoringObjectDisplayName)','asset_org':'$($AssignmentGroup)','PKICertPath':'$($PKICertPath)'}"
	}

#write-host ""
#$info | ConvertTo-Json -Compress
#write-host ""
write-host "ServiceNow Event info JSON payload `n $($info)"
Add-Event "ServiceNow Event info JSON payload `n $($info)"


# Gives http400 bad request with no JSON conversion nor hashtable at (@) sign, without the ConvertTo-Json at the end for $info and $body

$body = "{`"source`":`"SCOM-CONUS`",`"node`":`"$($Hostname)`",`"type`":`"Alert`",`"severity`":`"$($EventSeverity)`",`"description`":`"$($SNOWAlertName)`",`"resource`":`"$($AlertCategory)`",`"metric_name`":`"$($SNOWAlertName)`",`"additional_info`":`"$($info)`",`"event_class`":`"SCOM from Server $($Server) $($ServerIP)`",`"message_key`":`"SCOM-$($AlertFullName)-Alert`"}"

write-host -f green "Body BEFORE JSON conversion = `n $($body)"
write-host ""

$Convert = $body | ConvertTo-Json

write-host -f green "Body AFTER JSON conversion = `n $($Convert)"
write-host ""

# Old code from 11 Jul
$body = @{source="SCOM-CONUS";node="$Hostname";type="Alert";severity="$EventSeverity";description="$SNOWAlertName";resource="$AlertCategory";metric_name="$SNOWAlertName";additional_info=$Additional;event_class="SCOM from Server $Server $ServerIP";message_key="SCOM-$AlertFullName-Alert"} | ConvertTo-Json

# ServiceNow SME recommendation
$body = @{"source"="TCS";"event_class"="SCOM 2007 on scom.server.com";"resource"="C:";"node"="name.of.node.com";"metric_name"="Percentage Logical Disk Free Space";"type"="Disk space";"severity"="4";"description"="The disk C: on computer V-W2K8-dfg.dfg.com is running out of disk space. The value that exceeded the threshold is 41% free space."; "additional_info"='{"scom-severity":"Medium","metric-value":"38","os_type":"Windows.Server.2008"}'} | ConvertTo-Json;
write-host ""
write-host "ServiceNow Event Body payload `n `n $($body)"

Recommended trying to code around the JSON - 
Need to use back ticks (`) and single quotes, no convert-toJSON

$body = "{`"source`":`"SCOM-CONUS`",`"node`":`"$($Hostname)`",`"type`":`"Alert`",`"severity`":`"$($EventSeverity)`",`"description`":`"$($SNOWAlertName)`",`"resource`":`"$($AlertCategory)`",`"metric_name`":`"$($SNOWAlertName)`",`"additional_info`":$($info),`"event_class`":`"SCOM from Server $($Server) $($ServerIP)`",`"message_key`":`"SCOM-$($AlertFullName)-Alert`"}"

Back to original with JSON

#>

$info = @{"Alert Details"="$($Description)";"SCOM Alert Parameters"="$($AlertParameter)";"SCOM Alert ID"="$($AlertID)";"Hostname"="$($HostFinal)";"MonitoringObjectPath"="$($MonitoringObjectPath)";"MonitoringObjectDisplayName"="$($MonitoringObjectDisplayName)";"PrincipalName"="$($PrincipalName)";"DisplayName"="$($DisplayName)";"PKICertPath"="$($PKICertPath)";"asset_org"="$($AssignmentGroup)"} | ConvertTo-Json -Compress

write-host ""
write-host "============================================"
write-host "ServiceNow Event info JSON payload `n $($info)"
write-host "============================================"
Add-Event "ServiceNow Event info JSON payload `n $($info)"


<#
# NOTE * Uncomment as needed for debug write-host lines for REST body

# Debug verify Info and Body variables
$info

write-host "ServiceNow Event payload `n `n $($body)"
Add-Event "ServiceNow Event payload `n `n $($body)"

# Hashtable does not format correctly, nor array
# $body = @{source="SCOM";node="$Hostname";type="Alert";severity="$EventSeverity";description="$SNOWAlertName";resource="$AlertCategory";metric_name="$SNOWAlertName";additional_info=$Additional;event_class="SCOM from Server $Server $ServerIP";message_key="SCOM-$AlertFullName-Alert"} | ConvertTo-Json

Use this method for $body variable
$body = "{`"source`":`"SCOM`",`"node`":`"$($Hostname)`",`"type`":`"$($Type)`",`"severity`":`"$($EventSeverity)`",`"description`":`"$($SNOWAlertName)`",`"resource`":`"$($AlertCategory)`",`"metric_name`":`"$($SNOWAlertName)`",`"additional_info`":`"$($info)`",`"event_class`":`"SCOM from Server $($Server) $($ServerIP)`",`"message_key`":`"SCOM-$($AlertFullName-Alert)`"}"

Recommended trying to code around the JSON - 
Need to use back ticks (`) and single quotes, no convert-toJSON

$body = "{`"source`":`"SCOM-CONUS`",`"node`":`"$($Hostname)`",`"type`":`"Alert`",`"severity`":`"$($EventSeverity)`",`"description`":`"$($SNOWAlertName)`",`"resource`":`"$($AlertCategory)`",`"metric_name`":`"$($SNOWAlertName)`",`"additional_info`":$($info),`"event_class`":`"SCOM from Server $($Server) $($ServerIP)`",`"message_key`":`"SCOM-$($AlertFullName)-Alert`"}"

# This caused 401's with $info missing -compress
$body = @{"source"="SCOM-CONUS";"node"="$($Hostname)";"type"="Alert";"severity"="$($EventSeverity)";"description"="$($SNOWAlertName)";"resource"="$($AlertCategory)";"metric_name"="$($SNOWAlertName)";"additional_info"="$($info)";"event_class"="SCOM from Server $($Server) $($ServerIP)";"message_key"="SCOM-$($AlertFullName)-Alert"} | ConvertTo-Json

$body = "{`"source`":`"SCOM-CONUS`",`"node`":`"$($Hostname)`",`"type`":`"Alert`",`"severity`":`"$($EventSeverity)`",`"description`":`"$($SNOWAlertName)`",`"resource`":`"$($AlertCategory)`",`"metric_name`":`"$($SNOWAlertName)`",`"additional_info`":$($info),`"event_class`":`"SCOM from Server $($Server) $($ServerIP)`",`"message_key`":`"SCOM-$($AlertFullName)-Alert`"}"

Back to original with JSON

# Previous array converted
# $body = @{"source"="SCOM-CONUS";"node"="$($Hostname)";"type"="Alert";"severity"="$($EventSeverity)";"description"="$($SNOWAlertName)";"resource"="$($AlertCategory)";"metric_name"="$($SNOWAlertName)";"additional_info"="$($info)";"event_class"="SCOM from Server $($Server) $($ServerIP)";"message_key"="SCOM-$($AlertFullName)-Alert"} | ConvertTo-Json
# 23 July
# $body = "{`"source`":`"SCOM-CONUS`",`"node`":`"$($HostFinal)`",`"type`":`"Alert`",`"severity`":`"$($EventSeverity)`",`"description`":`"$($SNOWAlertName)`",`"resource`":`"$($AlertCategory)`",`"metric_name`":`"$($SNOWAlertName)`",`"additional_info`":$($info),`"event_class`":`"SCOM from Server $($HostFinal) $($ServerIP)`",`"message_key`":`"SCOM-$($AlertFullName)-Alert`"}"

$body = @{"source"="SCOM-CONUS";"node"="$($Hostname)";"type"="Alert";"severity"="$($EventSeverity)";"description"="$($SNOWAlertName)";"resource"="$($AlertCategory)";"metric_name"="$($SNOWAlertName)";"additional_info"="$($info)";"event_class"="SCOM from Server $($Server) $($ServerIP)";"message_key"="SCOM-$($AlertFullName)-Alert"} | ConvertTo-Json

#>


$body = @{"source"="SCOM-CONUS";"node"="$($Hostname)";"type"="Alert";"severity"="$($EventSeverity)";"description"="$($SNOWAlertName)";"resource"="$($AlertCategory)";"metric_name"="$($SNOWAlertName)";"additional_info"="$($info)";"event_class"="SCOM from Server $($Server) $($ServerIP)";"message_key"="SCOM-$($AlertFullName)-Alert"} | ConvertTo-Json

write-host ""
write-host "============================================"
write-host "ServiceNow Event Body payload `n `n $($body)"
write-host "============================================"


<#
#=======================================================
# Craft auth and headers for RESTAPI

#
# NOTE * Uncomment as needed for debug write-host lines

write-host "Proxy = $Proxy"
write-host "ServiceNowURL = $ServiceNowURL"

#>

$base64Auth=[System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($ServiceNowUser):$($ServiceNowPassword)"))
write-host ""
# write-host "Base64Auth ServiceNow User = $($ServiceNowUser)`nBase64Auth ServiceNow Password = $($ServiceNowPassword)`nBase64Auth encoding string output = $($base64Auth)"
# Add-Event "Base64Auth ServiceNow User = $($ServiceNowUser)`nBase64Auth ServiceNow Password = $($ServiceNowPassword)`nBase64Auth encoding string output = $($base64Auth)"
# write-host "Base64Auth ServiceNow User = $($ServiceNowUser)`nBase64Auth encoding string output = $($base64Auth)"
# Add-Event "Base64Auth ServiceNow User = $($ServiceNowUser)`nBase64Auth encoding string output = $($base64Auth)"

$headers = @{"Content-Type"="application/json"; "Authorization"="Basic $base64Auth"}

# Include $proxy if specified for Invoke-RestMethod
# $Proxy

if ( $Null -eq $Proxy )
	{
	try
		{
		$result = Invoke-RestMethod -method "POST" -uri $ServiceNowURL -headers $headers -body $body
		#$result | ConvertTo-Json
		#$Result | Format-List -property *
		}
	catch
		{
		$errorMessage = $_.Exception.Message
			if ($_.Exception.InnerException)
			{
			$errorMessage = $_.Exception.InnerException.Message
			}
       
		# Test Incident returned
		Write-Host ""
		if ( $result -eq $null ) { write-host "Null catch response"}
		
		Write-Host -foreground red "Error: REST Response Error Message: $errorMessage"
		Write-Host ""
		Add-Event "Error: REST Response Error Message: $errorMessage"
		Add-SCOMAlertFields
		}
	}
	
if ( $Null -ne $Proxy )
	{
	try
		{
		$result=Invoke-RestMethod -method "POST" -uri $ServiceNowURL -headers $headers -body $body -Proxy $Proxy
		#$result  | ConvertTo-Json;
		}
	catch
		{
		$errorMessage = $_.Exception.Message
			if ($_.Exception.InnerException)
			{
			$errorMessage = $_.Exception.InnerException.Message

			$errormsg = $_.ToString()
			$exception = $_.Exception
			$stacktrace = $_.ScriptStackTrace
			$failingline = $_.InvocationInfo.Line
			$positionmsg = $_.InvocationInfo.PositionMessage
			$pscommandpath = $_.InvocationInfo.PSCommandPath
			$failinglinenumber = $_.InvocationInfo.ScriptLineNumber
			write-host -foreground red "Errormsg $errormsg `n Exception $exception `n Scriptname $scriptname `n Failinglinenumber $failinglinenumber `n Failingline $failingline `n PSCommandPath $pscommandpath `n Positionmsg $pscommandpath `n Stacktrace $stacktrace"
			write-host ""
			Add-Event "Errormsg $errormsg `n Exception $exception `n Scriptname $scriptname `n Failinglinenumber $failinglinenumber `n Failingline $failingline `n PSCommandPath $pscommandpath `n Positionmsg $pscommandpath `n Stacktrace $stacktrace"
			}
       
		# Test Incident returned
		Write-Host ""
		if ( $result -eq $null ) { write-host "Null catch response"}
		
		Write-Host -foreground red "Error: REST Response Error Message: $errorMessage"
		Write-Host ""
		Add-Event "Error: REST Response Error Message: $errorMessage"
		Add-SCOMAlertFields
		}
	}


<#
# Additional debug to see SNow SysId for the injection
#======================================================
# Based on Incident, this was how we verified insertion, as INCxxxxxx was the value

$response.result.sys_id
write-host -f yellow "Event injection output result = $($result.result.sys_id)"
 
# If full debug required of invoke-restmedhod 
# $result | fl -property *
 
# Full object output debug
$Convert = $result | ConvertTo-Json
 
# Debug
# $Convert
# write-host -f yellow "View JSON converted result = $Convert"
 
# Parse event to see if alert created
$EventSysID = $result.result.sys_id
# Debug to hard code Sys_ID
$EventSysID
write-host -f yellow "Event SysID = $EventSysID"
 
# Add additional SNOW URL parameters
$EventURL = "##SERVICENOWURL##"
$URL = $EventURL + $EventSysID + "&sysparm_display_value=false&sysparm_exclude_reference_link=false"
write-host -f yellow "SNOW Event table URL query link = $URL"
write-host ""
 
#=======================================================
# SNOW requires time to submit event and build alert
#=======================================================

# Parse event to see if alert created
$EventSysID = $result.result.sys_id
# Debug to hard code Sys_ID
# $EventSysID
# write-host ""
 
#>

# Parse event to see if alert created
$EventSysID = $result.result.sys_id


if ( $null -eq $EventSysID )
	{
	write-host -f yellow "Event SysID = NULL"
	write-host ""

	#Log an event for script ending and total execution time.
	$EndTime = Get-Date
	$ScriptTime = ($EndTime - $StartTime).TotalSeconds
	write-host -f red "Script Completed. `n Script Runtime: ($ScriptTime) seconds."
	Add-Event "Script Completed. `n Script Runtime: ($ScriptTime) seconds."

	exit $0
	}
	
if ( $null -ne $EventSysID )
	{
	# Add additional SNOW URL parameters
	$EventURL = "https://$SNowHostName/api/now/table/em_event?sysparm_query=sys_id="
	# write-host -f green "SNow Event URL = $($EventURL)"
	# Add-Event "SNow Event URL = $($EventURL)"
	$URL = $EventURL + $EventSysID + "&sysparm_display_value=false&sysparm_exclude_reference_link=false"
	# write-host -f green "SNow Full Event URL with SysID = $($URL)"
	# Add-Event "SNow Full Event URL with SysID = $($URL)"
	}

<#
#
# NOTE * Uncomment as needed for debug write-host lines
write-host -f yellow "SNOW Event table URL query link = $URL"
write-host ""
#>
 
if ( $null -eq $Proxy )
	{
	try
		{
		$EventCreated = Invoke-RestMethod -method "GET" -uri $URL -headers $Headers
		}
	catch
		{
		$errorMessage = $_.Exception.Message
			if ($_.Exception.InnerException)
			{
			$errorMessage = $_.Exception.InnerException.Message
			}

		# Test Incident returned
		Write-Host ""
		if ( $EventCreated -eq $null ) { write-host "EventCreated REST NULL catch response"}
		
		Write-Host -foreground red "EventCreated Error: REST Response Error Message: $errorMessage"
		Write-Host ""
		Add-Event "EventCreated Error: REST Response Error Message: $errorMessage"
		}
	}
if ( $null -ne $Proxy )
	{
	try
		{
		$EventCreated = Invoke-RestMethod -method "GET" -uri $URL -headers $Headers -Proxy $Proxy
		}
	catch
		{
		$errorMessage = $_.Exception.Message
			if ($_.Exception.InnerException)
			{
			$errorMessage = $_.Exception.InnerException.Message
			}

		# Test Incident returned
		Write-Host ""
		if ( $EventCreated -eq $null ) { write-host "EventCreated REST Null catch response"}
		
		Write-Host -foreground red "EventCreated Error: REST Response Error Message: $errorMessage"
		Write-Host ""
		Add-Event "EventCreated Error: REST Response Error Message: $errorMessage"
		}
	}


if ( $EventSeverity -ne "Clear" )
{
# Test part one - SNOW Event table - check for alert table link
#=======================================================
if ( $Null -eq $EventCreated.result.alert.link ) 
	{
	$step = 0
	Do
		{
		if ( $null -eq $Proxy )
			{
			$EventCreated = Invoke-RestMethod -method "GET" -uri $URL -headers $Headers
			}
		if ( $null -ne $Proxy )
			{
			$EventCreated = Invoke-RestMethod -method "GET" -uri $URL -headers $Headers -Proxy $Proxy
			}
		$EventCreated = Invoke-RestMethod -method "GET" -uri $URL -headers $Headers
		Start-Sleep -Seconds 10
		write-host -f red "SNOW Event table query link NOT created"
		$step++
	 	}
	Until (($null -ne ($EventCreated.result.alert.link)) -or ($step -gt 6))
	write-host ""
	write-host -f yellow "SNOW Alert table URL link = $($EventCreated.result.alert.link)"
	write-host ""
	Add-Event "SNOW Alert table URL link = $($EventCreated.result.alert.link)"
	}

if ( $Null -eq $EventCreated.result.alert.link ) 
	{
	write-host -f yellow "SNOW Alert table query link NOT created after 30 seconds"
	write-host ""
	Add-Event "SNOW Alert table query link NOT created after 30 seconds"
	}


# Test part two - SNOW Alert table - check for link, Alert#
#=======================================================
$AlertURL = $EventCreated.result.alert.link

if ( $Null -ne $AlertURL )
	{
	write-host -f yellow "SNOW Alert table URL query link = $AlertURL"
	write-host ""
	}

if ( $Null -eq $AlertURL )
	{
	if ( $null -eq $Proxy )
		{
		$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers
		}
	if ( $null -ne $Proxy )
		{
		$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers -Proxy $Proxy
		}
	write-host -f red "NO SNOW Alert table URL query link after 30+ seconds"
	Add-Event "NO SNOW Alert table URL query link after 30+ seconds"

	if ( $Null -ne $AlertURL )
		{
		$step = 0
		Do
			{
			if ( $null -eq $Proxy )
				{
				$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers
				}
			if ( $null -ne $Proxy )
				{
				$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers -Proxy $Proxy
				}

			Start-Sleep -Seconds 10
			write-host -f red "SNOW Alert table query link NOT created"
			$step++
		 	}
		Until ( ($null -ne ($AlertURL) ) -or ($step -gt 6) )
		write-host ""
		write-host -f yellow "SNOW Alert table URL query link = $AlertURL"
		write-host ""
		}
	if ( $Null -ne $($AlertCreated.result.number) )
		{
		write-host -f green "SNOW Alert number = $($AlertCreated.result.number)"
		write-host ""
		}
	if ( $null -eq $($AlertCreated.result.number) )
		{
		if ( $null -eq $Proxy )
			{
			$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers
			}
		if ( $null -ne $Proxy )
			{
			$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers -Proxy $Proxy
			}
		write-host -f red "NO SNOW Alert table URL query link after 30+ seconds"
		Add-Event "NO SNOW Alert table URL query link after 30+ seconds"
		}
	}

if ( $Null -ne $($AlertCreated.result.number) )
	{
	write-host -f yellow "SNOW Alert number = $($AlertCreated.result.number)"
	write-host ""
	# Construct ServiceNow Alert public URL
	# Found this was the error:  $AlertPublicSysID = $AlertCreated.result.SysID
	$AlertPublicSysID = $AlertCreated.result.sys_id
	write-host "Alert Public SysID = $($AlertPublicSysID)"
	write-host ""
	$AlertPublicSysURL = "https://$SNowHostName/now/nav/ui/classic/params/target/em_alert.do%3Fsys_id%" + $AlertPublicSysID
	write-host "Alert Public URL = $($AlertPublicSysURL)"
	write-host ""
	Add-Event "Alert Public URL = $($AlertPublicSysURL)"
	}

if ( $Null -eq $($AlertCreated.result.number) )
	{
	if ( $null -eq $Proxy )
		{
		$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers
		}
	if ( $null -ne $Proxy )
		{
		$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers -Proxy $Proxy
		}

	if ( $null -ne $AlertURL )
		{
		$step = 0
		Do
			{
			if ( $null -eq $Proxy )
				{
				$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers
				}
			if ( $null -ne $Proxy )
				{
				$AlertCreated = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers -Proxy $Proxy
				}
			Start-Sleep -Seconds 10
			write-host -f red "SNOW Alert table query link NOT created"
			$step++
		 	}
		Until ( ($null -ne ($AlertURL) ) -or ($step -gt 6) )
		write-host ""
		write-host -f yellow "SNOW Alert table URL query link = $AlertURL"
		write-host ""
		# Construct ServiceNow Alert public URL
		# Found this was the error:  $AlertPublicSysID = $AlertCreated.result.SysID
		$AlertPublicSysID = $AlertCreated.result.sys_id
		write-host "Alert Public SysID = $($AlertPublicSysID)"
		write-host ""
		$AlertPublicSysURL = "https://$SNowHostName/now/nav/ui/classic/params/target/em_alert.do%3Fsys_id%" + $AlertPublicSysID
		write-host "Alert Public URL = $($AlertPublicSysURL)"
		write-host ""
		Add-Event "Alert Public URL = $($AlertPublicSysURL)"
		}
	}

if ( $Null -eq $AlertCreated.result.number )
	{
	write-host -f red "SNOW Alert number NOT yet processed in 30+ seconds, processed as null"
	write-host ""
	Add-Event "SNOW Alert number NOT yet processed in 30+ seconds, processed as null"
	}



# Test part three - SNOW Incident table - check for link, INC#
#=======================================================
$IncidentURL = $AlertCreated.result.incident.link

Add-Event "AlertCreated Result output = $AlertCreated.result"
Add-Event "AlertCreated Result Incident = $AlertCreated.result.incident"
Add-Event "AlertCreated Result Incident link = $IncidentURL"


if ( $Null -ne $IncidentURL )
	{
	write-host -f yellow "SNOW Incident table URL query link = $IncidentURL"
	write-host ""
	# Construct ServiceNow Incident public URL
	# This failed - $IncidentPublicSysID = $IncidentURL.result.sys_id
	$IncidentPublicSysID = $IncidentCreated.result.sys_id
	write-host "Incident Public SysID = $($IncidentPublicSysID)"
	write-host ""
	Add-Event "Incident Public SysID = $($IncidentPublicSysID)"
	
	$IncidentPublicSysURL = "https://$SNowHostName/now/cwf/agent/record/incident/" + $IncidentPublicSysID
	write-host "First check for Incident creation - Incident Public URL = $($IncidentPublicSysURL)"
	write-host ""
	Add-Event "First check for Incident creation - Incident Public URL = $($IncidentPublicSysURL)"
	}

if ( $Null -eq $IncidentURL )
	{
	write-host -f red "Retrying SNOW Alert table URL query as no Incident link exists"

	if ( $null -eq $Proxy )
		{
		$AlertURLRetry = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers
		}
	if ( $null -ne $Proxy )
		{
		$AlertURLRetry = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers -Proxy $Proxy
		}

	$IncidentURL = $AlertURLRetry.result.incident.link
	write-host ""
	write-host -f red "SNOW Alert table link NOT created in AlertURL"

	if ( $null -eq $Proxy )
		{
		$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers
		}
	if ( $null -ne $Proxy )
		{
		$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers -Proxy $Proxy
		}

	$step = 0
	Do
		{
		if ( $null -eq $Proxy )
			{
			$AlertURLRetry = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers
			}
		if ( $null -ne $Proxy )
			{
			$AlertURLRetry = Invoke-RestMethod -method "GET" -uri $AlertURL -headers $Headers -Proxy $Proxy
			}
		Start-Sleep -Seconds 10
		write-host ""
		write-host -f red "SNOW Alert table link NOT created in AlertURL"
		$step++
	 	}
	Until ( ($Null -ne ($IncidentURL) ) -or ($step -gt 6) )

	write-host ""
	write-host -f yellow "Retried 60 seconds - SNOW Alert table URL Incident link = $($AlertURLRetry.result.incident.link)"
	write-host ""

	if ( $Null -ne $($IncidentCreated.result.number) )
		{
		write-host -f green "SCOM alert created SNOW Incident = $($IncidentCreated.result.number)"
		write-host ""

		# Construct ServiceNow Incident public URL
		# This failed - $IncidentPublicSysID = $IncidentURL.result.sys_id
		$IncidentPublicSysID = $IncidentCreated.result.sys_id
		write-host "Incident Public SysID = $($IncidentPublicSysID)"
		write-host ""
		$IncidentPublicSysURL = "https://$SNowHostName/now/cwf/agent/record/incident/" + $IncidentPublicSysID
		write-host "Second Check for incident creation - Incident Public URL = $($IncidentPublicSysURL)"
		write-host ""
		Add-Event "Second Check for incident creation - Incident Public URL = $($IncidentPublicSysURL)"
		}

	if ( $Null -eq $($IncidentCreated.result.number) )
		{
		write-host -f red "Retried 2 minutes - SNOW Incident number NOT yet processed in 2+ minutes, processed as null"
		write-host ""
		$step = 0
		Do
			{
			if ( $null -eq $Proxy )
				{
				$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers
				}
			if ( $null -ne $Proxy )
				{
				$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers -Proxy $Proxy
				}
			Start-Sleep -Seconds 10
			write-host -f red "2nd retry set - SNOW Alert table URL Incident link NOT created"
			$step++
		 	}
		Until ( ($Null -ne ($IncidentURL) ) -or ($step -gt 6) )
		write-host ""
		write-host -f yellow "2nd retry set, 30 seconds more - SNOW Alert table URL Incident link = $($IncidentCreated.result.alert.link)"
		write-host ""
		}
	}

if ( $Null -eq $($IncidentCreated.result.number) )
	{
	write-host ""
	write-host -f yellow "3rd retry set - nothing after 3 minutes, SCOM alert created SNOW Incident = $($IncidentCreated.result.number)"
	write-host ""

	if ( $null -eq $Proxy )
		{
		$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers
		}
	if ( $null -ne $Proxy )
		{
		$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers -Proxy $Proxy
		}
	write-host -f yellow "SNOW Incident result number URL REST query initiated"

	if ( $Null -ne $IncidentURL )
		{
		$step = 0
		Do
			{
			if ( $null -eq $Proxy )
				{
				$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers
				}
			if ( $null -ne $Proxy )
				{
				$IncidentCreated = Invoke-RestMethod -method "GET" -uri $IncidentURL -headers $Headers -Proxy $Proxy
				}
			Start-Sleep -Seconds 10
			write-host -f red "SNOW Alert table query link NOT created"
			$step++
		 	}
		Until ( ($null -ne ($IncidentURL) ) -or ($step -gt 6) )
		write-host ""
		write-host -f green "SNOW Incident result number = $($IncidentCreated.result.number)"

		# Construct ServiceNow Incident public URL
		# This failed - $IncidentPublicSysID = $IncidentURL.result.sys_id
		$IncidentPublicSysID = $IncidentCreated.result.sys_id
		write-host "Incident Public SysID = $($IncidentPublicSysID)"
		write-host ""
		$IncidentPublicSysURL = "https://$SNowHostName/now/cwf/agent/record/incident/" + $IncidentPublicSysID
		write-host "Third check for Incident Creation - Incident Public URL = $($IncidentPublicSysURL)"
		write-host ""
		Add-Event "Third check for Incident creation - Incident Public URL = $($IncidentPublicSysURL)"

		if ( $null -eq $($IncidentCreated.result.number) )
			{
			write-host -f green "3rd retry set, still no incident after 2 more minutes"
			Add-Event "3rd retry set, still no incident after 2 more minutes"
			}
		}
	}
}
 
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
 
# Change resolution state of alert to acknowledged (249)
$ResolutionState = 249
# Change resolution state of alert to 'Assigned to Engineering" (248)
$ResolutionState = 248
 
#>
 
if ( $Null -ne $($IncidentCreated.result.number) )
	{
	$TicketID = $IncidentCreated.result.number
	}
if ( $Null -ne $($AlertCreated.result.number) )
	{
	$TicketID = $AlertCreated.result.number
	}

$ResolutionState = 249

$IncidentPublicSysURL = "https://$SNowHostName/now/cwf/agent/record/incident/" + $IncidentPublicSysID
write-host "Incident Creation - Incident Public URL = $($IncidentPublicSysURL)"
write-host ""
Add-Event "Incident creation - Incident Public URL = $($IncidentPublicSysURL)"

Add-SCOMAlertFields
 
#============================================================
$Result = "GOOD"
$Message = "Completed ServiceNow Event creation for ($date)"

#Write-Host -f green "Completed ServiceNow Event creation for ($date) `n $body"
Add-Event "Completed ServiceNow Event creation for ($date) `n $body"

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
#Get-SNowParameters
New-SNowEvent



# End of script section
#=================================================================================
#Log an event for script ending and total execution time.
$EndTime = Get-Date
$ScriptTime = ($EndTime - $StartTime).TotalSeconds
Add-Event "New-SNowEvent Script Completed. `n Script Runtime: ($ScriptTime) seconds."
#=================================================================================
# End of script