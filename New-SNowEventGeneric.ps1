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
		 Position=2)]
		 [ValidateNotNullorEmpty()]
		 [String]$AssignmentGroup,
     [Parameter(
	 	 Mandatory=$true,
         ValueFromPipeline=$true,
	 	 Position=3)]
		 [ValidateNotNullorEmpty()]
		 [String]$Team
)


# Hard code URL, Proxy, CallerID SNOW variables into script
#===============================================

# Don't forget to replace SNOW DEV URL with Prod before going to production events
$ServiceNowURL="https://##CUSTOMER##.servicenowservices.com/api/now/table/em_event" 
$ServiceNowURL="##ServiceNowURL##"
# OAUTH URL /oauth_token.do
$Proxy = "##PROXY##"

# Assume module NOT loaded into current PowerShell profile
Import-Module -Name CredentialManager



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
# Retrieve SNOW credential from Credential Manager
#===============================================
# Example
# $Credential = Get-StoredCredential -Target "SNOW_Account"
#
# ID, Password, and Caller_ID are provided by SNOW team
#
# Get-StoredCredential -Target "ServiceNowCredential"
# Example
# $Credential = Get-StoredCredential -Target "SNOW_Account"
# $Credential = Get-StoredCredential -Target "ServiceNowCredential"
# $Credential = Get-StoredCredential -Target "svc_rest_scom"
# ID, Password, and Caller_ID are provided by ##CUSTOMER## team
#>


$Credential = Get-StoredCredential -Target "SNOW-TEST-CRED"

$ServiceNowUser = $Credential.Username
$ServiceNowPassword = $Credential.GetNetworkCredential().Password

if ( $Null -eq $Credential )
	{
	write-host -f red "ServiceNow Credential NOT stored on server"
	write-host ""
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNow Credential NOT stored on server")
	}



<#
# Assuming No changes, inputs passed to SCOM channel for SNOW event creation

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


# write-host -f green "End Global section"
#Add-Event "End Global section"
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
#>

# Change resolution state of alert to acknowledged (249)
$AlertResolutionState = 249

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
	write-host -f red "Exiting - ServiceNow (SNow) Event NOT created as SCOM alert closed with TicketID = $TicketID"
	Add-Event "Exiting - ServiceNow (SNow) Incident NOT created as SCOM alert closed with TicketID = $TicketID"
	exit $0
	}


<# Resolve alert?
#===========================
Get-SCOMAlert -Name "$AlertName" -ResolutionState 0 | Resolve-SCOMAlert -ticketID $TicketID `
	-Owner "$AssignmentGroup" `
	-Comment "Resolve ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert"

	Add-Event "Resolved SCOM alert $TicketID for group $AssignmentGroup"
#>

<#
#=======================================================
# Get-SCOM alert and update alert
# Did event created become an alert?
#================================
#>

# Alert section
if ( $Null -ne $($AlertCreated.result.number) )
	{
	$TicketID = $AlertCreated.result.number
	# Debug
	# $TicketID
	# Update ticket, assignment group, resolution state -AND ( $_.ResolutionState -eq 0 ) } `
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-Comment "ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert `
	Updated SCOM alert $AlertName, Server = $Hosts, with TicketID $TicketID for group $AssignmentGroup"

	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = $TicketID, for group $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 249 Acknowledged, TicketID = NO SNOW alert, for group $AssignmentGroup"
	}
if ( $null -eq $($AlertCreated.result.number) )
	{
	$TicketID = "NO SNOW alert created after 6 plus attempts over 60 seconds"
 
	# Update ticket, assignment group, resolution state -AND ( $_.ResolutionState -eq 0 ) } `
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-Comment "ServiceNow SCOM alert automation - Set SCOM Alert Ticket, Owner, Resolution state in current alert `
	Updated SCOM alert $AlertName, Server = $Hosts, with 'NO SNOW Alert' for TicketID $TicketID, for group $AssignmentGroup"

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
	| Set-SCOMAlert -ticketID $TicketID -Owner "$AssignmentGroup"  -ResolutionState $ResolutionState `
	-Comment "ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert`
	Updated SCOM alert $AlertName, Server = $Hosts, with 'NO SNOW Incident' for TicketID $TicketID, for group $AssignmentGroup"
	# ResolutionState 248 = Assigned to Engineering
	write-host ""
	write-host -f yellow "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 248 Assigned to Engineering, SNOW Incident $TicketID, for group $AssignmentGroup"
	Add-Event "Updated SCOM alert $AlertName, Server = $Hosts, Resolution State 248 Assigned to Engineering, SNOW Incident $TicketID, for group $AssignmentGroup"
	}
if ( $null -eq $($IncidentCreated.result.number) )
	{
	$TicketID = "NO SNOW INC"
	write-host -f red "NO SNOW incident created after 20+ plus attempts over 3 minutes"

	# Update ticket, assignment group, resolution state
	Get-SCOMAlert -Id $AlertID | where { ( $_.Name -eq "$AlertName" ) } `
	| Set-SCOMAlert -ticketID $TicketID -Owner "$TEAM $AssignmentGroup"  -ResolutionState $ResolutionState `
	-Comment "ServiceNow SCOM alert automation - Set Ticket, Owner, Resolution state in current alert`
	Updated SCOM alert $AlertName, Server = $Hosts, with 'NO SNOW Incident' for TicketID $TicketID, for group $AssignmentGroup"
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
#>

#write-host -f green $ServiceNowURL
#write-host -f green "ServiceNow (SNow) URL specified = $($ServiceNowURL)"
#Add-Event "ServiceNow (SNow) URL specified = $($ServiceNowURL)"

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
	Add-Event "Exiting - NO ServiceNow URL specified"
	exit $0
	}


<#
# Pre-req for CredentialManager powershell (posh) module
# Assume module NOT loaded into current PowerShell profile

* Uncomment as needed for debug write-host lines
#>

# Verify Credential Manager snap in installed
$CredMgrModuleBase = Get-Module -Name CredentialManager
 
if ( $Null -ne $CredMgrModuleBase.ModuleBase )
	{
	write-host -f yellow "CredentialManager PoSH Module Installed, ModuleBase = $($CredMgrModuleBase.ModuleBase)"
	Add-Event "ServiceNow Credential PowerShell module installed, ModuleBase = $($CredMgrModuleBase.ModuleBase)"
	}
 
if ( $Null -eq $CredMgrModuleBase.ModuleBase )
	{
	write-host -f red "CredentialManager PoSH Module NOT Installed"
	Add-Event "ServiceNow Credential PowerShell module NOT installed"
	exit $0
	}
 
<#
# Verify SNOW credential exists in Credential Manager
#===============================================
# Example
# $Credential = Get-StoredCredential -Target "SNOW_Account"
#
# ID, Password, and Caller_ID are provided by SNOW team
 
# From Global section
# $Credential = Get-StoredCredential -Target "ServiceNowCredential"
# $ServiceNowUser = $Credential.Username
# $ServiceNowPassword = $Credential.GetNetworkCredential().Password
#>
 
if ( $null -ne $Credential )
	{
	Write-host -f green "Stored Credential variable exists"
	Write-host ""
	Add-Event "Stored Credential variable exists"
	}
 
# Test Credential variables for User password are provided
#===============================================
 
if ( $Null -eq $Credential.UserName )
	{
	write-host -f red "ServiceNow Credential NOT stored on server - credential $($Credential), username $($Credential.UserName)"
	Add-Event "ServiceNow Credential NOT stored on server - credential $($Credential), username $($Credential.UserName)"
	exit $0
	}
 
 
# Test Credential variables for User password are provided
#===============================================
if ( $Null -eq $ServiceNowUser )
	{
	write-host -f red "ServiceNow User NOT stored on server"
	Add-Event "ServiceNow User NOT stored on server"
	}
if ( $Null -eq $ServiceNowPassword )
	{
	write-host -f red "ServiceNow Password NOT stored on server"
	Add-Event "ServiceNow Password NOT stored on server"
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
	[String]$AssignmentGroup,
	[String]$Team
#>


# Set up EventData variable with SCOM to SNow fields
#===============================================


<#
# Multiple locations for HostName in alerts based on class, path, and other variables
# Hostname
# Figure out hostname based on alert values
#===============================================
#>

$MonitoringObjectPath = ($Alert |select MonitoringObjectPath).MonitoringObjectPath
$MonitoringObjectDisplayName = ($Alert |select MonitoringObjectDisplayName).MonitoringObjectDisplayName	
$PrincipalName = ($Alert |select PrincipalName).PrincipalName
$DisplayName = ($Alert |select DisplayName).DisplayName
$PKICertPath = ($Alert |select Path).Path

# Update tests with else for PowerShell script
if ( $MonitoringObjectPath -ne $null ) { $Hostname = $MonitoringObjectPath }
if ( $MonitoringObjectDisplayName -ne $null ) { $Hostname = $MonitoringObjectDisplayName }
if ( $PrincipalName -ne $null ) { $Hostname = $PrincipalName }
if ( $DisplayName -ne $null ) { $Hostname = $DisplayName }
if ( $PKICertPath -ne $null ) { $Hostname = $PKICertPath }

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


# Combined event
write-host ""
write-host -f yellow "SCOM Alert - Hostname = $Hosts, Server IP = $ServerIP, AlertName = $AlertName, Alert ID = $AlertID"
Add-Event "SCOM Alert - Hostname = $Hosts, Server IP = $ServerIP, AlertName = $AlertName, Alert ID = $AlertID"


# Get ResolutionState
$alertResolutionState = $Alert.ResolutionState

# Evaluate alert closed before SCOM channel SNOW script executed
#===============================================

if ( $AlertResolutionState -eq 255 )
	{
	write-host "ServiceNow (SNow) Event NOT needed as SCOM alert is CLOSED"
	Add-Event "ServiceNow (SNow) Event NOT needed as SCOM alert is CLOSED"
	exit $0
	}


<# 
# Optional $Description enrichment
$AlertParameters = $Alert.Parameters
$AlertID = $Alert.ID.Guid
$AlertManagementGroup =  $Alert.ManagementGroup.Name
$AlertCategory = $Alert.Category

# Debug
$ServerIP
$MonitoringObjectPath
$MonitoringObjectDisplayName
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

# Set Alert Parameters, ID, ManagementGroup, Category from SCOM alert to use in event fields
$AlertParameters = $Alert.Parameters

# Used with manual testing $Alerts
#$AlertID = $Alert.ID.Guid

$AlertManagementGroup =  $Alert.ManagementGroup.Name
$AlertCategory = $Alert.Category

# Set SCOM Pack, Class, Object into Message_key
$AlertFullName = $Alert.MonitoringObjectFullName

# Set AlertName for Testing
$SNOWAlertName = "$Team ##COMPANY##: SCOM - $AlertName"

# Alert description
$AlertDescription = $Alert.Description

# Display AlertDescription (Debug)
# $AlertDescription

<# Determine SCOM Alert Description exclude JSON special characters
#===============================================
# NOTE * Uncomment as needed for debug write-host lines
#>

# write-host -f green "Begin SCOM Alert Description JSON audit"

$Description = $Alert.Description

$Description = $Description -replace "^@", ""
$Description = $Description -replace "=", ":"
$Description = $Description -replace ";", ","
$Description = $Description -replace "`", "#"
$Description = $Description -replace "{", "*"
$Description = $Description -replace "}", "*"
$Description = $Description -replace "\n", "\\n"
$Description = $Description -replace "\r", "\\r"
$Description = $Description -replace "\\\\n", "\\n"

<#
# write-host -f green "SCOM Alert Description formatted for JSON"
# write-host -f green $Description
# Add-Event $Description
#>

# Determine SCOM Alert Severity to ITSM tool impact
#===============================================
$Severity = $Alert.Severity
 
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

write-host ""
write-host -f yellow "SCOM Alert ID = $AlertID, Severity = $EventSeverity, ServerName = $Hosts, Server IP = $ServerIP, Alert Parameters = $AlertParameters, SCOM management group = $AlertManagementGroup, Alert Category = $AlertCategory"
Add-Event "SCOM Alert ID = $AlertID, Severity = $EventSeverity, ServerName = $Hosts, Server IP = $ServerIP, Alert Parameters = $AlertParameters, SCOM management group = $AlertManagementGroup, Alert Category = $AlertCategory"


<#
# Build SNow event payload
#=======================================================
#$info = @{Test="12345"} | ConvertTo-Json
# $info = "{ SCOM Alert Parameters: $AlertParameters , SCOM Alert ID: $AlertID }" | ConvertTo-Json

# SNOW provided event body example
#$body = @{source="SCOM";node="$Hostname";type="Alert";severity="$EventSeverity";Description="This is a SCOM alert -  $Description";resource="NIC2";metric_name="Adapter2";additional_info=$info;event_class="SCOM from $ServerIP";message_key="SCOM-$Hostname-Alert"} | ConvertTo-Json

#>

#$info = @("Alert Details: $Description","SCOM Alert Parameters: $AlertParameters","SCOM Alert ID: $AlertID","Hostname: $Server","MonitoringObjectPath: $MonitoringObjectPath","MonitoringObjectDisplayName: $MonitoringObjectDisplayName","PrincipalName: $PrincipalName","DisplayName: $DisplayName","PKICertPath: $PKICertPath" )
$info = @("SCOM Management Group: $AlertManagementGroup","Hostname: $Server","SCOM Alert Parameters: $AlertParameters","SCOM Alert ID: $AlertID","Alert Details: $Description","MonitoringObjectPath: $MonitoringObjectPath","MonitoringObjectDisplayName: $MonitoringObjectDisplayName","PrincipalName: $PrincipalName","DisplayName: $DisplayName","PKICertPath: $PKICertPath" )

write-host "ServiceNow Event payload `n `n $($info)"
Add-Event "ServiceNow Event payload `n `n $($info)"

# JSON audit
$EventInfo = $info -replace "^@", ""
$EventInfo = $info -replace "=", ":"
$EventInfo = $info -replace ";", ","
$EventInfo = $info -replace "`", "#"
$EventInfo = $info -replace "{", "*"
$EventInfo = $info -replace "}", "*"
$EventInfo = $info -replace "\n", "\\n"
$EventInfo = $info -replace "\r", "\\r"
$EventInfo = $info -replace "\\\\n", "\\n"

<#
#
# NOTE * Uncomment as needed for debug write-host lines
 
#Debug
#$EventInfo
 
write-host "MessageKey = $Hosts"
write-host "HostName = $HostName"
#>

$body = @{source="SCOM";node="$Hostname";type="Alert";severity="$EventSeverity";description="$SNOWAlertName";resource="$AlertCategory";metric_name="$SNOWAlertName";additional_info=$info;event_class="SCOM from Server $Server $ServerIP";message_key="SCOM-$AlertFullName-Alert"} | ConvertTo-Json
write-host "ServiceNow Event payload `n `n $($body)"

<#
#
# NOTE * Uncomment as needed for debug write-host lines

# Debug verify Info and Body variables
$info

write-host "ServiceNow Event payload `n `n $($body)"
Add-Event "ServiceNow Event payload `n `n $($body)"

$body
#>

<#
#=======================================================
# Craft auth and headers for RESTAPI

#
# NOTE * Uncomment as needed for debug write-host lines

write-host "Proxy = $Proxy"
write-host "ServiceNowURL = $ServiceNowURL"

#>

$base64Auth=[System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($ServiceNowUser):$($ServiceNowPassword)"))
$headers = @{"Content-Type"="application/json"; "Authorization"="Basic $base64Auth"}

#Include $proxy if specified
$Proxy

if ( $null -eq $Proxy )
	{
	$result=Invoke-RestMethod -method "POST" -uri $ServiceNowURL -headers $headers -body $body
	#$result | ConvertTo-Json;
	}
if ( $Null -ne $Proxy )
	{
	$result=Invoke-RestMethod -method "POST" -uri $ServiceNowURL -headers $headers -body $body -Proxy $Proxy
	#$result  | ConvertTo-Json;
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
write-host -f yellow "View JSON converted result = $Convert"

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

#>


#=======================================================
# SNOW requires time to submit event and build alert
#=======================================================

# Parse event to see if alert created
$EventSysID = $result.result.sys_id
# Debug to hard code Sys_ID
# $EventSysID
write-host ""
write-host -f yellow "Event SysID = $EventSysID"

# Add additional SNOW URL parameters
$EventURL = "https://##SERVICENOWURL##/api/now/table/em_event?sysparm_query=sys_id="
$URL = $EventURL + $EventSysID + "&sysparm_display_value=false&sysparm_exclude_reference_link=false"

<#
#
# NOTE * Uncomment as needed for debug write-host lines
write-host -f yellow "SNOW Event table URL query link = $URL"
write-host ""
#>

if ( $null -eq $Proxy )
	{
	$EventCreated = Invoke-RestMethod -method "GET" -uri $URL -headers $Headers
	}
if ( $null -ne $Proxy )
	{
	$EventCreated = Invoke-RestMethod -method "GET" -uri $URL -headers $Headers -Proxy $Proxy
	}



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
	}
 
if ( $Null -eq $EventCreated.result.alert.link ) 
	{
	write-host -f yellow "SNOW Alert table query link NOT created after 30 seconds"
	write-host ""
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
		}
	}

if ( $Null -ne $($AlertCreated.result.number) )
	{
	write-host -f yellow "SNOW Alert number = $($AlertCreated.result.number)"
	write-host ""
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
		}
	}
 
if ( $Null -eq $AlertCreated.result.number )
	{
	write-host -f red "SNOW Alert number NOT yet processed in 30+ seconds, processed as null"
	write-host ""
	}

 
 
# Test part three - SNOW Incident table - check for link, INC#
#=======================================================
$IncidentURL = $AlertCreated.result.incident.link
 
if ( $Null -ne $IncidentURL )
	{
	write-host -f yellow "SNOW Incident table URL query link = $IncidentURL"
	write-host ""
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
		if ( $null -eq $($IncidentCreated.result.number) )
			{
			write-host -f green "3rd retry set, still no incident after 2 more minutes"
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

if ( $null -ne $($IncidentCreated.result.number) )
	{
	$TicketID = $IncidentCreated.result.number
	}
if ( $null -ne $($AlertCreated.result.number) )
	{
	$TicketID = $AlertCreated.result.number
	}
	
$ResolutionState = 249


Add-SCOMAlertFields

#============================================================
$Result = "GOOD"
$Message = "Completed ServiceNow Event creation for ($date)"

#Write-Host -f green "Completed ServiceNow Event creation for ($date) `n $EventData"
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