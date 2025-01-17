<# 
#================================================================
# ServiceNow changing SNOW password script for SCOM alert integration
#================================================================

#================================================================
Deletes existing credential, then creates ServiceNow (SNow) credential on SCOM management servers.
Stored credential is under local system account.  
Requires SysInternals PSEXEC to store credentials under local system.

Has parameters to specify SNOW ServiceNow environment (test/prod).

Initiates SNow event script with SCOM admin specific alert, to create an incident in the ServiceNow environment.
 
#================================================================

# Test session is running as local system

#================================================================
# Replace the account name(s) for Test or Prod
#
# Replace ServiceNowCredential and ServiceNowProdCredential

Test example
	New-StoredCredential -Target "ServiceNowCredential" -UserName "ServiceNowProdCredential" -Password $Password -Persist "LocalMachine"
Prod example
	New-StoredCredential -Target "ServiceNowProdCredential" -UserName "ServiceNowProdCredential" -Password $Password -Persist "LocalMachine"

Replace ##ServiceNowAssignmentGroup## with organization/team to take action on password

#================================================================

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
     [String]$Password
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
$ScriptName = "New-SNowPassword.ps1"

if ( $SNowENV -eq "Test" ) { $EventID = "713" }
if ( $SNowENV -eq "Prod" ) { $EventID = "714" }


# Create new object for MOMScript API, or SCOM alert properties
$momapi = New-Object -comObject MOM.ScriptAPI

# Begin logging script starting into event log
# write-host "Script is starting. `n Running as ($whoami)."
$momapi.LogScriptEvent($ScriptName,$EventID,0,"$($ScriptName) Script is starting. `n Running as ($whoami).")
#=================================================================================

# PropertyBag Script section - Monitoring scripts get this
#=================================================================================
# Load SCOM PropertyBag function
$bag = $momapi.CreatePropertyBag()

$date = get-date -uFormat "%Y-%m-%d"


<# 
#================================================================
# AESMP changing SNOW password
#================================================================

#================================================================
# Create Test credential on SCOM mgmt servers
# 
# Open PSEXEC -S to create PowerShell session as local system
#================================================================

# Test session is running as local system
#>
# whoami
if ( ( whoami | where { $_ -like "NT Service\system" } ).Count -gt 0 ) { write-host "Running as Local System" }
if ( ( whoami | where { $_ -like "NT Service\system" } ).Count -eq 0 ) { write-host "NOT Running as Local System" ; exit }


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


# Gather SCOM alert for testing incident with new password

get-scomalert -Name "Tool Test Alert*" | Set-SCOMAlert -ResolutionState 0 -Comment "ServiceNow $($SNowENV) SCOM password automation test, changing resolutionState to new (0)"

# Setup alerts for incident
$Alerts = get-scomalert -Name "Tool Test Alert*" -ResolutionState (0..254)
$AlertID = $Alert[0].ID
$AlertName = $Alert[0].Name


# Create Test credential
#================================================================

if ( $SNowENV -eq "Test" )
	{
	Remove-StoredCredential -Target "ServiceNowCredential"

	New-StoredCredential -Target "ServiceNowCredential" -UserName "ServiceNowCredential" -Password $Password -Persist "LocalMachine"

	Get-StoredCredential -Target "ServiceNowCredential"

	# Check credential matches
	$Credential = Get-StoredCredential -Target "ServiceNowCredential"
	$ServiceNowUser = $Credential.Username
	$ServiceNowPassword = $Credential.GetNetworkCredential().Password
	$ServiceNowUser ; $ServiceNowPassword
	
	# Log details
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"Credential = $($Credential)")
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNow User = $($ServiceNowUser)")
	#$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNow password = $($ServiceNowPassword)")
	
	# Verify encoding
	$base64Auth=[System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($ServiceNowUser):$($ServiceNowPassword)"))

	Write-Host "Base64auth = $($base64Auth)"
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"Base64auth = $($base64Auth)")
	
	cd "D:\SCOMScripts\SCOMtoServiceNow"
	.\New-SNowEvent.ps1 -SNowEnv $SnowENV -AlertName $AlertName -AlertID $AlertID -AssignmentGroup "##ServiceNowAssignmentGroup##" -Team SYM

	get-scomalert -Name "Tool Test Alert*" | Set-SCOMAlert -ResolutionState 255 -Comment "ServiceNow TEST SCOM password automation test, closing alerts ResolutionState to Closed (255)"

	}


# Create Prod credential
#================================================================

if ( $SNowENV -eq "Prod" )
	{
	Remove-StoredCredential -Target "ServiceNowProdCredential"

	New-StoredCredential -Target "ServiceNowProdCredential" -UserName "ServiceNowProdCredential" -Password $Password -Persist "LocalMachine"

	Get-StoredCredential -Target "ServiceNowProdCredential"

	# Check credential matches
	$Credential = Get-StoredCredential -Target "ServiceNowCredential"
	$ServiceNowUser = $Credential.Username
	$ServiceNowPassword = $Credential.GetNetworkCredential().Password
	$ServiceNowUser ; $ServiceNowPassword

	# Log details
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"Credential = $($Credential)")
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNow User = $($ServiceNowUser)")
	#$momapi.LogScriptEvent($ScriptName,$EventID,0,"ServiceNow password = $($ServiceNowPassword)")
	
	# Verify encoding
	$base64Auth=[System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($ServiceNowUser):$($ServiceNowPassword)"))

	Write-Host "Base64auth = $($base64Auth)"
	$momapi.LogScriptEvent($ScriptName,$EventID,0,"Base64auth = $($base64Auth)")

	cd "D:\SCOMScripts\SCOMtoServiceNow"
	.\New-SNowEvent.ps1 -SNowEnv $SnowENV -AlertName $AlertName -AlertID $AlertID -AssignmentGroup "##ServiceNowAssignmentGroup##" -Team SYM

	get-scomalert -Name "Tool Test Alert*" | Set-SCOMAlert -ResolutionState 255 -Comment "ServiceNow Prod SCOM password automation test, closing alerts ResolutionState to Closed (255)"

	}


# End of script section
#=================================================================================
#Log an event for script ending and total execution time.
$EndTime = Get-Date
$ScriptTime = ($EndTime - $StartTime).TotalSeconds
Add-Event "New-SNowPassword Script Completed. `n Script Runtime: ($ScriptTime) seconds."
#=================================================================================
# End of script