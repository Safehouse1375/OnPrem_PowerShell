#Modifications: remove migrated domain checks and just remove profile.
#Create local log file of actions

#### Customer-specific variables ####

# Migrated profile name - if we create a new profile we'll name it this
$MigratedProfile = "Savannah"

#### Constant(ish) variables ####

# Regex pattern for SIDs
$PatternSID = 'S-1-12-1-\d+-\d+\-\d+\-\d+$'	#matches AAD accounts


# NEED TO COMMENT OUT THE VERSION OF OUTLOOK THAT IS NOT IN USE
# Outlook registry location (2016, 2019, O365)
#$OutlookReg = "Software\Microsoft\Office\16.0\Outlook"

# Outlook profiles registry location (2016, 2019, O365)
#$ProfilesRootReg = "Software\Microsoft\Office\16.0\Outlook\Profiles"

# Outlook profiles registry location (2013)
$ProfilesRootReg = "Software\Microsoft\Office\15.0\Outlook\Profiles"

# Outlook registry location (2013)
$OutlookReg = "Software\Microsoft\Office\15.0\Outlook"


# Name of default profile registry entry
$DefaultProfileEntry = "DefaultProfile"

# Location of registry setting for deciding whether to prompt user to select an Outlook profile
$ProfilePromptRootReg = "Software\Microsoft\Exchange\Client\Options"

# Name of registry entry for deciding whether to prompt user to select an Outlook profile
$ProfilePromptEntry = "PickLogonProfile"

# Log of events for outputting to text file.
$global:EventLog = "[" + (Get-Date -Format G) + "] Beginning profile reset. `n"

#Define log file path
$logPath = "C:\Savannah\OutlookReset\"
$logFile = "OutlookProfileResetLog_" + (Get-Date).ToString("yyy-MM-dd_HHmm") + ".txt"

If(!(Test-Path($logPath))){
	
	$message = "[" + (Get-Date -Format G) + "] First run! Creating log file folder `n"
    Write-Host $message
	$EventLog += $message
	
	New-Item $logPath -ItemType Directory -ErrorAction SilentlyContinue
}

### Function to create a new blank Outlook profile and set as default
function New-Profile
{
    param($SID, $EventLog, $ProfilesRootReg, $MigratedProfile, $OutlookReg, $DefaultProfileEntry, $ProfilePromptRootReg, $ProfilePromptEntry)

 	# Migrated profile path
    $migratedProfileReg = "registry::HKEY_USERS\" + $SID + "\" + $ProfilesRootReg + "\" + $MigratedProfile
    
	
    <# Check if profile path exists
    if (Test-Path $migratedProfileReg)
    {
        # Delete existing profile entry
		$message = "[" + (Get-Date -Format G) + "] Profile named " + $MigratedProfile + " detected! Deleting this profile `n"
        Write-Host $message
		$EventLog += $message
		Try {
			Remove-Item $migratedProfileReg -Recurse
		} Catch {
			$exceptionDetail = $_.Exception.Message
			$message = "[" + (Get-Date -Format G) + "] Error Deleting existing profile: " + $exceptionDetail + " `n"
			$EventLog += $message
		}
    }#>

    # Create new profile
	$message = "[" + (Get-Date -Format G) + "] Creating profile: " + $MigratedProfile + " at path " + $migratedProfileReg + " `n"
	Write-Host $message
	$EventLog += $message
	Try {
		New-Item $migratedProfileReg | Out-Null
	} Catch {
		$exceptionDetail = $_.Exception.Message
		$message = "[" + (Get-Date -Format G) + "] Error creating profile: " + $exceptionDetail + " `n"
		$EventLog += $message
	}

    # Set migrated profile as default
    $userOutlookReg = "registry::HKEY_USERS\" + $SID + "\" + $OutlookReg
    $defaultProfile = Get-ItemPropertyValue -Path $userOutlookReg -Name $DefaultProfileEntry -ErrorAction SilentlyContinue
    
	$message = "[" + (Get-Date -Format G) + "] Setting " + $MigratedProfile + " as default profile `n"
	Write-Host $message
	$EventLog += $message
		
    if ($null -eq $defaultProfile)
    {
		Try {
			New-ItemProperty -Path $userOutlookReg -Name $DefaultProfileEntry -Value $MigratedProfile -PropertyType String | Out-Null
		} Catch {
			$exceptionDetail = $_.Exception.Message
			$message = "[" + (Get-Date -Format G) + "] Error setting default profile: " + $exceptionDetail + " `n"
			$EventLog += $message
		}
    } else{	
		Try {
			Set-ItemProperty -Path $userOutlookReg -Name $DefaultProfileEntry -Value $MigratedProfile
			$message = "[" + (Get-Date -Format G) + "] Replaced default profile at " + $userOutlookReg + "\" + $DefaultProfileEntry + " with value: " + $MigratedProfile + "`n"
			$EventLog += $message
		} Catch {
			$exceptionDetail = $_.Exception.Message
			$message = "[" + (Get-Date -Format G) + "] Error relacing default profile: " + $exceptionDetail + " `n"
			$EventLog += $message
		}
    }

    # Configure Outlook to always open the default profile
    $userProfilePromptReg = "registry::HKEY_USERS\" + $SID + "\" + $ProfilePromptRootReg
    $prompt = Get-ItemPropertyValue -Path $userProfilePromptReg -Name $ProfilePromptEntry -ErrorAction SilentlyContinue
	
	$message = "[" + (Get-Date -Format G) + "] Configuring Outlook to load " + $MigratedProfile + " profile automatically `n"
	Write-Host $message
	$EventLog += $message
	
    if ($null -eq $prompt)
    {
		Try {
			New-Item $userProfilePromptReg -Force | Out-Null
			New-ItemProperty -Path $userProfilePromptReg -Name $ProfilePromptEntry -Value "0" -PropertyType String | Out-Null
		
			$message = "[" + (Get-Date -Format G) + "] Running New-ItemProperty cmdlet `n"
			Write-Host $message
			$EventLog += $message
		} Catch {
			$exceptionDetail = $_.Exception.Message
			$message = "[" + (Get-Date -Format G) + "] Error setting Outlook to open default profile: " + $exceptionDetail + " `n"
			$EventLog += $message
		}
    }
    else
    {
		Try {
			Set-ItemProperty -Path $userProfilePromptReg -Name $ProfilePromptEntry -Value "0"
		
			$message = "[" + (Get-Date -Format G) + "] Running Set-ItemProperty cmdlet `n"
			Write-Host $message
			$EventLog += $message
		} Catch {
			$exceptionDetail = $_.Exception.Message
			$message = "[" + (Get-Date -Format G) + "] Error setting Outlook to open default profile: " + $exceptionDetail + " `n"
			$EventLog += $message
		}
    }
	
	return $EventLog
}



### Function to process individual Outlook profiles for a single user
function Process-User
{
    param($SID, $EventLog, $ProfilesRootReg, $MigratedProfile, $OutlookReg, $DefaultProfileEntry, $ProfilePromptRootReg, $ProfilePromptEntry)

    # Check we can find Outlook registry keys to make sure it's the correct version of Outlook
    $userOutlookReg = "registry::HKEY_USERS\" + $SID + "\" + $OutlookReg
	$userProfilePath = "registry::HKEY_USERS\" + $SID + "\" + $ProfilesRootReg #Sometimes the Outlookreg key will exist but the "Profiles" path won't. This accounts for both
	
    if (!(Test-Path $userOutlookReg) -Or !(Test-Path $userProfilePath))
    {
        Write-Host Unable to find Outlook registry settings at $userOutlookReg
        Write-Host Outlook may not be installed, it may be the wrong version, or this user might not have used it yet.
		
		$message = "[" + (Get-Date -Format G) + "] Unable to find Outlook registry settings at " + $userOutlookReg + " `n"
		Write-Host $message
		$EventLog += $message
		
		$message = "[" + (Get-Date -Format G) + "] Outlook may not be installed, it may be the wrong version, or this user might not have used it yet `n"
		Write-Host $message
		$EventLog += $message
		
        return $EventLog
    } else {

		# Check for default profile
		try
		{
			$userOutlookReg = "registry::HKEY_USERS\" + $SID + "\" + $OutlookReg
			$defaultProfile = Get-ItemPropertyValue -Path $userOutlookReg -Name $DefaultProfileEntry -ErrorAction SilentlyContinue
			$defaultProfileReg = "registry::HKEY_USERS\" + $SID + "\" + $ProfilesRootReg + "\" + $defaultProfile
			
		}
		catch
		{
			$defaultProfile = $null
		}

		
		# Check the if default profile exists and matches the migrated name
		if ($null -ne $defaultProfile -and $defaultProfile -eq $MigratedProfile -and (Test-Path $defaultProfileReg))
		{
			$message = "[" + (Get-Date -Format G) + "] Found " + $MigratedProfile + " profile. No need to create a new one.`n"
			Write-Host $message
			$EventLog += $message
			
			return $EventLog
			
		}else{ # Our migrated profile is not default, so we'll assume they need a new one
			Try {
				$EventLog = New-Profile -SID $SID -EventLog $EventLog -ProfilesRootReg $ProfilesRootReg -MigratedProfile $MigratedProfile -OutlookReg $OutlookReg -DefaultProfileEntry $DefaultProfileEntry -ProfilePromptRootReg $ProfilePromptRootReg -ProfilePromptEntry $ProfilePromptEntry
			} Catch {
				$exceptionDetail = $_.Exception.Message
				$message = "[" + (Get-Date -Format G) + "] Error running New-Profile function: " + $exceptionDetail + " `n"
				$EventLog += $message
			}
			
				return $EventLog
		}
	}
}



### Main part of script starts -  find and iterate over each user registry hive

 
# Get Username, SID, and location of ntuser.dat for all users
$profileList = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object {$_.PSChildName -match $PatternSID} | 
    Select-Object  @{name="SID";expression={$_.PSChildName}}, 
            	   @{name="UserHive";expression={"$($_.ProfileImagePath)\ntuser.dat"}}, 
            	   @{name="Username";expression={$_.ProfileImagePath -replace '^(.*[\\\/])', ''}}

# Get all user SIDs found in HKEY_USERS (ntuder.dat files that are loaded)
$loadedHives = Get-ChildItem Registry::HKEY_USERS | Where-Object {$_.PSChildname -match $PatternSID} | Select-Object @{name="SID";expression={$_.PSChildName}}

# If we're running unattended as SYSTEM and nobody's logged on, there may be no loaded registry hives
if ($null -eq $loadedHives)
{
    $unloadedHives = $profileList
}else{
    # Get all users that are not currently logged on
    $unloadedHives = Compare-Object $profileList.SID $loadedHives.SID | Select-Object @{name="SID";expression={$_.InputObject}}
}

$message = "[" + (Get-Date -Format G) + "] Found " + $profileList.Count + " user profiles on machine. Looping through them now.`n"
Write-Host $message
$EventLog += $message

# Loop through each profile on the machine
foreach ($user in $profileList)
{
	$message = "[" + (Get-Date -Format G) + "] Start user ----------" + $user.Username + "---------- `n"
	Write-Host $message
	$EventLog += $message
	
    # Load User ntuser.dat if it's not already loaded
    if ($user.SID -in $unloadedHives.SID)
    {	
		$message = "[" + (Get-Date -Format G) + "] Loading Registry Hive " + $user.UserHive + " `n"
		Write-Host $message
		$EventLog += $message
      
        reg load HKU\$($user.SID) $($user.UserHive) | Out-Null
    }


    # Process this user
    $EventLog = Process-User -SID $user.SID -EventLog $EventLog -ProfilesRootReg $ProfilesRootReg -MigratedProfile $MigratedProfile -OutlookReg $OutlookReg -DefaultProfileEntry $DefaultProfileEntry -ProfilePromptRootReg $ProfilePromptRootReg -ProfilePromptEntry $ProfilePromptEntry
	
    # Unload ntuser.dat    
    if ($user.SID -in $unloadedHives.SID)
    {
        ### Garbage collection and closing of ntuser.dat ###
		$message = "[" + (Get-Date -Format G) + "] Closing Registry Hive " + $user.UserHive + " `n"
		Write-Host $message
		$EventLog += $message
        
		[gc]::Collect()      
        reg unload HKU\$($user.SID) | Out-Null
    }
	$message = "[" + (Get-Date -Format G) + "] Finished User ----------" + $user.Username + "---------- `n"
	Write-Host $message
	$EventLog += $message
}

$EventLog | Out-File "$($logPath)$($logFile)"
