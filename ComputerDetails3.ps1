#===================================================================================================
#  Get Computer Details
#  Jared K. Thomason
#  01/05/2012
#===================================================================================================

#Tells the Script to continue on errors without displaying the error
$erroractionpreference = "SilentlyContinue"

#===================================================================================================
# ************ You may want to check the following parameters before you run the script ************
# Directory for excell files to be stored.
$ExcelFileDirectory = "C:\Scripts\POWERSHELL\ComputerDetails\logs"

# Location and file name for the log file.
$logfile = C:\Scripts\POWERSHELL\ComputerDetails\logs\AMC_Servers.txt

# Canonical name of the OU you want to search through.
$ouToSearch = "faa.gov/Programs/National/Data_Centers/Servers/AMC_Servers"
#===================================================================================================

# START OF WRITE-LOG FUNCTION **********************************************************************
# This is a function used to write information to a log file.
function write-log([string]$info){
	if($loginitialized -eq $false){
		$FileHeader > $logfile
		$script:loginitialized = $True
	}
	$info >> $logfile
}
# END OF FUNCTION **********************************************************************************

# Log file Information *****************************************************************************
# ********** You may want to make sure this is the location you want to use ************************
$script:logfile = "C:\Scripts\POWERSHELL\ComputerDetails\logs\AMCComputerDetails-$(get-date -format MM-dd-yyyy-HH-mm-ss).txt"
$script:separator = @"
$("-" * 25)
"@
$script:loginitialized = $false
$script:FileHeader = @"
$separator
***Application Information***
Filename:  ComputerDetails3.ps1
Created by:  Jared
Last Modified:  $(Get-Date -Date (get-item C:\Scripts\POWERSHELL\ComputerDetails\ComputerDetails3.ps1).LastWriteTime -f MM/dd/yyyy)
"@

# START OF Get-OSWMI FUNCTION ********************************************************************
# A function used to get the OS information for the server using WMI
function Get-OSWMI ([string]$strComputer) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-OSWMI function"
	$WmiObject = "OperatingSystem"

	$WmiObject | 
		% { Set-Variable -name $_ -value (gwmi Win32_$_ -ComputerName $strComputer) }

	$strOS = ($OperatingSystem.Caption).ToString()
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-OSWMI function - $strOS"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strOS
}
# END OF FUNCTION ********************************************************************************

# START OF Get-SERVERPORV FUNCTION ***************************************************************
# Used to determine whether the server is physical or virtual based on the 'v' or 'n' in the server naming standard
function Get-ServerPorV ([string]$strComputer) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-ServerType function"
	If ($strComputer.Substring(7,1) -eq "n") {
		$strPorV = "Physical"
	}
	elseif ($strComputer.Substring(7,1) -eq "v") {
		$strPorV = "Virtual"
	}
	else {
		$strPorV = "UNKNOWN"
	}
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-ServerType function - $strPorV"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strPorV
}
# END OF FUNCTION *******************************************************************************

# START OF Get-SERVERSTATUS FUNCTION ***************************************************************
# Used to determine whether the server is physical or virtual based on the 'v' or 'n' in the server naming standard
function Get-ServerSTATUS ([string]$strComputer) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-ServerType function"
	If ($strComputer.Substring(6,1) -eq "p") {
		$strSTATUS = "Prod"
	}
	elseif ($strComputer.Substring(6,1) -eq "t") {
		$strSTATUS = "Test"
	}
	elseif ($strComputer.Substring(6,1) -eq "d") {
		$strSTATUS = "Dev"
	}
	else {
		$strSTATUS = "UNKNOWN"
	}
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-ServerType function - $strSTATUS"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strSTATUS
}
# END OF FUNCTION *******************************************************************************

# START OF GET-OS FUNCTION **********************************************************************
function get-OS ($Server) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-OS function"
	$strComputer = [string]$Server.Name
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Calling the Get-OSWMI function to get the OS of the server"
	# Call a function to get the Operating System of the server.
	$error.clear()
	$strOS = Get-OSWMI $strComputer
	if ($error) {
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Could not use the Get-OSWMI function to get the OS of the server"
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") -       There are probably issues with WMI access"
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") -       Using the OS information stored in active directory"
		$strOS = [string]$server.OSName
	}
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-OS function - $strOS"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strOS
}
# END OF FUNCTION ********************************************************************************

# START OF GET-WMIAppList FUNCTION ***************************************************************
function get-WMIAppList ([string]$strComputer) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-WMIAppList Function"
	$applist = get-wmiobject -namespace "root/cimv2" -computername $strComputer win32_product | sort-object name
	$arrAppList = @()
	foreach ($app in $applist) {
		$arrAppList += $app.name + ","
	}
	$strAppList = [string]$arrAppList
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-WMIAppList function - $strAppList"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strAppList
}
# END OF FUNCTION ********************************************************************************

# START OF GET-REGUNINSTALLLIST FUNCTION *********************************************************
function get-RegUninstallList ([string]$strComputer) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-RegUninstallList function"
	# Branch of the Registry 
	$Branch='LocalMachine' 
	# Main Sub Branch you need to open 
	$SubBranch="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
	$EncounteredError = 0
	$Error.Clear()
	$registry=[microsoft.win32.registrykey]::OpenRemoteBaseKey('Localmachine',$strComputer)
	# Error Handling
	if ($error) {
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: Could not access the registry of $strComputer *******************"
		Write-Host "Could not access the registry of $strComputer"
		$EncounteredError = 1
	}
	Else {
		$Error.Clear()
		$registrykey=$registry.OpenSubKey($Subbranch)
		if ($error) {
			Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: Could not access the registry of strComputer *******************"
			Write-Host "Could not access the registry of $strComputer"
			$EncounteredError = 1
		}
		Else {
			$Error.Clear()
			$SubKeys=$registrykey.GetSubKeyNames()
			if ($error) {
				Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: Could not access the registry of $strComputer *******************"
				Write-Host "Could not access the registry of $strComputer"
				$EncounteredError = 1
			}
		}
	}

	# Drill through each key from the list and pull out the value of “DisplayName”
	if ($EncounteredError -eq 0) {
		$installedApps = @()
		Foreach ($key in $subkeys) {
			$exactkey=$key 
			$NewSubKey=$SubBranch+"\\"+$exactkey 
			$ReadUninstall=$registry.OpenSubKey($NewSubKey) 
			$Value=$ReadUninstall.GetValue("DisplayName")
			if ($Value -ne $null) {
				if ((-not $Value.contains(" (KB")) -and (-not $Value.contains("ArcSight SmartConnector")) -and (-not $Value.contains(" - KB")) -and (-not $Value.contains(" kb")) -and (-not $Value.contains("Windows Management Framework Core"))) {
					$installedApps += [string]$Value + ","
				}
			}
		}
	}
	$orderedInstalledApps = $installedApps | sort-object
	#Ensture the list of installed applications is a string.
	$strInstalledApps = [string]$orderedInstalledApps
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-RegUninstallList function - $strInstalledApps"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strInstalledApps
}
# END OF FUNCTION *********************************************************************************

# START OF GET-INSTALLEDAPPS FUNCTION *************************************************************
function get-InstalledApps ([string]$strComputer,[string]$strOS) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-InstalledApps function to determine the applications that are installed on the server"
	if ($strOS.contains("server 2008")) {
		$strInstalledApps = get-WMIAppList $strComputer
	}
	$strInstalledApps = Get-RegUninstallList $strComputer
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-InstalledApps function"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strInstalledApps
}
# END OF FUNCTION *********************************************************************************

# START OF GET-APPPOC FUNCTION ********************************************************************
function get-AppPOC ([string]$strNotes) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-AppPOC function"
	# Look for POC information in the Active Directory Description field.
	If ($strNotes.Contains("POC ")) {
		$tmpAppPOC = $strNotes.Replace('POC ', '!')
		$arrAppPOC = @()
		$arrAppPOC += $tmpAppPOC.Split("!")
		$strAppPOC = $arrAppPOC[-1]
	}
	Elseif ($strNotes.Contains("POC: ")) {
		$tmpAppPOC = $strNotes.Replace('POC: ', '!')
		$arrAppPOC = @()
		$arrAppPOC += $tmpAppPOC.Split("!")
		$strAppPOC = $arrAppPOC[-1]
	}
	Elseif ($strNotes.Contains("POC:")) {
		$tmpAppPOC = $strNotes.Replace('POC:', '!')
		$arrAppPOC = @()
		$arrAppPOC += $tmpAppPOC.Split("!")
		$strAppPOC = $arrAppPOC[-1]
	}
	Else {$strAppPOC = ""}
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-AppPOC function - $strAppPOC"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $strAppPOC
}
# END OF FUNCTION **********************************************************************************

# START OF GET-IPInfo FUNCTION ********************************************************************
function get-IPInfo ([string]$strComputer) {
	$Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $strComputer | ? {$_.IPEnabled}
	foreach ($Network in $Networks) {
		$IPAddress  = $Network.IpAddress[0]
		$SubnetMask  = $Network.IPSubnet[0]
		$DefaultGateway = $Network.DefaultIPGateway
		$DNSServers  = $Network.DNSServerSearchOrder
		$MACAddress  = $Network.MACAddress
		$WINSServers = $Network.WINSPrimaryServer + "," + $Network.WINSSecondaryServer
		$OutputObj  = New-Object -Type PSObject
		$OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
		$OutputObj | Add-Member -MemberType NoteProperty -Name SubnetMask -Value $SubnetMask
		$OutputObj | Add-Member -MemberType NoteProperty -Name Gateway -Value $DefaultGateway
		$OutputObj | Add-Member -MemberType NoteProperty -Name DNSServers -Value $DNSServers
		$OutputObj | Add-Member -MemberType NoteProperty -Name WINSServers -Value $WINSServers
		$OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress
		$OutputObj
	}
	return $Outputobj
}
# END OF FUNCTION **********************************************************************************

# START OF GET-BLInstalled FUNCTION ********************************************************************
function get-BLInstalled ([string]$strComputer) {
	$error.clear()
	$service = gwmi win32_service -ComputerName $strComputer -filter "name = 'RSCDsvc'"
	if($error) {
		$BLInstalled = "ERROR"}
	elseif ($service.status -eq "OK") {
		$BLInstalled = "YES" }
	else {
		$BLInstalled = "NO" }
	return $BLInstalled
}
# END OF FUNCTION **********************************************************************************

# START OF GET-NetBackupClientInfo FUNCTION ********************************************************************
function get-NetBackupClientInfo ([string]$strComputer) {
	$error.clear()
	$service = gwmi win32_service -ComputerName $strComputer -filter "name = 'NetBackup INET Daemon'"
	if($error) {
		$NetBackupClientInstalled = "ERROR"}
	elseif ($service.status -eq "OK") {
		$VersionInfo = C:\Scripts\POWERSHELL\ComputerDetails\psexec \\jamcdapnap090.amc.faa.gov -s -n 30 "C:\Program Files\VERITAS\NetBackup\bin\Admincmd\bpgetconfig" -g $strComputer -L | find "Version Name"
		$NetBackupClientVersion = $VersionInfo -replace "Version Name = ",""
		$NetBackupClientInstalled = "YES"
	}
	else {
		$NetBackupClientVersion = "N/A"
		$NetBackupClientInstalled = "NO"
	}
	$ClientObject = New-Object PsObject
	$ClientObject | Add-Member -memberType NoteProperty "NetBackupClientInstalled" -Value $NetBackupClientInstalled
	$ClientObject | Add-Member -memberType NoteProperty "NetBackupClientVersion" -Value $NetBackupClientVersion
	return $ClientObject
}
# END OF FUNCTION **********************************************************************************

# START OF Get-McAfeeInstalled FUNCTION ********************************************************************
function Get-McAfeeInstalled ([string]$strComputer) {
	$error.clear()
	$service = gwmi win32_service -ComputerName $strComputer -filter "name = 'McShield'"
	if($error) {
		$McAfeeInstalled = "ERROR"}
	elseif ($service.status -eq "OK") {
		$McAfeeInstalled = "YES" }
	else {
		$McAfeeInstalled = "NO" }
	return $McAfeeInstalled
}
# END OF FUNCTION **********************************************************************************

# START OF Get-ArcSightInstalled FUNCTION ********************************************************************
function Get-ArcSightInstalled ([string]$strComputer) {
	$error.clear()
	$service = gwmi win32_service -ComputerName $strComputer -filter "name = 'arc_nt_local'"
	if($error) {
		$ArcSightInstalled = "ERROR"}
	elseif ($service.status -eq "OK") {
		$ArcSightInstalled = "YES" }
	else {
		$ArcSightInstalled = "NO" }
	return $ArcSightInstalled
}
# END OF FUNCTION **********************************************************************************

# START OF Get-DellOpenManageInfo FUNCTION ********************************************************************
function Get-DellOpenManageInfo ([string]$strComputer) {
	$error.clear()
	$service = gwmi win32_service -ComputerName $strComputer -filter "name = 'Server Administrator'"
	if($error) {
		$DellOpenManageInstalled = "ERROR"}
	elseif ($service.status -eq "OK") {
		$DellOpenManageInstalled = "YES"
		$objOMSAVersionInfo = C:\Scripts\OpenManageDiscovery\psexec \\$strcomputer -s -n 30 omreport about
		foreach ($line in $objOMSAVersionInfo) {
			if ($line -match "Version") {
			$DellOpenManageVersion = $line -replace "Version      : ","" }}}
	else {
		$DellOpenManageInstalled = "NO"
		$DellOpenManageVersion = "N/A" }
	$ClientObject = New-Object PsObject
	$ClientObject | Add-Member -memberType NoteProperty "DellOpenManageInstalled" -Value $DellOpenManageInstalled
	$ClientObject | Add-Member -memberType NoteProperty "DellOpenManageVersion" -Value $DellOpenManageVersion
	return $ClientObject
}
# END OF FUNCTION **********************************************************************************

# START OF Get-PhysicalDiskInfo FUNCTION ********************************************************************
function Get-PhysicalDiskInfo ([string]$strComputer) {
	$objControllerIDInfo = C:\Scripts\POWERSHELL\ComputerDetails\psexec \\$strcomputer -s -n 30 omreport storage controller
	foreach ($line in $objControllerIDInfo) {
		if ($line -match "ID                                            :"){
			$controllerId = $line -replace "ID                                            : ",""
			$objPhysicalDiskInfo = C:\Scripts\OpenManageDiscovery\psexec \\$strcomputer -s -n 30 omreport storage pdisk controller=$controllerID
			foreach ($line in $objPhysicalDiskInfo) {
				if ($line -match "ID                        :") {
					$PhysicalDiskID = $line -replace "ID                        : ","" }
				elseif ($line -match "Capacity                  :") {
					$PhysicalDiskSize = $line -replace "Capacity                  : ",""
					$PhysicalDiskSizeInfo += $PhysicalDiskID + " - " + $PhysicalDiskSize + ","}
				elseif ($line -match "Vendor ID                 :") {
					$NewPhysicalDiskVendor = ($line -replace "Vendor ID                 : ","")
					if ($NewPhysicalDiskVendor.Trim() -ne $TrimPhysicalDiskVendor) {
						$TrimPhysicalDiskVendor = $NewPhysicalDiskVendor.Trim()
						$PhysicalDiskVendor += $TrimPhysicalDiskVendor + ","}}
				elseif ($line -match "Bus Protocol              :"){
					$NewPhysicalDiskType = ($line -replace "Bus Protocol              : ","")
					if ($NewPhysicalDiskType.Trim() -ne $TrimNewPhysicalDiskType) {
						$TrimNewPhysicalDiskType = $NewPhysicalDiskType.Trim()
						$PhysicalDiskType += $TrimNewPhysicalDiskType + ","}}
				elseif ($line -match "Product ID                :") {
					$NewPhysicalDiskProductID = ($line -replace "Product ID                : ","")
					if ($NewPhysicalDiskProductID.Trim() -ne $TrimNewPhysicalDiskProductID) {
						$TrimNewPhysicalDiskProductID = $NewPhysicalDiskProductID.Trim()
						$PhysicalDiskProductID += $TrimNewPhysicalDiskProductID + ","}}
				elseif ($line -match "Serial No.                :") {
					$NewPhysicalDiskSerialNumber = ($line -replace "Serial No.                : ","")
					$TrimPhysicalDiskSerialNumber = $NewPhysicalDiskSerialNumber.Trim()
					$PhysicalDiskSerialNumber += $TrimPhysicalDiskSerialNumber + "," }
				elseif ($line -match "Part Number               :") {
					$NewPhysicalDiskPartNumber = $line -replace "Part Number               : ",""
					$TrimPhysicalDiskPartNumber = $NewPhysicalDiskPartNumber.Trim()
					$PhysicalDiskPartNumber += $TrimPhysicalDiskPartNumber + "," }
			}
		}
	}
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskSizeInfo" -Value $PhysicalDiskSizeInfo.TrimEnd(",")
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskVendor" -Value $PhysicalDiskVendor.TrimEnd(",")
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskType" -Value $PhysicalDiskType.TrimEnd(",")
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskProdID" -Value $PhysicalDiskProductID.TrimEnd(",")
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskSN" -Value $PhysicalDiskSerialNumber.TrimEnd(",")
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskPN" -Value $PhysicalDiskPartNumber.TrimEnd(",")
	return $ClientObject
}
# END OF FUNCTION **********************************************************************************

# START OF GET-DETAILS FUNCTION ********************************************************************
function get-details ($ServerInfo) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-Details Function"
	$strComputer = $ServerInfo.dnshostname
	$strNotes = [string]$serverInfo.Description
	$strPorV = Get-ServerPorV $strComputer
	$strSTATUS = Get-ServerStatus $strComputer
	$strOS = Get-OS $ServerInfo
	$objNetworkInfo = Get-IPInfo $strComputer
	numnetworks = 1
	$strIP = ""
	$strSubnetMask = ""
	$strGateway = ""
	$strDNSServers = ""
	$strWINSServers = ""
	$strMACAddress = ""
	foreach ($Network in $objNetworkInfo) {
		if ($Network.IPAddress -ne "0.0.0.0") {
		$strIP += $Network.IPAddress + ","
		$strSubnetMask += $Network.SubnetMask + ","
		$strGateway += $Network.Gateway + ","
		$strDNSServers += $Network.DNSServers + ","
		$strWINSServers += $Network.WINSServers + ","
		$strMACAddress += $Network.MACAddress + ","}
	}
	$strTrimIP = $strIP.TrimEnd(",")
	$strTrimSubnetMask = $strSubnetMask.TrimEnd(",")
	$strTrimGateway = $strGateway.TrimEnd(",")
	$strTrimDNSServers = $strDNSServers.TrimEnd(",")
	$strTrimWINSServers = $strWINSServers.TrimEnd(",")
	$strTrimMACAddress = $strMACAddress.TrimEnd(",")
	$strMake = (get-wmiobject Win32_ComputerSystem -computername jamcdapnap090.amc.faa.gov).Manufacturer
	$strModel = (get-wmiobject Win32_ComputerSystem -computername jamcdapnap090.amc.faa.gov).Model
	$strMemory = (gwmi Win32_ComputerSystem -computername jamcdapnap090.amc.faa.gov).TotalPhysicalMemory / 1024 / 1024 / 1024
	$strFixMemory = "{0:f}" -f ($strMemory)
	$strProcessors = (gwmi Win32_ComputerSystem -computername $strComputer).NumberOfProcessors
	$Processors = Get-WmiObject Win32_processor -computername $strComputer
	foreach ($Processor in $Processors) {
		$strProcessorType = $Processor.Name
		$strProcessorSpeed = "{0:f}" -f ($Processor.MaxClockSpeed / 1000)
	}
	$DiskInfo = Get-WmiObject win32_LogicalDisk -computername $strComputer
	foreach ($Disk in $DiskInfo) {
		If (($Disk.Size -ne $null) -and ($Disk.Size -ne "0") -and ($Disk.DriveType -eq "3")) {
			$strDiskSize = $Disk.Size / 1024 / 1024 / 1024
			$strFixDiskSize = "{0:f}" -f $strDiskSize
			$strLogicalDisk += $Disk.DeviceID + " - " + $strFixDiskSize + ","}
	}
	$strTrimLogicalDisk = $strLogicalDisk.TrimEnd(",")
	$DiskInfo2 = get-wmiobject Win32_DiskDrive -computername $strComputer
	$NumDisks = 0
	foreach ($Disk in $DiskInfo2) {
		if ($Disk.MediaType -eq "Fixed hard disk media") {
			$Numdisks = $NumDisks + 1
			$strDiskSize = "{0:f}" -f ($Disk.Size / 1024 / 1024 / 1024)
			$strDiskIndex = $Disk.Index
			$strDiskSizeInfo += "Disk " + $strDiskIndex + " - " + $strDiskSize + " GB - " + $Disk.Partitions + " Partitions"
		}
	}
	$strTrimDiskSize = ($strDiskSizeInfo.TrimEnd(","))
	$PhysicalDiskInfo = Get-PhysicalDiskInfo $strComputer
	write-host $PhysicalDiskInfo.PhDiskSizeInfo
	$strBLInstalled = Get-BLInstalled $strComputer
	$NetBackupClientInfo = get-NetBackupClientInfo $strComputer
	$strNetBackupClientInstalled = $NetBackupClientInfo.NetBackupClientInstalled
	$strNetBackupClientVersion = $NetBackupClientInfo.NetBackupClientVersion
	$strMcAfeeInstalled = Get-McAfeeInstalled $strComputer
	$strArcSightInstalled = Get-ArcSightInstalled $strComputer
	$DellOpenManageInfo = Get-DellOpenManageInfo $strComputer
	$strDellOpenManageInstalled = $DellOpenManageInfo.DellOpenManageInstalled
	$strDellOpenManageVersion = $DellOpenManageInfo.DellOpenManageVersion
	$strInstalledApps = Get-InstalledApps $strComputer $strOS
	$strAppPOC = Get-AppPOC $strNotes

	# Create Object to hold information
	$ClientObject = New-Object PsObject
	$ClientObject | Add-Member -memberType NoteProperty "ServerName" -Value $strComputer
	$ClientObject | Add-Member -memberType NoteProperty "ADDescription" -Value $strNotes
	$ClientObject | Add-Member -memberType NoteProperty "PorV" -Value $strPorV
	$ClientObject | Add-Member -memberType NoteProperty "STATUS" -Value $strSTATUS
	$ClientObject | Add-Member -memberType NoteProperty "OS" -Value $strOS
	$ClientObject | Add-Member -memberType NoteProperty "IP" -Value $strTrimIP
	$ClientObject | Add-Member -memberType NoteProperty "Mask" -Value $strTrimSubnetMask
	$ClientObject | Add-Member -memberType NoteProperty "Gateway" -Value $strTrimGateway
	$ClientObject | Add-Member -memberType NoteProperty "DNSServers" -Value $strTrimDNSServers
	$ClientObject | Add-Member -memberType NoteProperty "WINSServers" -Value $strTrimWINSServers
	$ClientObject | Add-Member -memberType NoteProperty "MACAddress" -Value $strTrimMACAddress
	$ClientObject | Add-Member -memberType NoteProperty "Make" -Value $strMake
	$ClientObject | Add-Member -memberType NoteProperty "Model" -Value $strModel
	$ClientObject | Add-Member -memberType NoteProperty "Memory" -Value $strFixMemory
	$ClientObject | Add-Member -memberType NoteProperty "Processors" -Value $strProcessors
	$ClientObject | Add-Member -memberType NoteProperty "ProcessorType" -Value $strProcessorType
	$ClientObject | Add-Member -memberType NoteProperty "ProcessorSpeed" -Value $strProcessorSpeed
	$ClientObject | Add-Member -memberType NoteProperty "LogicalDiskInfo" -Value $strTrimDiskSize
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskSizeInfo" -Value $PhysicalDiskInfo.PhDiskSizeInfo
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskVendor" -Value $PhysicalDiskInfo.PhDiskVendor
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskType" -Value $PhysicalDiskInfo.PhDiskType
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskProdID" -Value $PhysicalDiskInfo.PhDiskProdID
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskSN" -Value $PhysicalDiskInfo.PhDiskSN
	$ClientObject | Add-Member -memberType NoteProperty "PhDiskPN" -Value $PhysicalDiskInfo.PhDiskPN
	$ClientObject | Add-Member -memberType NoteProperty "BladeLogic" -Value $strBLInstalled
	$ClientObject | Add-Member -memberType NoteProperty "NetBackupClient" -Value $strNetBackupClientInstalled
	$ClientObject | Add-Member -memberType NoteProperty "NetBackupClientVersion" -Value $strNetBackupClientVersion
	$ClientObject | Add-Member -memberType NoteProperty "McAfee" -Value $strMcAfeeInstalled
	$ClientObject | Add-Member -memberType NoteProperty "ArcSight" -Value $strArcSightInstalled
	$ClientObject | Add-Member -memberType NoteProperty "DellOpenManage" -Value $strDellOpenManageInstalled
	$ClientObject | Add-Member -memberType NoteProperty "DellOpenManageVersion" -Value $strDellOpenManageVersion
	$ClientObject | Add-Member -memberType NoteProperty "InstalledApps" -Value $strInstalledApps
	$ClientObject | Add-Member -memberType NoteProperty "AppPOC" -Value $strAppPOC

	# Return Information
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Get-Details Function"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	return $ClientObject
}
# END OF FUNCTION *********************************************************************************

# START OF CREATE-SPREADSHEET FUNCTION ********************************************************************
function Create-Spreadsheet ($colComputerDetails) {
	# Create Workbook *********************************************************************************
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Create-Spreadsheet function"
	$Excel = New-Object -comobject Excel.Application
	$Excel.visible = $True
	$Excel.displayalerts=$False
	$WorkBook = $Excel.Workbooks.Add()
	$WorkSheet = $WorkBook.Worksheets.Item(1)

	# Write Column Headings ****************************************************************************
	$WorkSheet.Cells.Item(1,1) = "Server Name"
	$WorkSheet.Cells.Item(1,2) = "AD Description"
	$WorkSheet.Cells.Item(1,3) = "System Name"
	$WorkSheet.Cells.Item(1,4) = "Operating System"
	$WorkSheet.Cells.Item(1,5) = "IP Address"
	$WorkSheet.Cells.Item(1,6) = "Subnet Mask"
	$WorkSheet.Cells.Item(1,7) = "Default Gateway"
	$WorkSheet.Cells.Item(1,8) = "DNS Servers"
	$WorkSheet.Cells.Item(1,9) = "WINS Servers"
	$WorkSheet.Cells.Item(1,10) = "MAC Address"
	$WorkSheet.Cells.Item(1,11) = "Make"
	$WorkSheet.Cells.Item(1,12) = "Model"
	$WorkSheet.Cells.Item(1,13) = "Memory (GB)"
	$WorkSheet.Cells.Item(1,14) = "Processors"
	$WorkSheet.Cells.Item(1,15) = "Processor Type"
	$WorkSheet.Cells.Item(1,16) = "Processor Speed GHz"
	$WorkSheet.Cells.Item(1,17) = "Logical Disk"
	$WorkSheet.Cells.Item(1,18) = "Physical Disk Sizes"
	$WorkSheet.Cells.Item(1,19) = "Physical Disk Vendor"
	$WorkSheet.Cells.Item(1,20) = "Physical Disk Type"
	$WorkSheet.Cells.Item(1,21) = "Physical Disk Product IDs"
	$WorkSheet.Cells.Item(1,22) = "Physical Disk Serial Numbers"
	$WorkSheet.Cells.Item(1,23) = "Physical Disk Part Numbers"
	$WorkSheet.Cells.Item(1,24) = "Status"
	$WorkSheet.Cells.Item(1,25) = "Physical or Virtual"
	$WorkSheet.Cells.Item(1,26) = "Rack Name"
	$WorkSheet.Cells.Item(1,27) = "Installed Apps"
	$WorkSheet.Cells.Item(1,28) = "Application POC"
	$WorkSheet.Cells.Item(1,29) = "Data Services POC"
	$WorkSheet.Cells.Item(1,30) = "BladeLogic"
	$WorkSheet.Cells.Item(1,31) = "NetBackup Client"
	$WorkSheet.Cells.Item(1,32) = "NetBackup Client Version"
	$WorkSheet.Cells.Item(1,33) = "McAfee"
	$WorkSheet.Cells.Item(1,34) = "ArcSight"
	$WorkSheet.Cells.Item(1,35) = "Dell OpenManage"
	$WorkSheet.Cells.Item(1,36) = "Dell OpenManage Version"

	# Format the column headings **********************************************************************
	$range = $WorkSheet.UsedRange
	$range.Interior.ColorIndex = 19
	$range.Font.ColorIndex = 11
	$range.Font.Bold = $True

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Creating the excel spreadsheet"
	# Create and write the Excel Spreadsheet.  Do it all at once so the Excel spreadsheet is not open for an extended period of time.
	Foreach ($Server in $colComputerDetails) {
		$sname = $server.servername
		$WorkSheet.Cells.Item($intRow,1) = $Server.ServerName
		$WorkSheet.Cells.Item($intRow,2) = $Server.ADDescription
		$WorkSheet.Cells.Item($intRow,4) = $Server.OS
		$WorkSheet.Cells.Item($intRow,5) = $Server.IP
		$WorkSheet.Cells.Item($intRow,6) = $Server.Mask
		$WorkSheet.Cells.Item($intRow,7) = $Server.Gateway
		$WorkSheet.Cells.Item($intRow,8) = $Server.DNSServers
		$WorkSheet.Cells.Item($intRow,9) = $Server.WINSServers
		$WorkSheet.Cells.Item($intRow,10) = $Server.MACAddress
		$WorkSheet.Cells.Item($intRow,11) = $Server.Make
		$WorkSheet.Cells.Item($intRow,12) = $Server.Model
		$WorkSheet.Cells.Item($intRow,13) = $Server.Memory
		$WorkSheet.Cells.Item($intRow,14) = $Server.Processors
		$WorkSheet.Cells.Item($intRow,15) = $Server.ProcessorType
		$WorkSheet.Cells.Item($intRow,16) = $Server.ProcessorSpeed
		$WorkSheet.Cells.Item($intRow,17) = $Server.LogicalDisk
		$WorkSheet.Cells.Item($intRow,18) = $Server.PhDiskSizeInfo
		$WorkSheet.Cells.Item($intRow,19) = $Server.PhDiskVendor
		$WorkSheet.Cells.Item($intRow,20) = $Server.PhDiskType
		$WorkSheet.Cells.Item($intRow,21) = $Server.PhDiskProdID
		$WorkSheet.Cells.Item($intRow,22) = $Server.PhDiskSN
		$WorkSheet.Cells.Item($intRow,23) = $Server.PhDiskPN
		$WorkSheet.Cells.Item($intRow,24) = $Server.STATUS
		$WorkSheet.Cells.Item($intRow,25) = $Server.PorV
		$WorkSheet.Cells.Item($intRow,27) = $Server.InstalledApps
		$WorkSheet.Cells.Item($intRow,28) = $Server.AppPOC
		$WorkSheet.Cells.Item($intRow,30) = $Server.BladeLogic
		$WorkSheet.Cells.Item($intRow,31) = $Server.NetBackupClient
		$WorkSheet.Cells.Item($intRow,32) = $Server.NetBackupClientVersion
		$WorkSheet.Cells.Item($intRow,33) = $Server.McAfee
		$WorkSheet.Cells.Item($intRow,34) = $Server.ArcSight
		$WorkSheet.Cells.Item($intRow,35) = $Server.DellOpenManage
		$WorkSheet.Cells.Item($intRow,36) = $Server.DellOpenManageVersion
		$intRow = $intRow + 1
	}	
	# Close and save the Excel Spreadsheet. **********************************************************
	$range.EntireColumn.AutoFit()
	$Workbook.SaveAs("$ExcelFileDirectory\CSA-ComputerDetails2.xls")
	$Excel.quit()
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Completed the Create-Spreadsheet function"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
}
# END OF FUNCTION ****************************************************************************************

# Main Script ********************************************************************************************
Write-log "**************************************************************************************************************"
Write-log "**************************************************************************************************************"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Acquiring Computer Details for the active and available servers in $ouToSearch"
# Get the information for each computer in the specified OU and all sub OUs
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Getting the list of computers from active directory"
$ADComputerList = Get-QADComputer -SearchRoot $ouToSearch -SearchScope 'Subtree' | sort-object name
$ComputerList = Get-Content C:\Scripts\POWERSHELL\ComputerDetails\list.txt
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - List has been acquired"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"

# Set some variables *******************************************************************************
#Integer used for the row numbers in the spreadsheet
$intRow = 2
# Ping command that will used to determine the status of servers.
$ping = new-object System.Net.NetworkInformation.Ping
# Create Array to be used to temporarily store data
$colComputerDetails = @()
$SkippedServers = @()

# Start processing each Computer Object that was found in the specified OU *************************
Foreach ($server in $ComputerList) {
	Write-log "**************************************************************************************************************"
	Write-log "**************************************************************************************************************"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Processing $Server"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Pulling Server Name information"
	$Servershort = $Server -replace ".amc.faa.gov",""
	$ServerNameForSearch = "FAA\" + $Servershort + "$"
	$ServerInfo = Get-QADComputer "$ServerNameForSearch"
	# Pull out the Server name and FQDN
	$strComputer = [string]$server
	Write-Host "Gatering Data for $strComputer"
	$strFQDN = [string]$serverInfo.dNSHostName
	$strDateCreated = [string]$serverInfo.whenCreated
	$strCN = [string]$serverInfo.canonicalName
	$strDescription = [string]$serverInfo.description

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining whether the server object is disabled in Active Directory"
	# Determine whether or not the server is disabled
	$strServerDisabled = [string]$serverInfo.AccountIsDisabled

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining the object type to ensure it is a computer object"
	# Determine the object type
	$strServerType = [string]$ServerInfo.Type

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Making sure the server is pingable"
	# Ping the server to see if you can connect. This will ping both the FQDN and the display name.
	$Reply = $ping.send($strFQDN)
	$strReplyStatus = [string]$reply.status
	if ($strReplyStatus -ne "success") {
		$Reply = $ping.send($strComputer)
		$strReplyStatus = [string]$reply.status
	}

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Deciding whether or not to continue to process this server"
	# If the server is not disabled, is a Computer Object, and is pingable, continue to process the server.
	If (($strServerDisabled -ne "True") -and ($strServerType -eq "computer") -and ($strReplyStatus -eq "success")){
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Continuing to process this server"
		# Call the get-details function
		$ComputerDetails = get-details $ServerInfo
		$colComputerDetails += $ComputerDetails
	}
	Elseif ($strServerDisabled -eq "True") {
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: $strComputer is disabled in Active Directory and has been ignored *******************"
		Write-Host "$strComputer is disabled."
		$SkippedServers += $strComputer
	}
	Elseif ($strServerType -ne "computer") {
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: $strComputer is not a server and has been ignored *******************"
		Write-Host "$strComputer is not a server."
		$SkippedServers += $strComputer
	}
	Elseif ($strReplyStatus -ne "Success") {
		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: $strComputer cannot be reached and has been ignored *******************"
		Write-Host "$strComputer cannot be reached"
		$SkippedServers += $strComputer
	}

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Resetting variables"
	# Reset the Variables before proceeding with the next computer
		$strComputer = ""
		$strPorV = ""
		$strOS = ""
		$installedApps = ""
		$strAppPOC = ""
		$strNotes = ""
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Proceeding with the next server in the list"
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $Separator"
}

Create-Spreadsheet $colComputerDetails
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - **** The following servers were skipped ****"
# List the skipped servers *************************************************************************
Foreach ($Server in $SkippedServers) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $Server"
}
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Computer Details have been acquired for the active and available servers in $ouToSearch"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"