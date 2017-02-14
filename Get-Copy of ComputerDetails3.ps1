#==============================================================================
#  Get Computer Details
#  Jared K. Thomason
#  12/29/2011
#==============================================================================

#Tells the Script to continue on errors without displaying the error
$erroractionpreference = "SilentlyContinue"

#==============================================================================
#You may want to check the following parameters before you run the script.

#Directory for excell files to be stored.
$ExcelFileDirectory = "c:"

#Location and file name for the log file.
$logfile = C:\computerdetails.txt

#Canonical name of the OU you want to search through.
$ouToSearch = "faa.gov/Programs/National/Data_Centers/Servers/AWA_Servers"

#==============================================================================

# FUNCTION ********************************************************************
# This is a function used to write information to a log file.
function write-log([string]$info){
	if($loginitialized -eq $false){
		$FileHeader > $logfile
		$script:loginitialized = $True
	}
	$info >> $logfile
}

# Log file Information
$script:logfile = "C:\Scripts\POWERSHELL\ComputerDetails\logs\ComputerDetails-$(get-date -format MM-dd-yyyy-HH-mm-ss).txt"
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

# FUNCTION ********************************************************************
# A function used to get the OS information for the server using WMI
function Get-OSInformation ([string]$ComputerName) {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting the Get-OSInformation function"
	# WMI Used in Script
	$WmiObject = "ComputerSystem",
	"OperatingSystem",
	"NetworkAdapterConfiguration", 
	"SystemEnclosure",
	"LogicalDisk",
	"Processor",
	"DisplayConfiguration"

	$WmiObject | 
		% { Set-Variable -name $_ -value (gwmi Win32_$_ -ComputerName $ComputerName) }

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Getting OS"
	# Get OperatingSystem
	$OS = ($OperatingSystem.Caption).ToString()
#	$OSLanguage = ($OperatingSystem.OSLanguage).ToString()


Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - OS is $OS"
	# Create Object to hold information
	$ClientObject = New-Object PsObject
	$ClientObject | Add-Member -memberType NoteProperty "OperatingSystem" -Value $OS

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Returning $ClientObject"
	# Return Information
	return $ClientObject
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - End of the Get-OSInformation function"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
}

# FUNCTION ********************************************************************
function get-details ([string]$strComputer) {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining whether the server is a physical or a virtual server"
#Determine whether it is a physical or a virtual server
	If ($strComputer.Substring(7,1) -eq "n") {
		$strPorV = "Physical"
	}
	elseif ($strComputer.Substring(7,1) -eq "v") {
		$strPorV = "Virtual"
	}
	else {
		$strPorV = ""
	}

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Calling the Get-OSInformation function to get the OS of the server"
# Call a function to get the Operating System of the server.
	$error.clear()
	$ServerInformation = Get-OSInformation $strComputer
	if ($error) {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Could not use the Get-OSInformation function to get the OS of the server"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") -       There are probably problems with WMI access"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") -       Using the OS information stored in active directory"
		$strOS = [string]$server.OSName
	}
	else {
		$strOS = [string]$ServerInformation.OperatingSystem
	}

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining the applications that are found in the uninstall directory in the registry"
# Start searching the registry for applications that can be uninstalled.  This will help us determine what applications
# are actually installed.
	if ($strOS.contains("server 2008")) {
		$applist = get-wmiobject -namespace "root/cimv2" -computername $strComputer win32_product | sort-object name
		$arrAppList = @()
		foreach ($app in $applist) {
			$arrAppList += $app.name + ","
		}
		$strAppList = [string]$arrAppList
		Write-host $strAppList
	}
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

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting to drill through and acquire the names of the applications"
	# Drill through each key from the list and pull out the value of 
	# “DisplayName”
	if ($EncounteredError -eq 0) {
		$installedApps = @()
		Foreach ($key in $subkeys) {
			$exactkey=$key 
			$NewSubKey=$SubBranch+"\\"+$exactkey 
			$ReadUninstall=$registry.OpenSubKey($NewSubKey) 
			$Value=$ReadUninstall.GetValue("DisplayName")
			if ($Value -ne $null) {
				if ((-not $Value.contains(" (KB")) -and (-not $Value.contains("ArcSight SmartConnector")) -and (-not $Value.contains(" - KB")) -and (-not $Value.contains(" kb")) -and (-not $Value.contains("Windows Management Framework Core"))) {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Application Name - $Value"
					$installedApps += [string]$Value + ","
				}
			}
		}
	}
$orderedInstalledApps = $installedApps | sort-object
#Ensture the list of installed applications is a string.
	$strInstalledApps = [string]$orderedInstalledApps
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Finished searching the registry"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Using the Active Directory description field to put in the notes column"
#Use the Active Directory Description Field for Notes
	$strNotes = [string]$server.Description

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Attempting to acquire POC information using the Notes"
#Look for POC information in the Active Directory Description field.
	If ($strNotes.Contains("POC "))
	{
		$tmpAppPOC = $strNotes.Replace('POC ', '!')
		$arrAppPOC = @()
		$arrAppPOC += $tmpAppPOC.Split("!")
		$strAppPOC = $arrAppPOC[-1]
	}
	Elseif ($strNotes.Contains("POC: "))
	{
		$tmpAppPOC = $strNotes.Replace('POC: ', '!')
		$arrAppPOC = @()
		$arrAppPOC += $tmpAppPOC.Split("!")
		$strAppPOC = $arrAppPOC[-1]
	}
	Elseif ($strNotes.Contains("POC:"))
	{
		$tmpAppPOC = $strNotes.Replace('POC:', '!')
		$arrAppPOC = @()
		$arrAppPOC += $tmpAppPOC.Split("!")
		$strAppPOC = $arrAppPOC[-1]
	}
	Else
	{$strAppPOC = ""}

	# Create Object to hold information
	$ClientObject = New-Object PsObject
	$ClientObject | Add-Member -memberType NoteProperty "ComputerName" -Value $strComputer
	$ClientObject | Add-Member -memberType NoteProperty "PorV" -Value $strPorV
	$ClientObject | Add-Member -memberType NoteProperty "OS" -Value $strOS
	$ClientObject | Add-Member -memberType NoteProperty "InstalledApps" -Value $strInstalledApps
	$ClientObject | Add-Member -memberType NoteProperty "AppPOC" -Value $strAppPOC
	$ClientObject | Add-Member -memberType NoteProperty "Notes" -Value $strNotes

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Returning $ClientObject"
	# Return Information
	return $ClientObject
}

# Main Script *****************************************************************
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Acquiring Computer Details for the active and available servers in $ouToSearch"
#Get the information for each computer in the specified OU and all sub OUs
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Getting the list of computers from active directory"
$ComputerList = Get-QADComputer -SearchRoot $ouToSearch -SearchScope 'Subtree' | sort-object name
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - List has been acquired"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"

#Set some variables-------------------------------------------
#Integer used for the row numbers in the spreadsheet
$intRow = 2
#Ping command that will used to determine the status of servers.
$ping = new-object System.Net.NetworkInformation.Ping
#Create Array to be used to temporarily store data
$colComputerDetails = @()
$SkippedServers = @()

#Start processing each Computer Object that was found in the specified OU.
Foreach ($server in $ComputerList)
{
Write-log "**************************************************************************************************************"
Write-log "**************************************************************************************************************"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Processing $Server"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Pulling Server Name information"
#Pull out the Server name and FQDN
	$strComputer = [string]$server.Name
	Write-Host "Gatering Data for $strComputer"
	$strFQDN = [string]$server.dNSHostName

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining whether server is disabled"
#Determine whether or not the server is disabled
	$strServerDisabled = [string]$server.AccountIsDisabled

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining the object type to ensure it is a computer object"
#Determine the object type
	$strServerType = [string]$Server.Type

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Making sure the server is pingable"
#Ping the server to see if you can connect. This will ping both the FQDN and the display name.
	$Reply = $ping.send($strFQDN)
	$strReplyStatus = [string]$reply.status
	if ($strReplyStatus -ne "success") {
		$Reply = $ping.send($strComputer)
		$strReplyStatus = [string]$reply.status
	}

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Deciding whether or not to continue to process this server"
#If the server is not disabled, is a Computer Object, and is pingable, continue to process the server.
	If (($strServerDisabled -ne "True") -and ($strServerType -eq "computer") -and ($strReplyStatus -eq "success")){
/*
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining whether the server is a physical or a virtual server"
#Determine whether it is a physical or a virtual server
		If ($strComputer.Substring(7,1) -eq "n") {
			$strPorV = "Physical"
		}
		elseif ($strComputer.Substring(7,1) -eq "v") {
			$strPorV = "Virtual"
		}
		else {
			$strPorV = ""
		}

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Calling the Get-OSInformation function to get the OS of the server"
# Call a function to get the Operating System of the server.
		$error.clear()
		$ServerInformation = Get-OSInformation $strComputer
		if ($error) {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Could not use the Get-OSInformation function to get the OS of the server"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") -       There are probably problems with WMI access"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") -       Using the OS information stored in active directory"
			$strOS = [string]$server.OSName
		}
		else {
			$strOS = [string]$ServerInformation.OperatingSystem
		}

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining the applications that are found in the uninstall directory in the registry"
# Start searching the registry for applications that can be uninstalled.  This will help us determine what applications
# are actually installed.
		if ($strOS.contains("server 2008")) {
			$applist = get-wmiobject -namespace "root/cimv2" -computername $strComputer win32_product | sort-object name
			$arrAppList = @()
			foreach ($app in $applist) {
				$arrAppList += $app.name + ","
			}
			$strAppList = [string]$arrAppList
			Write-host $strAppList
		}
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

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Starting to drill through and acquire the names of the applications"
		# Drill through each key from the list and pull out the value of 
		# “DisplayName”
		if ($EncounteredError -eq 0) {
			$installedApps = @()
			Foreach ($key in $subkeys) 
			{ 
				$exactkey=$key 
				$NewSubKey=$SubBranch+"\\"+$exactkey 
				$ReadUninstall=$registry.OpenSubKey($NewSubKey) 
				$Value=$ReadUninstall.GetValue("DisplayName")
				if ($Value -ne $null) {
					if ((-not $Value.contains(" (KB")) -and (-not $Value.contains("ArcSight SmartConnector")) -and (-not $Value.contains(" - KB")) -and (-not $Value.contains(" kb")) -and (-not $Value.contains("Windows Management Framework Core"))) {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Application Name - $Value"
						$installedApps += [string]$Value + ","
					}
				}
			}
		}
$orderedInstalledApps = $installedApps | sort-object
#Ensture the list of installed applications is a string.
		$strInstalledApps = [string]$orderedInstalledApps
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Finished searching the registry"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Using the Active Directory description field to put in the notes column"
#Use the Active Directory Description Field for Notes
		$strNotes = [string]$server.Description

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Attempting to acquire POC information using the Notes"
#Look for POC information in the Active Directory Description field.
		If ($strNotes.Contains("POC "))
		{
			$tmpAppPOC = $strNotes.Replace('POC ', '!')
			$arrAppPOC = @()
			$arrAppPOC += $tmpAppPOC.Split("!")
			$strAppPOC = $arrAppPOC[-1]
		}
		Elseif ($strNotes.Contains("POC: "))
		{
			$tmpAppPOC = $strNotes.Replace('POC: ', '!')
			$arrAppPOC = @()
			$arrAppPOC += $tmpAppPOC.Split("!")
			$strAppPOC = $arrAppPOC[-1]
		}
		Elseif ($strNotes.Contains("POC:"))
		{
			$tmpAppPOC = $strNotes.Replace('POC:', '!')
			$arrAppPOC = @()
			$arrAppPOC += $tmpAppPOC.Split("!")
			$strAppPOC = $arrAppPOC[-1]
		}
		Else
		{$strAppPOC = ""}
*/

$ComputerDetails = get-details $strComputer
$strComputer = $ComputerDetails.ComputerName
$strPorV = $ComputerDetails.PorV
$strOS = $ComputerDetails.OS
$strInstalledApps = $ComputerDetails.InstalledApps
$strAppPOC = $ComputerDetails.AppPOC
$strNotes = $ComputerDetails.Notes

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Putting the details that have been acquired in an array to be used later"
#Store the Computer Details we have acquired in an array to be used later.
		$objComputerDetails = New-Object System.Object
		$objComputerDetails | Add-Member -type NoteProperty -name ServerName -value $strComputer
		$objComputerDetails | Add-Member -type NoteProperty -name PorV -value $strPorV
		$objComputerDetails | Add-Member -type NoteProperty -name OS -value $strOS
		$objComputerDetails | Add-Member -type NoteProperty -name InstalledApp -value $strInstalledApps
		$objComputerDetails | Add-Member -type NoteProperty -name AppPOC -value $strAppPOC
		$objComputerDetails | Add-Member -type NoteProperty -name Notes -value $strNotes
		$colComputerDetails += $objComputerDetails
		Write-Host $colComputerDetails[-1].ServerName

#		$WorkSheet.Cells.Item($intRow,1) = $strComputer
#		$WorkSheet.Cells.Item($intRow,5) = $strPorV
#		$WorkSheet.Cells.Item($intRow,6) = $strOS
#		$WorkSheet.Cells.Item($intRow,7) = $strInstalledApps
#		$WorkSheet.Cells.Item($intRow,8) = $strAppPOC
#		$WorkSheet.Cells.Item($intRow,9) = $strNotes
#		$intRow = $intRow + 1
	}
	Elseif ($strServerDisabled -eq "True") {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: $strComputer is disabled *******************"
		Write-Host "$strComputer is disabled."
		$SkippedServers += $strComputer
	}
	Elseif ($strServerType -ne "computer") {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: $strComputer is not a server *******************"
		Write-Host "$strComputer is not a server."
		$SkippedServers += $strComputer
	}
	Elseif ($strReplyStatus -ne "Success") {
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - ERROR: $strComputer cannot be reached *******************"
		Write-Host "$strComputer cannot be reached"
		$SkippedServers += $strComputer
	}

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Resetting variables"
#Reset the Variables before proceeding with the next computer
		$strComputer = ""
		$strPorV = ""
		$strOS = ""
		$installedApps = ""
		$strAppPOC = ""
		$strNotes = ""
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Proceeding with the next server in the list"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $Separator"
}

#Create Workbook----------------------------------------------
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Creating Workbook in Excel"
$Excel = New-Object -comobject Excel.Application
$Excel.visible = $True
$Excel.displayalerts=$False
$WorkBook = $Excel.Workbooks.Add()
$WorkSheet = $WorkBook.Worksheets.Item(1)
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Workbook has been created"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"

#Write Column Headings-----------------------------------------
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Writing Column Headings in the Excel Workbook"
$WorkSheet.Cells.Item(1,1) = "Server Name"
$WorkSheet.Cells.Item(1,2) = "System Name"
$WorkSheet.Cells.Item(1,3) = "Data Services POC"
$WorkSheet.Cells.Item(1,4) = "Rack"
$WorkSheet.Cells.Item(1,5) = "Physical or Virtual"
$WorkSheet.Cells.Item(1,6) = "OS"
$WorkSheet.Cells.Item(1,7) = "Installed App"
$WorkSheet.Cells.Item(1,8) = "Application POC"
$WorkSheet.Cells.Item(1,9) = "Notes"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Column Headings have been written"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"

#Format the column headings-----------------------------------
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Formatting the column headings"
$range = $WorkSheet.UsedRange
$range.Interior.ColorIndex = 19
$range.Font.ColorIndex = 11
$range.Font.Bold = $True
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Column headings have been formatted"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Creating the excel spreadsheet"
#Create and write the Excel Spreadsheet.  Do it all at once so the Excel spreadsheet is not open for an extended period of time.
Foreach ($Server in $colComputerDetails) {
$sname = $server.servername
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Adding details for $sname to the spreadsheet"
		$WorkSheet.Cells.Item($intRow,1) = $Server.ServerName
		$WorkSheet.Cells.Item($intRow,5) = $Server.PorV
		$WorkSheet.Cells.Item($intRow,6) = $Server.OS
		$WorkSheet.Cells.Item($intRow,7) = $Server.InstalledApp
		$WorkSheet.Cells.Item($intRow,8) = $Server.AppPOC
		$WorkSheet.Cells.Item($intRow,9) = $Server.Notes
		$intRow = $intRow + 1
}	

Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Spreadsheet has been created"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Saving the spreadsheet"
#Close and save the Excel Spreadsheet.
$range.EntireColumn.AutoFit()
$Workbook.SaveAs("$ExcelFileDirectory\AWA-ComputerDetails2.xls")
$Excel.quit()
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - **** The following servers were skipped ****"
Foreach ($Server in $SkippedServers) {
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $Server"
}
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Computer Details have been acquired for the active and available servers in $ouToSearch"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - $separator"
