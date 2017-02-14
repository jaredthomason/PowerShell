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
#Canonical name of the OU you want to search through.
$ouToSearch = "faa.gov/Programs/National/Data_Centers/Servers/AWA_Servers"
#==============================================================================

#Get the information for each computer in the specified OU and all sub OUs
$ComputerList = Get-QADComputer -SearchRoot $ouToSearch -SearchScope 'Subtree' | sort-object name

#Create Workbook----------------------------------------------
$Excel = New-Object -comobject Excel.Application
$Excel.visible = $True
$Excel.displayalerts=$False
$WorkBook = $Excel.Workbooks.Add()
$WorkSheet = $WorkBook.Worksheets.Item(1)

#Write Colum Headings-----------------------------------------
$WorkSheet.Cells.Item(1,1) = "Server Name"
$WorkSheet.Cells.Item(1,2) = "System Name"
$WorkSheet.Cells.Item(1,3) = "Data Services POC"
$WorkSheet.Cells.Item(1,4) = "Rack"
$WorkSheet.Cells.Item(1,5) = "Physical or Virtual"
$WorkSheet.Cells.Item(1,6) = "OS"
$WorkSheet.Cells.Item(1,7) = "Installed App"
$WorkSheet.Cells.Item(1,8) = "Application POC"
$WorkSheet.Cells.Item(1,9) = "Notes"

#Format the column headings-----------------------------------
$range = $WorkSheet.UsedRange
$range.Interior.ColorIndex = 19
$range.Font.ColorIndex = 11
$range.Font.Bold = $True

#Integer used for the row numbers in the spreadsheet
$intRow = 2

#Ping command that will used to determine the status of servers.
$ping = new-object System.Net.NetworkInformation.Ping

#Create Array to be used to temporarily store data
$colComputerDetails = @()

#Start processing each Computer Object that was found in the specified OU.
Foreach ($server in $ComputerList)
{
#Determine the Server name
	$strComputer = [string]$server.Name
	Write-Host "Gatering Data for $strComputer"
	$strFQDN = [string]$server.dNSHostName

#Determine whether or not the server is disabled
	$strServerDisabled = [string]$server.AccountIsDisabled

#Determine the object type
	$strServerType = [string]$Server.Type

#Ping the server to see if you can connect. This will ping both the FQDN and the display name.
	$Reply = $ping.send($strFQDN)
	$strReplyStatus = [string]$reply.status
	if ($strReplyStatus -ne "success") {
		$Reply = $ping.send($strComputer)
		$strReplyStatus = [string]$reply.status
	}

#If the server is not disabled, is a Computer Object, and is pingable, continue to process the server.
	If (($strServerDisabled -ne "True") -and ($strServerType -eq "computer") -and ($strReplyStatus -eq "success")){

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

#Call another script to get the Operating System of the server.
		$error.clear()
		$ServerInformation = C:\Scripts\POWERSHELL\ComputerDetails\get-inventory.ps1 -client $strComputer
		if ($error) {
			$strOS = [string]$server.OSName
		}
		else {
			$strOS = [string]$ServerInformation.OperatingSystem
		}

#Start searching the registry for applications that can be uninstalled.  This will help us determine what applications
#are actually installed.
		# Branch of the Registry 
		$Branch='LocalMachine' 
		# Main Sub Branch you need to open 
		$SubBranch="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
		$EncounteredError = 0
		$Error.Clear()
		$registry=[microsoft.win32.registrykey]::OpenRemoteBaseKey('Localmachine',$strComputer)

#Error Handling
		if ($error) {
			Write-Host "Could not access the registry of $strComputer"
			$EncounteredError = 1
		}
		Else {
			$Error.Clear()
			$registrykey=$registry.OpenSubKey($Subbranch)
			if ($error) {
				Write-Host "Could not access the registry of $strComputer"
				$EncounteredError = 1
			}
			Else {
				$Error.Clear()
				$SubKeys=$registrykey.GetSubKeyNames()
				if ($error) {
					Write-Host "Could not access the registry of $strComputer"
					$EncounteredError = 1
				}
			}
		}

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
				    if (($Value -notlike "* (KB*") -or ($Value -ne "ArcSight SmartConnector")) {
					    $installedApps += $Value
					    $installedApps += ","
				    }
			    }
			}
		}
#Ensture the list of installed applications is a string.
		$strInstalledApps = [string]$installedApps

#Use the Active Directory Description Field for Notes
		$strNotes = [string]$server.Description

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

		$WorkSheet.Cells.Item($intRow,1) = $strComputer
		$WorkSheet.Cells.Item($intRow,5) = $strPorV
		$WorkSheet.Cells.Item($intRow,6) = $strOS
		$WorkSheet.Cells.Item($intRow,7) = $strInstalledApps
		$WorkSheet.Cells.Item($intRow,8) = $strAppPOC
		$WorkSheet.Cells.Item($intRow,9) = $strNotes
		$intRow = $intRow + 1
	}
	Elseif ($strServerDisabled -eq "True") {Write-Host "$strComputer is disabled."}
	Elseif ($strServerType -ne "computer") {Write-Host "$strComputer is not a server."}
	Elseif ($strReplyStatus -ne "Success") {Write-Host "$strComputer cannot be reached"}

#Reset the Variables before proceeding with the next computer
		$strComputer = ""
		$strPorV = ""
		$strOS = ""
		$installedApps = ""
		$strAppPOC = ""
		$strNotes = ""
}

#Create and write the Excel Spreadsheet.  Do it all at once so the Excel spreadsheet is not open for an extended period of time.
#Foreach ($Server in $colComputerDetails) {
#		$WorkSheet.Cells.Item($intRow,1) = $Server.ServerName
#		$WorkSheet.Cells.Item($intRow,5) = $Server.PorV
#		$WorkSheet.Cells.Item($intRow,6) = $Server.OS
#		$WorkSheet.Cells.Item($intRow,7) = $Server.InstalledApp
#		$WorkSheet.Cells.Item($intRow,8) = $Server.AppPOC
#		$WorkSheet.Cells.Item($intRow,9) = $Server.Notes
#		$intRow = $intRow + 1
#}	

#Close and save the Excel Spreadsheet.
$range.EntireColumn.AutoFit()
$Workbook.SaveAs("$ExcelFileDirectory\AWA-ComputerDetails2.xls")
$Excel.quit()
