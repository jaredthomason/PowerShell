#==============================================================================
#  Get Computer Details
#  Jared K. Thomason
#  12/29/2011
#==============================================================================

$erroractionpreference = "SilentlyContinue"
$ScriptDir = "c:"
$ouToSearch = "faa.gov/LOB/ATO/AWA/- AWA_Servers"
$ComputerList = Get-QADComputer -SearchRoot $ouToSearch -SearchScope 'Subtree' | sort-object name

#Create Workbook----------------------------------------------
$Excel = New-Object -comobject Excel.Application
$Excel.visible = $True
$Excel.displayalerts=$False
$WorkBook = $Excel.Workbooks.Add()
$WorkSheet = $WorkBook.Worksheets.Item(1)

$WorkSheet.Cells.Item(1,1) = "Server Name"
$WorkSheet.Cells.Item(1,2) = "System Name"
$WorkSheet.Cells.Item(1,3) = "Data Services POC"
$WorkSheet.Cells.Item(1,4) = "Rack"
$WorkSheet.Cells.Item(1,5) = "Physical or Virtual"
$WorkSheet.Cells.Item(1,6) = "OS"
$WorkSheet.Cells.Item(1,7) = "Installed App"
$WorkSheet.Cells.Item(1,8) = "Application POC"
$WorkSheet.Cells.Item(1,9) = "Notes"

$range = $WorkSheet.UsedRange
$range.Interior.ColorIndex = 19
$range.Font.ColorIndex = 11
$range.Font.Bold = $True

$intRow = 2

Foreach ($server in $ComputerList)
{
	Write-Host $server.Name
	If (($server.AccountIsDisabled -ne "True") -and ($Server.Type -eq "computer")){
	$strComputer = [string]$server.Name
	If ($strComputer.Substring(7,1) -eq "p") {
		$strPorV = "Physical"
	}
	elseif ($strComputer.Substring(7,1) -eq "v") {
		$strPorV = "Virtual"
	}
	else {
		$strPorV = ""
	}
	$strOS = [string]$server.OSName

	# Branch of the Registry 
	$Branch='LocalMachine' 

	# Main Sub Branch you need to open 
	$SubBranch="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
	$EncounteredError = 0
	$Error.Clear()
	$registry=[microsoft.win32.registrykey]::OpenRemoteBaseKey('Localmachine',$strComputer)
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
	# “DisplayName” – Write to the Host console the name of the computer 
	# with the application beside it
	if ($EncounteredError -eq 0) {
	$installedApps = @()
	Foreach ($key in $subkeys) 
	{ 
	    $exactkey=$key 
	    $NewSubKey=$SubBranch+"\\"+$exactkey 
	    $ReadUninstall=$registry.OpenSubKey($NewSubKey) 
	    $Value=$ReadUninstall.GetValue("DisplayName") 
	    $installedApps += $Value
	}
	}
	$strNotes = [string]$server.Description
	If ($strNotes.Contains("POC"))
	{
	$tmpAppPOC = $strNotes.Replace('POC', '!')
	$arrAppPOC = @()
	$arrAppPOC += $tmpAppPOC.Split("!")
	$strAppPOC = $arrAppPOC[-1]
	}
	Else
	{$strAppPOC = ""}

	$WorkSheet.Cells.Item($intRow,1) = $strComputer
	$WorkSheet.Cells.Item($intRow,5) = $strPorV
	$WorkSheet.Cells.Item($intRow,6) = $strOS
	$WorkSheet.Cells.Item($intRow,7) = $installedApps
	$WorkSheet.Cells.Item($intRow,8) = $strAppPOC
	$WorkSheet.Cells.Item($intRow,9) = $strNotes
	$intRow = $intRow + 1
	}
	$strComputer = ""
	$strPorV = ""
	$strOS = ""
	$installedApps = ""
	$strAppPOC = ""
	$strNotes = ""
}
$range.EntireColumn.AutoFit()
$Workbook.SaveAs("$ScriptDir\ComputerDetails.xls")
$Excel.quit()