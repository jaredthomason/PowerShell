#========================================================================
#  Check for patch
#  Austin J. Hageman
#  11/05/08
#========================================================================

$erroractionpreference = "SilentlyContinue"
$ScriptDir = split-path -parent $MyInvocation.MyCommand.Path

#Get computer list from AD----------------------------------------------
$strCategory = "computer"
$ADList = @()

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchScope = "Subtree"
$objSearcher.PageSize = 1000
$objSearcher.Filter = ("(objectCategory=$strCategory)")
$colProplist = "name","description"
foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://OU=- AJF-Citrix,OU=Servers,OU=Abacus,OU=National,OU=Programs,DC=faa,DC=gov")
$objSearcher.SearchRoot = $objDomain
$colResults = $objSearcher.FindAll()
foreach ($objResult in $colResults)
    {$objComputer = $objResult.Properties; $ADList += ,($objComputer.name,$objComputer.description)}

$objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://OU=AJF-Citrix,OU=Servers,OU=Abacus,OU=National,OU=Programs,DC=faa,DC=gov")
$objSearcher.SearchRoot = $objDomain
$colResults = $objSearcher.FindAll()
foreach ($objResult in $colResults)
    {$objComputer = $objResult.Properties; $ADList += ,($objComputer.name,$objComputer.description)}
	
#Create Workbook----------------------------------------------
$Excel = New-Object -comobject Excel.Application
$Excel.visible = $True
$Excel.displayalerts=$False
$WorkBook = $Excel.Workbooks.Add()
$WorkSheet = $WorkBook.Worksheets.Item(1)

$WorkSheet.Cells.Item(1,1) = "Machine Name"
$WorkSheet.Cells.Item(1,2) = "PatchStatus"
$WorkSheet.Cells.Item(1,3) = "Report Time Stamp"
$WorkSheet.Cells.Item(1,4) = "Description"
$WorkSheet.Cells.Item(1,6) = "Patched"
$WorkSheet.Cells.Item(1,7) = "Not Patched"
$WorkSheet.Cells.Item(1,8) = "Not Available"

$range = $WorkSheet.UsedRange
$range.Interior.ColorIndex = 19
$range.Font.ColorIndex = 11
$range.Font.Bold = $True

$WorkSheet.Cells.Item(2,6) = "0"
$WorkSheet.Cells.Item(2,7) = "0"
$WorkSheet.Cells.Item(2,8) = "0"

$WorkSheet.Cells.Item(2,6).Interior.ColorIndex = 4
$WorkSheet.Cells.Item(2,7).Interior.ColorIndex = 3
$WorkSheet.Cells.Item(2,8).Interior.ColorIndex = 6

$intRow = 2

$notavailable = 0
$patched = 0
$notpatched = 0

Foreach ($ADobj in $ADList)
#Foreach ($strComputer in get-content $ScriptDir\MachineList.Txt)
{
	$strComputer = [string]$ADobj[0]
	$strDescription = [string]$ADobj[1]
	
	$WorkSheet.Cells.Item($intRow,1) = $strComputer
	$WorkSheet.Cells.Item($intRow,4) = $strDescription
	
	$PatchStatus = Get-WMIObject Win32_QuickFixEngineering -computer $strcomputer |where {$_.HotFixID -eq "KB958644"}
	if ($error) {
		#echo $error[0].FullyQualifiedErrorId
		$error.Clear()
		$notavailable += 1
		$WorkSheet.Cells.Item(2,8) = $notavailable
		$WorkSheet.Cells.Item($intRow,2).Interior.ColorIndex = 6
		$WorkSheet.Cells.Item($intRow,2) = "N/A"
	}
	Else
	{
		If($PatchStatus -eq $Null) {
			$notpatched += 1
			$WorkSheet.Cells.Item(2,7) = $notpatched
			$WorkSheet.Cells.Item($intRow,2).Interior.ColorIndex = 3
			$WorkSheet.Cells.Item($intRow,2) = "NO"
		}
		Else {
			$patched += 1
			$WorkSheet.Cells.Item(2,6) = $patched
			$WorkSheet.Cells.Item($intRow,2).Interior.ColorIndex = 4
			$WorkSheet.Cells.Item($intRow,2) = "YES"
		}
	}
	$WorkSheet.Cells.Item($intRow,3) = Get-Date
	$intRow = $intRow + 1
}
$range.EntireColumn.AutoFit()
$datestring = get-date -format "yyyyMMddHHmmss"
$Workbook.SaveAs("$ScriptDir\patchresults_$datestring.xls")
$Excel.quit()
