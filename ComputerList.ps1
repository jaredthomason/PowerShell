#===================================================================================================
#  Get Computer List
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
$logfile = C:\Scripts\POWERSHELL\ComputerDetails\logs\ACT_Servers.txt

# Canonical name of the OU you want to search through.
$ouToSearch = "faa.gov/Programs/National/Data_Centers/Servers/ACT_Servers"
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
$script:logfile = "C:\Scripts\POWERSHELL\ComputerDetails\logs\ATCComputerDetails-$(get-date -format MM-dd-yyyy-HH-mm-ss).txt"
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
	$WorkSheet.Cells.Item(1,2) = "Server FQDN"

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
		$WorkSheet.Cells.Item($intRow,2) = $Server.ServerFQDN
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
$ComputerList = Get-QADComputer -SearchRoot $ouToSearch -SearchScope 'Subtree' | sort-object name
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
	# Pull out the Server name and FQDN
	$strComputer = [string]$server.Name
	Write-Host "Gatering Data for $strComputer"
	$strFQDN = [string]$server.dNSHostName

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining whether the server object is disabled in Active Directory"
	# Determine whether or not the server is disabled
	$strServerDisabled = [string]$server.AccountIsDisabled

	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Determining the object type to ensure it is a computer object"
	# Determine the object type
	$strServerType = [string]$Server.Type

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
	Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Server is enabled and available."

		Write-log "$(get-date -format "MM-dd-yyyy HH:mm:ss") - Putting the details that have been acquired in an array to be used later"
		# Store the Computer Details we have acquired in an array to be used later.
		$objComputerDetails = New-Object System.Object
		$objComputerDetails | Add-Member -type NoteProperty -name ServerName -value $strComputer
		$objComputerDetails | Add-Member -type NoteProperty -name ServerFQDN -value $strFQDN
		$colComputerDetails += $objComputerDetails
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
		$strFQDN = ""
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