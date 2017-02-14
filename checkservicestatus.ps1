######################################################################################
# Service status report
# --------------------------
# This is to take report of particular service status in multiple server
#
# ---------------------------------------------------------------------------------
# Author: Senthil Kumar C
# Input server list in C:\Systems.txt  and Output will be saved in C:\Report.txt
######################################################################################

function RSCDsvc {
Process{

Trap {
Continue}


$pingresult = gwmi -Query "select * from win32_pingstatus where address = '$_'" -ErrorAction Stop

$obj = New-Object psobject
$obj | Add-Member noteproperty ServerName $_

if($pingresult.statuscode -eq 0) {

Trap {
Continue}


$service = gwmi win32_service -ComputerName $_ -filter "name = 'RSCDsvc'" -ErrorAction Stop

$obj | Add-Member Noteproperty ServiceName ($service.name)
$obj | Add-Member Noteproperty Status ($service.status)
$obj | Add-Member Noteproperty Reachable Reachable

}

else {


$obj | Add-Member noteproperty Reachable Notreachable

}

write-output $obj

}
}

gc C:\Scripts\POWERSHELL\CheckServiceStatus\systems.txt | RSCDsvc | ft ServerName,ServiceName,Status,Reachable -AutoSize | Out-File C:\Scripts\POWERSHELL\CheckServiceStatus\Report.txt
