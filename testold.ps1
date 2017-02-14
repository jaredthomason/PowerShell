$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "   Machine Name   "
$c.Cells.Item(1,2) = "    IP Address    "
$c.Cells.Item(1,3) = "     MAC Address     "
$c.Cells.Item(1,4) = "   Last Boot Time   "

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit()
$m = 2

$x = get-content C:\Scripts\powershell\lastreboot\Clientlist.txt
foreach ($i in $x)
{$y = get-wmiobject Win32_NetworkAdapterConfiguration -computername $i -Filter "IPenabled = 'True'"
foreach ($j in $y)
{$c.Cells.Item($m, 1) = $j.DNSHostName
$c.Cells.Item($m, 2) = $j.IPAddress
$c.Cells.Item($m, 3) = $j.MACAddress}
$date = new-object -com WbemScripting.SWbemDateTime
$z = get-wmiobject Win32_OperatingSystem -computername $i
foreach ($k in $z)
{$date.value = $k.lastBootupTime
If ($k.Version -eq "5.2.3790" )
{$c.Cells.Item($m, 4) = $Date.GetVarDate($True)}
Else
{$c.Cells.Item($m, 4) = $Date.GetVarDate($False)}
}
$m = $m + 1
}