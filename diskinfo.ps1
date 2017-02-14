$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "   Server Name   "
$c.Cells.Item(1,2) = "    Disk Letter  "
$c.Cells.Item(1,3) = " Total Size "
$c.Cells.Item(1,4) = "   Used       "
$c.Cells.Item(1,5) = "   Free   "
$c.Cells.Item(1,6) = "   Total Files"
$c.Cells.Item(1,7) = " Volume Name   "

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit()
$m = 2

$x = get-content C:\Scripts\powershell\diskinfo\list.txt
Write-Host "Getting List from File"


foreach ($i in $x)
{
  $disks =  gwmi -computername $i win32_logicaldisk -filter "drivetype=3"
  foreach ($disk in $disks) {
   write-host "disk id:" $disk.deviceid
   
   
     if ($disk.deviceid -eq "D:")
	 
	 {
     $used = ([int64]$disk.size - [int64]$disk.freespace)
	 $c.Cells.Item($m, 1)= $i
     $c.Cells.Item($m, 2) = $disk.deviceid
	 $c.Cells.Item($m, 3)= "{0:0.0} gb" -f ($disk.size/1gb)
	 $c.Cells.Item($m, 4)= "{0:0.0} gb" -f ($used/1gb)
	 $c.Cells.Item($m, 5)= "{0:0.0} gb" -f ($disk.freespace/1gb)
	 $c.Cells.Item($m, 6)= "Getting SIZE"
	 $filesi=0
	 
#     $GciFiles = get-Childitem \\$i\d$ -Recurse -force
#     foreach ($file in $GciFiles) {$filesi++}
#      {$GciFiles |sort |ft name, attributes -auto}
#	   $c.Cells.Item($m, 6)= $filesi 
#	   $c.Cells.Item($m, 7)= $disk.VolumeName
	  
     $m = $m + 1
	 }
 	 Write-Host "Incrementing M for next line"
    write-host " Trying NEXT in File"
	Write-Host "-----------------------------------------------"
}
}
