

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "   Computer Name    "
$c.Cells.Item(1,2) = "        OS Version                                                   "
$c.Cells.Item(1,3) = "   Hotfix              "
$c.Cells.Item(1,4) = "   TItle 4 "
$c.Cells.Item(1,5) = "   Install ORG               "
$c.Cells.Item(1,6) = "   Product ID                   "
$c.Cells.Item(1,7) = "   Install Key                                     "

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit()
$m = 2
$e = 3

$y = get-content C:\Scripts\powershell\cdkey\list.txt
Write-Host "Getting List from File"

function Get-WindowsKey {
    param ($targets = ".")
    $hklm = 2147483650
    $regPath = "Software\Microsoft\Windows NT\CurrentVersion"
    $regValue = "DigitalProductId"
    Foreach ($target in $targets) {
        $e = $e + 1
        $productKey = $null
        $win32os = $null
        $wmi = [WMIClass]"\\$target\root\default:stdRegProv"
        $data = $wmi.GetBinaryValue($hklm,$regPath,$regValue)
        $binArray = ($data.uValue)[52..66]
        $charsArray = "B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9"
        ## decrypt base24 encoded binary data
        For ($i = 24; $i -ge 0; $i--) {
            $k = 0
            For ($j = 14; $j -ge 0; $j--) {
                $k = $k * 256 -bxor $binArray[$j]
                $binArray[$j] = [math]::truncate($k / 24)
                $k = $k % 24
            }
            $productKey = $charsArray[$k] + $productKey
            If (($i % 5 -eq 0) -and ($i -ne 0)) {
                $productKey = "-" + $productKey
            }
        }
        $win32os = Get-WmiObject Win32_OperatingSystem -computer $target
        $obj = New-Object Object
        $obj | Add-Member Noteproperty Computer -value $target
        $obj | Add-Member Noteproperty Caption -value $win32os.Caption
        $obj | Add-Member Noteproperty CSDVersion -value $win32os.CSDVersion
        $obj | Add-Member Noteproperty OSArch -value $win32os.OSArchitecture
        $obj | Add-Member Noteproperty BuildNumber -value $win32os.BuildNumber
        $obj | Add-Member Noteproperty RegisteredTo -value $win32os.RegisteredUser
        $obj | Add-Member Noteproperty ProductID -value $win32os.SerialNumber
        $obj | Add-Member Noteproperty ProductKey -value $productkey
        $obj
        $c.Cells.Item($e,1) = $target
        $c.Cells.Item($e,2) = $win32os.Caption
        $c.Cells.Item($e,3) = $win32os.CSDVersion
        $c.Cells.Item($e,4) = $win32os.OSArchitecture
        $c.Cells.Item($e,5) = $win32os.RegisteredUser
        $c.Cells.Item($e,6) = $win32os.SerialNumber 
        $c.Cells.Item($e,7) = $productkey

    }
}

Get-WindowsKey $y

   





















