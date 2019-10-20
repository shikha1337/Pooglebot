
$file = Read-Host "Excel Sheet Location"

$excel = New-Object -ComObject Excel.Application
$wb = $excel.workbooks.open($file)
$sheet = $wb.Worksheets.Item(1)
$rowmax = ($sheet.UsedRange.Rows).count
$ip = New-Object -TypeName psobject
$ip | Add-Member -MemberType NoteProperty -Name "Google Bot IP" -Value $null
$array = @()
for ($i = 2;$i -le $rowmax; $i++)
{ $temp = $ip | Select-Object *
  $temp."Google Bot Ip" = $sheet.Cells.Item($i,1).Text
  $array += $temp
}
$errpref = $ErrorActionPreference
$ErrorActionPreference = "silentlycontinue"
$array_2 = @("IP")
foreach ($x in $array)
{$ip = $x | select -ExpandProperty "Google Bot Ip"
 $check = Resolve-DnsName $ip | Select -ExpandProperty NameHost -Verbose false
 if ($check -like "*.googlebot.com" -or $check -like "*.google.com")
 { 
    $array_2 += $ip
    
 }
 $excel.Quit()

}
$save = Read-Host "Enter Name for saved file"
if (Test-Path $save)
    { 
        del $save
        $array_2 | foreach { Add-Content -Path $save -Value $_}
    }
else
    {
        $array_2 | foreach { Add-Content -Path $save -Value $_}

    }
$ErrorActionPreference = $errpref
