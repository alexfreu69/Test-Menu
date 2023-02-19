function Read1()
{
    for ($i=1;$i -lt 10;$i++){
        Get-Content -LiteralPath c:\temp\mappe1.xls -Encoding Byte
    }
}

function Read2()
{
    for ($i=1;$i -lt 10;$i++){
        [System.IO.File]::ReadAllBytes("c:\temp\mappe1.xls")
    }
}


#(Measure-Command {Read1}).Milliseconds
#(Measure-Command {Read2}).Milliseconds

function Test-Menu($filename)
{
    $ret = $false
    [byte[]] $b = [System.IO.File]::ReadAllBytes($filename)
    [byte[]] $SIG = 0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1
    [byte[]] $x=$b[0..7]
    if([Linq.Enumerable]:: SequenceEqual($x, $SIG))
    {
        Write-Host $b[0x1E] # Sector Size   
        if($b[0x1E] -eq 9)
        {
            $sectorsize = 512

        }
        elseif ($b[0x1E] -eq 12)
        {
            $sectorsize = 4096
        }   
        $NumDirSects = [System.BitConverter]::ToUInt32($b,0x28)
        $NumFATSects = [System.BitConverter]::ToUInt32($b,0x2C)
        $NumMiniFATSects = [System.BitConverter]::ToUInt32($b,0x40)
        $DirSect1 = [System.BitConverter]::ToUInt32($b,0x30)
        $MiniFATSect1 = [System.BitConverter]::ToUInt32($b,0x3C)
        $DiFATSect1 = [System.BitConverter]::ToUInt32($b,0x44)
        Write-Host $NumDirSects, $NumFATSects, $NumMiniFATSects,$DirSect1,$MiniFATSect1,$DiFATSect1
    }
    return $ret
}

Test-Menu "c:\temp\mappe1.xls"