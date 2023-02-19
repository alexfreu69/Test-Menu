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
    Write-Host ([Linq.Enumerable]:: SequenceEqual($x, $SIG))
    Write-Host $x
    Write-Host $Sig
    return $ret
}

Test-Menu "c:\temp\mappe1.xls"