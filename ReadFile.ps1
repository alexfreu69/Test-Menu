function Test-Menu()
{
    [CmdletBinding()]
    param($filename)

    If ($PSBoundParameters['Debug']) {$DebugPreference = 'Continue'}

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

        $RecordType = [System.BitConverter]::ToUInt16($b,$sectorsize)
        $BIFF = ""
        switch ($RecordType)
        {
            0x0009 {$BIFF="BIFF2"}
            0x0209 {$BIFF="BIFF3"}
            0x0409 {$BIFF="BIFF4"}
            0x0809 {$BIFF="BIFF5"}
        }

        $loc = $sectorsize
        $StartMenuRecs = 0
        do {
          $RecWord = [System.BitConverter]::ToUInt16($b,$loc)
          $loc += 2
          $LenWord = [System.BitConverter]::ToUInt16($b,$loc)
          $loc += 2
          Write-Debug ($RecWord.ToString() + " " + $LenWord.ToString())
          $loc += $LenWord 
          if ($RecWord -eq 193) # 193 = MMS
          {
            $StartMenuRecs = $loc
            break
          }
        } until ($RecWord -eq 10)

        $MenuEditCount=0
        if($StartMenuRecs -ne 0)
        {
            $loc = $StartMenuRecs
            do
            {
                $RecWord = [System.BitConverter]::ToUInt16($b,$loc)
                $loc += 2
                $LenWord = [System.BitConverter]::ToUInt16($b,$loc)
                $loc += 2
                Write-Debug ($RecWord.ToString() + " " + $LenWord.ToString())
                $loc += $LenWord
                if ($RecWord -eq 194 -or $RecWord -eq 195) { $MenuEditCount++ }

            } while ($RecWord -eq 194 -or $RecWord -eq 195)
        }
        return ($MenuEditCount -gt 0)
    }
    return $ret
}



Test-Menu "c:\temp\mappe1.xls" -Debug
