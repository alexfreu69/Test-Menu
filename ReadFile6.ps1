function Test-Menu()
{
    [CmdletBinding()]
    param($filename)


    If ($PSBoundParameters['Debug']) {$DebugPreference = 'Continue'}

    $ret = $false
    [byte[]] $b = [System.IO.File]::ReadAllBytes($filename)
    [byte[]] $SIG = 0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1
    [byte[]] $x=$b[0..7]

    $FileLength=$b.Length
    $Fragmented=$false

    Write-Debug "File Length: $FileLength"

    if([Linq.Enumerable]:: SequenceEqual($x, $SIG))
    {
         

        $Sectorsize = 1 -shl $b[0x1E]
        $MiniSectorSize = 1 -shl $B[0x20]

        Write-Debug "Sector Size: $sectorsize"
        Write-Debug "Mini Sector Size: $MiniSectorSize"

        $NumDirSects = [System.BitConverter]::ToUInt32($b,0x28)
        $NumFATSects = [System.BitConverter]::ToUInt32($b,0x2C)
        $NumMiniFATSects = [System.BitConverter]::ToUInt32($b,0x40)
        $DirSect1 = [System.BitConverter]::ToUInt32($b,0x30)
        $MiniStrMax = [System.BitConverter]::ToUInt32($b,0x38)
        $MiniFATSect1 = [System.BitConverter]::ToUInt32($b,0x3C)
        $DiFATSect1 = [System.BitConverter]::ToUInt32($b,0x44)
        $NumDiFATSects = [System.BitConverter]::ToUInt32($b,0x48)
        

        $DirectoryEntries = ($DirSect1 + 1) * $sectorsize

        Write-Debug ("Start of Directory: $([string]::Format("0x{0:x}",$DirectoryEntries))")
        Write-Debug ("Mini Stream Cutoff Size: $MiniStrMax")

        $MiniStreamStartSect = [System.BitConverter]::ToUInt32($b,$DirectoryEntries + 116)
        

        $WorkbookEntry = $DirectoryEntries + 128 # 2nd Entry should be \Root Entry\Workbook
        [byte[]] $WB = $b[$WorkbookEntry..($WorkbookEntry+15)]
        [byte[]] $OLDWB = $b[$WorkbookEntry..($WorkbookEntry+7)]
        [byte[]] $WBSIG = 0x57,0x00,0x6F,0x00,0x72,0x00,0x6B,0x00,0x62,0x00,0x6F,0x00,0x6F,0x00,0x6B,0x00
        [byte[]] $OLDSIG = 0x42,0x00,0x6F,0x00,0x6F,0x00,0x6B,0x00

        $WorkbookFound = [Linq.Enumerable]::SequenceEqual($WB, $WBSIG)
        $OldWorkbookFound = [Linq.Enumerable]::SequenceEqual($OLDWB, $OLDSIG)

        #$t=[System.BitConverter]::ToString($WB)

        Write-Debug $OldWorkbookFound
        
        if (!$WorkbookFound -and !$OldWorkbookFound) { 
            Write-Debug "Workbook substream not found!"
            return $false
        }


        $WBStreamSize =[System.BitConverter]::ToUInt32($b,$WorkbookEntry + 120)

        Write-Debug ("Stream Size: $WBStreamSize")

        $WBStreamStartSect = [System.BitConverter]::ToUInt32($b,$WorkbookEntry + 116)

        [byte[]] $BIFFStream = @()

        [byte[]] $FAT = @()
        for ($i=0;$i -lt $NumFATSects;$i++)
        {
            $DiFATEntry = [System.BitConverter]::ToUInt32($b,76+4*$i)
            $FATStart = ($DiFATEntry + 1 ) * $sectorsize
            $FAT += $b[$FATstart..($FATStart+$sectorsize-1)]
        }

        if ($WBStreamSize -ge $MiniStrMax)
        { 

            Write-Debug "Using FAT"

            # Write-Debug "FAT: $FAT"
            Write-Debug "Assembling Workbook Stream ..."

            [byte[]] $BIFFStream = [byte[]]::new([Math]::Ceiling($WBStreamSize / $sectorsize)*$Sectorsize)

            [uint32] $nextp = $WBStreamStartSect
            $n=0
            do
            {
                [uint32] $p = [uint32] $nextp
                [uint32] $nextp = [System.BitConverter]::ToUInt32($FAT,$p * 4)
                # Write-Debug "Sector: $p $nextp"
                [byte[]]$arr=$b[(($p+1)*$sectorsize)..(($p+2)*$sectorsize-1)]
                #$BIFFStream += $b[(($p+1)*$sectorsize)..(($p+2)*$sectorsize-1)] # Slow!!!
                $arr.CopyTo($BIFFStream,$n)
                $n+=$sectorsize

            }
            until ([uint32]$nextp -eq [uint32]"0xFFFFFFFE")

        } 

        else

        {

            Write-Debug "Using MiniFAT"
            [byte[]] $MiniFAT = @()
            [uint32] $nextp = $MiniFATSect1
            do
            {
                [uint32] $p = [uint32] $nextp
                [uint32] $nextp = [System.BitConverter]::ToUInt32($FAT,$p * 4)
                # Write-Debug "Sector: $p $nextp"
                $MiniFAT += $b[(($p+1)*$sectorsize)..(($p+2)*$sectorsize-1)]
            }
            until ([uint32]$nextp -eq [uint32]"0xFFFFFFFE")

            # Write-Debug ("MiniFAT: $MiniFAT")

            Write-Debug ("Mini Stream Start Sector: $MiniStreamStartSect")

            [byte[]] $MiniStream=@()
            [uint32] $nextp = $MiniStreamStartSect
            do
            {
                [uint32] $p = [uint32] $nextp
                [uint32] $nextp = [System.BitConverter]::ToUInt32($FAT,$p * 4)
                # Write-Debug "Sector: $p $nextp"
                $MiniStream += $b[(($p+1)*$sectorsize)..(($p+2)*$sectorsize-1)]
            }
            until ([uint32]$nextp -eq [uint32]"0xFFFFFFFE")

            # Write-Debug ("MiniStream: $MiniStream")
            # Write-Debug ("MiniStream Length: $($MiniStream.Length)")


            [uint32] $nextp = $WBStreamStartSect
            do
            {
                [uint32] $p = [uint32] $nextp
                [uint32] $nextp = [System.BitConverter]::ToUInt32($MiniFAT,$p * 4)
                # Write-Debug "Sector: $p $nextp"
                $BIFFStream += $MiniStream[($p*$MiniSectorSize)..(($p+1)*$MiniSectorSize-1)]
            }
            until ([uint32]$nextp -eq [uint32]"0xFFFFFFFE")
        }

        # Write-Debug ("BIFF: $BIFFStream")

        $RecordType = [System.BitConverter]::ToUInt16($BIFFStream,0)
        $BIFF = ""
        switch ($RecordType)
        {
            0x0009 {$BIFF="BIFF2"}
            0x0209 {$BIFF="BIFF3"}
            0x0409 {$BIFF="BIFF4"}
            0x0809 {$BIFF="BIFF5"}
        }
        Write-Debug $BIFF

        $loc = 0
        $StartMenuRecs = 0
        do {
          $RecWord = [System.BitConverter]::ToUInt16($BIFFStream,$loc)
          $loc += 2
          $LenWord = [System.BitConverter]::ToUInt16($BIFFStream,$loc)
          $loc += 2
          Write-Debug "$RecWord $LenWord"
          $loc += $LenWord 
          if ($RecWord -eq 193) # 193 = MMS
          {
            $StartMenuRecs = $loc
            break
          }

          # Write-Debug "Location: $loc"
          if ($loc -gt $FileLength)
          {
            $Fragmented = $true
            break
          }
        } until ($RecWord -eq 10)

        if ($Fragmented)
        {
            Write-Debug "File is fragmented. Unable to parse BIFF record."
            return $false
        }

        $MenuEditCount=0
        if($StartMenuRecs -ne 0)
        {
            $loc = $StartMenuRecs
            do
            {
                $RecWord = [System.BitConverter]::ToUInt16($BIFFStream,$loc)
                $loc += 2
                $LenWord = [System.BitConverter]::ToUInt16($BIFFStream,$loc)
                $loc += 2
                Write-Debug "$RecWord $LenWord"
                $loc += $LenWord
                if ($RecWord -eq 194 -or $RecWord -eq 195) { $MenuEditCount++ }

            } while ($RecWord -eq 194 -or $RecWord -eq 195)
        }
        Write-Debug $MenuEditCount
        return ($MenuEditCount -gt 0)
    }
    return $ret
}



Test-Menu "C:\TEMP\CS0024669\test5.XLS" -Debug 


