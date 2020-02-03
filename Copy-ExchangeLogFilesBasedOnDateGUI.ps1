Function SKFileCopy {
    [CmdletBinding()]
    Param(
        [String]$Activity = "Browsing each server to get Log files copied",
        [String()]$ServerNamesList = @("E2010","E2016-01","E2016-02"),
        [switch]$ListFilesOnly = $False,
        [string]$SourceDirectory = "C:\Program Files\Microsoft\Exchange Server\V14\Logging\RPC Client Access\",
        [string]$SourceFiles = "*.LOG",
        [string]$StrTarget= "c:\temp\ExchangeRCALogs\",
        [datetime]$StartDate = (Get-date -Year 2020 -Month 01 -day 20),
        [datetime]$EndDate = (Get-date -Year 2020 -Month 01 -day 24)
    )
    #Variable defined to describe the activity type - used for Write-Progress
    ##### Parameterized ## $Activity = "Browsing each server to get Log files copied" 
    #Variable defined to hold the list of servers we want the files copied from - Hostnames only, use your own servers list
    ##### Parameterized ## $ServerNamesList = @("E2010","E2016-01","E2016-02")
    #Variable to tell the script whether we only list the files or we copy them to the target - Set to $True => List only, set to $False => Copy files to the defined $StrTarget
    ##### Parameterized ## $ListFilesOnly = $False
    #Variable to point to the Exchange Install path on the remote servers - some are located on D:, some on C:, some below Program Files, others on the Root of another Driv. Get yours on your servers by checking on Powershell the $ExchangeInstallPath variable
    #### Replaced by parameter SourceDirectory ### $SourceExchangeInstallPath = "C`$\Program Files\Microsoft\Exchange Server\V14"
    #Variable to point to the Logging subfolder with wildcard on the LOG files we need for this program
    ##### Rpelaced by parameter SourceDirectory and Sourcefiles ### $ExchangeRCALogsPath = $SourceExchangeInstallPath + "\Logging\RPC Client Access\*.LOG"
    #Variable to point to the target directory
    ##### Replaced by parameter ### $StrTarget= "c:\temp\ExchangeRCALogs\"
    #Storing StartDate and EndDate in PowerShell format, using -Year -Month -Day to get rid of Date format MM/DD/YYYY DD/MM/YYYY ambiguity
    ##### Replaced by parameret ###  $StartDate = (Get-date -Year 2020 -Month 01 -day 20)
    ##### Replaced by parameret ###  $EndDate = (Get-date -Year 2020 -Month 01 -day 24)

    if (-not (Test-Path $StrTarget)){
        Write-Host "$StrTarget does not exist... creating it."
        New-Item -ItemType Directory -Path $StrTarget
    } Else {
        Write-Host "$StrTarget is here ! Continuing ..."
    }

    #Test if Source Directory ends with backslash - if it doesn't, append backslash
    If ($SourceDirectory -notmatch "\\$"){
        $SourceDirectory += "\"
    }

    $SourceDirectoryAndFilesToCopy = $SourceDirectory + $SourceFiles

    $ServerCounter = 0
    write-Host "Now Browsing each server..."
    Foreach ($ServerName in $ServerNamesList){
        Write-Progress -Activity $Activity -Status "Working on server $ServerName" -PercentComplete $($ServerCounter / $($ServerNamesList.Count)*100)
        sleep 2
        $StrSource ="\\$ServerName\$SourceDirectoryAndFilesToCopy"
        $FilesSearchResults = Get-ChildItem $StrSource | Where-Object {($_.LastWriteTime.Date -ge $StartDate.Date) -and ($_.LastWriteTime.Date -le $EndDate.Date)}

        Write-Host "Found $($FilesSearchResults.count) files." -BackgroundColor Red
        If ($ListFilesOnly){
            Write-Host "-ListFilesOnly switch specified - Listing files only on server $ServerName!" -BackgroundColor Green
            Write-Host "would copy from $StrSource to $($StrTarget)$($ServerName)_<FileName>"
        } Else {
            Write-Host "-ListFilesOnly switch NOT specified - Copying files to $StrTarget from server $ServerName" -BackgroundColor Green
            Write-Host "Copying now..." -BackgroundColor Blue -ForegroundColor Yellow
            $FilesSearchResults | Foreach {Copy-Item -Path $_ -Destination "$($StrTarget)$($ServerName)_$($_.Name)"}
        }
        $ServerCounter ++
    }
}
