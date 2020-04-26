#########################  Creating a GUI for a script #########################

#region Step ? - place functions for the form...

Function Title1 {
    [CmdletBinding()]
    param (
        [parameter(position = 1)]
        [string]$Title,
        [int]$TotalLength = 100, 
        [string]$Back = "Blue",
        [string]$Fore = "Yellow"
    )
    Write-Host "BOOKMARK inside TITLE1"
    $TitleLength = $Title.Length
    [string]$StarsBeforeAndAfter = ""
    $RemainingLength = $TotalLength - $TitleLength
    If ($($RemainingLength % 2) -ne 0) {
        $Title = $Title + " "
    }
    $Counter = 0
    For ($i=1;$i -le $(($RemainingLength)/2);$i++) {
        $StarsBeforeAndAfter += "*"
        $counter++
    }
    
    $Title = $StarsBeforeAndAfter + $Title + $StarsBeforeAndAfter
    Write-host
    Write-Host $Title -BackgroundColor $Back -foregroundcolor $Fore
    Write-Host
    
}


Function StatusLabel {
    [CmdletBinding()]
    Param(  [parameter(Position = 1)][string]$msg,
            [parameter(Position = 2)][string]$LabelObjectName = "lblProgressBar"
    )
    Write-Host "BOOKMARK inside StatusLabel"
    # Trick to enable a Label to update during work :
    # Follow with "Dispatcher.Invoke("Render",[action][scriptblobk]{})" or [action][scriptblock]::create({})
    $wpf.$LabelObjectName.Content = $Msg
    $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
}

Function Update-WPFProgressBarAndStatus {
    [CmdletBinding(DefaultParameterSetName = "Normal")]
    Param(  [parameter(Position = 1, ParameterSetName = "Normal")][string]$msg="Message",
            [parameter(Position=2, ParameterSetName = "Normal")][int]$p=50,
            #[parameter(Position = 3, ParameterSetName = "Normal")][string]$status="Working...",
            [parameter(position = 4, ParameterSetName = "Normal")][string]$color,
            [parameter(position = 5, ParameterSetName = "Normal")][string]$ProgressBarName = "ProgressBar",
            [parameter(position = 5, ParameterSetName = "Reset")][switch]$Reset,
            [parameter(position = 6, ParameterSetName = "Normal")][array]$wpf = $wpf
            )
            Write-Host "BOOKMARK inside Update-WPFProgressBarAndStatus"
    IF ($Reset){
        write-Host "BOOKMARK - inside WPF Progress bar update - RESET requested" -fore green
        $wpf.$FormName.IsEnabled = $true
        $wpf.$ProgressBarName.Value = 0
        StatusLabel "Ready !"
    } Else {
        write-Host "BOOKMARK - inside WPF Progress bar update - setting progress bar value and title etx...." -fore green

        If ($color){
            $wpf.$ProgressBarName.Color = $Color
        }
        $wpf.$ProgressBarName.Value = $p
        #$wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
        # $wpf.$ProgressBarName.Foreground
        write-Host "ATTENTION" -ForegroundColor red
        Write-Host $msg -ForegroundColor red
        Title1 $msg; StatusLabel $msg
        # If ($p -eq 100){
        #     $msg = "Done!"
        #     $Status = "Done!"
        #     Title1 $msg; StatusLabel $msg
        # }
        #Commenting the below as it's already in the main script...
        #Write-progress -Activity $msg -Status $status -PercentComplete $p
    }
}


Function MsgBox {
    [CmdletBinding()]
    Param(
        [Parameter(Position=0)][String]$msg = "Do you want to continue ?",
        [Parameter(Position=1)][String]$Title = "Question...",
        [Parameter(Position=2)]
            [ValidateSet("OK","OKCancel","YesNo","YesNoCancel")]
                [String]$Button = "YesNo",
        [Parameter(Position=3)]
            [ValidateSet("Asterisk","Error","Exclamation","Hand","Information","None","Question","Stop","Warning")]
                [String]$Icon = "Question"
    )
    Add-Type -AssemblyName presentationframework, presentationcore
    [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)
}

#endregion Step ??


#region Step 1 - Convert your script to a Function

Function SKFileCopy {
    [CmdletBinding()]
    Param(
        [String]$Activity = "Browsing each server to get Log files copied",
        [String[]]$ServerNamesList = @("E2010","E2016-01","E2016-02"),
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
    ##### Replaced by paramere ###  $StartDate = (Get-date -Year 2020 -Month 01 -day 20)
    ##### Replaced by paramere ###  $EndDate = (Get-date -Year 2020 -Month 01 -day 24)

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

    #Test also if Target Directory ends with backslash - if it doesn't, append backslash
    If ($StrTarget -notmatch "\\$"){
        $StrTarget += "\"
    }

    $SourceDirectoryAndFilesToCopy = $SourceDirectory + $SourceFiles

    #Replace colons, most likely there should be one colon after the drive letter. If not, we keep the SourceDirectory as is
    $SourceDirectoryAndFilesToCopy = $SourceDirectoryAndFilesToCopy -replace ":","`$"

    $ServerCounter = 0
    write-Host "Now Browsing each server..."
    Foreach ($ServerName in $ServerNamesList){
        #Updating cmd shell progress bar as we go through servers...
        $Status = "Working on server $ServerName"
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $($ServerCounter / $($ServerNamesList.Count)*100)
        #Updating Form (for GUI only) progress bar as we go through servers...
        try {
            Update-WPFProgressBarAndStatus -msg $Status -p $($ServerCounter / $($ServerNamesList.Count)*100)
        } catch {
            # no action on GUI progress bar if not in GUI
        }
        Start-Sleep 2
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
    #Updating cmd shell final progress bar to get to 100% (because $ServerCounter++ is after the first Write-Progress)
    Write-Progress -Activity $Activity -Status "Done !" -PercentComplete $($ServerCounter / $($ServerNamesList.Count)*100)
    #Updating WPF progress bar as well to get to 100%
    Update-WPFProgressBarAndStatus -msg "Done !" -p 100
}

#endRegion step1


#region Step 2 - Create the form and assign the script to a button (or buttons to run the script with and without a switch)

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"
<Window x:Name="frmSKFileCentralize" x:Class="CopyFilesFromMultipleServersToCentralLocation.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:CopyFilesFromMultipleServersToCentralLocation"
        mc:Ignorable="d"
        Title="SK Copy - Copy Files From Different Servers" Height="531.362" Width="829.1">
    <Grid>
        <TextBox x:Name="txtServersListCSV" HorizontalAlignment="Left" Height="133" TextWrapping="Wrap" Text="Server1, Server2, Server3" VerticalAlignment="Top" Width="378" Margin="10,41,0,0"/>
        <Label x:Name="lblServersListCSV" Content="List of servers here: comma separated or load list from txt file" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
        <Button x:Name="btnLoadFromFile" Content="[Optional] Load servers list from file" HorizontalAlignment="Left" Margin="393,41,0,0" VerticalAlignment="Top" Width="159" Height="90" Cursor="Hand" VerticalContentAlignment="Center">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <TextBox x:Name="txtSourceFolderPath" HorizontalAlignment="Left" Height="69" Margin="10,233,0,0" TextWrapping="Wrap" Text="c:\Program Files\Microsoft\Exchange Sever\V14\Logging\RPC Client Access" VerticalAlignment="Top" Width="378"/>
        <Label x:Name="lblRemoteSourceFolderPath" Content="Source Directory on Remote Servers:" HorizontalAlignment="Left" Margin="10,202,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblFileExtension" Content="Files to look for:" HorizontalAlignment="Left" Margin="404,202,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtFileExtension" HorizontalAlignment="Left" Height="69" Margin="404,233,0,0" TextWrapping="Wrap" Text="*.LOG" VerticalAlignment="Top" Width="191"/>
        <Label x:Name="lblTargetCentralDir" Content="Central Target Directory" HorizontalAlignment="Left" Margin="10,319,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtTargetCentralDir" HorizontalAlignment="Left" Height="47" Margin="10,350,0,0" TextWrapping="Wrap" Text="C:\temp\ExchangeRCALogs\" VerticalAlignment="Top" Width="378"/>
        <Button x:Name="btnTestFiles" Content="Test Files Source" HorizontalAlignment="Left" Margin="454,349,0,0" VerticalAlignment="Top" Width="123" Height="48" Cursor="Hand" VerticalContentAlignment="Center">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <Button x:Name="btnCopyFiles" Content="Copy Files" HorizontalAlignment="Left" Margin="633,350,0,0" VerticalAlignment="Top" Width="123" Height="48" Cursor="Hand" VerticalContentAlignment="Center">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <ProgressBar x:Name="ProgressBar" HorizontalAlignment="Left" Height="16" Margin="26,464,0,0" VerticalAlignment="Top" Width="766"/>
        <Label x:Name="lblProgressBar" Content="Ready." HorizontalAlignment="Left" Margin="26,433,0,0" VerticalAlignment="Top" Width="766"/>
        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="152" Margin="583,22,0,0" VerticalAlignment="Top" Width="173"/>
        <DatePicker x:Name="EndDatePicker" HorizontalAlignment="Left" Margin="601,136,0,0" VerticalAlignment="Top" DisplayDate="2020-01-01"/>
        <Label x:Name="lblEndDate" Content="End date of files to copy:" HorizontalAlignment="Left" Margin="601,105,0,0" VerticalAlignment="Top" Width="155"/>
        <DatePicker x:Name="StartDatePicker" HorizontalAlignment="Left" Margin="601,62,0,0" VerticalAlignment="Top" DisplayDate="2020-01-01"/>
        <Label x:Name="lblStartDate" Content="Start date of files to copy:" HorizontalAlignment="Left" Margin="601,31,0,0" VerticalAlignment="Top" Width="155"/>

    </Grid>
</Window>

"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean

# Read the XAML code
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)

# Populate the Hash table $wpf with the Names / Values pairs using the form control names
# Form control objects will be available as $wpf.<Form control name> like $wpf.RunButton for example...
# Adding an event like Click or MouseOver will be with $wpf.RunButton.addClick({Code})
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

# Seen another method where the developper creates variables for each control instead of using a hash table
# $wpf {Key name, Value}, he uses Set-Variable "var_$($_.Name)" with value $TempForm.FindName($_.Name) instead of $HashTable.Add($_.Name,$tempForm.FindName($_.Name)):
#
#       $NamedNodes = $xaml.SelectNodes("//*[@Name]") 
#       $NamedNodes | Foreach-Object {Set-Variable -Name "var_$($.Name)" -Value $tempform.FindName($_.Name) -ErrorAction Stop}
#
# that way, each control will be accessible with the variable name named $var_<control name> like $var_btnQuery
# we would add events like Click or MouseOver using $var_btnQuery.addClick({Code})
# more info there: https://adamtheautomator.com/build-powershell-gui/

#Get the form name to be used as parameter in functions external to form...
$FormName = $NamedNodes[0].Name

#region Other functions needed to manage the form...


#endregion
#end of region Other functions needed...

#Define events functions
#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.$FormName.Add_Loaded({
    #Load default variables for the form fields, even if these are already on the XAML code (I prefer like this)
    $wpf.txtServersListCSV.Text = "HarounElPoussah, HarounElPoussah, HarounElPoussah"
    $wpf.txtSourceFolderPath.Text = "C:\temp\ExchangeRCALogsSubset"
    $wpf.txtTargetCentralDir.Text = "C:\temp\GUIScriptTest" + (Get-Date -Format ddMMyyyyhhmmss)
    $wpf.StartDatePicker.SelectedDate = Get-Date -Year 2020 -Month 01 -Day 20
    $wpf.EndDatePicker.SelectedDate = Get-Date -Year 2020 -Month 01 -Day 30
})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.$FormName.Add_ContentRendered({
    #Update-Cmd
})
$wpf.$FormName.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})
#endregion Load, Draw and closing form events
#End of load, draw and closing form events

#region Button events
$wpf.btnLoadFromFile.add_Click({
    MsgBox -msg "Function not implemented yet :-)" -Button OK -Icon Information -Title "Coming soon..."
})
$wpf.btnCopyFiles.add_Click({
    MsgBox -msg "Function not implemented yet :-)" -Button OK -Icon Information -Title "Coming soon..."
})
$wpf.btnTestFiles.add_Click({
    # Update progressbar label ...
    $wpf.lblProgressBar.Content = "Please wait..."
    $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
    if ([string]::IsNullOrEmpty($wpf.StartDatePicker.SelectedDate) -or [string]::IsNullOrEmpty($wpf.EndDatePicker.SelectedDate)) {
        MsgBox -msg "Either Start Date or End Date or both have not been selected ... please select date and try again" -Title "Missing dates" `
            -Button OK -Icon Error
    } Else {
        #Transforming CSV files list into an array of servers
        [array]$ServersList = $wpf.txtServersListCSV.Text -split "\s+|,\s*" -ne ''
        #Calling the SKFileCopy functions, passing the form inputs as arguments
        SKFileCopy -ServerNamesList $ServersList `
            -ListFilesOnly `
            -SourceDirectory $wpf.txtSourceFolderPath.Text `
            -SourceFiles $wpf.txtFileExtension.Text `
            -Strtarget $wpf.txtTargetCentralDir.Text `
            -StartDate $wpf.StartDatePicker.SelectedDate `
            -EndDate $wpf.EndDatePicker.SelectedDate
    }
    #Reset Progress Bar Label
    Update-WPFProgressBarAndStatus -Reset
})
#endregion


#HINT: to update progress bar and/or label during WPF Form treatment, add the following:
# ... to re-draw the form and then show updated controls in realtime ...
$wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


# Load the form:
# Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.$FormName.Dispatcher.InvokeAsync({
    $wpf.$FormName.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null

#endregion Step 2