param ([Parameter(Mandatory = $true,
    HelpMessage = "Enter the name of the Quicken data file; e.g., Home.qdf : ")]
    [System.IO.FileInfo]
    $FileName,

    [Parameter(Mandatory = $false)]
    [Switch]
    $Speak,

    [Parameter(Mandatory = $false)]
    [Switch]
    $DebugMessages)

    write-host "the value of Filename is $Filename"
    write-host "the value of Speak is $Speak"
<#
Invoke like;
powershell.exe -noprofile -file $runThis  -Filename "Home.qdf" -Speak
#>
if ($DebugMessages){Set-PSDebug -strict -trace 2}
($ThisVersion="V3.0.0")
<#
The name of this script is "LoadQuickenDb.ps1"
2017-08-20 - Copyright 2017 FAJ

One day I should put in the standard established for Powershell Headers.
Mod 2017-10-15 - Push *.dat files to the dat subfolder.
Mod 2017-11-19 - When Quicken exits bring this window to the foreground.
Added function 'Show-Process($Process, [Switch]$Maximize)'
Mod 2018-05-24 'Loop on Read-Host
2019-02-19 FAJ
    Changed launch of Powershell using an Alias in profile.ps1 as
      #function fQuickenHome {$command='C:\Users\Super` Computer\Dropbox\Private\Q\LoadQuickenDb.ps1 home.qdf -Speak';start-process powershell -argumentlist $command;remove-variable command}
      function fQuickn ($arg="home") {$command=join-path $env:HOMEPATH -ChildPath \Dropbox\Private\Q\LoadQuickenDb.ps1; `
      start-process powershell -Args "-noprofile -command & {. '$($command)' $arg.qdf -speak }";remove-variable command}
      set-alias -name Quickn -value fQuickn -Option Readonly -passthru | format-list
      or from a shortcut as C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe "& 'C:\Users\Super Computer\Dropbox\Private\Q\LoadQuickenDb.ps1'  home.qdf -Speak"
    Added tracing of arguments
2019-02-23 FAJ
    If file already in destination folder show dates
    Added comment lines so they show in log.
2019-02-24 FAJ
    $SourceDir is now based on "$MyInvocation.MyCommand.Path". Assumption is the script and the files reside in the Repository Workspace.
2019-04-09 FAJ V2.15
    Are we running CORE?
    using if ($psversiontable.psedition -ne "CORE") but this may change in the future when Powershell is renamed.
    Perhaps the executable name should be tested.
    Then
        Cannot set $bSayit to $true
        Cannot load [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($DestinationPath, 'OnlyErrorDialogs', 'SendToRecycleBin')

2019-05-31 FAJ V3.0.0
    Added params to this script.
    It should be called like this
    powershell.exe -noprofile -file $FullPathToScriptFile -Filename DataFileLikeHOME.QDF -Speak
    *** We assume the DataFile is in the folder (the REPOSITORY) where the script resides.
#>

<#
This script invokes Quicken and requires 2 arguments on the command line invoking it.
The first argument is the name of a Quicken data file.
The second argument is the a string with indicates if you want to enable text-to-speech prompts; Speak | NoSpeak
Here is an example of the Target property in a SHORTCUT on my Desktop;
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe "& C:\Users\Super` Computer\Documents\Dropbox\Private\Q\LoadQuickenDb.ps1 home.qdf -Speak"
Notice
1)the &
2)the escape after the word 'Super' as there is a space in the path.


This script invokes Quicken with a COPY of your data file. I refer to this file as $FileName

I keep my data file(s) in a folder called Dropbox\Private\Q\. I refer to this folder and file as the repository workspace- $SourceDir and $SourcePath
The copy is placed in the run-time workspace. This workspace is on the local machine and is not in a path that uses any cloud services.
Quicken uses $env:homepath/Documents/Quicken as the run-time workspace.
I refer to this folder and file as $DestinationDir and $DestinationPath
If it already is on the desktop you can overwrite it or use the existing file on the desktop.

When Quicken exits you decide if you want to delete the copy or move it back to the repository. Moving to back to the repository OVERWRITES the original file.
If you delete the file it is moved to the RecycleBin.
#>

# Begin
#Before starting the Transcript we need to know where the run-time workspace is located.
Write-Host "**************Beginning Script" -ForegroundColor "Yellow"
Set-StrictMode -Version latest #Before careful. I don't know all the implications of this setting!

#Find the run-time workspace.
$DestinationDir = Join-Path $env:HOMEDRIVE$env:HOMEPATH "Documents\Quicken" #This is where Quicken likes the run-time file to be.
write-host "Does destination folder exist?"
if (test-path $DestinationDir) {#it exists
    write-host "Destination folder $DestinationDir exists"
}
else {#destination folder does not exist
    #it may be better to exit the program then to create this Folder.
    write-host "the destination folder does not exist"
    write-host "creating destination folder"
    new-item -path $DestinationDir -itemtype directory # was it created?
}
$TranscriptName="Powershell.out"
$TranscriptNameOld=$TranscriptName + ".old"

if (test-path (Join-path $DestinationDir $TranscriptNameold)) {remove-item (Join-path $DestinationDir $TranscriptNameold)}
if (test-path (Join-path $DestinationDir $TranscriptName)) {rename-item -path (Join-path $DestinationDir $TranscriptName) -newname (Join-path $DestinationDir $TranscriptNameold)}
Start-Transcript -path (join-path $DestinationDir $TranscriptName) -IncludeInvocationHeader #-OutputDirectory $DestinationDir

#Write-Host "**************Beginning Script**************" -ForegroundColor yellow
write-host (get-process -id $pid).processname
get-process -id $pid | format-list -property *
write-host $(Get-Date)
write-host -foregroundColor yellow "Version:$ThisVersion"
"The current directory is {0}" -f (get-location)
"The Process Id is {0}" -f $pid
"MyInvocation follows"
$MyInvocation
"MyInvocation.MyCommand.path follows"
$MyInvocation.MyCommand.path
#Set-PSDebug -Step


$ToneGood = 500
$ToneBad = 100
$ToneDuration = 500

Try {
    if ( (get-process "qw"-ErrorAction SilentlyContinue) -ne $null ) { write-host -ForegroundColor Red "Quicken is running"; read-host "press RETURN to exit"; exit }

    if ($speak -and $psversiontable.psedition -ne "CORE") {
        [bool]$bSayit = $true
        #https://msdn.microsoft.com/en-us/library/system.speech.synthesis.speechsynthesizer(v=vs.110).aspx
        add-type -assemblyname system.speech
        $oSynth = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $oSynth.rate = 3 # range -10 to 10
        $SayIt = "Quicken"
        $SayIt = "$env:USERNAME     you are invoking Quicken"
        #$oSynth.SpeakAsync($SayIt)
    }
    else { [bool]$bSayit = $false }

    $SayIt = "Using {0}" -f $Filename
    write-host -ForegroundColor yellow $SayIt
    if ($bSayIt) { $oSynth.SpeakAsync($SayIt) }

    $SourceDir= split-path $MyInvocation.MyCommand.path -Parent
    Write-host -ForegroundColor yellow "The path to the Repository is $SourceDir"

    #Now test the SourceDir exists. If it doesn't then exit.
    if (!(Test-Path $SourceDir)) {read-host "The path to the Repository Workspace $SourceDir is incorrect"; exit}

    #Does the Quicken data file exist?
    $SourcePath = Join-Path $SourceDir $FileName
    if (!(Test-Path $SourcePath)) {read-host "The Repository file path $SourcePath does not exist."; exit}
    else {
        write-host -ForegroundColor Yellow "The file in the Repository is $SourcePath"
        get-item $SourcePath | format-list Fullname, CreationTime, LastWriteTime, LastAccessTime
    }

    $DestinationPath = Join-Path $DestinationDir $FileName #full path to Quicken data file.
    if (test-path $DestinationPath) {
        $TempFolderName = Split-path $DestinationDir -leaf
        $Sayit = "The data file is already in the $tempFolderName folder. Overwrite It? "
        if ($bSayit) {$oSynth.SpeakAsync($Sayit)}
        Write-Host $SayIt -ForegroundColor Yellow
        get-item $DestinationPath | format-list Fullname, CreationTime, LastWriteTime, LastAccessTime
        Do { $MyResponse = Read-host "$Filename exists in folder $DestinationDir - Overwrite? [y(es)/n(o)]"}
        while ("y", "n" -notcontains $MyResponse)
        #$MyResponse = read-host "$Filename exists in folder $DestinationDir - Overwrite? [y(es)/n(o)]"
        if ( $MyResponse.tolower() -eq "y") {
            write-host -ForegroundColor Yellow "Overwriting as requested"
            Copy-Item $SourcePath $DestinationDir
        }
        else {write-host -ForegroundColor Yellow "Using existing file"}
    }
    else {
        write-host -ForegroundColor Yellow "copying $SourcePath to $DestinationDir"
        #just trying to leave audit trail - experimental
        ($thisCmd="Copy-Item $SourcePath $DestinationDir")
        #Invoke-expression $thisCmd # Ithought this was working but it isn't
        Copy-Item $SourcePath $DestinationDir
        if ($?){"{0} completed" -f $thisCmd}
    }

    "Launch Quicken by referencing the data file in $DestinationPath"
    $ExitCode = 1
    do {
        #$LastExitCode = 0
        if ($bSayit) {$oSynth.SpeakAsync(("{0}     you are invoking Quicken with {1}" -f $env:USERNAME, $Filename))}

        cmd /C "$DestinationPath" #launch Quicken using file association and WAIT for it to exit.
        #Start-Process -wait "$DestinationPath" #launch Quicken using file association and WAIT for it to exit.

        $ExitCode = $LastExitCode
        write-host "LastExitCode $ExitCode"
        if ($ExitCode -ne 0) {
            $Sayit = "Oops, Quicken stopped with code $ExitCode "
            $oSynth.SpeakAsync($SayIt)
            write-host -ForegroundColor "Red" $SayIt
            $MyResponse = Read-Host "Do you want to restart Quicken? 'Y[es]/N[o]'"
            if ($MyResponse -eq "n") {
                $ExitCode = 0
            }

        }
    } until ($ExitCode -eq 0)

    #Bring this windows back to the foreground.
    function Show-Process($Process, [Switch]$Maximize) {
        $sig = '
        [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
        '

        if ($Maximize) { $Mode = 3 } else { $Mode = 4 }
        $type = Add-Type -MemberDefinition $sig -Name WindowAPI -PassThru
        $hwnd = $process.MainWindowHandle
        $null = $type::ShowWindowAsync($hwnd,$mode)
        $null = $type::SetForegroundWindow($hwnd)
    }
    #Set-PSDebug -Step
    #start-sleep -Seconds 1
    #if ($bSayit) {$oSynth.SpeakAsync("You should see me now.")}
    Show-Process -Process (get-process -id $pid) -Maximize

<#     sleep -Seconds 2
    [void][reflection.assembly]::loadwithpartialname("system.windows.forms")
    $altkeys = @(0xA4, 0x09)
    [system.windows.forms.sendkeys]::sendwait('%{TAB}')
 #>
    #At this point Quicken has exited. Now decide what to do with the data file we where working with.
    if ($bSayit) {$oSynth.SpeakAsync("Do you want to move $($Filename) to the repository?")}
    Do { $MyResponse = Read-host "Move $Filename to reposity [y(es)/n(o)]"}
    while ("y", "n" -notcontains $MyResponse)
    #$MyResponse = read-host "Move $Filename to reposity [y(es)/n(o)]"
    if ($MyResponse.tolower() -eq "y") {
        $Sayit = "Moving '$Filename' to the repository "
        if ($bSayIt) {$oSynth.SpeakAsync($SayIt)}
        move-Item $DestinationPath $SourceDir -force
        write-host  -foregroundColor Yellow "$($Sayit) at $(Get-Date) " # "V2.15.3"
        [console]::beep($ToneGood, 500)
    }
    else {
        if ($bSayit) {$oSynth.SpeakAsync("Do you want to move $($Filename) to the recycle-bin?")}
        Do { $MyResponse = Read-host "Move $($Filename) to the recycle-bin? [y(es)/n(o)]"}
        while ("y", "n" -notcontains $MyResponse)
        #$MyResponse = read-host "Move $($Filename) to the recycle-bin? [y(es)/n(o)]"
        if ( $MyResponse.tolower() -eq "y") {
            $SayIt = "MOVING $($Filename) to the recycle-bin  "
            write-host  -foregroundColor Yellow "$SayIt at $(Get-Date) " # "V2.15.3"
            if ($bSayIt) {$oSynth.Speak($SayIt)}

            if ($psversiontable.psedition -ne "CORE") {
                Add-Type -AssemblyName Microsoft.VisualBasic
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($DestinationPath, 'OnlyErrorDialogs', 'SendToRecycleBin')
            }
            else {
                remove-item $DestinationPath
            }
            [console]::beep($ToneGood, $ToneDuration)
        }
        else {
            $SayIt = "Leaving the working file"
            if ($bSayIt) {$oSynth.Speak($SayIt)}
            write-host  -foregroundColor Yellow $SayIt
            #Read-Host "$SayIt"
            [console]::beep($ToneBad, $ToneDuration)
        }
    }
    #Push *.dat files to subdirectory
    # $DatFolder=join-path $DestinationDir "Dat"
    # if (!(test-path $DatFolder)) {New-Item -ItemType "directory" -path $DatFolder}
    # $SayIt = "moving D A T files to subdirectory"
    # if ($bSayIt) {$oSynth.Speak($SayIt)}
    # write-host $SayIt
    # copy-Item -path (join-path $DestinationDir ($filename.basename + "*.dat")) -destination $DatFolder
    # remove-Item -path (join-path $DestinationDir ($filename.basename + "*.dat"))

    explorer.exe $DestinationDir #spawn file-manager
}  #end try
catch {
    $Sayit = "Something went wrong!"
    write-host -ForegroundColor Red $SayIt
    if ($bSayIt) {$oSynth.Speak($SayIt)}
    [console]::beep($ToneBad, $ToneDuration)
    [console]::beep($ToneBad, $ToneDuration)
    [console]::beep($ToneBad, $ToneDuration)
    Read-Host -prompt "This gives you a chance to see what went wrong."
}
Finally {
    $Sayit = "Fin E"
    write-host -ForegroundColor yellow $SayIt
    if ($bSayIt) {
        $oSynth.Speak($SayIt)
        #$oSynth = $null
        $oSynth.Dispose()
    }
    Stop-Transcript
    #Read-Host -prompt "This gives you a chance to see what went wrong."
}

