# TODO:
# [ ] correct the name of installer: teamsbootstrapper
# [ ] clear browser cache: clear all profiles cache: User Data\Local State.profile.profile.profiles_order
# [ ] clear browser cache: clear new edge cache
# [ ] deep clear: clear reg/appdata/outlook Meeting/Presence addon/other residual data for new teams
# [v] Operation: enable dev mode.
# [ ] Detect current installation
# [ ] Show version of downloaded Teams msix
# [ ] change a way for install. installer -p --> installer -p -o "MSIX_PATH"
# [v] Show tips if install/uninstall failed.
# [v] Use Stop-RelatedProccess in Clear-BrowserCache
# [ ] Auto increase version

$DEBUG = $false
$TEAMS_INSTALLER = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
$TEAMS_MSIX_X64 = "https://go.microsoft.com/fwlink/?linkid=2196106"
$TEAMS_MSIX_X86 = "https://go.microsoft.com/fwlink/?linkid=2196060&clcid=0x409"

$CURRENT_PRINT_LEVEL = 0
$NON_ADMIN = $false
$CACHE = $null

function Write-AppHeader{
    Write-Host ""
    Write-Host -ForegroundColor Yellow "+- TeamsUtils for New Teams T2.1 ----------------+"
    Write-Host -ForegroundColor Yellow "|  author   : Nicky Yang                         |"
    Write-Host -ForegroundColor Yellow "|  version  : v0.1.9                             |"
    Write-Host -ForegroundColor Yellow "|  release  : 2023/11/20                         |"
    Write-Host -ForegroundColor Yellow "+------------------------------------------------+"
    Write-Host ""
}

function Write-AdminBanner{
    if($Script:NON_ADMIN -eq $true){
        Write-Host -BackgroundColor Red -ForegroundColor White   "                                                 "
        Write-Host -BackgroundColor Red -ForegroundColor White   "               ** NON-ADMIN MODE **              "
        Write-Host -BackgroundColor Red -ForegroundColor White   "                                                 "
        Write-Host ""
    }else{
        Write-Host -BackgroundColor Green -ForegroundColor White "                                                 "
        Write-Host -BackgroundColor Green -ForegroundColor White "                 ** ADMIN MODE **                "
        Write-Host -BackgroundColor Green -ForegroundColor White "                                                 "
        Write-Host ""
    }

}

function Write-DeepCleanTip{
    Write-Host ""
    Write-Host -BackgroundColor Red -ForegroundColor White " [Error] Operation failed. Please try rebooting your computer or use the Deep clean. "
    Write-Host ""
}

function Test-Administrator
{  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    $isAdmin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if(-not $isAdmin){
        Write-Host -BackgroundColor Red -ForegroundColor White " NOTE: You are not run as administrator, some function will be limited."
        # Write-Host -ForegroundColor Red "NOTE: You are not run as administrator, some function will be limited."
        $yn = Get-Confirm -prompt "# Do you want to run as administrator?"

        if($yn){
            $sh = new-object -com 'Shell.Application'
            $sh.ShellExecute('powershell', "-NoExit $($MyInvocation.ScriptName)", '', 'runas')
            exit
        }else{
            $script:NON_ADMIN = $true
        }
    }
}

function Get-Cache{
    try{
        $cache = "$PSScriptRoot\TeamsUtilsCache"
        # write-host $cache
        New-Item $cache -ItemType Directory *> $null
    }catch{
        # do nothing
        Write-Host "Cannot create $cache"
    }
    $script:CACHE = $cache
    return $cache
}

function Remove-Cache {
    try{
        Remove-Item $CACHE -Force -Recurse *> $null
    }catch{
        # do nothing
    }
}

function Get-Installer{
    $installer = "$CACHE/installer.exe"
    if(-not (Test-Path $installer)){
        Write-Operation -Message "Download the installer" -ForegroundColor Cyan

        try{                
            Invoke-WebRequest $script:TEAMS_INSTALLER -OutFile $installer
            Write-Host -ForegroundColor Green "[OK]"
        }catch{
            Write-Host -ForegroundColor Red "[FAILED]"
            Write-Error $_.Exception.Message
        }
    }
    return $installer
}

$all_options = @("1", "0")
$admin_options = @("4","6","7")

function Get-CleanSelection() {
    $selection = $null
    While ($true) {
        Write-Host -ForegroundColor Cyan "1. Clear New Teams cache."
        Write-Host -ForegroundColor Cyan "0. Cancel"
        Write-Host ""
        $selection = Read-Host "# Which process do you want to perform (1/0)?"
        if ($all_options.Contains($selection)) {
            if ($admin_options.Contains($selection) -and $Script:NON_ADMIN){
                Write-Host -ForegroundColor Red "`n# [Error] Need admin permission.`n"
            }else{
                return $selection
            }
        }else{
            Write-Host -ForegroundColor Red "`n# [Error] Wrong selection.`n"
        }
    }
}

function Get-Confirm{
    param (
        [string[]]$prompt
    )

    Write-Host -ForegroundColor Yellow $prompt -NoNewline
    $yn = Read-Host " (Y/N)"

    if($yn.ToUpper() -eq "Y"){
        return $true
    }else{
        return $false
    }

}

function Write-DoItMessage{
    Write-Host ""
    Write-Host "# OK, run the operation now!"
    Write-Host "# --------------------"
    Write-Host ""
}

function Write-Operation {
    param (
        [string]$Message,
        [int]$Level = $script:CURRENT_PRINT_LEVEL,
        [string]$ForegroundColor,
        [Switch]$NoStatus
    )
    $CHAR_PER_LINE = 70

    $prefix = (" "*($Level-1)*2+"* ")
    Write-Host $prefix -NoNewline -ForegroundColor $ForegroundColor
    Write-Host $Message -NoNewline -ForegroundColor $ForegroundColor
    $leftchar = $CHAR_PER_LINE - $prefix.Length - 2 - $Message.Length -2
    if($NoStatus){
        Write-Host ""
    }else{
        Write-Host (" " + "."*$leftchar + " ") -NoNewline
    }
}

# Operations

function Stop-Proc{
    param(
        [string]$ProcName
    )

    Write-Operation -Message "Stop Process: $ProcName" -ForegroundColor Cyan

    try {
        (Get-Process -ProcessName $ProcName | Stop-Process -Force) *> $null
        Start-Sleep -Seconds 3
        Write-Host -ForegroundColor Green "[OK]"
    }
    catch {
        Write-Host -ForegroundColor Cyan "[FAILED]"
        Write-Error "# [Warning] Operation failed."
        Write-Error  $_.Exception.Message
    }
}

function Stop-RelatedProccess {
    param(
        [switch]$brwoser = $false,
        [switch]$teamsdep = $false,
        [switch]$teams = $false
    )
    if($teams){
        $procs = "ms-teams"
        $procs.Foreach( {Stop-Proc -ProcName $_} )
    }

    if($teamsdep){
        $procs = "ms-teams", "outlook"
        $procs.Foreach( {Stop-Proc -ProcName $_} )
    }

    if($brwoser){
        $procs = "chrome", "msedge", "iexplorer", "MicrosoftEdge"
        $procs.Foreach( {Stop-Proc -ProcName $_} )
    }
}

function Clear-TeamsCache() {
    $script:CURRENT_PRINT_LEVEL = 1

    Write-Operation -Message "Clear Teams Cache" -ForegroundColor Gray -NoStatus
    
    $script:CURRENT_PRINT_LEVEL = 2

    Stop-RelatedProccess -teamsdep

    Write-Operation -Message "Clear Teams cache data" -ForegroundColor Cyan

    try {
        (Get-ChildItem -Path $env:LOCALAPPDATA\"Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams" | Remove-Item -Confirm:$false -Recurse) *> $null
        
        Write-Host -ForegroundColor Green "[OK]"
    }
    catch {
        Write-Host -ForegroundColor Red "[FAILED]"
        Write-Error "# [Warning] Operation failed."
        Write-Error  $_.Exception.Message
    }
}

function Main{
    if(-not $DEBUG){
        Remove-Cache
    }
    Get-Cache

    Write-AppHeader

    # Unmark if needs run in administrator
    Test-Administrator

    if($Script:NON_ADMIN -eq $true){
        Clear-Host
        Write-AppHeader
    }
    Write-AdminBanner

    $selection = Get-CleanSelection

    Write-Host ""
    try {

        if ($selection -eq "1" -and (Get-Confirm("# NOTE: It will CLOSE running Teams/Outlook before the operation, press Y to continue."))) {
            # Clear Teams cache
            Write-DoItMessage
            Clear-TeamsCache
        }
        else{        
            Write-Host "# Canceled."
        }

        Write-Host ""
        Write-Host "# --------------------"
        Write-Host "# Finished, Bye!"
        Write-Host ""
    }
    catch {
        Write-Host ""
        Write-Host "# --------------------"
        Write-Host "# Unknown error!"
        Write-Host ""
    }
    finally {
        if(-not $DEBUG){
            Remove-Cache
        }
        Read-Host  "# Press Enter to exit"
    }
}

Main