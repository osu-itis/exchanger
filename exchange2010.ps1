#Set installation source to same directory as script execution
$sourcePath = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host 'Using ' -NoNewline
Write-Host $sourcePath -ForegroundColor DarkGreen -NoNewline
Write-Host ' as the installation source.'

# Detect correct OS here and exit if no match
$wmiOS = Get-WMIObject win32_OperatingSystem
$OScap = $wmiOS.Caption
$OSver = $wmiOS.Version
[array]$wmiProc = Get-WmiObject win32_Processor
if ($wmiProc[0].Architecture -eq '9')
    {
    if ($OScap -match 'Windows 7')
        {$os = 'Win7'}
    elseif (($OSver.Contains('6.1')) -and ($OScap -match '2008'))
        {$os = 'R2'}
    elseif (($OSver.Contains('6.0')) -and ($OScap -match '2008'))
        {$os = 'R1'}
    else
        {
        Write-Host 'The script requires Windows Server 2008 with SP2, R2, or Windows 7 for the management tools, which this is not.' -ForegroundColor Red -BackgroundColor Black
        break
        }
    }
else
    {
    Write-Host 'Exchange 2010 requires x64 architecture, which this is not.' -ForegroundColor Red -BackgroundColor Black
    break
    }

#Region Installation files and properties
$fileWinRM = @{'filename'='Windows6.0-KB968930-x64.msu';
    'shortname'='WinRM';
    'displayname'='Windows Remote Management Framework';
    'url'='http://download.microsoft.com/download/2/8/6/28686477-3242-4E96-9009-30B16BED89AF/Windows6.0-KB968930-x64.msu';
    'size'='14MB'}
$fileNET35 = @{'filename'='dotnetfx35.exe';
    'shortname'='.NET 3.5';
    'displayname'='.NET 3.5 SP1';
    'url'='http://download.microsoft.com/download/2/0/E/20E90413-712F-438C-988E-FDAA79A8AC3D/dotnetfx35.exe';
    'size'='235MB'}
$fileNET35HF = @{'filename'='NDP35SP1-KB958484-x64.exe';
    'shortname'='.NET 3.5 hotfix';
    'displayname'='.NET 3.5 hotfix';
    'url'='http://download.microsoft.com/download/B/4/2/B42197BD-AEE1-4FE6-8CB3-29D60D0C3727/NDP35SP1-KB958484-x64.exe';
    'size'='1.4MB'}
#This is the SP1 version of the filter pack, but there are no changes in SP1 that impact Exchange,
#so the RTM version is also sufficient.
$fileOFP = @{'filename'='2010FilterPack64bit.exe';
    'shortname'='Office 2010 Filter Pack';
    'displayname'='Office 2010/2007 Filter Pack';
    'url'='http://download.microsoft.com/download/0/A/2/0A28BBFA-CBFA-4C03-A739-30CCA5E21659/FilterPack64bit.exe';
    'size'='4MB'}
$fileKB977624 = @{'filename'='Windows6.0-KB977624-v2-x64.msu';
    'shortname'='KB977624';
    'displayname'='KB 977624 v2';
    'url'='https://premier.microsoft.com/kb/977624';
    'size'='3MB'}
$fileKB979744 = @{'filename'='Windows6.0-KB979744-v2-x64.msu';
    'shortname'='KB979744';
    'displayname'='KB 979744 v2';
    'url'='http://connect.microsoft.com/VisualStudio/Downloads/DownloadDetails.aspx?DownloadID=27109';
    'size'='10MB'}
$fileKB979744R2 = @{'filename'='Windows6.1-KB979744-v2-x64.msu';
    'shortname'='KB979744';
    'displayname'='KB 979744 v2';
    'url'='http://connect.microsoft.com/VisualStudio/Downloads/DownloadDetails.aspx?DownloadID=27109';
    'size'='10MB'}  
$fileKB979917 = @{'filename'='Windows6.0-KB979917-x64.msu';
    'shortname'='KB979917';
    'displayname'='KB 979917';
    'url'='http://archive.msdn.microsoft.com/Project/Download/FileDownload.aspx?ProjectName=KB979917&DownloadId=12756';
    'size'='3MB'}
$fileKB973136 = @{'filename'='Windows6.0-KB973136-x64.msu';
    'shortname'='KB973136';
    'displayname'='KB 973136';
    'url'='https://connect.microsoft.com/VisualStudio/Downloads/DownloadDetails.aspx?DownloadID=20922';
    'size'='700KB'}
$fileKB977592 = @{'filename'='Windows6.0-KB977592-x64.msu';
    'shortname'='KB977592';
    'displayname'='KB 977592';
    'url'='https://premier.microsoft.com/kb/977592';
    'size'='300KB'}
$fileKB979099 = @{'filename'='Windows6.1-KB979099-x64.msu';
    'shortname'='KB979099';
    'displayname'='KB 979099';
    'url'='http://download.microsoft.com/download/1/F/B/1FB7F377-CB25-4D51-B4A7-D3F05B7A55CA/Windows6.1-KB979099-x64.msu';
    'size'='2MB'}
$fileKB983440 = @{'filename'='Windows6.1-KB983440-x64.msu';
    'shortname'='KB983440';
    'displayname'='KB 983440';
    'url'='http://archive.msdn.microsoft.com/KB983440/Release/ProjectReleases.aspx?ReleaseId=4410';
    'size'='3MB'}
$fileKB977020 = @{'filename'='Windows6.1-KB977020-v2-x64.msu';
    'shortname'='KB977020';
    'displayname'='KB 977020 v2';
    'url'='http://connect.microsoft.com/VisualStudio/Downloads/DownloadDetails.aspx?DownloadID=27977';
    'size'='600KB'}
$fileUCMARuntime = @{'filename'='UCMARuntimeSetup.exe';
    'shortname'='UCMARuntime';
    'displayname'='Unified Communications Managed API 2.0';
    'url'='http://download.microsoft.com/download/5/8/4/58494AD4-4091-457F-B23D-E57D211D2B5D/UcmaRuntimeSetup.exe';
    'size'='16MB'}
$fileSSRT = @{'filename'='SpeechPlatformRuntime.msi';
    'shortname'='SpeechPlatform';
    'displayname'='Speech Platform - Server Runtime';
    'url'='http://download.microsoft.com/download/0/4/0/040235F1-3798-4B10-BB36-FAF870A8D559/Runtime/x64/SpeechPlatformRuntime.msi';
    'size'='3MB'}
#EndRegion  

Function InstallApp($app)
    {
    if (!($installedUpdates))
        {
        Write-Host 'Loading list of installed hotfixes...'
        $script:installedUpdates = Get-WmiObject win32_quickfixengineering
        }
    switch ($app)
        {
        'WinRM' 
            {
            $appArray = $fileWinRM
            $kb = 'KB968930'
            }
        'NET35'
            {
            $appArray = $fileNET35
            $checkExpression = "test-path 'HKLM:Software\Microsoft\NET Framework Setup\NDP\v3.5'"
            }
        'NET35HF'
            {
            $appArray = $fileNET35HF
            $checkExpression = "test-path 'HKLM:SOFTWARE\Wow6432Node\Microsoft\Updates\Microsoft .NET Framework 3.5 SP1\SP1\KB958484'"
            }
        'OFP'
            {
            $appArray = $fileOFP
            $checkExpression = "test-path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{95140000-2000-0409-1000-0000000FF1CE}'"
            }
        'UCMA'
            {
            $appArray = $fileUCMARuntime
            $checkExpression = "test-path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{7EB901DD-CB50-4046-A434-3E9A112E8F86}'"
            }
        'SSRT'
            {
            $appArray = $fileSSRT
            $checkExpression = "test-path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{3B433087-E62E-4BF5-97F9-4AF6E1C2409C}'"
            }
        'KB977624'
            {
            $appArray = $fileKB977624
            $kb = 'KB977624'
            }
        'KB979744'
            {
            $appArray = $fileKB979744
            $kb = 'KB979744'
            }
        'KB979744R2'
            {
            $appArray = $fileKB979744R2
            $kb = 'KB979744'
            }
        'KB973136'
            {
            $appArray = $fileKB973136
            $kb = 'KB973136'
            }
        'KB977592'
            {
            $appArray = $fileKB977592
            $kb = 'KB977592'
            }
        'KB979099'
            {
            $appArray = $fileKB979099
            $kb = 'KB979099'
            }
        'KB983440'
            {
            $appArray = $fileKB983440
            $kb = 'KB983440'
            }
        'KB977020'
            {
            $appArray = $fileKB977020
            $kb = 'KB977020'
            }

        }
    trap
        {
        Write-Host ''
        Write-Host "There was a problem downloading or installing $($appArray.shortname)." -ForegroundColor Red
        Write-Host ''
        break
        }
    #Check for existing installation
    Write-Host "Verifying $($appArray.displayname) is installed..." -NoNewline
    if (($app -eq 'WinRM') -or ($app -like 'KB*'))
        {
        if ($installedUpdates -match $kb)
            {
            $bInstalled = $true
            }
        else
            {
            $bInstalled = $false
            }
        }
    else
        {
        if (Invoke-Expression $checkExpression)
            {
            $bInstalled = $true
            }
        else
            {
            $bInstalled = $false
            }
        }
    if ($bInstalled)
        {
        Write-Host "$($appArray.displayname) is installed." -ForegroundColor Green
        return
        }
    Write-Host "$($appArray.displayname) is not installed." -ForegroundColor Yellow
    Write-Host "Installing $($appArray.displayname)..." -NoNewline
    
    #Install app:  Check for existing installation file.
    $fullPath = $sourcePath+"\$($appArray.filename)"
    if (!(Test-Path $fullPath))
        {
        Write-Host ''
        Write-Host "$($appArray.filename) not found in source path." -ForegroundColor Yellow
        if (($app -eq 'KB977624') -or ($app -eq 'KB977592'))
            {
            Write-Host "$($apparray.displayname) can only be downloaded from the Premier website," -ForegroundColor Red
            Write-Host "$($apparray.url) or by requesting it from Microsoft." -ForegroundColor Red
            break
            }
        if (($app -eq 'KB979744') -or ($app -eq 'KB979744R2') -or ($app -eq 'KB979917') -or ($app -eq 'KB973136') -or ($app -eq 'KB983440') -or ($app -eq 'KB977020'))
            {
            Write-Host "$($apparray.displayname) must be manually downloaded: " -ForegroundColor Red
            Write-Host "$($apparray.url)"
            break
            }

        $dl = Read-Host "Do you want to download it now? ($($appArray.size))(Y/N)"
        if ($dl -ne 'y')
            {
            Write-Host "You have chosen to not download the $($appArray.shortname) installation file."
            Write-Host "Put $($appArray.filename) in the source directory and run the script again."
            break
            }
        else
            {
            Write-Host "Downloading $($appArray.shortname)..." -NoNewline
            $dlClient = New-Object System.Net.WebClient
            $dlClient.DownloadFile($appArray.url,$fullPath)
            if (!(Test-Path $fullPath))
                {
                Write-Host ''
                Write-Host "There was a problem downloading $($appArray.shortname)." -ForegroundColor Red
                Write-Host ''
                }
            else
                {
                Write-Host 'done.' -ForegroundColor Green
                }
            }
        }
    
    #Install app: Run installation.
    if ($app -eq 'WinRM')
        {
        $expression = "wusa $fullPath /quiet"
        Invoke-Expression $expression
        Write-Host 'External update process started...Be patient, it takes time.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host 'When the WinRM installation is complete, the system will automatically reboot.'
        Write-Host 'Then you can rerun the script to continue.  This script will now end.'
        break
        }
    else
        {
        if ($app -eq 'NET35HF')
            {$arguments = '/passive /norestart'}
        else
            {$arguments = '/quiet /norestart'}
        $process = [System.Diagnostics.Process]::Start($fullPath,$arguments)
        $process.WaitForExit()
        Write-Host "$($appArray.displayname) installation complete." -ForegroundColor Green
        }
    }   
 
Function InstallNET35()
    {
    InstallApp 'NET35'
    If (-not($os -eq 'Win7'))
        {
        InstallApp 'NET35HF'
        }
    }
 
function InstallHotfixes()
    {

    if ($os -eq 'R1')
        {
        InstallApp 'KB977624'
        InstallApp 'KB979744'
        InstallApp 'KB979917'
        InstallApp 'KB973136'
        InstallApp 'KB977592'
        }
    elseif (($os -eq 'R2') -and (($opt -eq 2) -or ($opt -eq 6) -or ($opt -eq 7)))
        {
        InstallApp 'KB979099'
        InstallApp 'KB979744R2'
        InstallApp 'KB983440'
        InstallApp 'KB977020'
        }
    elseif (($os -eq 'R2') -and (($opt -eq 1) -or ($opt -eq 3)))
        {
        InstallApp 'KB979099'
        }
    elseif ($os -eq 'Win7')
        {
        #InstallApp 'KB977020'
        #InstallApp 'KB983440'
        }
    }

Function SetTCPSharing()
    {
    trap
        {
        Write-Host ''
        Write-Host 'There was problem setting the NET TCP Port Sharing service to Automatic startup.' -ForegroundColor Red
        Write-Host 'The service must be set to Automatic for Exchange setup to be successful.' -ForegroundColor Red
        Write-Host ''
        return
        }   
    #Set NETTCPPortSharing to Automatic
    Write-Host 'Configuring the NET TCP Port Sharing service...' -NoNewline
    Set-Service NetTcpPortSharing -StartupType Automatic
    Write-Host 'done.' -ForegroundColor Green
    }
 
Function EnableRemoting()
    {
    trap
        {
        Write-Host ''
        Write-Host 'There was problem configuring the system for remote PowerShell.' -ForegroundColor Red
        Write-Host ''
        return
        }
    #Enable Remote PowerShell for Exchange administration from workstations
    Write-Host 'Enabling system for remote PowerShell connections...'
    Enable-PSRemoting -force
    Write-Host 'Remote PowerShell configuration is done.' -ForegroundColor Green
    }
    
Function EnableFirewall()
    {
    trap
        {
        Write-Host ''
        Write-Host 'There was problem starting the Windows Firewall service.' -ForegroundColor Red
        Write-Host 'The firewall service must be running during Exchange setup.  It can be stopped after it completes.' -ForegroundColor Red
        Write-Host ''
        return
        }   
    #Ensure Windows Firewall is running or Exchange install will fail
    Write-Host 'Starting the Windows Firewall service...' -NoNewline
    Set-Service 'MpsSvc' -StartupType Automatic -Status Running
    Write-Host 'done.' -ForegroundColor Green
    }

if ($os -eq 'Win7')
    {
    $tools = '. dism /Online /Enable-Feature /FeatureName:IIS-WebServerRole /FeatureName:IIS-WebServerManagementTools /FeatureName:IIS-IIS6ManagementCompatibility /FeatureName:IIS-Metabase /FeatureName:IIS-LegacySnapIn; . dism /Online /Disable-Feature /FeatureName:IIS-WebServer /NoRestart'
    }
elseif ($os -eq 'R1')
    {
    $ht = '. ServerManagerCmd.exe -i NET-Framework RSAT-ADDS Web-Server Web-Basic-Auth Web-Windows-Auth Web-Metabase Web-Net-Ext Web-Lgcy-Mgmt-Console WAS-Process-Model RSAT-Web-Server'
    $cas = '. ServerManagerCmd.exe -i NET-Framework RSAT-ADDS Web-Server Web-Basic-Auth Web-Windows-Auth Web-Metabase Web-Net-Ext Web-Lgcy-Mgmt-Console WAS-Process-Model RSAT-Clustering RSAT-Web-Server Web-ISAPI-Ext Web-Digest-Auth Web-Dyn-Compression NET-HTTP-Activation RPC-over-HTTP-proxy Web-WMI'
    $mbx = '. ServerManagerCmd.exe -i NET-Framework RSAT-ADDS Web-Server Web-Basic-Auth Web-Windows-Auth Web-Metabase Web-Net-Ext Web-Lgcy-Mgmt-Console WAS-Process-Model RSAT-Clustering RSAT-Web-Server'
    $um = '. ServerManagerCmd.exe -i NET-Framework RSAT-ADDS Web-Server Web-Basic-Auth Web-Windows-Auth Web-Metabase Web-Net-Ext Web-Lgcy-Mgmt-Console WAS-Process-Model RSAT-Web-Server Desktop-Experience'
    $edge = '. ServerManagerCmd.exe -i NET-Framework RSAT-ADDS ADLDS'
    $typical = '. ServerManagerCmd.exe -i NET-Framework RSAT-ADDS Web-Server Web-Basic-Auth Web-Windows-Auth Web-Metabase Web-Net-Ext Web-Lgcy-Mgmt-Console WAS-Process-Model RSAT-Clustering RSAT-Web-Server Web-ISAPI-Ext Web-Digest-Auth Web-Dyn-Compression NET-HTTP-Activation RPC-over-HTTP-proxy Web-WMI'
    $tools = '. ServerManagerCmd.exe -i Web-Lgcy-Mgmt-Console'
    }
elseif ($os -eq 'R2')
    {
    $ht = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server -restart'
    $cas = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Web-ISAPI-Ext,Web-Digest-Auth,Web-Dyn-Compression,NET-HTTP-Activation,RPC-Over-HTTP-Proxy,Web-WMI -restart'
    $mbx = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server -restart'
    $um = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Desktop-Experience -restart'
    $edge = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,ADLDS -restart'
    $typical = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Web-ISAPI-Ext,Web-Digest-Auth,Web-Dyn-Compression,NET-HTTP-Activation,RPC-Over-HTTP-Proxy,Web-WMI -restart'
    $tools = 'Add-WindowsFeature NET-Framework,Web-Lgcy-Mgmt-Console -restart'
    Import-Module ServerManager
    }
$opt = 'None'

if ($os -eq 'R1')
    {
    InstallApp 'WinRM'
    }

clear
if ($opt -ne 'None') {write-host 'Last command: '$opt -foregroundcolor Yellow}
write-host
write-host 'Exchange Server 2010 Prerequisites Installation'
write-host 'Please select which role you are going to install:'
write-host
if ($os -eq 'Win7')
    {
    Write-Host '1-7) Not applicable on this system' -ForegroundColor Gray
    Write-Host '8)  Management Tools'
    Write-Host '9-12) Not applicable on this system' -ForegroundColor Gray
    }
else
    {
    write-host '1)  Hub Transport'
    write-host '2)  Client Access Server'
    write-host '3)  Mailbox'
    write-host '4)  Unified Messaging'
    write-host '5)  Edge Transport'
    write-host '6)  Typical (CAS\HT\Mailbox)'
    write-host '7)  Client Access and Hub Transport'
    Write-Host '8)  Management Tools'
    write-host
    write-host '9)  Configure NetTCP Port Sharing service'
    write-host '    Required for the Client Access Server role' -foregroundcolor yellow
    write-host '    Automatically set for options 2,6, and 7' -foregroundcolor yellow
    write-host '10) Install 2010 Office System Converter: Microsoft Filter Pack'
    write-host '    Required if installing Hub Transport or Mailbox Server roles' -foregroundcolor yellow
    write-host '    Automatically set for options 1,3,6, and 7' -foregroundcolor yellow
    Write-Host '11) Enable PowerShell Remoting'
    Write-Host '    Automatically set for options 1,2,3,4,6, and 7' -ForegroundColor Yellow
    write-host
    }
write-host '13) Restart the System'
write-host '14) End'
write-host
Write-Host 'Note: Using ' -NoNewline
Write-Host $sourcePath -ForegroundColor DarkGreen -NoNewline
Write-Host ' as the installation source.'
$opt = Read-Host 'Select an option.. ? '
 
switch ($opt)
    {
    1{
        InstallNET35; InstallApp 'OFP'; EnableFirewall; EnableRemoting; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $ht
        }
    2{
        InstallNET35; EnableFirewall; EnableRemoting; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $cas
        SetTCPSharing
        }
    3{
        InstallNET35; InstallApp 'OFP'; EnableFirewall; EnableRemoting; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $mbx
        }
    4{
        InstallNET35; InstallApp 'UCMA'; InstallApp 'SSRT'; EnableRemoting; EnableFirewall; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $um
        }
    5{
        InstallNET35; EnableFirewall; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $edge
        }
    6{
        InstallNET35; InstallApp 'OFP'; EnableFirewall; EnableRemoting; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $typical
        SetTCPSharing
        }
    7{
        InstallNET35; InstallApp 'OFP'; EnableFirewall; EnableRemoting; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $cas
        SetTCPSharing
        }
    8{
        InstallNET35; InstallHotfixes
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $tools
        }
    9 { SetTCPSharing }
    10 { InstallApp 'OFP' }
    11 { EnableRemoting }
    13 { Restart-Computer }
    14 {write-host 'Exiting...'}
    default {write-host "You haven't selected any of the available options."}
    }

