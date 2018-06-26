#Set installation source to same directory as script execution
$sourcePath = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host 'Using ' -NoNewline
Write-Host $sourcePath -ForegroundColor DarkGreen -NoNewline
Write-Host ' as the installation source.'

# Detect correct OS here and exit if no match
$wmiOS = Get-WMIObject win32_OperatingSystem
$OScap = $wmiOS.Caption
$OSver = $wmiOS.Version
$OSsp = $wmiOS.ServicePackMajorVersion
[array]$wmiProc = Get-WmiObject win32_Processor
if ($wmiProc[0].Architecture -eq '9')
{
    if (($OSver.Contains('6.1')) -and ($OScap -match '2008') -and ($OSsp -eq '1'))
    {
        Write-Host '2008R2 SP1 detected'
    }
    else
    {
        Write-Host 'The script requires Windows Server 2008R2 SP1' -ForegroundColor Red -BackgroundColor Black
        break
    }
}
else
{
    Write-Host 'Exchange 2010 requires x64 architecture' -ForegroundColor Red -BackgroundColor Black
    break
}

#Region Installation files and properties
#This is the SP1 version of the filter pack, but there are no changes in SP1 that impact Exchange,
#so the RTM version is also sufficient.
$fileOFP = @{'filename'='2010FilterPack64bit.exe';
    'shortname'='Office 2010 Filter Pack';
    'displayname'='Office 2010/2007 Filter Pack';
    'url'='http://download.microsoft.com/download/0/A/2/0A28BBFA-CBFA-4C03-A739-30CCA5E21659/FilterPack64bit.exe';
    'size'='4MB'}
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
    switch ($app)
    {
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
    if (Invoke-Expression $checkExpression)
        { $bInstalled = $true }
    else
        { $bInstalled = $false }
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
    $arguments = '/quiet /norestart'
    $process = [System.Diagnostics.Process]::Start($fullPath,$arguments)
    $process.WaitForExit()
    Write-Host "$($appArray.displayname) installation complete." -ForegroundColor Green
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

$ht = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server -restart'
$cas = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Web-ISAPI-Ext,Web-Digest-Auth,Web-Dyn-Compression,NET-HTTP-Activation,RPC-Over-HTTP-Proxy,Web-WMI -restart'
$mbx = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server -restart'
$um = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Desktop-Experience -restart'
$edge = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,ADLDS -restart'
$typical = 'Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Web-ISAPI-Ext,Web-Digest-Auth,Web-Dyn-Compression,NET-HTTP-Activation,RPC-Over-HTTP-Proxy,Web-WMI -restart'
$tools = 'Add-WindowsFeature NET-Framework,Web-Lgcy-Mgmt-Console -restart'

Import-Module ServerManager
$opt = 'None'

clear
if ($opt -ne 'None') {Write-Host 'Last command: '$opt -foregroundcolor Yellow}
Write-Host
Write-Host 'Exchange Server 2010 Prerequisites Installation'
Write-Host 'Please select which role you are going to install:'
Write-Host
Write-Host '1)  Hub Transport'
Write-Host '2)  Client Access Server'
Write-Host '3)  Mailbox'
Write-Host '4)  Unified Messaging'
Write-Host '5)  Edge Transport'
Write-Host '6)  Typical (CAS\HT\Mailbox)'
Write-Host '7)  Client Access and Hub Transport'
Write-Host '8)  Management Tools'
Write-Host
Write-Host '9)  Configure NetTCP Port Sharing service'
Write-Host '    Required for the Client Access Server role' -foregroundcolor yellow
Write-Host '    Automatically set for options 2,6, and 7' -foregroundcolor yellow
Write-Host '10) Install 2010 Office System Converter: Microsoft Filter Pack'
Write-Host '    Required if installing Hub Transport or Mailbox Server roles' -foregroundcolor yellow
Write-Host '    Automatically set for options 1,3,6, and 7' -foregroundcolor yellow
Write-Host '11) Enable PowerShell Remoting'
Write-Host '    Automatically set for options 1,2,3,4,6, and 7' -ForegroundColor Yellow
Write-Host
Write-Host '13) Restart the System'
Write-Host '14) End'
Write-Host
Write-Host 'Note: Using ' -NoNewline
Write-Host $sourcePath -ForegroundColor DarkGreen -NoNewline
Write-Host ' as the installation source.'
$opt = Read-Host 'Select an option.. ? '

switch ($opt)
{
    1{
        InstallApp 'OFP'; EnableFirewall; EnableRemoting
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $ht
        }
    2{
        EnableFirewall; EnableRemoting
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $cas
        SetTCPSharing
        }
    3{
        InstallApp 'OFP'; EnableFirewall; EnableRemoting
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $mbx
        }
    4{
        InstallApp 'UCMA'; InstallApp 'SSRT'; EnableRemoting; EnableFirewall
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $um
        }
    5{
        EnableFirewall
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $edge
        }
    6{
        InstallApp 'OFP'; EnableFirewall; EnableRemoting
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $typical
        SetTCPSharing
        }
    7{
        InstallApp 'OFP'; EnableFirewall; EnableRemoting
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $cas
        SetTCPSharing
        }
    8{
        Write-Host 'Beginning Windows components installation...'
        Invoke-Expression $tools
        }
    9 { SetTCPSharing }
    10 { InstallApp 'OFP' }
    11 { EnableRemoting }
    13 { Restart-Computer }
    14 { Write-Host 'Exiting...' }
    default { Write-Host "You haven't selected any of the available options." }
}

