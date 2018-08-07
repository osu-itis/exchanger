# Install Exchange 2010 with SP3, CU20, CAS/Tools/MRS
# loops are for suckers

param (
    [string]$filepath = 'C:\exchange\',
    [switch]$prepare = $false
)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010


Write-Host 'This process will install Exchange 2010 with SP3 and CU20'
Write-Host 'The CAS role as well as management tools will also be installed and MRS enabled'

if ($prepare -eq $true)
{
    # Prepare
    Write-Host 'Prepare AD: ' -NoNewline
    #$exp = Start-Process "$filepath\setup.com" -ArgumentList @('/p', '/on:"Contoso"') -wait -NoNewWindow -PassThru
    if($exp.ExitCode -ne 0) {
        Write-Host 'FAIL' -ForegroundColor Red -BackgroundColor Black
        break
    }
    else { Write-Host 'OK' }
}

# Install
Write-Host 'Install Exchange 2010 CAS, Tools: ' -NoNewline
$instex = Start-Process "$filepath\setup.com" -ArgumentList @('/m:Install', '/r:c,t') -wait -NoNewWindow -PassThru
if($instex.ExitCode -ne 0) {
    Write-Host 'FAIL' -ForegroundColor Red -BackgroundColor Black
    break
}
else { Write-Host 'OK' }

if ($prepare -eq $true)
{
    # Prepare SP3
    Write-Host 'Prepare SP3: ' -NoNewline
    #$exp3 = Start-Process "$filepath\sp3\setup.com" -ArgumentList @('/p', '/on:"Contoso"' -wait -NoNewWindow -PassThru
    if($exp3.ExitCode -ne 0) {
        Write-Host 'FAIL' -ForegroundColor Red -BackgroundColor Black
        break
    }
    else { Write-Host 'OK' }
}

# Install SP3
Write-Host 'Install SP3: ' -NoNewline
$sp3 = Start-Process "$filepath\sp3\setup.com" -ArgumentList @('/m:Upgrade', '/InstallWindowsComponents') -wait -NoNewWindow -PassThru
if($sp3.ExitCode -ne 0) {
    Write-Host 'FAIL' -ForegroundColor Red -BackgroundColor Black
    break
}
else { Write-Host 'OK' }

# Install CU20
Write-Host 'Install CU20: ' -NoNewline
$cup = Start-Process C:\Windows\System32\msiexec.exe -ArgumentList @("/p $filepath\Exchange2010-KB4073537-x64-en.msp", "/passive") -wait -NoNewWindow -PassThru
if(($cup.ExitCode -ne 0) -Or ($cup.ExitCode -ne 1641) -Or ($cup.ExitCode -ne 3010)) {
    Write-Host 'FAIL' -ForegroundColor Red -BackgroundColor Black
    break
}
else { Write-Host 'OK' }

# Enable MRS
Write-Host 'Enabling MRSProxy: ' -NoNewline
Set-WebServicesVirtualDirectory -Identity "EWS (Default Web Site)" -MRSProxyEnabled $true

$mrs = Get-WebServicesVirtualDirectory -Identity "EWS (Default Web Site)" -MRSMaxConnections 20
if($mrs.MRSProxyEnabled -eq $true) { Write-Host 'OK' }
else {
    Write-Host 'FAIL'
    break
}

Write-Host 'You must now reboot.'

