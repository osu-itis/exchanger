import-module updateservices

# get all update guids for a given computer from WSUS

$wsusserver = ""
$computername = ""

$target_computer = Get-WsusComputer -nameincludes $computername

$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver,$False)

$updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Installed

$tar = $wsus.GetComputerTarget($target_computer.id)
$tar.GetUpdateInstallationInfoPerUpdate($updatescope) | format-table -Property UpdateId
