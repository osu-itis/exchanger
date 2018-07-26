import-module updateservices

# given a list of KB numbers approve the appropriate update in WSUS

$wsusserver = ""
$arch = "x64"
$targetgroup = ""

$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver,$False,8530)

$kbs = Get-Content "kbs.txt"

foreach($kb in $kbs)
{
    $updates = $wsus.searchupdates($kb)
    foreach($update in $updates)
    {
        if (($update.Title -Match $arch) -or ($update.title -match "Exchange Server 2010"))
        {
            get-wsusupdate -updateid $update.id.updateid | Approve-WsusUpdate -Action Install -TargetGroupName $targetgroup
            write-host $update.id.updateid
        }
    }
}

