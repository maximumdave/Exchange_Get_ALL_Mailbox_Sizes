add-pssnapin *exch*

$allusers = get-aduser -Filter * -Properties proxyaddresses | where-object {$_.proxyaddresses -ne $null -and $_.DistinguishedName -notlike "*OU=Student-Accounts*"}

($allusers).Count

$i=0
$array = $null

$array = foreach ($user in $allusers)
{
    $allstat = $null
    
    if (get-mailbox -ResultSize unlimited -Identity $user.ObjectGUID.tostring())
    {
        $type = "Regular"
        $allstat = get-mailbox -ResultSize unlimited -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    elseif (get-mailbox -ResultSize unlimited -Archive -identity $user.ObjectGUID.tostring())
    {
        $type = "Archive"
        $allstat = get-mailbox -ResultSize unlimited -Archive -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    elseif (get-mailbox -ResultSize unlimited -Arbitration -identity $user.ObjectGUID.tostring())
    {
        $type = "Arbitration"
        $allstat = get-mailbox -ResultSize unlimited -Arbitration -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics        
    }
    elseif (get-mailbox -ResultSize unlimited -AuditLog -Identity $user.ObjectGUID.tostring())
    {
        $type = "AuditLog"
        $allstat = get-mailbox -ResultSize unlimited -AuditLog -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    elseif (get-mailbox -ResultSize unlimited -AuxAuditLog -Identity $user.ObjectGUID.tostring())
    {
        $type = "AuxAuditLog"
        $allstat = get-mailbox -ResultSize unlimited -AuxAuditLog -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    elseif (get-mailbox -ResultSize unlimited -publicfolder -identity $user.ObjectGUID.tostring())
    {
        $type = "Public Folder"
        $allstat = get-mailbox -ResultSize unlimited -PublicFolder -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    elseif (get-mailbox -ResultSize unlimited -Migration -identity $user.ObjectGUID.tostring())
    {
        $type = "Migration"
        $allstat = get-mailbox -ResultSize unlimited -Migration -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    elseif (get-mailbox -ResultSize unlimited -RemoteArchive -identity $user.ObjectGUID.tostring())
    {
        $type = "Remote Archive"
        $allstat = get-mailbox -ResultSize unlimited -RemoteArchive -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    elseif (get-mailbox -ResultSize unlimited -Monitoring -identity $user.ObjectGUID.tostring())
    {
        $type = "Monitoring"
        $allstat = get-mailbox -ResultSize unlimited -Monitoring -Identity $user.ObjectGUID.ToString() | Get-MailboxStatistics
    }
    else
    {
        Write-host "ERROR. Undefined mailbox type, or no mailbox"
    }
    
    New-Object PSObject -Property @{
        DisplayName = $allstat.DisplayName
        StorageLimitStatus = $allstat.StorageLimitStatus
        LastLoginTime = $allstat.LastLogonTime
        Database = $allstat.Database
        ServerName = $allstat.ServerName
        Type = $type
        "Total Item Size (MB)" = [math]::Round(($allstat.TotalItemSize.ToString().Split("(")[1].SPlit(" ")[0].Replace(",","")/1MB),2)
        ItemCount = $allstat.ItemCount
    }

    $i++
    $percent = ($i / $allusers.count) * 100
    Write-host "Progress:" $i "of" $allusers.count "or" $percent "%"
}

$outpath = "C:\Apps\Output\test.csv"
$array | select DisplayName,StorageLimitStatus,LastLoginTime,Database,ServerName,Type,"Total Item Size (MB)",ItemCount | export-csv $outpath -NoTypeInformation
