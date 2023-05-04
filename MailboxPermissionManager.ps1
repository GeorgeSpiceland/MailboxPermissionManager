param (
    [Parameter()][switch]$GetMailboxes,
    [Parameter()][switch]$GetMailboxPermission,
    [Parameter()][switch]$GetMailboxFolders,
    [Parameter()][string]$Identity
)

$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
If ($isconnected -ne "True") {
    Connect-ExchangeOnline
} else {
    Write-Host $getsessions.Name
}

$mailboxes = Get-EXOMailbox -ResultSize unlimited -Properties Name,DistinguishedName,Guid,DisplayName

if ($GetMailboxes.IsPresent){
    foreach ($mailbox in $mailboxes){
        Write-Host $mailbox.DisplayName
    }
}

if ($GetMailboxPermission.IsPresent){
    foreach ($mailbox in $mailboxes){
        if ($mailbox.DisplayName -eq $Identity){
            Get-MailboxPermission $mailbox
        }
    }
}

if ($GetMailboxFolders.IsPresent){
    foreach ($mailbox in $mailboxes){
        if($Identity.Length -gt 0){
            if($mailbox.DisplayName -eq $Identity){
                Write-Host $mailbox.DisplayName
                Get-MailboxFolder -Identity $Identity -GetChildren
            }
        } else {
            Write-Host $mailbox.Displayname
            Write-Host "------"
            Get-MailboxFolder -GetChildren
            Write-Host ""
        }
    }
}