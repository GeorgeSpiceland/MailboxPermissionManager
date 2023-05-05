param (
    [Parameter()][switch]$GetMailboxes,
    [Parameter()][switch]$GetMailboxPermission,
    [Parameter()][switch]$GetMailboxFolders,
    [Parameter()][string]$Identity
)

function Set-MailboxComboList {
    param (
        $mailboxCombo
    )
    foreach ($mailbox in $mailboxes){
        $mailboxCombo.Items.Add($mailbox.DisplayName)
    }
    
}

function Set-FolderCombo {
    param (
        $folderCombo
    )
    $folderCombo.Items.Add("Inbox")
    $folderCombo.Items.Add("Calendar")
    $folderCombo.Items.Add("Contacts")
}

Add-Type -assembly System.Windows.Forms
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = 'Mailbox Permission Manager'
$main_form.Width = 600
$main_form.Height = 400

$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
If ($isconnected -ne "True") {
    Connect-ExchangeOnline
} else {
    Write-Host $getsessions.Name
    #Remove-PSSession $getsessions.Name
    #Connect-ExchangeOnline
}

$mailboxes = Get-EXOMailbox -ResultSize unlimited -Properties Name,DistinguishedName,Guid,DisplayName

$mailboxCombo = New-Object System.Windows.Forms.ComboBox
$mailboxCombo.Width = 300
$mailboxCombo.Location = New-Object System.Drawing.Point(60,10)
Set-MailboxComboList($mailboxCombo)

$folderCombo = New-Object System.Windows.Forms.ComboBox
$folderCombo.Width = 300
$folderCombo.Location = New-Object System.Drawing.Point(60,40)
Set-FolderCombo($folderCombo)

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Folder Permissions"
$Label2.Location = New-Object System.Drawing.Point(0,70)
$Label2.AutoSize = $true

$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = ""
$Label3.Location = New-Object System.Drawing.Point(10,90)
$Label3.AutoSize = $true


$FolderListButton = New-Object System.Windows.Forms.Button
$FolderListButton.Location = New-Object System.Drawing.Size(400,10)
$FolderListButton.Size = New-Object System.Drawing.Size(130,23)
$FolderListButton.Text = "Get Folder Permissions"

$FolderListButton.Add_Click(
    {
        $identityPath = $mailboxCombo.SelectedItem + ":\" + $folderCombo.SelectedItem
        $OutputObject = Get-MailboxFolderPermission -Identity $identityPath | Out-String
        $Label3.Text = $OutputObject
    }
)

if ($GetMailboxes.IsPresent){
    foreach ($mailbox in $mailboxes){
        Write-Host $mailbox.DisplayName
        Set-MailboxComboList($mailboxCombo)
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

$main_form.Controls.Add($Label2)
$main_form.Controls.Add($Label3)
$main_form.Controls.Add($FolderListButton)
$main_form.Controls.Add($mailboxCombo)
$main_form.Controls.Add($folderCombo)
$main_form.ShowDialog()