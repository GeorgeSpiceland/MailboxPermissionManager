param (
    [Parameter()][switch]$GetMailboxes,
    [Parameter()][switch]$GetMailboxPermission,
    [Parameter()][switch]$GetMailboxFolders,
    [Parameter()][string]$Identity
)

function Set-MailboxComboList {
    param (
        $ComboBox
    )
    foreach ($mailbox in $mailboxes){
        $ComboBox.Items.Add($mailbox.PrimarySmtpAddress)
    }
    
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

$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Width = 300
$ComboBox.Location = New-Object System.Drawing.Point(60,10)
Set-MailboxComboList($ComboBox)

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Folder List"
$Label2.Location = New-Object System.Drawing.Point(0,40)
$Label2.AutoSize = $true

$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = ""
$Label3.Location = New-Object System.Drawing.Point(110,40)
$Label3.AutoSize = $true


$FolderListButton = New-Object System.Windows.Forms.Button
$FolderListButton.Location = New-Object System.Drawing.Size(400,10)
$FolderListButton.Size = New-Object System.Drawing.Size(120,23)
$FolderListButton.Text = "Get Folders"

$FolderListButton.Add_Click(
    {
        $Label3.Text = Get-MailboxFolder -Identity $ComboBox.SelectedItem -GetChildren
    }
)

if ($GetMailboxes.IsPresent){
    foreach ($mailbox in $mailboxes){
        Write-Host $mailbox.DisplayName
        Set-MailboxComboList($ComboBox)
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
$main_form.Controls.Add($ComboBox)
$main_form.ShowDialog()