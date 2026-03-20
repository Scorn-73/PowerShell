#Set which user you are coppying mailboxes for

$SourceUser = Read-Host -Prompt 'Enter the email address of the source user'
$TargetUser = Read-Host -Prompt 'Enter the email address of the target user'

#List mailboxes the user has access to

$AdminEmail = Read-Host -Prompt 'Enter the email for the Admin account of the client'

Connect-ExchangeOnline -UserPrincipalName $AdminEmail
$Mailboxes = Get-Mailbox | Get-MailboxPermission -User $SourceUser -ResultSize unlimited | Where-Object { 

$_.AccessRights -contains "FullAccess" -or
$_.AccessRights -contains "SendAs" -or
$_.AccessRights -contains "SendOnBehalf"

}

# Check if any mailboxes were found
if ($Mailboxes.Count -eq 0) 
{
    Write-Output "No mailbox permisisons found for $SourceUser"

    return
}


#Goes through each mailbox and sets the permission for the user

foreach ($Mailbox in $Mailboxes) 
{
    $MailboxIdentity = $Mailbox.Identity
    
    #Remove Full Access permissions
    Add-mailboxPermission -Identity $MailboxIdentity -user $TargetUser -AccessRights FullAccess
    
    #Remove Send As permissions
    Add-RecipientPermission -Identity $MailboxIdentity -Trustee $TargetUser -AccessRights Sendas
    
    #Remove Send On Behalf permissions
    Set-Mailbox -Identity $MailboxIdentity -GrantSendOnBehalfTo @{Add="$TargetUser"}
}

# --- Copy calendar permissions ---
$calendarPermissions = Get-MailboxFolderPermission "$sourceUser`:\Calendar" |
    Where-Object { $_.User -ne "Default" -and $_.User -ne "Anonymous" }

foreach ($permission in $calendarPermissions) {
    Write-Host "Granting $($permission.AccessRights) to $($permission.User) on $targetUser Calendar"
    Add-MailboxFolderPermission "$targetUser`:\Calendar" `
        -User $permission.User `
        -AccessRights $permission.AccessRights
}

Write-Output "All Permissions coppied from all mailboxes for $SourceUser to $TargetUser"