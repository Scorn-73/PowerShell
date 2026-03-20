#Closure template
#Check variables are correct first
#check correct username is made from the full name

#Set Variables
$User                = Read-Host Prompt "Enter the full name of the user"
$Client              = "CLIENT"
$EmailDomain         = "INPUTCLIENTDOMAIN"
$UserFirstName       = $user.split(' ')[0]
$UserLastName        = $user.split(' ')[1]
$username            = $UserFirstName[0] +"."+ $UserLastName
$Useremail           = $Username + $EmailDomain
$AdminEmail          = "INPUTADMINEMAIL"
$CurrentDate         = Get-Date -Format "dd.MM.yy"
$Ticketnumber        = Read-host Prompt "Enter Ticket Number"
$UserInfo            = Get-ADUser -Filter "EmailAddress -eq '$useremail'"
$UserDN              = $UserInfo.DistinguishedName
$closedaccountDN     = "INPUT DN FOR CLOSED ACCOUNTS FOLDER IN AD"
$ChallowUsername     = whoami
$ChallowInitialLower = $ChallowUsername.split('w')[1]
$ChallowInitialUpper = $ChallowInitialLower.ToUpper()
$message = "O365 Admin Credentials for "+$Client
$o365AdminCred = Get-Credential -Message $message $AdminEmail

#connect to exchange and graph
Start-Sleep -Seconds 2
Connect-Graph -Scopes User.ReadWrite.All, Organization.Read.All -NoWelcome
Connect-ExchangeOnline -Credential $o365AdminCred -ShowBanner:$false

$userlicensedetail = Get-mguserlicensedetail -UserId $Useremail
$Skuid             = $userlicensedetail.SkuId


Write-Output "
Closing full account for $User

"

Write-Output "
------ Retrieving mailbox permissions ------"
$Mailboxes = Get-Mailbox | Get-MailboxPermission -User $UserEmail -ResultSize unlimited | Where-Object {

    $_.AccessRights -contains "FullAccess" -or
    $_.AccessRights -contains "SendAs" -or
    $_.AccessRights -contains "SendOnBehalf"
}

# Check if any mailboxes were found
if ($Mailboxes.Count -eq 0){
    Write-Output "No mailbox permisisons found for $UserEmail"    
}else{

    #Export the results to a csv
    $Mailboxes | Export-csv -Path "C:\support\mailboxes\$user.csv"

    #Goes through each mailbox and removes the permission for the user
    Write-Output "
    ------ Removing Mailbox permissions ------"
    foreach ($Mailbox in $Mailboxes) {
        $MailboxIdentity = $Mailbox.Identity
    
        #Remove Full Access permissions
        Remove-mailboxPermission -Identity $MailboxIdentity -user $UserEmail -AccessRights FullAccess -Confirm:$False
    
        #Remove Send As permissions
        Remove-RecipientPermission -Identity $MailboxIdentity -Trustee $UserEmail -AccessRights Sendas -Confirm:$False
    
        #Remove Send On Behalf permissions
        Set-Mailbox -Identity $MailboxIdentity -GrantSendOnBehalfTo @{remove=$UserEmail} -Confirm:$False
    }
    Write-Output "All Permissions removed from all mailboxes for $UserEmail"
}



#Convert to shared mailbox

Set-Mailbox -Identity $Username -Type Shared

#Removes O365 Group Ownership
Write-Output "
------ Checking for O365 Group Memberships ------"
$O365GroupsOwner = Get-UnifiedGroup | Where-Object {
    (Get-UnifiedGroupLinks $_.Identity -LinkType Owners).PrimarySMTPAddress -contains $UserEmail
}

if ($O365GroupsOwner.Count -eq 0) {
    Write-Output "

    No Office 365 Group Ownership Found

    "
}
Else{

    Foreach ($O365GroupOwner in $O365GroupsOwner){

        Remove-UnifiedGroupLinks -Identity $O365GroupOwner.Identity -LinkType Owners -Links $UserEmail -Confirm:$false

        Write-Output "
    
        Removed Ownership from the following Office 365 Groups:
        $($O365GroupOwner)

        "
    }
}


#Remove Office 365 Group Membership
$O365GroupsMember = Get-UnifiedGroup | Where-Object {
    (Get-UnifiedGroupLinks $_.Identity -LinkType Members).PrimarySMTPAddress -contains $UserEmail
}

if ($O365GroupsMember.Count -eq 0) {
    Write-Output "

    No Office 365 Group Membership Found

    "
}
Else{

    Foreach ($O365GroupMember in $O365GroupsMember){

        Remove-UnifiedGroupLinks -Identity $O365GroupMember.Identity -LinkType Members -Links $UserEmail -Confirm:$false

        Write-Output "
    
    Removed Membership from the following Office 365 Groups:
    $($O365GroupMember)

        "
    }
}

Write-Output "
------ blocking account and removing licence ------"

#Blocks User account in O365
$params = @{
	accountEnabled = $false
}
Update-MgUser -UserId "$UserEmail" -BodyParameter $params

#Remove Licence applied to the user
Set-MgUserLicense -UserId "$UserEmail" -RemoveLicenses @($Skuid) -AddLicenses @() |Out-Null

Write-Output "
------ Closing account in AD ------"

### AD Closure

###Set Description for User
$CurrentDescription = (Get-ADUser -Identity "$UserDN" -Property Description).Description
$NewDescription = "AC Closed $CurrentDate $ChallowInitialUpper $TicketNumber $CurrentDescription"
Set-ADUser -Identity "$UserDN" -Description $NewDescription

###Hide from Address List
Get-ADuser -Identity $UserDN -property msExchHideFromAddressLists |  Set-ADObject -Replace @{msExchHideFromAddressLists=$true}

###Remove Group Membership
Get-AdPrincipalGroupMembership -Identity $UserDN | Where-Object -Property Name -ne -value "Domain Users" | Remove-AdGroupMember -Members $UserDN -Confirm:$False

#Disable Account in AD
Disable-ADAccount -Identity $UserDN -Confirm:$False

###Move User in AD to Closed accounts group
#Splits the DN of the User to just have their Department
Move-ADObject -Identity "$UserDN" -TargetPath "$closedaccountDN"

Write-host "------ The Account for $User has been closed. They have been moved into the Closed Accounts folder ------" -ForegroundColor Cyan
Write-Host "Deactivate user in CRM and add Lyn as a resource" -BackgroundColor White -ForegroundColor DarkRed
Start-ADSyncSyncCycle -PolicyType Delta