#Set client domain variable
$ClientDomain = "clientdomain.com"

#Connect-ExchangeOnline
$FullAccessResults = @()
$SendAsResults = @()
$SendOnBehalfResults = @()
$mailboxes = Get-Mailbox -ResultSize Unlimited -Filter ('RecipientTypeDetails -eq "UserMailbox"')
  
foreach ($mailbox in $mailboxes)   
{$FullAccessSearch = Get-MailboxPermission -Identity $mailbox.alias -ResultSize Unlimited | ?{($_.IsInherited -eq $false) -and ($_.User -ne "NT AUTHORITY\SELF") -and ($_.AccessRights -like "FullAccess")}
    
     foreach ($entry in $FullAccessSearch) {
        # Handle mailbox UPN
        if ($mailbox.UserPrincipalName -notlike "*@$ClientDomain") {
            try {
                $MGUser = Get-MgUser -UserId $mailbox.UserPrincipalName
                $upn = $MGUser.UserPrincipalName
            } catch {
                $upn = $mailbox.UserPrincipalName
            }
        } else {
            $upn = $mailbox.UserPrincipalName
        }

        # Handle delegate UPN
        $userName = $entry.User.ToString()
        if ($userName -notlike "*@$ClientDomain") {
            try {
                $MGUser = Get-MgUser -UserId $userName
                $Userupn = $MGUser.UserPrincipalName
            } catch {
                $Userupn = $userName
            }
        } else {
            $Userupn = $userName
        }

        $FullAccessResults += [PSCustomObject]@{
            Identity     = $upn
            User         = $Userupn
            AccessRights = ($entry.AccessRights -join ', ')
        }
    }
}

foreach ($mailbox in $mailboxes)   
{$SendAsSearch = Get-RecipientPermission -Identity $mailbox.alias | Where-Object { $_.AccessRights -like "*send*" -and -not ($_.Trustee -match "NT AUTHORITY") -and ($_.IsInherited -eq $false)}
    
     foreach ($entry in $SendAsSearch) {
        # Handle mailbox UPN
        if ($mailbox.UserPrincipalName -notlike "*@$ClientDomain") {
            try {
                $MGUser = Get-MgUser -UserId $mailbox.UserPrincipalName
                $upn = $MGUser.UserPrincipalName
            } catch {
                $upn = $mailbox.UserPrincipalName
            }
        } else {
            $upn = $mailbox.UserPrincipalName
        }

        # Handle delegate UPN
        $userName = $entry.Trustee.ToString()
        if ($userName -notlike "*@$ClientDomain") {
            try {
                $MGUser = Get-MgUser -UserId $userName
                $Userupn = $MGUser.UserPrincipalName
            } catch {
                $Userupn = $userName
            }
        } else {
            $Userupn = $userName
        }

        $SendAsResults += [PSCustomObject]@{
            Identity     = $upn
            User         = $Userupn
            AccessRights = ($entry.AccessRights -join ', ')
        }
    }
}

foreach ($mailbox in $mailboxes)
{$SendOnBehalfSearch = Get-Mailbox -ResultSize Unlimited -Filter ('RecipientTypeDetails -eq "UserMailbox"') | select Alias, GrantSendOnBehalfTo

     foreach ($entry in $SendOnBehalfSearch) {
        # Handle mailbox UPN
        if ($mailbox.UserPrincipalName -notlike "*@$ClientDomain") {
            try {
                $MGUser = Get-MgUser -UserId $mailbox.UserPrincipalName
                $upn = $MGUser.UserPrincipalName
            } catch {
                $upn = $mailbox.UserPrincipalName
            }
        } else {
            $upn = $mailbox.UserPrincipalName
        }

        # Handle delegate UPN
        $userName = $entry.User.ToString()
        if ($userName -notlike "*@$ClientDomain") {
            try {
                $MGUser = Get-MgUser -UserId $userName
                $Userupn = $MGUser.UserPrincipalName
            } catch {
                $Userupn = $userName
            }
        } else {
            $Userupn = $userName
        }

        $SendOnBehalfResults += [PSCustomObject]@{
            Identity     = $upn
            User         = $Userupn
            AccessRights = ($entry.AccessRights -join ', ')
        }
    }
}


$FullAccessResults | Export-Csv -Path "C:\Support\Exports\Full Access Mailbox Permissions.csv" -NoTypeInformation
$SendAsResults | Export-Csv -Path "C:\Support\Exports\Send As Mailbox Permissions.csv" -NoTypeInformation
$SendOnBehalf| Export-Csv -Path "C:\Support\Exports\Send On Behalf Mailbox Permissions.csv" -NoTypeInformation