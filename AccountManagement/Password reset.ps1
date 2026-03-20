### Reset Password + Generate Secret Link

$UserPrincipalName = Read-Host -Prompt "Enter the user's email (UPN)"

$PasswordReport = @()

# --- Connect ---
Start-Sleep -Seconds 2
Connect-Graph -Scopes User.ReadWrite.All -NoWelcome

# --- Password Generator ---
function Get-Password {
    param (
        [int]$len = 10,
        [bool]$ucfirst = $true,
        [bool]$spchar = $true
    )

    if ($len -lt 12 -or ($len % 2) -ne 0) {
        $len = 12
    }

    $length = $len - 2
    $conso = @('b','c','d','f','g','h','j','k','l','m','n','p','r','s','t','v','w','x','y','z')
    $vowel = @('a','e','i','o','u')
    $spchars = @('!','@','#','$','%','_','&','*','-','+','?','=','£')
    $password = ''
    $max = $length / 2

    for ($i = 1; $i -le $max; $i++) {
        $password += $conso[(Get-Random -Minimum 0 -Maximum 20)]
        $password += $vowel[(Get-Random -Minimum 0 -Maximum 5)]
    }

    if ($spchar) {
        $password = $password.Substring(0, $password.Length - 1) + $spchars[(Get-Random -Minimum 0 -Maximum $spchars.Length)]
    }

    $password += (Get-Random -Minimum 10 -Maximum 100)

    if ($ucfirst) {
        $password = $password.Substring(0, 1).ToUpper() + $password.Substring(1)
    }

    return $password
}

# --- Passphrase (sanitised) ---
function Get-Passphrase {
    $start = "Account"
    $end = get-random -maximum 9999 -minimum 1000
    return $start + $end
}

# --- Generate Password ---
$password = Get-Password -len 14
$securePassword = ConvertTo-SecureString -AsPlainText -Force $password

Write-Output "Resetting password for $UserPrincipalName"

# --- Reset AD Password ---
try {
    Set-ADAccountPassword -Identity $UserPrincipalName -Reset -NewPassword $securePassword
    Enable-ADAccount -Identity $UserPrincipalName
}
catch {
    Write-Warning "Failed to reset password: $_"
    return
}

# --- Force Sync to Azure AD ---
Write-Output "Triggering AD Sync..."
Start-ADSyncSyncCycle -PolicyType Delta

# --- OneTimeSecret Link ---
$url = "https://eu.onetimesecret.com/api/v1/share"

$Secret       = "$password"
$passphrase   = Get-Passphrase
$ttlInSeconds = 604800  # 7 days

$linkbody = "kind=share&secret=$Secret&passphrase=$passphrase&ttl=$ttlInSeconds"
$linkheaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}

$response = Invoke-RestMethod $url -Method 'POST' -Headers $linkheaders -Body $linkbody

$secretkey = $response.secret_key
$secreturl = "https://eu.onetimesecret.com/secret/$secretkey"

# --- Output ---
$result = [PSCustomObject]@{
    UserEmail  = $UserPrincipalName
    Password   = $password
    SecretLink = $secreturl
    Passphrase = $passphrase
}

$PasswordReport += $result

Write-Host "Password reset complete." -ForegroundColor Green
Write-Output $PasswordReport