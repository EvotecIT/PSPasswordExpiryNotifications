function Find-AllUsers {
    [CmdletBinding()]
    param(
        [string] $AdditionalProperties,
        [hashtable] $WriteParameters
    )
    $Properties = @(
        'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress', 'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired'
        if ($AdditionalProperties) {
            $AdditionalProperties
        }
    )
    try {
        $Users = Get-ADUser -filter { Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0 -and PasswordNotRequired -ne $True } -Properties $Properties -ErrorAction Stop

        $ProcessedUsers = foreach ($_ in $Users) {
            if ($null -ne $_.Manager) {
                $Manager = Get-ADUser $_.Manager -Properties Mail
            } else {
                $Manager = $null
            }

            if ($AdditionalProperties) {
                $EmailTemp = $_.$AdditionalProperties
                if ($EmailTemp -like '*@*') {
                    $EmailAddress = $EmailTemp
                } else {
                    $EmailAddress = $_.EmailAddress
                }
            } else {
                $EmailAddress = $_.EmailAddress
            }


            [PSCustomobject] @{
                UserPrincipalName   = $_.UserPrincipalName
                SamAccountName      = $_.SamAccountName
                DisplayName         = $_.DisplayName
                GivenName           = $_.GivenName
                Surname             = $_.Surname
                EmailAddress        = $EmailAddress
                PasswordExpired     = $_.PasswordExpired
                PasswordLastSet     = $_.PasswordLastSet
                PasswordNotRequired = $_.PasswordNotRequired
                "Manager"           = $Manager.Name
                "ManagerEmail"      = $Manager.Mail
                "DateExpiry"        = ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))
                "DaysToExpire"      = (NEW-TIMESPAN -Start (GET-DATE) -End ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))).Days
            }
        }
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Color @WriteParameters '[e] Error: ', $ErrorMessage -Color White, Red
        Exit
    }
    return $ProcessedUsers
}