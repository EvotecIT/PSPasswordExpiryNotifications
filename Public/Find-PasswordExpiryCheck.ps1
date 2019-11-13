function Find-PasswordExpiryCheck {
    [CmdletBinding()]
    param(
        [string] $AdditionalProperties,
        [System.Collections.IDictionary] $WriteParameters,
        [System.Collections.IDictionary] $CachedUsers
    )
    if ($null -eq $WriteParameters) {
        $WriteParameters = @{
            ShowTime   = $true
            LogFile    = ""
            TimeFormat = "yyyy-MM-dd HH:mm:ss"
        }
    }


    $Properties = @(
        'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress', 'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired', 'Enabled', 'PasswordNeverExpires', 'Mail'
        if ($AdditionalProperties) {
            $AdditionalProperties
        }
    )
    # We're caching all users to make sure it's speedy gonzales when querying for Managers
    if (-not $CachedUsers) {
        $CachedUsers = [ordered] @{ }
    }
    $Forest = Get-ADForest

    $Users = @(
        try {

            foreach ($Domain in $Forest.Domains) {
                Write-Color @WriteParameters -Text "[i] Processing ", "$($Domain)", " for users in forest ", $Forest.Name -Color White, Yellow, White, Yellow, White, Yellow, White

                $Server = Get-ADDomainController -Discover -DomainName $Domain -ErrorAction Stop
                #$Users = Get-ADUser -Server $Server -Filter { Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0 -and PasswordNotRequired -ne $True } -Properties $Properties -ErrorAction Stop

                # We query all users instead of using filter. Since we need manager field and manager data this way it should be faster (query once - get it all)
                $DomainUsers = Get-ADUser -Server $Server -Filter '*' -Properties $Properties -ErrorAction Stop
                foreach ($_ in $DomainUsers) {
                    Add-Member -InputObject $_ -Value $Domain -Name 'Domain' -Force -Type NoteProperty
                    $CachedUsers["$($_.DistinguishedName)"] = $_
                    # We reuse filtering
                    if ($_.Enabled -eq $true -and $_.PasswordNeverExpires -eq $false -and $_.PasswordLastSet -gt 0 -and $_.PasswordNotRequired -ne $true) {
                        $_
                    }
                }
            }
        } catch {
            $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
            Write-Color @WriteParameters '[e] Error: ', $ErrorMessage -Color White, Red
        }
    )
    $ProcessedUsers = foreach ($_ in $Users) {
        <#
        if ($LargeScope) {
            # if large scope is used it means the domain has most likely a lot of users. This means it's inefficient to quetrry

        } else {
            if ($null -ne $_.Manager) {
                $Manager = Get-ADUser $_.Manager -Properties Mail -Server $Server
            } else {
                $Manager = $null
            }
        }
        #>

        $UserManager = $CachedUsers["$($_.Manager)"]


        #$Manager = $UserManager.DisplayName


        if ($AdditionalProperties) {
            # fix trhis for a user
            $EmailTemp = $_.$AdditionalProperties
            if ($EmailTemp -like '*@*') {
                $EmailAddress = $EmailTemp
            } else {
                $EmailAddress = $_.EmailAddress
            }
            # Fix this for manager as well
            if ($UserManager) {
                if ($UserManager.$AdditionalProperties -like '*@*') {
                    $UserManager.Mail = $UserManager.$AdditionalProperties
                }
            }
        } else {
            $EmailAddress = $_.EmailAddress
        }

        if ($_."msDS-UserPasswordExpiryTimeComputed" -ne 9223372036854775807) {
            # This is standard situation where users password is expiring as needed
            try {
                $DateExpiry = ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))
            } catch {
                $DateExpiry = $_."msDS-UserPasswordExpiryTimeComputed"
            }
            try {
                $DaysToExpire = (New-TimeSpan -Start (Get-Date) -End ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))).Days
            } catch {
                $DaysToExpire = $null
            }
            $PasswordNeverExpires = $_.PasswordNeverExpires
        } else {
            # This is non-standard situation. This basically means most likely Fine Grained Group Policy is in action where it makes PasswordNeverExpires $true
            # Since FGP policies are a bit special they do not tick the PasswordNeverExpires box, but at the same time value for "msDS-UserPasswordExpiryTimeComputed" is set to 9223372036854775807
            $DateExpiry = $null
            $DaysToExpire = $null
            $PasswordNeverExpires = $true
        }

        [PSCustomobject] @{
            UserPrincipalName    = $_.UserPrincipalName
            Domain               = $_.Domain
            SamAccountName       = $_.SamAccountName
            DisplayName          = $_.DisplayName
            GivenName            = $_.GivenName
            Surname              = $_.Surname
            EmailAddress         = $EmailAddress
            PasswordExpired      = $_.PasswordExpired
            PasswordLastSet      = $_.PasswordLastSet
            PasswordNotRequired  = $_.PasswordNotRequired
            PasswordNeverExpires = $PasswordNeverExpires
            "Manager"            = $UserManager.Name
            "ManagerEmail"       = $UserManager.Mail
            "DateExpiry"         = $DateExpiry
            "DaysToExpire"       = $DaysToExpire
        }
        #$CachedUsers["$($_.DistinguishedName)"] = $UserToReturn
    }
    $ProcessedUsers
}

#$Test = Find-PasswordExpiryCheck -AdditionalProperties 'extensionAttribute13'
#$Test | Format-Table -AutoSize *