function Find-AllUsers () {
    $users = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0 -and PasswordNotRequired -ne $True } -Properties Manager, DisplayName, GivenName, Surname, SamAccountName, EmailAddress, msDS-UserPasswordExpiryTimeComputed, PasswordExpired, PasswordLastSet, PasswordNotRequired
    $users = $users | Select-Object UserPrincipalName, SamAccountName, DisplayName, GivenName, Surname, EmailAddress, PasswordExpired, PasswordLastSet, PasswordNotRequired,
    @{Name = "Manager"; Expression = { (Get-ADUser $_.Manager).Name }},
    @{Name = "ManagerEmail"; Expression = { (Get-ADUser -Properties Mail $_.Manager).Mail  }},
    @{Name = "DateExpiry"; Expression = { ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")) }},
    @{Name = "DaysToExpire"; Expression = { (NEW-TIMESPAN -Start (GET-DATE) -End ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))).Days }}
    return $users
}