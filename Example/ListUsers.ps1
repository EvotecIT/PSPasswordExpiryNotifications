Import-Module .\PSPasswordExpiryNotifications.psd1 -Force
Clear-Host
#Find-PasswordExpiryCheck | Sort-Object -Property UserPrincipalName | Format-Table *

Find-PasswordExpiryCheck | Where-Object { $_.PasswordNeverExpires -eq $true } | Sort-Object -Property PasswordLastSet | Format-Table *