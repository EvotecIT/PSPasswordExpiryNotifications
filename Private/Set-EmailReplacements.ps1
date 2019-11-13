function Set-EmailReplacements {
    [CmdletBinding()]
    param(
        [string] $Replacement,
        [PSCustomObject] $User,
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $FormattingParameters,
        [System.Collections.DictionaryEntry] $Day
    )

    $Replacement = $Replacement -replace "<<DisplayName>>", $user.DisplayName
    $Replacement = $Replacement -replace "<<DateExpiry>>", $user.DateExpiry
    $Replacement = $Replacement -replace "<<GivenName>>", $user.GivenName
    $Replacement = $Replacement -replace "<<Surname>>", $user.Surname
    $Replacement = $Replacement -replace "<<TimeToExpire>>", $Day.Value
    $Replacement = $Replacement -replace "<<ManagerDisplayName>>", $user.Manager
    $Replacement = $Replacement -replace "<<ManagerEmail>>", $user.ManagerEmail
    return $Replacement
}
