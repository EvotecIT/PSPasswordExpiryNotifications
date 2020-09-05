function Set-EmailReplacements {
    [CmdletBinding()]
    param(
        [string] $Replacement,
        [PSCustomObject] $User,
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $FormattingParameters,
        [int] $Day
    )

    $Replacement = $Replacement -replace "<<DisplayName>>", $user.DisplayName
    $Replacement = $Replacement -replace "<<DateExpiry>>", $user.DateExpiry
    $Replacement = $Replacement -replace "<<GivenName>>", $user.GivenName
    $Replacement = $Replacement -replace "<<Surname>>", $user.Surname
    $Replacement = $Replacement -replace "<<TimeToExpire>>", $Day
    $Replacement = $Replacement -replace "<<ManagerDisplayName>>", $user.Manager
    $Replacement = $Replacement -replace "<<ManagerEmail>>", $user.ManagerEmail

    if ($FormattingParameters.Conditions) {
        foreach ($Key in $FormattingParameters.Conditions.Keys) {
            $Found = $false
            $ReplaceFrom = "<<$Key>>"
            $DefaultReplaceTo = $FormattingParameters.Conditions["$Key"]['DefaultCondition']
            foreach ($Condition in $FormattingParameters.Conditions["$Key"].Keys | Where-Object { $_ -ne 'DefaultCondition' }) {
                if ($FormattingParameters.Conditions["$Key"]["$Condition"]) {
                    foreach ($SubKey in $FormattingParameters.Conditions["$Key"]["$Condition"].Keys) {
                        if ($SubKey -eq $User.$Condition) {
                            $Replacement = $Replacement -replace "$ReplaceFrom", $FormattingParameters.Conditions["$Key"]["$Condition"][$SubKey]
                            $Found = $true
                        }
                    }
                }
            }
            if (-not $Found) {
                $Replacement = $Replacement -replace "$ReplaceFrom", $DefaultReplaceTo
            }
        }
    }
    return $Replacement
}