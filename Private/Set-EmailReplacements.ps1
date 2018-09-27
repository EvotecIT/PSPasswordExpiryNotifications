function Set-EmailReplacements($Replacement, $User, $EmailParameters, $FormattingParameters, $Day) {
    <#
    $body = "<p><i>$TableWelcomeMessage</i>"
    if ($($TableData | Measure-Object).Count -gt 0) {
        $body += $TableData | ConvertTo-Html | Out-String
        $body = $body -replace " Added", "<font color=`"green`"><b> Added</b></font>"
        $body = $body -replace " Removed", "<font color=`"red`"><b> Removed</b></font>"
        $body = $body -replace " Deleted", "<font color=`"red`"><b> Deleted</b></font>"
        $body = $body -replace " Changed", "<font color=`"blue`"><b> Changed</b></font>"
        $body = $body -replace " Change", "<font color=`"blue`"><b> Change</b></font>"
        $body = $body -replace " Disabled", "<font color=`"red`"><b> Disabled</b></font>"
        $body = $body -replace " Enabled", "<font color=`"green`"><b> Enabled</b></font>"
        $body = $body -replace " Locked out", "<font color=`"red`"><b> Locked out</b></font>"
        $body = $body -replace " Lockouts", "<font color=`"red`"><b> Lockouts</b></font>"
        $body = $body -replace " Unlocked", "<font color=`"green`"><b> Unlocked</b></font>"
        $body = $body -replace " Reset", "<font color=`"blue`"><b> Reset</b></font>"
        $body += "</p>"
    }
    else {
        $body += "<br><i>No changes happened during that period.</i></p>"
    }
    #>
    $Replacement = $Replacement -replace "<<DisplayName>>", $user.DisplayName
    $Replacement = $Replacement -replace "<<DateExpiry>>", $user.DateExpiry
    $Replacement = $Replacement -replace "<<GivenName>>", $user.GivenName
    $Replacement = $Replacement -replace "<<Surname>>", $user.Surname
    $Replacement = $Replacement -replace "<<TimeToExpire>>", $Day.Value
    $Replacement = $Replacement -replace "<<ManagerDisplayName>>", $user.Manager
    $Replacement = $Replacement -replace "<<ManagerEmail>>", $user.ManagerEmail
    return $Replacement
}
