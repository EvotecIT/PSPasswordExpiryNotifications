function Invoke-ReminderToUsers {
    [cmdletBinding()]
    param(
        [System.Collections.IDictionary] $RemindersToUsers,
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $FormattingParameters,
        [System.Collections.IDictionary] $ConfigurationParameters,
        [Array] $Users
    )
    Invoke-ReminderToUsersInternal -Rule $RemindersToUsers -EmailParameters $EmailParameters -ConfigurationParameters $ConfigurationParameters -FormattingParameters $FormattingParameters -Users $Users
    if (-not $Limits.TestingLimitReached) {
        foreach ($Rule in $RemindersToUsers.Rules) {
            Invoke-ReminderToUsersInternal -Rule $Rule -EmailParameters $EmailParameters -ConfigurationParameters $ConfigurationParameters -FormattingParameters $FormattingParameters -Users $Users
            if ($Limits.TestingLimitReached) {
                break
            }
        }
    }
}