function Get-LowestHighestDays {
    [cmdletBinding()]
    param(
        [System.Collections.IDictionary] $RemindersToUsers
    )

    $HighestLowest = Get-LowestHighestInternal -Rule $RemindersToUsers
    foreach ($Rule in $RemindersToUsers.Rules) {
        $PotentialHighestLowest = Get-LowestHighestInternal -Rule $Rule
        if ($null -eq $PotentialHighestLowest.DayHighest) {
            continue
        }
        if ($null -eq $HighestLowest.DayHighest) {
            $HighestLowest = $PotentialHighestLowest
        } else {
            if ($PotentialHighestLowest.DayHighest -gt $HighestLowest.DayHighest) {
                $HighestLowest.DayHighest = $PotentialHighestLowest.DayHighest
            }
            if ($PotentialHighestLowest.DayLowest -lt $HighestLowest.DayLowest) {
                $HighestLowest.DayLowest = $PotentialHighestLowest.DayLowest
            }
        }
    }
    return $HighestLowest
}