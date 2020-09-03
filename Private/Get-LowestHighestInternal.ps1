function Get-LowestHighestInternal {
    [cmdletBinding()]
    param(
        [System.Collections.IDictionary] $Rule
    )
    if ($Rule.Enable) {
        if ($Rule.Reminders -is [System.Collections.IDictionary]) {
            $DayHighest = Get-HashMaxValue -hashTable $Rule.Reminders
            $DayLowest = Get-HashMaxValue -hashTable $Rule.Reminders -Lowest
        } else {
            [Array] $OrderedDays = $Rule.Reminders | Sort-Object -Unique
            if ($OrderedDays.Count -gt 0) {
                $DayHighest = $OrderedDays[-1]
                $DayLowest = $OrderedDays[0]
            }
        }
    }
    [ordered] @{
        DayHighest = $DayHighest
        DayLowest  = $DayLowest
    }
}