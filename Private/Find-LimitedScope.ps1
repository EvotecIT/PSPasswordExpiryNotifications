function Find-LimitedScope {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $ConfigurationParameters,
        [System.Collections.IDictionary] $CachedUsers
    )
    $Forest = Get-ADForest
    $UsersInGroups = if ($ConfigurationParameters.RemindersSendToManager.LimitScope) {
        foreach ($Group in $ConfigurationParameters.RemindersSendToManager.LimitScope.Groups) {
            foreach ($Domain in $Forest.Domains) {
                $Server = Get-ADDomainController -Discover -DomainName $Domain
                try {
                    $GroupMembers = Get-ADGroupMember -Identity $Group -Server $Server -ErrorAction Stop -Recursive
                    #$GroupMembers

                    foreach ($_ in $GroupMembers) {
                        $CachedUsers["$($_.distinguishedName)"]
                    }

                } catch {
                    $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                    Write-Color @WriteParameters '[e] Error: ', $ErrorMessage -Color White, Red
                    continue
                }
            }
        }
    }
    $UsersInGroups
}