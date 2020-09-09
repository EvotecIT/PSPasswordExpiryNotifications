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
                    #$GroupMembers = Get-ADGroupMember -Identity $Group -Server $($Server.HostName) -ErrorAction Stop -Recursive
                    $GroupMembers = Get-ADGroup -Identity $Group -Server $($Server.HostName) -ErrorAction Stop -Properties Members

                    #foreach ($_ in $GroupMembers) {
                    #    $CachedUsers["$($_.distinguishedName)"]
                    #}
                    foreach ($_ in $GroupMembers.Members) {
                        $CachedUsers["$($_)"]
                    }
                } catch {
                    $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                    Write-Color @WriteParameters '[e] Managers Limited Scope Error: ', $ErrorMessage -Color White, Red
                    continue
                }
            }
        }
    }
    $UsersInGroups
}