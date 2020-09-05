﻿Function Start-PasswordExpiryCheck {
    [CmdletBinding()]
    param (
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $FormattingParameters,
        [System.Collections.IDictionary] $ConfigurationParameters
    )
    $time = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    Test-Prerequisits

    $WriteParameters = $ConfigurationParameters.DisplayConsole
    # This takes care of additional fields for all rules (native and additional)
    $FieldName = @(
        $ConfigurationParameters.RemindersSendToUsers.UseAdditionalField
        foreach ($Rule in $ConfigurationParameters.RemindersSendToUsers.Rules | Where-Object { $_.Enable -eq $true }) {
            $Rule.UseAdditionalField
        }
    ) | Sort-Object -Unique

    $Today = Get-Date
    $CachedUsers = [ordered] @{ }
    $CachedUsersPrepared = [ordered] @{ }
    $CachedManagers = [ordered] @{ }

    [Array] $ConditionProperties = if ($FormattingParameters.Conditions) {
        foreach ($Key in $FormattingParameters.Conditions.Keys) {
            foreach ($Condition in $FormattingParameters.Conditions["$Key"].Keys | Where-Object { $_ -ne 'DefaultCondition' }) {
                $Condition
            }
        }
    }
    $Users = Find-PasswordExpiryCheck -AdditionalProperties $FieldName -ConditionProperties $ConditionProperties -WriteParameters $WriteParameters -CachedUsers $CachedUsers -CachedUsersPrepared $CachedUsersPrepared -CachedManagers $CachedManagers | Sort-Object DateExpiry

    # Build a report for expired users
    $UsersExpired = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $Today }

    $UsersNotified = @(
        #region Send Emails to Users
        Invoke-ReminderToUsers -RemindersToUsers $ConfigurationParameters.RemindersSendToUsers -EmailParameters $EmailParameters -ConfigurationParameters $ConfigurationParameters -FormattingParameters $FormattingParameters -EmailBody $EmailBody -Users $Users
        <#
        $Rule = $ConfigurationParameters.RemindersSendToUsers
        if ($Rule.Enable -eq $true) {
            Write-Color @WriteParameters '[i] Starting processing ', 'Users', ' section' -Color White, Yellow, White

            if ($Rule.Reminders -is [System.Collections.IDictionary]) {
                [Array] $DaysToExpire = ($Rule.Reminders).Values | Sort-Object -Unique
            } else {
                [Array] $DaysToExpire = $Rule.Reminders | Sort-Object -Unique
            }
            $Count = 0
            foreach ($u in $Users) {
                if ($TestingLimitReached -eq $true) {
                    break
                }
                if ($u.DaysToExpire -in $DaysToExpire) {
                    if ($u.EmailAddress -like '*@*') {
                        $Count++
                        Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expires in ", "$($u.DaysToExpire)", " days (", "$($u.DateExpiry)", ")." -Color White, Yellow, White, Red, White, Red, White
                        $TemporaryBody = Set-EmailReplacements -Replacement $EmailBody -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters -Day $u.DaysToExpire
                        $EmailSubject = Set-EmailReplacements -Replacement $EmailParameters.EmailSubject -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters -Day $u.DaysToExpire
                        #$u.DaysToExpire = $Day.Value

                        if ($Rule.RemindersDisplayOnly -eq $true) {
                            Write-Color @WriteParameters -Text "[i] Pretending to send email to ", "$($u.EmailAddress)", " ...", "Success" -Color White, Green, White, Green
                            $EmailSent = [ordered] @{ }
                            $EmailSent.Status = $false
                            $EmailSent.SentTo = 'N/A'
                        } else {
                            $EmailSplat = @{
                                EmailParameters = $EmailParameters
                                Body            = $TemporaryBody
                                Subject         = $EmailSubject
                            }
                            if ($FormattingParameters.CompanyBranding.Inline) {
                                $EmailSplat.InlineAttachments = @{ logo = $FormattingParameters.CompanyBranding.Logo }
                            }
                            if ($Rule.SendToDefaultEmail -eq $false) {
                                Write-Color @WriteParameters -Text "[i] Sending email to ", "$($u.EmailAddress)", " ..." -Color White, Green -NoNewLine
                                $EmailSplat.To = $u.EmailAddress
                            } else {
                                Write-Color @WriteParameters -Text "[i] Sending email to users is disabled. Sending email to default value: ", "$($EmailParameters.EmailTo) ", "..." -Color White, Yellow, White -NoNewLine
                            }
                            $EmailSent = Send-Email @EmailSplat
                            if ($EmailSent.Status -eq $true) {
                                Write-Color -Text "Done" -Color "Green"
                            } else {
                                Write-Color -Text "Failed!" -Color "Red"
                            }
                        }
                        Add-Member -InputObject $u -NotePropertyName "EmailSent" -NotePropertyValue $EmailSent.Status
                        Add-Member -InputObject $u -NotePropertyName "EmailSentTo" -NotePropertyValue $EmailSent.SentTo
                    } else {
                        Add-Member -InputObject $u -NotePropertyName "EmailSent" -NotePropertyValue $false
                        Add-Member -InputObject $u -NotePropertyName "EmailSentTo" -NotePropertyValue 'Not available'
                        Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expires in ", "$($u.DaysToExpire)", " days (", "$($u.DateExpiry)", "). However user has no email address and will be skipped." -Color White, Yellow, White, Red, White, Red, White
                    }
                    $u
                    if ($Rule.SendCountMaximum -eq $Count) {
                        Write-Color @WriteParameters -Text "[i] Sending email to maximum number of users ", "$($Rule.SendCountMaximum) ", "has been reached. Skipping..." -Color White, Yellow, White
                        $TestingLimitReached = $true
                        break
                    }
                }
            }
            #}
            Write-Color @WriteParameters '[i] Ending processing ', 'Users', ' section' -Color White, Yellow, White
        } else {
            Write-Color @WriteParameters '[i] Skipping processing ', 'Users', ' section' -Color White, Yellow, White
        }
        #>
    )
    #endregion

    #region Send Emails to Managers
    $ManagersReceived = if ($ConfigurationParameters.RemindersSendToManager.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Managers', ' section' -Color White, Yellow, White
        # preparing email
        $EmailSubject = $ConfigurationParameters.RemindersSendToManager.ManagersEmailSubject
        $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters
        $EmailReportBranding = Set-EmailReportBranding -FormattingOptions $FormattingParameters
        $EmailBody += Set-EmailFormatting -Template $FormattingParameters.TemplateForManagers `
            -FormattingParameters $FormattingParameters `
            -ConfigurationParameters $ConfigurationParameters `
            -AddAfter $EmailReportBranding

        # preparing manager lists
        if ($ConfigurationParameters.RemindersSendToManager.LimitScope.Groups) {
            # send emails to managers only if those people are in limited scope groups
            [Array] $LimitedScopeMembers = Find-LimitedScope -ConfigurationParameters $ConfigurationParameters -CachedUsers $CachedUsersPrepared
            [Array] $UsersWithManagers = foreach ($_ in $UsersNotified) {
                if ($LimitedScopeMembers.UserPrincipalName -contains $_.UserPrincipalName) {
                    if ($null -ne $_.ManagerEmail) {
                        $_
                    }
                }
            }
        } else {
            [Array] $UsersWithManagers = foreach ($_ in $UsersNotified) {
                if ($null -ne $_.ManagerEmail) {
                    $_
                }
            }
            # $UsersWithManagers = $UsersNotified | Where-Object { $null -ne $_.ManagerEmail }
        }
        # Find managers with emails. Make sure only unique is added to the list
        $ManagersEmails = [System.Collections.Generic.List[string]]::new()
        foreach ($u in $UsersWithManagers) {
            if ($ManagersEmails -notcontains $u.ManagerEmail) {
                $ManagersEmails.Add($u.ManagerEmail)
            }
        }
        Write-Color @WriteParameters '[i] Preparing package for managers with emails ', "$($UsersWithManagers.Count) ", 'users to process with', ' manager filled in', ' where unique managers ', "$($ManagersEmails.Count)" -Color White, Yellow, White, Yellow, White, Yellow
        # processing one manager at time
        $Count = 0
        foreach ($m in $ManagersEmails) {
            $Count++

            # preparing users belonging to manager
            $ColumnNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'

            [Array] $UsersNotifiedManagers = $UsersWithManagers | Where-Object { $_.ManagerEmail -eq $m }
            [string] $ManagerDN = $UsersNotifiedManagers[0].ManagerDN
            $ManagerFull = $CachedManagers[$ManagerDN]
            if ($ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.IncludeNames -ne '') {
                $UsersNotifiedManagers = $UsersNotifiedManagers | Select-Object $ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.IncludeNames
            } else {
                $UsersNotifiedManagers = $UsersNotifiedManagers | Select-Object 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'EmailSent', 'EmailSentTo'
            }
            if ($ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.Enabled -eq $true) {
                foreach ($u in $UsersNotifiedManagers) {
                    Write-Color @WriteParameters -Text '[-] User ', "$($u.DisplayName) ", " Managers Email (", "$($m)", ')' -Color White, Yellow, White, Yellow, White
                }
            }

            if ($ConfigurationParameters.RemindersSendToManager.RemindersDisplayOnly -eq $true) {
                Write-Color @WriteParameters -Text "[i] Pretending to send email to manager email ", "$($m)", " ...", "Success" -Color White, Green, White, Green
                $EmailSent = @{ }
                $EmailSent.Status = $false
                $EmailSent.SentTo = 'N/A'
            } else {
                $TemporaryBody = $EmailBody
                $TemporaryBody = Set-EmailBodyReplacementTable -Body $TemporaryBody -TableName 'ManagerUsersTable' -TableData $UsersNotifiedManagers
                $TemporaryBody = Set-EmailReplacements -Replacement $TemporaryBody -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters #-Day ''

                if ($ConfigurationParameters.Debug.DisplayTemplateHTML -eq $true) {
                    Get-HTML -text $TemporaryBody
                }
                $EmailSplat = @{
                    EmailParameters = $EmailParameters
                    Body            = $TemporaryBody
                    Subject         = $EmailSubject
                }
                if ($FormattingParameters.CompanyBranding.Inline) {
                    $EmailSplat.InlineAttachments = @{ logo = $FormattingParameters.CompanyBranding.Logo }
                }
                if ($ConfigurationParameters.RemindersSendToManager.SendToDefaultEmail -eq $false) {
                    Write-Color @WriteParameters -Text "[i] Sending email to managers email ", "$($m)", " ..." -Color White, Green -NoNewLine
                    $EmailSplat.To = $m
                } else {
                    Write-Color @WriteParameters -Text "[i] Sending email to managers is disabled. Sending email to default value: ", "$($EmailParameters.EmailTo) ", "..." -Color White, Yellow, White -NoNewLine
                }
                $EmailSent = Send-Email @EmailSplat
                if ($EmailSent.Status -eq $true) {
                    Write-Color -Text "Done" -Color "Green"
                } else {
                    Write-Color -Text "Failed!" -Color "Red"
                }
            }

            $ManagerFull['EmailSent'] = $EmailSent.Status
            $ManagerFull['EmailSentTo'] = $EmailSent.SentTo

            if ($ConfigurationParameters.RemindersSendToManager.Reports.IncludeManagersPasswordNotificationsSent.IncludeNames.Count -gt 0) {
                ([PSCustomObject] $ManagerFull) | Select-Object -Property $ConfigurationParameters.RemindersSendToManager.Reports.IncludeManagersPasswordNotificationsSent.IncludeNames
            } else {
                ([PSCustomObject] $ManagerFull) | Select-Object -Property 'UserPrincipalName', 'Domain', 'DisplayName', 'SamAccountName', 'EmailSent', 'EmailSentTo'
            }


            if ($ConfigurationParameters.RemindersSendToManager.SendCountMaximum -eq $Count) {
                Write-Color @WriteParameters -Text "[i] Sending email to maximum number of managers ", "$($ConfigurationParameters.RemindersSendToManager.SendCountMaximum) ", " has been reached. Skipping..." -Color White, Yellow, White -NoNewLine
                break
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Managers', ' section' -Color White, Yellow, White
    } else {
        Write-Color @WriteParameters '[i] Skipping processing ', 'Managers', ' section' -Color White, Yellow, White
    }
    #endregion Send Emails to Managers


    if ($ConfigurationParameters.DisableExpiredUsers.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Disable Expired Users', ' section' -Color White, Yellow, White
        foreach ($U in $UsersExpired) {
            if ($ConfigurationParameters.DisableExpiredUsers.DisplayOnly) {
                Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expired on (", "$($u.DateExpiry)", "). Pretending to disable acoount..." -Color White, Yellow, White, Red, White, Red, White
            } else {
                Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expired on (", "$($u.DateExpiry)", "). Disabling..." -Color White, Yellow, White, Red, White, Red, White
                Disable-ADAccount -Identity $u.SamAccountName -Confirm:$false
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Disable Expired Users', ' section' -Color White, Yellow, White
    }

    #region Send Emails to Admins
    if ($ConfigurationParameters.RemindersSendToAdmins.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Administrators', ' section' -Color White, Yellow, White

        $SummaryDays = Get-LowestHighestDays -RemindersToUsers $ConfigurationParameters.RemindersSendToUsers
        $DayHighest = $SummaryDays.DayHighest
        $DayLowest = $SummaryDays.DayLowest
        if ($null -eq $DayHighest -or $null -eq $DayLowest) {
            # Skip reports because reminders are not set at all - weird
            <#
            $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeSummary.Enabled = $false
            $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.Enabled = $false
            $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeManagersPasswordNotificationsSent.Enabled = $false
            $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.Enabled = $false
            $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.Enabled = $false
            $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.Enabled = $false
            #>
        }
        $DateCountdownStart = (Get-Date).AddDays($DayHighest).Date
        $DateIminnent = (Get-Date).AddDays($DayLowest).Date

        $ColumnNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames -gt 0) {
            $UsersNotified = $UsersNotified | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames
        } else {
            $UsersNotified = $UsersNotified | Select-Object $ColumnNames, 'EmailSent', 'EmailSentTo'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames.Count -gt 0) {
            $ExpiringIminent = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateIminnent -and $_.PasswordExpired -eq $false -and $_.PasswordNeverExpires -eq $false } | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames
        } else {
            $ExpiringIminent = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateIminnent -and $_.PasswordExpired -eq $false -and $_.PasswordNeverExpires -eq $false } | Select-Object $ColumnNames
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames -gt 0) {
            $ExpiringCountdownStarted = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateCountdownStart -and $_.PasswordExpired -eq $false -and $_.PasswordNeverExpires -eq $false } | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames
        } else {
            $ExpiringCountdownStarted = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateCountdownStart -and $_.PasswordExpired -eq $false -and $_.PasswordNeverExpires -eq $false } | Select-Object $ColumnNames
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.IncludeNames -gt 0) {
            $UsersExpired = $UsersExpired | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.IncludeNames
        } else {
            $UsersExpired = $UsersExpired | Select-Object $ColumnNames
        }

        $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters
        $EmailBody += "<body>"
        $EmailBody += Set-EmailReportBranding -FormattingOptions $FormattingParameters
        $EmailBody += Set-EmailReportDetails -FormattingOptions $FormattingParameters `
            -ReportOptions $ReportOptions `
            -TimeToGenerate $Time.Elapsed `
            -CountUsersCountdownStarted $($ExpiringCountdownStarted.Count) `
            -CountUsersImminent $($ExpiringIminent.Count) `
            -CountUsersAlreadyExpired $($UsersExpired.Count) -CountUsersNotified $($UsersNotified.Count)
        $time.Stop()

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeSummary.Enabled -eq $true) {
            $SummaryOfUsers = $Users | Group-Object DaysToExpire `
            | Select-Object @{Name = 'Days to Expire'; Expression = { [int] $($_.Name) } }, @{Name = 'Users with Days to Expire'; Expression = { [int] $($_.Count) } }
            $SummaryOfUsers = $SummaryOfUsers | Sort-Object -Property 'Days to Expire'

            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Summary of Expiring Users' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $SummaryOfUsers `
                -TableMessageWelcome "Summary of days to expire and it's count" `
                -TableMessageNoData 'There were no users that have days of expiring.'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Password Notifcations Sent' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $UsersNotified `
                -TableMessageWelcome "Following users had their password notifications sent" `
                -TableMessageNoData 'No users required nofifications.'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeManagersPasswordNotificationsSent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Password Notifcations Sent to Managers' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $ManagersReceived `
                -TableMessageWelcome "Following managers had their password bundle notifications sent" `
                -TableMessageNoData 'No managers required nofifications.'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Users expiring imminent' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $ExpiringIminent `
                -TableMessageWelcome "Following users expiring imminent (Less than $DayLowest day(s)" `
                -TableMessageNoData 'No users expiring.'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Expiring Couintdown Started' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $ExpiringCountdownStarted `
                -TableMessageWelcome "Following users expiring countdown started (Less than $DayHighest day(s))" `
                -TableMessageNoData 'There were no users that had their coundown started.'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Users are already expired' -Color White, Yellow
            if ($ConfigurationParameters.DisableExpiredUsers.Enable -eq $true -and -not $ConfigurationParameters.DisableExpiredUsers.DisplayOnly -eq $true) {
                $EmailBody += Set-EmailBody -TableData $UsersExpired -TableMessageWelcome "Following users are already expired (and were disabled...)" -TableMessageNoData "No users that are expired."
            } else {
                $EmailBody += Set-EmailBody -TableData $UsersExpired -TableMessageWelcome "Following users are already expired (and still enabled...)" -TableMessageNoData "No users that are expired and enabled."
            }
        }
        $EmailBody += "</body>"
        if ($ConfigurationParameters.Debug.DisplayTemplateHTML -eq $true) {
            Get-HTML -text $EmailBody
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.RemindersDisplayOnly -eq $true) {
            Write-Color @WriteParameters -Text "[i] Pretending to send email to admins email ", "$($ConfigurationParameters.RemindersSendToAdmins.AdminsEmail) ", "...", 'Success' -Color White, Yellow, White, Green
        } else {
            Write-Color @WriteParameters -Text "[i] Sending email to administrators on email address ", "$($ConfigurationParameters.RemindersSendToAdmins.AdminsEmail) ", "..." -Color White, Yellow, White -NoNewLine
            $EmailSplat = @{
                EmailParameters = $EmailParameters
                Body            = $EmailBody
                Subject         = $ConfigurationParameters.RemindersSendToAdmins.AdminsEmailSubject
                To              = $ConfigurationParameters.RemindersSendToAdmins.AdminsEmail
            }
            if ($FormattingParameters.CompanyBranding.Inline) {
                $EmailSplat.InlineAttachments = @{ logo = $FormattingParameters.CompanyBranding.Logo }
            }
            $EmailSent = Send-Email @EmailSplat
            if ($EmailSent.Status -eq $true) {
                Write-Color -Text "Done" -Color "Green"
            } else {
                Write-Color -Text "Failed! Error: $($EmailSent.Error)" -Color "Red"
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Administrators', ' section' -Color White, Yellow, White

    } else {
        Write-Color @WriteParameters '[i] Skipping processing ', 'Administrators', ' section' -Color White, Yellow, White

    }
    #endregion Send Emails to Admins
}
