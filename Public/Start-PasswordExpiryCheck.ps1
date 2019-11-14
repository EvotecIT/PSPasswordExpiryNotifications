Function Start-PasswordExpiryCheck {
    [CmdletBinding()]
    param (
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $FormattingParameters,
        [System.Collections.IDictionary] $ConfigurationParameters
    )
    $time = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    Test-Prerequisits

    $WriteParameters = $ConfigurationParameters.DisplayConsole
    $FieldName = $ConfigurationParameters.RemindersSendToUsers.UseAdditionalField

    $Today = Get-Date
    $CachedUsers = [ordered] @{ }
    $Users = Find-PasswordExpiryCheck -AdditionalProperties $FieldName -WriteParameters $WriteParameters -CachedUsers $CachedUsers | Sort-Object DateExpiry
    $UsersWithEmail = @(
        $Users | Where-Object { $_.EmailAddress -like '*@*' }
    )
    $UsersExpired = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $Today }

    $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters

    $Image = Set-EmailReportBranding -FormattingOptions $FormattingParameters

    $EmailBody += Set-EmailFormatting -Template $FormattingParameters.Template -FormattingParameters $FormattingParameters `
        -ConfigurationParameters $ConfigurationParameters -Image $Image

    $UsersNotified = @(
        [bool] $TestingLimitReached = $false
        #region Send Emails to Users
        if ($ConfigurationParameters.RemindersSendToUsers.Enable -eq $true) {
            Write-Color @WriteParameters '[i] Starting processing ', 'Users', ' section' -Color White, Yellow, White
            foreach ($Day in $ConfigurationParameters.RemindersSendToUsers.Reminders.GetEnumerator()) {
                if ($TestingLimitReached -eq $true) {
                    break
                }
                $Date = (Get-Date).AddDays($Day.Value).Date
                $Count = 0
                foreach ($u in $UsersWithEmail) {
                    if ($u.DateExpiry.Date -eq $Date) {
                        $Count++
                        if ($u.EmailAddress -like '*@*') {
                            Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expires in ", "$($Day.Value)", " days (", "$($u.DateExpiry)", ")."  -Color White, Yellow, White, Red, White, Red, White
                            $TemporaryBody = Set-EmailReplacements -Replacement $EmailBody -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters -Day $Day
                            $EmailSubject = Set-EmailReplacements -Replacement $EmailParameters.EmailSubject -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters -Day $Day
                            $u.DaysToExpire = $Day.Value

                            if ($ConfigurationParameters.RemindersSendToUsers.RemindersDisplayOnly -eq $true) {
                                Write-Color @WriteParameters -Text "[i] Pretending to send email to ", "$($u.EmailAddress)", " ...", "Success"  -Color White, Green, White, Green
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
                                if ($ConfigurationParameters.RemindersSendToUsers.SendToDefaultEmail -eq $false) {
                                    Write-Color @WriteParameters -Text "[i] Sending email to ", "$($u.EmailAddress)", " ..."  -Color White, Green -NoNewLine
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
                            Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expires in ", "$($Day.Value)", " days (", "$($u.DateExpiry)", "). However user has no email address and will be skipped."  -Color White, Yellow, White, Red, White, Red, White
                        }
                        $u
                        if ($ConfigurationParameters.RemindersSendToUsers.SendCountMaximum -eq $Count) {
                            Write-Color @WriteParameters -Text "[i] Sending email to maximum number of users ", "$($ConfigurationParameters.RemindersSendToUsers.SendCountMaximum) ", "has been reached. Skipping..." -Color White, Yellow, White
                            $TestingLimitReached = $true
                            break
                        }
                    }
                }
            }
            Write-Color @WriteParameters '[i] Ending processing ', 'Users', ' section' -Color White, Yellow, White
        } else {
            Write-Color @WriteParameters '[i] Skipping processing ', 'Users', ' section' -Color White, Yellow, White
        }
    )
    #endregion

    #region Send Emails to Managers
    if ($ConfigurationParameters.RemindersSendToManager.Enable -eq $true) {
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
            [Array] $LimitedScopeMembers = Find-LimitedScope -ConfigurationParameters $ConfigurationParameters -CachedUsers $CachedUsers
            [Array] $UsersWithManagers = foreach ($_ in $UsersNotified) {
                if ($LimitedScopeMembers.EmailAddress -contains $_.EmailAddress) {
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
        $Managers = foreach ($u in $UsersWithManagers) {
            $u.ManagerEmail
        }
        $Managers = $Managers | Sort-Object -Unique
        Write-Color @WriteParameters '[i] Preparing package for managers with emails ', "$($UsersWithManagers.Count) ", 'users to process with', ' manager filled in', ' where unique managers ', "$($Managers.Count)" -Color  White, Yellow, White, Yellow, White, Yellow
        # processing one manager at time
        $Count = 0
        foreach ($m in $Managers) {
            $Count++
            # preparing users belonging to manager
            $ColumnNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'
            if ($ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.IncludeNames -ne '') {
                $UsersNotifiedManagers = $UsersWithManagers | Where-Object { $_.ManagerEmail -eq $m } | Select-Object $ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.IncludeNames
            } else {
                $UsersNotifiedManagers = $UsersWithManagers | Where-Object { $_.ManagerEmail -eq $m } | Select-Object 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'EmailSent', 'EmailSentTo'
            }
            if ($ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.Enabled -eq $true) {
                foreach ($u in $UsersNotifiedManagers) {
                    Write-Color @WriteParameters -Text '[-] User ', "$($u.DisplayName) ", " Managers Email (", "$($m)", ')'  -Color White, Yellow, White, Yellow, White
                }
            }
            if ($ConfigurationParameters.RemindersSendToManager.RemindersDisplayOnly -eq $true) {
                Write-Color @WriteParameters -Text "[i] Pretending to send email to manager email ", "$($m)", " ...", "Success"  -Color White, Green, White, Green
                $EmailSent = @{ }
                $EmailSent.Status = $false
                $EmailSent.SentTo = 'N/A'
            } else {
                $TemporaryBody = $EmailBody
                $TemporaryBody = Set-EmailBodyTableReplacement -Body $TemporaryBody -TableName 'ManagerUsersTable' -TableData $UsersNotifiedManagers
                $TemporaryBody = Set-EmailReplacements -Replacement $TemporaryBody -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters #-Day ''

                if ($ConfigurationParameters.Debug.DisplayTemplateHTML -eq $true) { Get-HTML -text $TemporaryBody }
                $EmailSplat = @{
                    EmailParameters = $EmailParameters
                    Body            = $TemporaryBody
                    Subject         = $EmailSubject
                }
                if ($FormattingParameters.CompanyBranding.Inline) {
                    $EmailSplat.InlineAttachments = @{ logo = $FormattingParameters.CompanyBranding.Logo }
                }
                if ($ConfigurationParameters.RemindersSendToManager.SendToDefaultEmail -eq $false) {
                    Write-Color @WriteParameters -Text "[i] Sending email to managers email ", "$($m)", " ..."  -Color White, Green -NoNewLine
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
                Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expired on (", "$($u.DateExpiry)", "). Pretending to disable acoount..."  -Color White, Yellow, White, Red, White, Red, White
            } else {
                Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expired on (", "$($u.DateExpiry)", "). Disabling..."  -Color White, Yellow, White, Red, White, Red, White
                Disable-ADAccount -Identity $u.SamAccountName -Confirm:$false
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Disable Expired Users', ' section' -Color White, Yellow, White
    }

    #region Send Emails to Admins
    if ($ConfigurationParameters.RemindersSendToAdmins.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Administrators', ' section' -Color White, Yellow, White
        $DayHighest = Get-HashMaxValue $ConfigurationParameters.RemindersSendToUsers.Reminders
        $DayLowest = Get-HashMaxValue $ConfigurationParameters.RemindersSendToUsers.Reminders -Lowest
        $DateCountdownStart = (Get-Date).AddDays($DayHighest).Date
        $DateIminnent = (Get-Date).AddDays($DayLowest).Date
        #Write-Color 'Day Highest ', $DayHighest, ' Day Lowest ', $DayLowest, ' Day Countdown Start ', $DateCountdownStart, ' Day Iminnet ', $DateIminnent -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow

        $ColumnNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames -ne '') {
            $UsersNotified = $UsersNotified | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames
        } else {
            $UsersNotified = $UsersNotified | Select-Object $ColumnNames, 'EmailSent', 'EmailSentTo'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames -ne '') {
            $ExpiringIminent = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateIminnent -and $_.PasswordExpired -eq $false } | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames
        } else {
            $ExpiringIminent = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateIminnent -and $_.PasswordExpired -eq $false } | Select-Object $ColumnNames
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames -ne '') {
            $ExpiringCountdownStarted = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateCountdownStart -and $_.PasswordExpired -eq $false } | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames
        } else {
            $ExpiringCountdownStarted = $Users | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $DateCountdownStart -and $_.PasswordExpired -eq $false } | Select-Object $ColumnNames
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.IncludeNames -ne '') {
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
            -CountUsersAlreadyExpired $($UsersExpired.Count)
        $time.Stop()

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Password Notifcations Sent' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $UsersNotified `
                -TableMessageWelcome "Following users had their password notifications sent" `
                -TableMessageNoData 'No users required nofifications.'
        }
        if ( $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Users expiring imminent' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $ExpiringIminent `
                -TableMessageWelcome "Following users expiring imminent (Less than $DayLowest day(s)" `
                -TableMessageNoData 'No users expiring.'
        }
        if (  $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.Enabled -eq $true) {
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
        if ($ConfigurationParameters.Debug.DisplayTemplateHTML -eq $true) { Get-HTML -text $EmailBody }

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