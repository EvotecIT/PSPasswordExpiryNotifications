Function Start-PasswordExpiryCheck {
    param (
        [hashtable] $EmailParameters,
        [hashtable] $FormattingParameters,
        [hashtable] $ConfigurationParameters
    )
    $time = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    Test-Prerequisits
    $WriteParameters = $ConfigurationParameters.DisplayConsole

    $Today = get-date
    $Users = Find-AllUsers | Sort-Object DateExpiry

    #$UsersWithoutEmail = $Users | Where-Object { $_.EmailAddress -eq $null }
    $UsersWithEmail = $Users | Where-Object { $_.EmailAddress -ne $null }
    $UsersExpired = $Users | Where-Object { $_.DateExpiry -lt $Today }

    $UsersNotified = @()
    $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters
    $EmailBody += Set-EmailReportBranding -FormattingOptions $FormattingParameters
    $EmailBody += Set-EmailFormatting -Template $FormattingParameters.Template -FormattingParameters $FormattingParameters -ConfigurationParameters $ConfigurationParameters

    #region Send Emails to Users
    if ($ConfigurationParameters.RemindersSendToUsers.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Users', ' section' -Color White, Yellow, White
        foreach ($Day in $ConfigurationParameters.RemindersSendToUsers.Reminders.GetEnumerator()) {
            $Date = (get-date).AddDays($Day.Value).Date
            foreach ($u in $UsersWithEmail) {
                if ($u.DateExpiry.Date -eq $Date) {

                    Write-Color @WriteParameters -Text "[i] User ", "$($u.DisplayName)", " expires in ", "$($Day.Value)", " days (", "$($u.DateExpiry)", ")."  -Color White, Yellow, White, Red, White, Red, White
                    $TemporaryBody = Set-EmailReplacements -Replacement $EmailBody -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters -Day $Day
                    $EmailSubject = Set-EmailReplacements -Replacement $EmailParameters.EmailSubject -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters -Day $Day
                    $u.DaysToExpire = $Day.Value

                    if ($ConfigurationParameters.RemindersSendToUsers.RemindersDisplayOnly -eq $true) {
                        Write-Color @WriteParameters -Text "[i] Pretending to send email to ", "$($u.EmailAddress)", " ...", "Success"  -Color White, Green, White, Green
                        $EmailSent = @{}
                        $EmailSent.Status = $false
                        $EmailSent.SentTo = 'N/A'
                    } else {
                        if ($ConfigurationParameters.RemindersSendToUsers.SendToDefaultEmail -eq $false) {
                            Write-Color @WriteParameters -Text "[i] Sending email to ", "$($u.EmailAddress)", " ..."  -Color White, Green -NoNewLine
                            $EmailSent = Send-Email `
                                -EmailParameters $EmailParameters `
                                -Body $TemporaryBody `
                                -Subject $EmailSubject -To $u.EmailAddress #-WhatIf
                        } else {
                            Write-Color @WriteParameters -Text "[i] Sending email to users is disabled. Sending email to default value: ", "$($EmailParameters.EmailTo) ", "..." -Color White, Yellow, White -NoNewLine
                            $EmailSent = Send-Email `
                                -EmailParameters $EmailParameters `
                                -Body $TemporaryBody `
                                -Subject $EmailSubject #-WhatIf
                        }
                        if ($EmailSent.Status -eq $true) {
                            Write-Color -Text "Done" -Color "Green"
                        } else {
                            Write-Color -Text "Failed!" -Color "Red"
                        }
                    }
                    $u | Add-member -NotePropertyName "EmailSent" -NotePropertyValue $EmailSent.Status
                    $u | Add-member -NotePropertyName "EmailSentTo" -NotePropertyValue $EmailSent.SentTo

                    $UsersNotified += $u

                }
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Users', ' section' -Color White, Yellow, White
    } else {
        Write-Color @WriteParameters '[i] Skipping processing ', 'Users', ' section' -Color White, Yellow, White
    }
    #endregion

    #region Send Emails to Managers
    if ($ConfigurationParameters.RemindersSendToManager.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Managers', ' section' -Color White, Yellow, White
        # preparing email
        $EmailSubject = $ConfigurationParameters.RemindersSendToManager.ManagersEmailSubject
        $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters
        $EmailBody += Set-EmailReportBranding -FormattingOptions $FormattingParameters
        $EmailBody += Set-EmailFormatting -Template $FormattingParameters.TemplateForManagers -FormattingParameters $FormattingParameters -ConfigurationParameters $ConfigurationParameters

        # preparing manager lists
        $Managers = @()
        $UsersWithManagers = $UsersNotified | Where-Object { $_.ManagerEmail -ne $null }
        foreach ($u in $UsersWithManagers) {
            $Managers += $u.ManagerEmail
        }
        $Managers = $Managers | Sort-Object | Get-Unique
        Write-Color @WriteParameters '[i] Preparing package for managers with emails ', "$($UsersWithManagers.Count) ", 'users to process with', ' manager filled in', ' where unique managers ', "$($Managers.Count)" -Color  White, Yellow, White, Yellow, White, Yellow
        # processing one manager at time
        foreach ($m in $Managers) {
            # preparing users belonging to manager
            $ColumnNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'
            if ($ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.IncludeNames -ne '') {
                $UsersNotifiedManagers = $UsersNotified | Where-Object { $_.ManagerEmail -eq $m } | Select-Object $ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.IncludeNames
            } else {
                $UsersNotifiedManagers = $UsersNotified | Where-Object { $_.ManagerEmail -eq $m } | Select-Object 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'EmailSent', 'EmailSentTo'
            }
            if ($ConfigurationParameters.RemindersSendToManager.Reports.IncludePasswordNotificationsSent.Enabled -eq $true) {
                foreach ($u in $UsersNotifiedManagers) {
                    Write-Color @WriteParameters -Text '[-] User ', "$($u.DisplayName) ", " Managers Email (", "$($m)", ')'  -Color White, Yellow, White, Yellow, White
                }
            }
            if ($ConfigurationParameters.RemindersSendToManager.RemindersDisplayOnly -eq $true) {
                Write-Color @WriteParameters -Text "[i] Pretending to send email to manager email ", "$($m)", " ...", "Success"  -Color White, Green, White, Green
                $EmailSent = @{}
                $EmailSent.Status = $false
                $EmailSent.SentTo = 'N/A'
            } else {
                $TemporaryBody = $EmailBody
                $TemporaryBody = Set-EmailBodyTableReplacement -Body $TemporaryBody -TableName 'ManagerUsersTable' -TableData $UsersNotifiedManagers
                $TemporaryBody = Set-EmailReplacements -Replacement $TemporaryBody -User $u -FormattingParameters $FormattingParameters -EmailParameters $EmailParameters -Day ''

                if ($ConfigurationParameters.Debug.DisplayTemplateHTML -eq $true) { Get-HTML -text $TemporaryBody }
                if ($ConfigurationParameters.RemindersSendToManager.SendToDefaultEmail -eq $false) {
                    Write-Color @WriteParameters -Text "[i] Sending email to managers email ", "$($m)", " ..."  -Color White, Green -NoNewLine
                    $EmailSent = Send-Email `
                        -EmailParameters $EmailParameters `
                        -Body $TemporaryBody `
                        -Subject $EmailSubject `
                        -To $m #-WhatIf
                } else {
                    Write-Color @WriteParameters -Text "[i] Sending email to managers is disabled. Sending email to default value: ", "$($EmailParameters.EmailTo) ", "..." -Color White, Yellow, White -NoNewLine
                    $EmailSent = Send-Email `
                        -EmailParameters $EmailParameters `
                        -Body $TemporaryBody `
                        -Subject $EmailSubject #-WhatIf
                }
                if ($EmailSent.Status -eq $true) {
                    Write-Color -Text "Done" -Color "Green"
                } else {
                    Write-Color -Text "Failed!" -Color "Red"
                }
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Managers', ' section' -Color White, Yellow, White
    } else {
        Write-Color @WriteParameters '[i] Skipping processing ', 'Managers', ' section' -Color White, Yellow, White
    }
    #endregion Send Emails to Managers

    #region Send Emails to Admins
    if ($ConfigurationParameters.RemindersSendToAdmins.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Administrators', ' section' -Color White, Yellow, White
        $DayHighest = get-HashMaxValue $ConfigurationParameters.RemindersSendToUsers.Reminders
        $DayLowest = get-HashMaxValue $ConfigurationParameters.RemindersSendToUsers.Reminders -Lowest
        $DateCountdownStart = (get-date).AddDays($DayHighest).Date
        $DateIminnent = (get-date).AddDays($DayLowest).Date
        #Write-Color 'Day Highest ', $DayHighest, ' Day Lowest ', $DayLowest, ' Day Countdown Start ', $DateCountdownStart, ' Day Iminnet ', $DateIminnent -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow

        $ColumnNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames -ne '') {
            $UsersNotified = $UsersNotified | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames
        } else {
            $UsersNotified = $UsersNotified | Select-Object $ColumnNames, 'EmailSent', 'EmailSentTo'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames -ne '') {
            $ExpiringIminent = $Users | Where-Object { $_.DateExpiry -lt $DateIminnent -and $_.PasswordExpired -eq $false } | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames
        } else {
            $ExpiringIminent = $Users | Where-Object { $_.DateExpiry -lt $DateIminnent -and $_.PasswordExpired -eq $false } | Select-Object $ColumnNames
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames -ne '') {
            $ExpiringCountdownStarted = $Users | Where-Object { $_.DateExpiry -lt $DateCountdownStart -and $_.PasswordExpired -eq $false } |  Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames
        } else {
            $ExpiringCountdownStarted = $Users | Where-Object { $_.DateExpiry -lt $DateCountdownStart -and $_.PasswordExpired -eq $false } |  Select-Object $ColumnNames
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.IncludeNames -ne '') {
            $UsersExpired = $UsersExpired | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.IncludeNames
        } else {
            $UsersExpired = $UsersExpired | Select-Object $ColumnNames
        }


        $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters
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
            $EmailBody += Set-EmailBody -TableData $UsersExpired `
                -TableMessageWelcome "Following users are already expired (and still enabled...)" `
                -TableMessageNoData "No users that are expired and enabled."
        }
        if ($ConfigurationParameters.Debug.DisplayTemplateHTML -eq $true) { Get-HTML -text $EmailBody }

        if ($ConfigurationParameters.RemindersSendToAdmins.RemindersDisplayOnly -eq $true) {
            Write-Color @WriteParameters -Text "[i] Pretending to send email to admins email ", "$($ConfigurationParameters.RemindersSendToAdmins.AdminsEmail) ", "...", 'Success' -Color White, Yellow, White, Green
        } else {
            Write-Color @WriteParameters -Text "[i] Sending email to administrators on email address ", "$($ConfigurationParameters.RemindersSendToAdmins.AdminsEmail) ", "..." -Color White, Yellow, White -NoNewLine
            $EmailSent = Send-Email -EmailParameters $EmailParameters `
                -Body $EmailBody `
                -Subject $ConfigurationParameters.RemindersSendToAdmins.AdminsEmailSubject `
                -To $ConfigurationParameters.RemindersSendToAdmins.AdminsEmail
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