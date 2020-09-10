Function Start-PasswordExpiryCheck {
    [CmdletBinding()]
    param (
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $FormattingParameters,
        [System.Collections.IDictionary] $ConfigurationParameters
    )
    $time = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start

    $WriteParameters = $ConfigurationParameters.DisplayConsole

    if ($WriteParameters.LogFile) {
        $Folder = $WriteParameters.LogFile | Split-Path
        if (-not (Test-Path -Path $Folder)) {
            $null = New-Item -ItemType Directory -Path $Folder -Force
            if (-not (Test-Path -Path $Folder)) {
                Write-Color "[e] Can't created $Folder for logging. Terminating..." -Color Red
                return
            }
        }
    }

    Test-Prerequisits

    # Overwritting whatever user set as this is what it should be, always for proper display
    if ($EmailParameters.EmailEncoding -or $EmailParameters.EmailSubjectEncoding -or $EmailParameters.EmailBodyEncoding) {
        Write-Color @WriteParameters '[e] Setting encoding was depracated. Its set automatically now to utf8' -Color Red
    }
    $EmailParameters.EmailEncoding = ""
    $EmailParameters.EmailSubjectEncoding = ""
    $EmailParameters.EmailBodyEncoding = ""

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
    [Array] $Users = Find-PasswordExpiryCheck -AdditionalProperties $FieldName -ConditionProperties $ConditionProperties -WriteParameters $WriteParameters -CachedUsers $CachedUsers -CachedUsersPrepared $CachedUsersPrepared -CachedManagers $CachedManagers | Sort-Object DateExpiry

    # This will make sure to catch only applicable users. Since there are multiple rules possible we can't use $Users as our source of truth
    $Script:UsersApplicable = [System.Collections.Generic.List[PSCustomObject]]::new()

    #region Send Emails to Users
    [Array] $UsersNotified = Invoke-ReminderToUsers -RemindersToUsers $ConfigurationParameters.RemindersSendToUsers -EmailParameters $EmailParameters -ConfigurationParameters $ConfigurationParameters -FormattingParameters $FormattingParameters -Users $Users

    # Build a report for expired users
    [Array] $UsersExpired = $Script:UsersApplicable | Where-Object { $null -ne $_.DateExpiry -and $_.DateExpiry -lt $Today }

    #region Send Emails to Managers
    [Array] $ManagersReceived = if ($ConfigurationParameters.RemindersSendToManager.Enable -eq $true) {
        Write-Color @WriteParameters '[i] Starting processing ', 'Managers', ' section' -Color White, Yellow, White
        # preparing email
        $EmailSubject = $ConfigurationParameters.RemindersSendToManager.ManagersEmailSubject
        $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters
        $EmailReportBranding = Set-EmailReportBranding -FormattingOptions $FormattingParameters
        $EmailBody += Set-EmailFormatting -Template $FormattingParameters.TemplateForManagers `
            -FormattingParameters $FormattingParameters `
            -ConfigurationParameters $ConfigurationParameters `
            -Image $EmailReportBranding

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
                    Write-Color -Text "Done" -Color "Green" -LogFile $WriteParameters.LogFile
                } else {
                    Write-Color -Text "Failed!" -Color "Red" -LogFile $WriteParameters.LogFile
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
        $Today = Get-Date

        $ColumnNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'PasswordNeverExpires'

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames -gt 0) {
            $UsersNotified = $UsersNotified | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.IncludeNames
        } else {
            $UsersNotified = $UsersNotified | Select-Object $ColumnNames, 'EmailSent', 'EmailSentTo'
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames.Count -gt 0) {
            $ExpiringIminent = $Script:UsersApplicable | Where-Object { $null -ne $_.DateExpiry -and ($_.DateExpiry -lt $DateIminnent -and $_.DateExpiry -gt $Today) -and $_.PasswordExpired -eq $false } | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.IncludeNames
        } else {
            $ExpiringIminent = $Script:UsersApplicable | Where-Object { $null -ne $_.DateExpiry -and ($_.DateExpiry -lt $DateIminnent -and $_.DateExpiry -gt $Today) -and $_.PasswordExpired -eq $false } | Select-Object $ColumnNames
        }

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames -gt 0) {
            $ExpiringCountdownStarted = $Script:UsersApplicable | Where-Object { $null -ne $_.DateExpiry -and ($_.DateExpiry -lt $DateCountdownStart -and $_.DateExpiry -gt $DateIminnent) -and $_.PasswordExpired -eq $false } | Select-Object $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.IncludeNames
        } else {
            $ExpiringCountdownStarted = $Script:UsersApplicable | Where-Object { $null -ne $_.DateExpiry -and ($_.DateExpiry -lt $DateCountdownStart -and $_.DateExpiry -gt $DateIminnent) -and $_.PasswordExpired -eq $false } | Select-Object $ColumnNames
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

        $FilePathExcel = Get-FileName -Extension 'xlsx' -Temporary

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeSummary.Enabled -eq $true) {
            $SummaryOfUsers = $Script:UsersApplicable | Group-Object DaysToExpire `
            | Select-Object @{Name = 'Days to Expire'; Expression = { [int] $($_.Name) } }, @{Name = 'Users with Days to Expire'; Expression = { [int] $($_.Count) } }
            $SummaryOfUsers = $SummaryOfUsers | Sort-Object -Property 'Days to Expire'

            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Summary of Expiring Users' -Color White, Yellow
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsHTML -ne $false) {
                $EmailBody += Set-EmailBody -TableData $SummaryOfUsers `
                    -TableMessageWelcome "Summary of days to expire and it's count" `
                    -TableMessageNoData 'There were no users that have days of expiring.'
            }
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsExcel) {
                $SummaryOfUsers | ConvertTo-Excel -FilePath $FilePathExcel -ExcelWorkSheetName 'Summary' -AutoFilter -AutoFit
            }
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Password Notifcations Sent' -Color White, Yellow
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsHTML -ne $false) {
                $EmailBody += Set-EmailBody -TableData $UsersNotified `
                    -TableMessageWelcome "Following users had their password notifications sent" `
                    -TableMessageNoData 'No users required nofifications.'
            }
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsExcel) {
                $UsersNotified | ConvertTo-Excel -FilePath $FilePathExcel -ExcelWorkSheetName 'NotificationsSent' -AutoFilter -AutoFit
            }
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeManagersPasswordNotificationsSent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Password Notifcations Sent to Managers' -Color White, Yellow
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsHTML -ne $false) {
                $EmailBody += Set-EmailBody -TableData $ManagersReceived `
                    -TableMessageWelcome "Following managers had their password bundle notifications sent" `
                    -TableMessageNoData 'No managers required nofifications.'
            }
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsExcel) {
                $ManagersReceived | ConvertTo-Excel -FilePath $FilePathExcel -ExcelWorkSheetName 'NotificationsSentManagers' -AutoFilter -AutoFit
            }
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Users expiring imminent' -Color White, Yellow
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsHTML -ne $false) {
                $EmailBody += Set-EmailBody -TableData $ExpiringIminent `
                    -TableMessageWelcome "Following users expiring imminent (Less than $DayLowest day(s)" `
                    -TableMessageNoData 'No users expiring.'
            }
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsExcel) {
                $ExpiringIminent | ConvertTo-Excel -FilePath $FilePathExcel -ExcelWorkSheetName 'ExpiringImminent' -AutoFilter -AutoFit
            }
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Expiring Couintdown Started' -Color White, Yellow
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsHTML -ne $false) {
                $EmailBody += Set-EmailBody -TableData $ExpiringCountdownStarted `
                    -TableMessageWelcome "Following users expiring countdown started (Less than $DayHighest day(s))" `
                    -TableMessageNoData 'There were no users that had their coundown started.'
            }
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsExcel) {
                $ExpiringCountdownStarted | ConvertTo-Excel -FilePath $FilePathExcel -ExcelWorkSheetName 'ExpiringCountdownStarted' -AutoFilter -AutoFit
            }
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Users are already expired' -Color White, Yellow
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsHTML -ne $false) {
                if ($ConfigurationParameters.DisableExpiredUsers.Enable -eq $true -and -not $ConfigurationParameters.DisableExpiredUsers.DisplayOnly -eq $true) {
                    $EmailBody += Set-EmailBody -TableData $UsersExpired -TableMessageWelcome "Following users are already expired (and were disabled...)" -TableMessageNoData "No users that are expired."
                } else {
                    $EmailBody += Set-EmailBody -TableData $UsersExpired -TableMessageWelcome "Following users are already expired (and still enabled...)" -TableMessageNoData "No users that are expired and enabled."
                }
            }
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsExcel) {
                $UsersExpired | ConvertTo-Excel -FilePath $FilePathExcel -ExcelWorkSheetName 'UsersExpired' -AutoFilter -AutoFit
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
            if ($ConfigurationParameters.RemindersSendToAdmins.ReportsAsExcel) {
                $EmailSplat.Attachment = $FilePathExcel
            }
            $EmailSent = Send-Email @EmailSplat
            if ($EmailSent.Status -eq $true) {
                Write-Color -Text "Done" -Color "Green" -LogFile $WriteParameters.LogFile
            } else {
                Write-Color -Text "Failed! Error: $($EmailSent.Error)" -Color "Red" -LogFile $WriteParameters.LogFile
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Administrators', ' section' -Color White, Yellow, White

    } else {
        Write-Color @WriteParameters '[i] Skipping processing ', 'Administrators', ' section' -Color White, Yellow, White

    }
    #endregion Send Emails to Admins
}
