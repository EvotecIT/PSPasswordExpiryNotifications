function Invoke-ReminderToUsersInternal {
    [cmdletBinding()]
    param(
        [System.Collections.IDictionary] $Rule,
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $FormattingParameters,
        [System.Collections.IDictionary] $ConfigurationParameters,
        $EmailBody,
        [Array] $Users
    )
    $Limits = @{
        TestingLimitReached = $false
    }
    if ($Rule.Enable -eq $true) {
        #$Today = Get-Date
        $EmailBody = Set-EmailHead -FormattingOptions $FormattingParameters
        $Image = Set-EmailReportBranding -FormattingOptions $FormattingParameters
        $EmailBody += Set-EmailFormatting -Template $FormattingParameters.Template -FormattingParameters $FormattingParameters -ConfigurationParameters $ConfigurationParameters -Image $Image

        Write-Color @WriteParameters '[i] Starting processing ', 'Users', ' section' -Color White, Yellow, White

        if ($Rule.Reminders -is [System.Collections.IDictionary]) {
            [Array] $DaysToExpire = ($Rule.Reminders).Values | Sort-Object -Unique
        } else {
            [Array] $DaysToExpire = $Rule.Reminders | Sort-Object -Unique
        }
        $Count = 0
        foreach ($u in $Users) {
            if ($Limits.TestingLimitReached -eq $true) {
                break
            }

            if ($null -eq $Rule.PasswordNeverExpires -or $null -eq $Rule.PasswordNeverExpiresDays -or $Rule.PasswordNeverExpires -eq $false) {
                # This is standard situation that user expires normally
                if ($u.PasswordNeverExpires -eq $true -or $u.PasswordAtNextLogon -eq $true) {
                    continue
                }
            } elseif ($Rule.PasswordNeverExpires -eq $true -and $Rule.PasswordNeverExpiresDays -is [int]) {
                # this is for situation where we want to monitor PasswordNeverExpires
                if ($u.PasswordAtNextLogon -eq $true) {
                    continue
                }
                if ($u.PasswordNeverExpires -eq $false) {
                    continue
                }
            } else {
                Write-Color @WriteParameters '[i] Something went wrong as there are no rules matching...' -Color Red
                continue
            }

            if ($u.PasswordNeverExpires -eq $true) {
                # If we're here it means we want to get only users that never expire
                $PretendedDateExpiration = ($u.PasswordLastSet).AddDays($Rule.PasswordNeverExpiresDays)
                $PretendedDaysToExpire = (New-TimeSpan -Start (Get-Date) -End ($PretendedDateExpiration)).Days
                # we overwrite dates
                $u.DateExpiry = $PretendedDateExpiration
                $u.DaysToExpire = $PretendedDaysToExpire
            }

            # this is standard way - check if user is expiring within the correct date
            if ($u.DaysToExpire -notin $DaysToExpire) {
                continue
            }

            # This makes sure to apply notifications
            if ($Rule.LimitGroup) {
                $Found = $false
                foreach ($LimitGroup in $Rule.LimitGroup) {
                    if ($LimitGroup -in $u.MemberOf) {
                        $Found = $true
                        break
                    }
                }
                if (-not $Found) {
                    continue
                }
            }
            if ($Rule.LimitOU) {
                $Found = $false
                foreach ($LimitOU in $Rule.LimitOU) {
                    if ($LimitOU -like $u.OrganizationalUnit) {
                        $Found = $true
                        break
                    }
                }
                if (-not $Found) {
                    continue
                }
            }

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
                $Limits.TestingLimitReached = $true
                break
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Users', ' section' -Color White, Yellow, White
    } else {
        Write-Color @WriteParameters '[i] Skipping processing ', 'Users', ' section' -Color White, Yellow, White
    }
}