Import-Module ActiveDirectory

$script:WriteParameters = @{
    ShowTime   = $true
    LogFile    = ""
    TimeFormat = "yyyy-MM-dd HH:mm:ss"
}

function Test-Key ($ConfigurationTable, $ConfigurationSection = "", $ConfigurationKey, $DisplayProgress = $false) {
    if ($ConfigurationTable -eq $null) { return $false }
    try {
        $value = $ConfigurationTable.ContainsKey($ConfigurationKey)
    } catch {
        $value = $false
    }
    if ($value -eq $true) {
        if ($DisplayProgress -eq $true) {
            Write-Color @script:WriteParameters -Text "[i] ", "Parameter in configuration of ", "$ConfigurationSection.$ConfigurationKey", " exists." -Color White, White, Green, White
        }
        return $true
    } else {
        if ($DisplayProgress -eq $true) {
            Write-Color @script:WriteParameters -Text "[i] ", "Parameter in configuration of ", "$ConfigurationSection.$ConfigurationKey", " doesn't exist." -Color White, White, Red, White
        }
        return $false
    }
}

function Test-Prerequisits() {
    try {
        $TestActiveDirectory = get-addomain
    } catch {
        if ($_.Exception -match "Unable to find a default server with Active Directory Web Services running.") {
            Write-Color @script:WriteParameters "[-] ", "Active Directory", " not found. Please run this script with access to ", "Domain Controllers." -Color White, Red, White, Red
        }
        Write-Color @script:WriteParameters "[-] ", "Error: $($_.Exception.Message)" -Color White, Red
        Exit
    }
}

Function Get-ModulesAvailability ([string]$Name) {
    if (-not(Get-Module -name $name)) {
        if (Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) {
            try {
                Import-Module -Name $name
                return $true
            } catch {
                return $false
            }
        } else { return $false } #module not available
    } else { return $true } #module already loaded
}

function Get-HashMaxValue($hashTable, [switch] $Lowest) {
    if ($Lowest) {
        return ($hashTable.GetEnumerator() | Sort-Object value -Descending | Select-Object -Last 1).Value
    } else {
        return ($hashTable.GetEnumerator() | Sort-Object value -Descending | Select-Object -First 1).Value
    }
}

function Send-Email ([hashtable] $EmailParameters, [string] $Body = "", $Attachment = $null, [string] $Subject = "", $To = "") {
    #     $SendMail = Send-Email -EmailParameters $EmailParameters -Body $EmailBody -Attachment $Reports -Subject $TemporarySubject
    #  Preparing the Email properties
    $SmtpClient = New-Object -TypeName system.net.mail.smtpClient
    $SmtpClient.host = $EmailParameters.EmailServer

    # Adding parameters to login to server
    $SmtpClient.Port = $EmailParameters.EmailServerPort
    if ($EmailParameters.EmailServerLogin -ne "") {
        $SmtpClient.Credentials = New-Object System.Net.NetworkCredential($EmailParameters.EmailServerLogin, $EmailParameters.EmailServerPassword)
    }
    $SmtpClient.EnableSsl = $EmailParameters.EmailServerEnableSSL
    $MailMessage = New-Object -TypeName system.net.mail.mailmessage
    $MailMessage.From = $EmailParameters.EmailFrom
    if ($To -ne "") {
        foreach ($T in $To) { $MailMessage.To.add($($T)) }
    } else {
        if ($EmailParameters.Emailto -ne "") {
            foreach ($To in $EmailParameters.Emailto) { $MailMessage.To.add($($To)) }
        }
    }
    if ($EmailParameters.EmailCC -ne "") {
        foreach ($CC in $EmailParameters.EmailCC) { $MailMessage.CC.add($($CC)) }
    }
    if ($EmailParameters.EmailBCC -ne "") {
        foreach ($BCC in $EmailParameters.EmailBCC) { $MailMessage.BCC.add($($BCC)) }
    }
    $Exists = Test-Key $EmailParameters "EmailParameters" "EmailReplyTo" -DisplayProgress $false
    if ($Exists -eq $true) {
        if ($EmailParameters.EmailReplyTo -ne "") {
            $MailMessage.ReplyTo = $EmailParameters.EmailReplyTo
        }
    }
    $MailMessage.IsBodyHtml = 1
    if ($Subject -eq "") {
        $MailMessage.Subject = $EmailParameters.EmailSubject
    } else {
        $MailMessage.Subject = $Subject
    }
    $MailMessage.Body = $Body
    $MailMessage.Priority = [System.Net.Mail.MailPriority]::$($EmailParameters.EmailPriority)

    #  Encoding
    $MailMessage.BodyEncoding = [System.Text.Encoding]::$($EmailParameters.EmailEncoding)
    $MailMessage.SubjectEncoding = [System.Text.Encoding]::$($EmailParameters.EmailEncoding)

    #  Attaching file (s)
    if ($Attachment -ne $null) {
        foreach ($Attach in $Attachment) {
            if (Test-Path $Attach) {
                $File = new-object Net.Mail.Attachment($Attach)
                $MailMessage.Attachments.Add($File)
            }
        }
    }

    #  Sending the Email
    try {
        $SmtpClient.Send($MailMessage)
        #$att.Dispose();
        $MailMessage.Dispose();
        $MailSentTo = "$($MailMessage.To) $($MailMessage.CC) $($MailMessage.BCC)"
        return @{
            Status = $True
            Error  = ""
            SentTo = $MailSentTo
        }
    } catch {
        $MailMessage.Dispose();
        return @{
            Status = $False
            Error  = $($_.Exception.Message)
            SentTo = ""
        }
    }

}
function Set-EmailHead($FormattingOptions) {
    if ($FormattingOptions.CompanyBrandingTemplate -eq '') { $FormattingOptions.CompanyBrandingTemplate = 'TemplateDefault' }

    if ($FormattingOptions.CompanyBrandingTemplate -eq 'TemplateDefault') {
        $head = @"
        <style>
        BODY {
            background-color: white;
            font-family: $($FormattingOptions.FontFamily);
            font-size: $($FormattingOptions.FontSize);
        }

        TABLE {
            border-width: 1px;
            border-style: solid;
            border-color: black;
            border-collapse: collapse;
            font-family: $($FormattingOptions.FontTableDataFamily);
            font-size: $($FormattingOptions.FontTableDataSize);
        }

        TH {
            border-width: 1px;
            padding: 3px;
            border-style: solid;
            border-color: black;
            background-color: #00297A;
            color: white;
            font-family: $($FormattingOptions.FontTableHeadingFamily);
            font-size: $($FormattingOptions.FontTableHeadingSize);
        }
        TR {
            font-family: $($FormattingOptions.FontTableDataFamily);
            font-size: $($FormattingOptions.FontTableDataSize);
        }

        TD {
            border-width: 1px;
            padding-right: 2px;
            padding-left: 2px;
            padding-top: 0px;
            padding-bottom: 0px;
            border-style: solid;
            border-color: black;
            background-color: white;
            font-family: $($FormattingOptions.FontTableDataFamily);
            font-size: $($FormattingOptions.FontTableDataSize);
        }

        H2 {
            font-family: $($FormattingOptions.FontHeadingFamily);
            font-size: $($FormattingOptions.FontHeadingSize);
        }

        P {
            font-family: $($FormattingOptions.FontFamily);
            font-size: $($FormattingOptions.FontSize);
        }
    </style>
"@

    } else {

    }
    return $Head
}
function Set-EmailReportDetails {
    param(
        $FormattingOptions,
        $ReportOptions,
        $TimeToGenerate,
        [int] $CountUsersImminent,
        [int] $CountUsersCountdownStarted,
        [int] $CountUsersAlreadyExpired
    )
    $DateReport = get-date
    # HTML Report settings
    $Report = @"
        <p>
            <strong>Report Time:</strong> $DateReport
            <br>
            <strong>Time to generate:</strong> $($TimeToGenerate.Hours) hours, $($TimeToGenerate.Minutes) minutes, $($TimeToGenerate.Seconds) seconds, $($TimeToGenerate.Milliseconds) milliseconds
            <br>
            <strong>Account Executing Report :</strong> $env:userdomain\$($env:username.toupper()) on $($env:ComputerName.toUpper())
            <br>
            <strong>Users expiring countdown started: </strong> $CountUsersCountdownStarted
            <br>
            <strong>Users expiring soon: </strong> $CountUsersImminent
            <br>
            <strong>Users already expired count: </strong> $CountUsersAlreadyExpired
            <br>
        </p>
"@
    foreach ($ip in $ReportOptions.MonitoredIps.Values) {
        $Report += "<li>ip:</strong> $ip</li>"
    }
    $Report += '</ul>'
    $Report += '</p>'
    return $Report
}
function Set-EmailReportBrading($FormattingOptions) {
    $Report = "<a style=`"text-decoration:none`" href=`"$($FormattingOptions.CompanyBranding.Link)`" class=`"clink logo-container`">" +
    "<img width=<fix> height=<fix> src=`"$($FormattingOptions.CompanyBranding.Logo)`" border=`"0`" class=`"company-logo`" alt=`"company-logo`">" +
    "</a>"
    if ($FormattingOptions.CompanyBranding.Width -ne "") {
        $report = $report -replace "width=<fix>", "width=$($FormattingOptions.CompanyBranding.Width)"
    } else {
        $report = $report -replace "width=<fix>", ""
    }
    if ($FormattingOptions.CompanyBranding.Height -ne "") {
        $report = $report -replace "height=<fix>", "height=$($FormattingOptions.CompanyBranding.Height)"
    } else {
        $report = $report -replace "height=<fix>", ""
    }
    return $Report
}
function Set-EmailReplacements($Replacement, $User, $EmailParameters, $FormattingParameters, $Day) {
    <#
    $body = "<p><i>$TableWelcomeMessage</i>"
    if ($($TableData | Measure-Object).Count -gt 0) {
        $body += $TableData | ConvertTo-Html | Out-String
        $body = $body -replace " Added", "<font color=`"green`"><b> Added</b></font>"
        $body = $body -replace " Removed", "<font color=`"red`"><b> Removed</b></font>"
        $body = $body -replace " Deleted", "<font color=`"red`"><b> Deleted</b></font>"
        $body = $body -replace " Changed", "<font color=`"blue`"><b> Changed</b></font>"
        $body = $body -replace " Change", "<font color=`"blue`"><b> Change</b></font>"
        $body = $body -replace " Disabled", "<font color=`"red`"><b> Disabled</b></font>"
        $body = $body -replace " Enabled", "<font color=`"green`"><b> Enabled</b></font>"
        $body = $body -replace " Locked out", "<font color=`"red`"><b> Locked out</b></font>"
        $body = $body -replace " Lockouts", "<font color=`"red`"><b> Lockouts</b></font>"
        $body = $body -replace " Unlocked", "<font color=`"green`"><b> Unlocked</b></font>"
        $body = $body -replace " Reset", "<font color=`"blue`"><b> Reset</b></font>"
        $body += "</p>"
    }
    else {
        $body += "<br><i>No changes happened during that period.</i></p>"
    }
    #>
    $Replacement = $Replacement -replace "<<DisplayName>>", $user.DisplayName
    $Replacement = $Replacement -replace "<<DateExpiry>>", $user.DateExpiry
    $Replacement = $Replacement -replace "<<GivenName>>", $user.GivenName
    $Replacement = $Replacement -replace "<<Surname>>", $user.Surname
    $Replacement = $Replacement -replace "<<TimeToExpire>>", $Day.Value
    $Replacement = $Replacement -replace "<<ManagerDisplayName>>", $user.Manager
    $Replacement = $Replacement -replace "<<ManagerEmail>>", $user.ManagerEmail
    return $Replacement
}
function Set-EmailFormatting ($Template, $FormattingParameters, $ConfigurationParameters) {
    $WriteParameters = $ConfigurationParameters.DisplayConsole

    $Template = $Template.Split("`n") # https://blogs.msdn.microsoft.com/timid/2014/07/09/one-liner-fun-with-multi-line-blocktext-and-split-split/

    $Body = ""
    Write-Color @WriteParameters -Text "[i] Preparing template ", "adding", " HTML ", "<BR>", " tags..." -Color White, Yellow, White, Yellow -NoNewLine
    foreach ($t in $Template) {
        $Body += "<br>$t"
    }
    Write-Color -Text "Done" -Color "Green"
    foreach ($style in $FormattingParameters.Styles.GetEnumerator()) {
        foreach ($value in $style.Value) {
            if ($value -eq "") { continue }
            Write-Color @WriteParameters -Text "[i] Preparing template ", "adding", " HTML ", "$($style.Name)", " tag for ", "$value", ' tags...' -Color White, Yellow, White, Yellow, White, Yellow -NoNewLine
            $Body = $Body.Replace($value, "<$($style.Name)>$value</$($style.Name)>")
            Write-Color -Text "Done" -Color "Green"
        }
    }

    foreach ($color in $FormattingParameters.Colors.GetEnumerator()) {
        foreach ($value in $color.Value) {
            if ($value -eq "") { continue }
            Write-Color @WriteParameters -Text "[i] Preparing template ", "adding", " HTML ", "$($color.Name)", " tag for ", "$value", ' tags...' -Color White, Yellow, White, $($color.Name), White, Yellow -NoNewLine
            $Body = $Body.Replace($value, "<span style=color:$($color.Name)>$value</span>")
            Write-Color -Text "Done" -Color "Green"
        }
    }
    foreach ($links in $FormattingParameters.Links.GetEnumerator()) {
        foreach ($link in $links.Value) {
            #write-host $link.Text
            #write-host $link.Link
            if ($link.Link -like "*@*") {
                Write-Color @WriteParameters -Text "[i] Preparing template ", "adding", " EMAIL ", "Links for", " $($links.Key)..." -Color White, Yellow, White, White, Yellow, White -NoNewLine
                $Body = $Body -replace "<<$($links.Key)>>", "<span style=color:$($link.Color)><a href='mailto:$($link.Link)?subject=$($Link.Subject)'>$($Link.Text)</a></span>"
            } else {
                Write-Color @WriteParameters -Text "[i] Preparing template ", "adding", " HTML ", "Links for", " $($links.Key)..." -Color White, Yellow, White, White, Yellow, White -NoNewLine
                $Body = $Body -replace "<<$($links.Key)>>", "<span style=color:$($link.Color)><a href='$($link.Link)'>$($Link.Text)</a></span>"
            }
            Write-Color -Text "Done" -Color "Green"
        }

    }

    if ($ConfigurationParameters.DisplayTemplateHTML -eq $true) { Get-HTML($Body) }
    return $Body
}
function Set-EmailBody($TableData, $TableWelcomeMessage) {
    $body = "<p><i>$TableWelcomeMessage</i>"
    if ($($TableData | Measure-Object).Count -gt 0) {
        $body += $TableData | ConvertTo-Html -Fragment | Out-String
        #$body = ($body -split '_')[0] -creplace '(?<=\w)([A-Z])', ' $1' # adds spaces before each Capital Letter
        $body = $body -replace " Added", "<font color=`"green`"><b> Added</b></font>"
        $body = $body -replace " Removed", "<font color=`"red`"><b> Removed</b></font>"
        $body = $body -replace " Deleted", "<font color=`"red`"><b> Deleted</b></font>"
        $body = $body -replace " Changed", "<font color=`"blue`"><b> Changed</b></font>"
        $body = $body -replace " Change", "<font color=`"blue`"><b> Change</b></font>"
        $body = $body -replace " Disabled", "<font color=`"red`"><b> Disabled</b></font>"
        $body = $body -replace " Enabled", "<font color=`"green`"><b> Enabled</b></font>"
        $body = $body -replace " Locked out", "<font color=`"red`"><b> Locked out</b></font>"
        $body = $body -replace " Lockouts", "<font color=`"red`"><b> Lockouts</b></font>"
        $body = $body -replace " Unlocked", "<font color=`"green`"><b> Unlocked</b></font>"
        $body = $body -replace " Reset", "<font color=`"blue`"><b> Reset</b></font>"
        $body += "</p>"
    } else {
        #$body += "<br><i>No changes happend during that period.</i></p>"
    }
    return $body
}
function Set-EmailBodyTableReplacement($Body, $TableName, $TableData) {
    $TableData = $TableData | ConvertTo-Html -Fragment | Out-String
    $Body = $Body -replace "<<$TableName>>", $TableData
    return $Body
}

function Get-HTML($text) {
    $text = $text.Split("`r")
    foreach ($t in $text) {
        Write-Host $t
    }
}

function Find-AllUsers () {
    $users = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0 -and PasswordNotRequired -ne $True } -Properties Manager, DisplayName, GivenName, Surname, SamAccountName, EmailAddress, msDS-UserPasswordExpiryTimeComputed, PasswordExpired, PasswordLastSet, PasswordNotRequired
    $users = $users | Select-Object UserPrincipalName, SamAccountName, DisplayName, GivenName, Surname, EmailAddress, PasswordExpired, PasswordLastSet, PasswordNotRequired,
    @{Name = "Manager"; Expression = { (Get-ADUser $_.Manager).Name }},
    @{Name = "ManagerEmail"; Expression = { (Get-ADUser -Properties Mail $_.Manager).Mail  }},
    @{Name = "DateExpiry"; Expression = { ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")) }},
    @{Name = "DaysToExpire"; Expression = { (NEW-TIMESPAN -Start (GET-DATE) -End ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))).Days }}
    return $users
}

Function Start-PasswordExpiryCheck ([hashtable] $EmailParameters, [hashtable] $FormattingParameters, [hashtable] $ConfigurationParameters) {
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
    $EmailBody += Set-EmailReportBrading -FormattingOptions $FormattingParameters
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
                            $EmailSent = Send-Email -EmailParameters $EmailParameters -Body $TemporaryBody -ReportOptions $DisplayParameters -Subject $EmailSubject -To $u.EmailAddress
                        } else {
                            Write-Color @WriteParameters -Text "[i] Sending email to users is disabled. Sending email to default value: ", "$($EmailParameters.EmailTo) ", "..." -Color White, Yellow, White -NoNewLine
                            $EmailSent = Send-Email -EmailParameters $EmailParameters -Body $TemporaryBody -ReportOptions $DisplayParameters -Subject $EmailSubject
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
        $EmailBody += Set-EmailReportBrading -FormattingOptions $FormattingParameters
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
                    $EmailSent = Send-Email -EmailParameters $EmailParameters -Body $TemporaryBody -ReportOptions $DisplayParameters -Subject $EmailSubject -To $m
                } else {
                    Write-Color @WriteParameters -Text "[i] Sending email to managers is disabled. Sending email to default value: ", "$($EmailParameters.EmailTo) ", "..." -Color White, Yellow, White -NoNewLine
                    $EmailSent = Send-Email -EmailParameters $EmailParameters -Body $TemporaryBody -ReportOptions $DisplayParameters -Subject $EmailSubject
                    if ($EmailSent.Status -eq $true) {
                        Write-Color -Text "Done" -Color "Green"
                    } else {
                        Write-Color -Text "Failed!" -Color "Red"
                    }

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
        $EmailBody += Set-EmailReportBrading -FormattingOptions $FormattingParameters
        $EmailBody += Set-EmailReportDetails -FormattingOptions $FormattingParameters `
            -ReportOptions $ReportOptions `
            -TimeToGenerate $Time.Elapsed `
            -CountUsersCountdownStarted $($ExpiringCountdownStarted.Count) `
            -CountUsersImminent $($ExpiringIminent.Count) `
            -CountUsersAlreadyExpired $($UsersExpired.Count)
        $time.Stop()

        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludePasswordNotificationsSent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Password Notifcations Sent' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $UsersNotified -TableWelcomeMessage "Following users had their password notifications sent"
        }
        if ( $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringImminent.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Users expiring imminent' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $ExpiringIminent -TableWelcomeMessage "Following users expiring imminent (Less than $DayLowest day(s)"
        }
        if (  $ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpiringCountdownStarted.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Expiring Couintdown Started' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $ExpiringCountdownStarted -TableWelcomeMessage "Following users expiring countdown started (Less than $DayHighest day(s))"
        }
        if ($ConfigurationParameters.RemindersSendToAdmins.Reports.IncludeExpired.Enabled -eq $true) {
            Write-Color @WriteParameters -Text '[i] Preparing data for report ', 'Users are already expired' -Color White, Yellow
            $EmailBody += Set-EmailBody -TableData $UsersExpired -TableWelcomeMessage "Following users are already expired (and still enabled...)"
        }
        if ($ConfigurationParameters.Debug.DisplayTemplateHTML -eq $true) { Get-HTML -text $EmailBody }

        if ($ConfigurationParameters.RemindersSendToAdmins.RemindersDisplayOnly -eq $true) {
            Write-Color @WriteParameters -Text "[i] Pretending to send email to admins email ", "$($ConfigurationParameters.RemindersSendToAdmins.AdminsEmail) ", "...", 'Success' -Color White, Yellow, White, Green
        } else {
            Write-Color @WriteParameters -Text "[i] Sending email to administrators on email address ", "$($ConfigurationParameters.RemindersSendToAdmins.AdminsEmail) ", "..." -Color White, Yellow, White -NoNewLine
            $EmailSent = Send-Email -EmailParameters $EmailParameters -Body $EmailBody -ReportOptions $DisplayParameters -Subject $ConfigurationParameters.RemindersSendToAdmins.AdminsEmailSubject -To $ConfigurationParameters.RemindersSendToAdmins.AdminsEmail
            if ($EmailSent.Status -eq $true) {
                Write-Color -Text "Done" -Color "Green"
            } else {
                Write-Color -Text "Failed!" -Color "Red"
            }
        }
        Write-Color @WriteParameters '[i] Ending processing ', 'Administrators', ' section' -Color White, Yellow, White

    } else {
        Write-Color @WriteParameters '[i] Skipping processing ', 'Administrators', ' section' -Color White, Yellow, White

    }
    #endregion Send Emails to Admins
}

Export-ModuleMember -function 'Start-PasswordExpiryCheck'