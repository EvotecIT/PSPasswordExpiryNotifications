Import-Module .\PSPasswordExpiryNotifications.psd1 -Force

$EmailParameters = @{
    EmailFrom            = "monitoring@domain.pl"
    EmailTo              = "przemyslaw.klys@domain.pl" # your default email field (IMPORTANT)
    EmailCC              = ""
    EmailBCC             = ""
    EmailReplyTo         = "helpdesk@domain.pl" # email to use when users press Reply
    EmailServer          = ""
    EmailServerPassword  = ""
    EmailServerPort      = "587"
    EmailServerLogin     = ""
    EmailServerEnableSSL = 1
    EmailEncoding        = "Unicode"
    EmailSubject         = "[Password Expiring] Your password will expire on <<DateExpiry>> (<<TimeToExpire>> days)"
    EmailPriority        = "Low" # Normal, High
}

$FormattingParameters = @{
    CompanyBrandingTemplate = 'TemplateDefault'
    CompanyBranding         = @{
        Logo   = "https://evotec.xyz/wp-content/uploads/2015/05/Logo-evotec-012.png"
        Width  = "200"
        Height = ""
        Link   = "https://evotec.xyz"
        Inline = $false
    }

    FontFamily              = "Calibri Light"
    FontSize                = "9pt"

    FontHeadingFamily       = "Calibri Light"
    FontHeadingSize         = "12pt"

    FontTableHeadingFamily  = "Calibri Light"
    FontTableHeadingSize    = "9pt"

    FontTableDataFamily     = "Calibri Light"
    FontTableDataSize       = "9pt"

    Colors                  = @{
        Red   = "reset it"
        Blue  = "please contact", "CTRL+ALT+DEL"
        Green = "+48 22 600 20 20"
    }
    Styles                  = @{
        B = "To change your password", "<<DisplayName>>", "Change a password" # BOLD
        I = "password" # Italian
        U = "Help Desk" # Underline
    }
    Links                   = @{
        ClickHere        = @{
            Link  = "https://password.evotec.pl"
            Text  = "Click Here"
            Color = "Blue"

        }
        ClickingHere     = @{ Link = "https://passwordreset.microsoftonline.com/"
            Text               = "clicking here"
            Color              = "Red"
        }
        VisitingPortal   = @{ Link = "https://evotec.xyz"
            Text                 = "visiting Service Desk Portal"
            Color                = "Red"
        }
        ServiceDeskEmail = @{
            Link    = "helpdesk@domain.pl" # if contains @ treated as email
            Text    = "Service Desk"
            Color   = "Red"
            Subject = "I need help with my password" # Email subject used for email links only / ignored for http/https links
        }
    }


    Template                = "
    Hello <<DisplayName>>,
    Your password is due to expire in <<TimeToExpire>> days.

    To change your password:
    - press CTRL+ALT+DEL -> Change a password...

    If you have forgotten you password and need to reset it, you can do this by <<ClickingHere>>
    In case of problems please contact HelpDesk by <<VisitingPortal>> or by sending an email to <<ServiceDeskEmail>>.

    Alternatively you can always call Service Desk at +48 22 600 20 20

    Kind regards,
    Evotec IT"

    TemplateForManagers     = "
    Hello <<ManagerDisplayName>>,

    Below you can find a list of users who are about to expire in next few days.

    <<ManagerUsersTable>>

    This is just an informational message.. There is no need to do anything about it unless you see some disprepency.

    Kind regards,
    Evotec IT"

}
$ConfigurationParameters = @{
    RemindersSendToUsers   = @{
        Enable               = $true # doesn't processes this section at all if $false
        RemindersDisplayOnly = $true # prevents sending any emails (good for testing) - including managers
        SendToDefaultEmail   = $true # if enabled $EmailParameters are used (good for testing)
        Reminders            = @{
            Notification1 = 1
            Notification2 = 21
            Notification3 = 34
        }
        #UseAdditionalField   = 'extensionAttribute13'
    }
    RemindersSendToManager = @{
        Enable               = $true # doesn't processes this section at all if $false
        RemindersDisplayOnly = $true # prevents sending any emails (good for testing)
        SendToDefaultEmail   = $true # if enabled $EmailParameters are used (good for testing)
        ManagersEmailSubject = "Summary of password reminders (for users you manage)"
        Reports              = @{
            IncludePasswordNotificationsSent = @{
                Enabled          = $true
                IncludeNames     = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'EmailSent', 'EmailSentTo'
                TextBeforeReport = '"Following users which you are listed as manager for have their passwords expiring soon:"'

            }
        }
    }
    RemindersSendToAdmins  = @{
        Enable               = $true # doesn't processes this section at all
        RemindersDisplayOnly = $true # prevents sending any emails (good for testing)
        AdminsEmail          = 'notifications@domain.pl', 'przemyslaw.klys@domain.pl'
        AdminsEmailSubject   = "[Reporting Evotec] Summary of password reminders"
        Reports              = @{
            IncludePasswordNotificationsSent = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'EmailSent', 'EmailSentTo'
            }
            IncludeExpiringImminent          = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'
            }
            IncludeExpiringCountdownStarted  = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'
            }
            IncludeExpired                   = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet'
            }
        }
    }

    DisplayConsole         = @{
        ShowTime   = $true
        LogFile    = ""
        TimeFormat = "yyyy-MM-dd HH:mm:ss"
    }
    Debug                  = @{
        DisplayTemplateHTML = $false
    }

}

Start-PasswordExpiryCheck $EmailParameters $FormattingParameters $ConfigurationParameters