Import-Module .\PSPasswordExpiryNotifications.psd1 -Force

$EmailParameters = @{
    EmailFrom                  = "monitoring@domain.pl"
    EmailTo                    = "przemyslaw.klys@domain.pl" # your default email field (IMPORTANT)
    EmailReplyTo               = "helpdesk@domain.pl" # email to use when users press Reply
    EmailServer                = ""
    EmailServerPassword        = ""
    EmailServerPort            = "587"
    EmailServerLogin           = ""
    EmailServerEnableSSL       = 1
    EmailSubject               = "[Password Expiring] Your password will expire on <<DateExpiry>> (<<TimeToExpire>> days)"
    EmailPriority              = "Low" # Normal, High
    EmailUseDefaultCredentials = $false
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
    <<Image>>

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
        #SendCountMaximum     = 3
        Rules                = @(
            # rules are new way to define things. You can define more than one rule and limit it per group/ou
            # the primary rule above can be set or doesn't have to, all parameters from rules below can be use across different rules
            # it's up to you tho that the notifications are unique - if you put 2 rules that do the same thing, 2 notifications will be sent
            [ordered] @{
                Enable               = $false # doesn't processes this section at all if $false
                RemindersDisplayOnly = $false # prevents sending any emails (good for testing) - including managers
                SendToDefaultEmail   = $true # if enabled $EmailParameters are used (good for testing)
                Reminders            = 1, 7, 14, 15
                UseAdditionalField   = 'extensionAttribute13'
                SendCountMaximum     = 3
            }
            [ordered]@{
                Enable               = $false # doesn't processes this section at all if $false
                RemindersDisplayOnly = $false # prevents sending any emails (good for testing) - including managers
                SendToDefaultEmail   = $true # if enabled $EmailParameters are used (good for testing)
                Reminders            = 3, 9
                UseAdditionalField   = 'extensionAttribute12'
                SendCountMaximum     = 3
            }
            [ordered] @{
                Enable                   = $true # doesn't processes this section at all if $false
                RemindersDisplayOnly     = $false # prevents sending any emails (good for testing) - including managers
                SendToDefaultEmail       = $true # if enabled $EmailParameters are used (good for testing)
                Reminders                = 0, 1, 2, 3, 4, 5, 12, 13, 14, 15, 28, 30 #50
                UseAdditionalField       = 'extensionAttribute13'
                #SendCountMaximum         = 3
                # this means we want to process only users that NeverExpire
                PasswordNeverExpires     = $true
                PasswordNeverExpiresDays = 30
                # limit group or limit OU can limit people with password never expire to certain users only
                LimitGroup               = @(
                    #'CN=GDS-PasswordExpiryNotifications,OU=Security,OU=Groups,OU=Production,DC=ad,DC=evotec,DC=xyz',
                    #'CN=GDS-TestGroup9,OU=Security,OU=Groups,OU=Production,DC=ad,DC=evotec,DC=xyz'
                )
                LimitOU                  = @(
                    #'OU=UsersNoSync,OU=Accounts,OU=Production,DC=ad,DC=evotec,DC=xyz'
                    #'*OU=Accounts,OU=Production,DC=ad,DC=evotec,DC=xyz'
                    'OU=UsersNoSync,OU=Accounts,OU=Production,DC=ad,DC=evotec,DC=xyz'
                )
            }
        )
    }
    RemindersSendToManager = @{
        Enable               = $true # doesn't processes this section at all if $false
        RemindersDisplayOnly = $true # prevents sending any emails (good for testing)
        SendToDefaultEmail   = $true # if enabled $EmailParameters are used (good for testing)
        ManagersEmailSubject = "Summary of password reminders (for users you manage)"
        Reports              = @{
            IncludePasswordNotificationsSent = @{
                Enabled          = $true
                IncludeNames     = 'UserPrincipalName', 'Domain', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'EmailSent', 'EmailSentTo'
                TextBeforeReport = '"Following users which you are listed as manager for have their passwords expiring soon:"'

            }
        }

        # You can use limit scope
        #LimitScope           = @{
        #    Groups = 'RecursiveGoup-FGP-Check'
        #}
        # SendCountMaximum     = 3
    }
    RemindersSendToAdmins  = @{
        Enable               = $true # doesn't processes this section at all
        RemindersDisplayOnly = $true # prevents sending any emails (good for testing)
        AdminsEmail          = 'notifications@domain.pl', 'przemyslaw.klys@domain.pl'
        AdminsEmailSubject   = "[Reporting Evotec] Summary of password reminders"
        ReportsAsExcel       = $true
        #ReportsAsHTML        = $true
        Reports              = @{
            IncludeSummary                           = @{
                Enabled = $true
            }
            IncludePasswordNotificationsSent         = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'PasswordNeverExpires', 'EmailSent', 'EmailSentTo'
            }
            IncludeExpiringImminent                  = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'PasswordNeverExpires'
            }
            IncludeExpiringCountdownStarted          = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'PasswordNeverExpires'
            }
            IncludeExpired                           = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'PasswordNeverExpires'
            }
            IncludeManagersPasswordNotificationsSent = @{
                Enabled      = $true
                IncludeNames = 'UserPrincipalName', 'Domain', 'DisplayName', 'DateExpiry', 'DaysToExpire', 'PasswordExpired', 'SamAccountName', 'Manager', 'ManagerEmail', 'PasswordLastSet', 'PasswordNeverExpires', 'EmailSent', 'EmailSentTo'
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