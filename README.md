<p align="center">
  <a href="https://www.powershellgallery.com/packages/PSPasswordExpiryNotifications"><img src="https://img.shields.io/powershellgallery/v/PSPasswordExpiryNotifications.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PSPasswordExpiryNotifications"><img src="https://img.shields.io/powershellgallery/vpre/PSPasswordExpiryNotifications.svg?label=powershell%20gallery%20preview&colorB=yellow"></a>
  <a href="https://github.com/EvotecIT/PSPasswordExpiryNotifications"><img src="https://img.shields.io/github/license/EvotecIT/PSPasswordExpiryNotifications.svg"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/PSPasswordExpiryNotifications"><img src="https://img.shields.io/powershellgallery/p/PSPasswordExpiryNotifications.svg"></a>
  <a href="https://github.com/EvotecIT/PSPasswordExpiryNotifications"><img src="https://img.shields.io/github/languages/top/evotecit/PSPasswordExpiryNotifications.svg"></a>
  <a href="https://github.com/EvotecIT/PSPasswordExpiryNotifications"><img src="https://img.shields.io/github/languages/code-size/evotecit/PSPasswordExpiryNotifications.svg"></a>
  <a href="https://github.com/EvotecIT/PSPasswordExpiryNotifications"><img src="https://img.shields.io/powershellgallery/dt/PSPasswordExpiryNotifications.svg"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>

# PSPasswordExpiryNotifications - PowerShell module

Following PowerShell Module provides different approach to scheduling password notifications for expiring Active Directory based accounts. While most of the scripts require knowledge on HTML... this one is just one config file and a bit of tingling around with texts. Whether this is good or bad it's up to you to decide. I do plan to add an option to use external HTML template if there will be requests for that.

## Links

- [Short description for this project at (screenshots and all)](https://evotec.xyz/just-different-approach-to-active-directory-password-notifications/)
- [Full Description (and a know-how) for this project](https://evotec.xyz/hub/scripts/pspasswordexpirynotifications-powershell-module/)

### Updates

- 1.6.7 - 2020.09.09
  - Fixed logging to file for status of sent emails
  - Added auto creation of logs directory if it's missing
- 1.6.6 - 2020.09.06
  - Added ability of template per rule
- 1.6.5 - 2020.09.06
  - Resolved issues with encoding, removed encoding setting due conflicts
    - [x] Set by default to UTF-8 which should resolve weird chars
  - Added filtering by group
    - [x] `LimitGroup` takes an array of DistinguishedNames - compares on eq (no wildcard)
  - Added filtering by OU
    - [x] `LimitOU` takes an array of DistinguishedNames - compares with like so wildcard is supported
  - Added ability to define multiple rules within one run
  - Added ability to send Admins Report as Excel
    - [x] ReportsAsExcel = $true
  - Added ability to hide Admins Report as HTML
    - [x] ReportsAsHTML = $false
  - Added ability to send expiration emails to accounts that never expire:
    - [x] PasswordNeverExpires     = $true
    - [x] PasswordNeverExpiresDays = 30

- 1.6.4 - 2020.02.17
  - Fixes to manager sent emails
  - Fixes to sending emails in some edge cases
  - More reports

- 1.6.1 - 2019.11.16
  - Some stuff was rewritten for faster processing
  - Package is now published without any dependencies
    - PSSharedGoods\PSWriteColor and other modules are used only as part of development
    - You can remove those modules if you don't use their other features as those needed functions are bundled in.
  - LimitScope added to Managers. It's possible now to send notifications to managers of users that are in a given group(s) only.
  - `<<Image>>` was added in earlier version as part of Template
  - EmailUseDefaultCredentials now available (couldn't get Emails to work on one of the servers). By default set to False, but can be set to True if you have issues to send email
  - Targets whole Forest, rather than just Domain. May add a feature to limit to only domain later on.

- 1.1 - 2019.10.19
  - New feature:
    - SendCountMaximum added - good for limiting test emails
    - DisableExpiredUsers section added
- 1.0 - 2019.05.22
  - New feature:
    - Adds UseAdditionalField (for example 'extensionAttribute13') - the way it works now is that if you define additional attribute it takes precedence in sending emails.
To understand it, imagine yourself a situation where two users exists - przemyslaw.klys@domain.com and adm.przemyslaw.klys@domain.com.
One with mailbox, the other oen without or even with mailbox.
You can put email in extensionAttribute13 przemyslaw.klys@domain.com which will cause an overwrite of default email for adm.przemyslaw.klys@domain.com which will allow sending notifications that otherwise wouldn't reach user or would be lost.
This also works great for scenarios with Azure AD where having 2 emails with same address is not possible.
- 0.7 - 2018.11.03
  - Small updates to email notification, ability to inline logo
- 0.6
  - Removed "hidden" accounts responsible for Trusts from report, added count of users to report details
- 0.5
  - Initial Release

### Sample user report

![image](https://evotec.xyz/wp-content/uploads/2018/05/img_5b05821cbc2f6.png)

### Sample manager report

![image](https://evotec.xyz/wp-content/uploads/2018/05/img_5b05816f62291.png)

### Sample admin report

![image](https://evotec.xyz/wp-content/uploads/2018/05/img_5b05807017c06.png)
