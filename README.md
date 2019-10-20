[![PowerShellGallery Version](https://img.shields.io/powershellgallery/v/PSPasswordExpiryNotifications.svg)](https://www.powershellgallery.com/packages/PSPasswordExpiryNotifications)
[![PowerShellGallery Preview Version](https://img.shields.io/powershellgallery/vpre/PSPasswordExpiryNotifications.svg?label=powershell%20gallery%20preview&colorB=yellow)](https://www.powershellgallery.com/packages/PSPasswordExpiryNotifications)
[![PowerShellGallery Platform](https://img.shields.io/powershellgallery/p/PSPasswordExpiryNotifications.svg)](https://www.powershellgallery.com/packages/PSPasswordExpiryNotifications)
![Top Language](https://img.shields.io/github/languages/top/evotecit/PSPasswordExpiryNotifications.svg)
![Code](https://img.shields.io/github/languages/code-size/evotecit/PSPasswordExpiryNotifications.svg)
[![PowerShellGallery Downloads](https://img.shields.io/powershellgallery/dt/PSPasswordExpiryNotifications.svg)](https://www.powershellgallery.com/packages/PSPasswordExpiryNotifications)

# PSPasswordExpiryNotifications - PowerShell module

Following PowerShell Module provides different approach to scheduling password notifications for expiring Active Directory based accounts. While most of the scripts require knowledge on HTML... this one is just one config file and a bit of tingling around with texts. Whether this is good or bad it's up to you to decide. I do plan to add an option to use external HTML template if there will be requests for that.

## Links

- [Short description for this project at (screenshots and all)](https://evotec.xyz/just-different-approach-to-active-directory-password-notifications/)
- [Full Description (and a know-how) for this project](https://evotec.xyz/hub/scripts/pspasswordexpirynotifications-powershell-module/)

### Updates

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