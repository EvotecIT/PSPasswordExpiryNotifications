@{
    AliasesToExport      = ''
    Author               = 'Przemyslaw Klys'
    CompanyName          = 'Evotec'
    CompatiblePSEditions = 'Desktop', 'Core'
    Copyright            = '(c) 2011 - 2019 Przemyslaw Klys. All rights reserved.'
    Description          = 'This module allows creation of password expiry emails for users, managers and administrators according to defined template.'
    FunctionsToExport    = 'Find-PasswordExpiryCheck', 'Start-PasswordExpiryCheck'
    GUID                 = '25ad8288-996d-446c-9109-c48ceda61d9c'
    ModuleVersion        = '1.6.8'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            Tags       = 'password', 'passwordexpiry', 'activedirectory', 'windows'
            ProjectUri = 'https://github.com/EvotecIT/PSPasswordExpiryNotifications'
            IconUri    = 'https://evotec.xyz/wp-content/uploads/2018/11/PSPasswordExpiryNotifications-Alternative.png'
        }
    }
    RequiredModules      = @{
        ModuleVersion = '0.0.173'
        ModuleName    = 'PSSharedGoods'
        GUID          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
    }, @{
        ModuleVersion = '0.1.10'
        ModuleName    = 'PSWriteExcel'
        GUID          = '82232c6a-27f1-435d-a496-929f7221334b'
    }
    RootModule           = 'PSPasswordExpiryNotifications.psm1'
}