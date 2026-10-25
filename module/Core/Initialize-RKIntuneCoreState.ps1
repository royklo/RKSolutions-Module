function Initialize-RKIntuneCoreState {
    <#
    .SYNOPSIS
        Creates the script-scope state the Intune anomaly core reads. A collection calls it first, so each
        collection starts with empty caches; RKSolutions.psm1 calls it once on load for Enrollment Flows.
    #>
    [CmdletBinding()]
    param()

    $script:BitLockerSettingPrefix     = 'device_vendor_msft_bitlocker_'
    $script:LapsSettingPrefix          = 'device_vendor_msft_laps_policies_'
    $script:LapsBackupDirectorySetting = 'device_vendor_msft_laps_policies_backupdirectory'
    $script:AllFilters = @{}
    $script:AllGroups  = @{}
    $script:ODataToFriendly = @{
        '#microsoft.graph.allDevicesAssignmentTarget'       = 'All Devices'
        '#microsoft.graph.allLicensedUsersAssignmentTarget' = 'All Users'
        '#microsoft.graph.groupAssignmentTarget'            = 'Group (Include)'
        '#microsoft.graph.exclusionGroupAssignmentTarget'   = 'Group (Exclude)'
    }
    # Admin-center names for compliance rules, keyed by the settingName suffix Graph returns.
    # Sources: learn.microsoft.com/intune/device-security/compliance/ref-windows-settings
    #          learn.microsoft.com/intune/device-security/compliance/monitor-policy (built-in policy)
    # Names must not contain ', ' - the Noncompliant tab splits reasons on it.
    $script:ComplianceSettingFriendlyNames = @{
        'RequireRemainContact'                        = 'Is active (no compliance check-in within the compliance status validity period)'
        'RequireDeviceCompliancePolicyAssigned'       = 'Has a compliance policy assigned'
        'RequireUserExistence'                        = 'Enrolled user exists (user deleted or disabled or without an Intune license)'
        'BitLockerEnabled'                            = 'Require BitLocker'
        'SecureBootEnabled'                           = 'Require Secure Boot to be enabled on the device'
        'CodeIntegrityEnabled'                        = 'Require code integrity'
        'EarlyLaunchAntiMalwareDriverEnabled'         = 'Early launch antimalware driver enabled'
        'RequireHealthyDeviceReport'                  = 'Reported healthy by Device Health Attestation'
        'OsMinimumVersion'                            = 'Minimum OS version'
        'OsMaximumVersion'                            = 'Maximum OS version'
        'MobileOsMinimumVersion'                      = 'Minimum OS required for mobile devices'
        'MobileOsMaximumVersion'                      = 'Maximum OS required for mobile devices'
        'ValidOperatingSystemBuildRanges'             = 'Valid operating system builds'
        'ConfigurationManagerComplianceRequired'      = 'Require device compliance from Configuration Manager'
        'PasswordRequired'                            = 'Require a password to unlock mobile devices'
        'PasswordBlockSimple'                         = 'Simple passwords'
        'PasswordRequiredType'                        = 'Password type'
        'PasswordMinimumCharacterSetCount'            = 'Password complexity'
        'PasswordMinimumLength'                       = 'Minimum password length'
        'PasswordMinutesOfInactivityBeforeLock'       = 'Maximum minutes of inactivity before password is required'
        'PasswordExpirationDays'                      = 'Password expiration (days)'
        'PasswordPreviousPasswordBlockCount'          = 'Number of previous passwords to prevent reuse'
        'PasswordRequiredToUnlockFromIdle'            = 'Require password when device returns from idle state'
        'StorageRequireEncryption'                    = 'Encryption of data storage on a device'
        'ActiveFirewallRequired'                      = 'Firewall (turned off or allowing all inbound traffic)'
        'TpmRequired'                                 = 'Trusted Platform Module (TPM)'
        'AntivirusRequired'                           = 'Antivirus (turned off or out of date in Windows Security)'
        'AntiSpywareRequired'                         = 'Antispyware (turned off or out of date in Windows Security)'
        'DefenderEnabled'                             = 'Microsoft Defender Antimalware (service turned off)'
        'DefenderVersion'                             = 'Microsoft Defender Antimalware minimum version'
        'SignatureOutOfDate'                          = 'Microsoft Defender Antimalware security intelligence up-to-date'
        'RtpEnabled'                                  = 'Real-time protection (turned off)'
        'DeviceThreatProtectionEnabled'               = 'Require the device to be at or under the machine risk score'
        'DeviceThreatProtectionRequiredSecurityLevel' = 'Require the device to be at or under the machine risk score'
    }
}
