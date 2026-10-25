# IntuneAnomalyCore.Tests.ps1 - unit tests for module/Core and the module's core adapters. No Graph connection.

BeforeAll {
    Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
    Import-Module (Join-Path $PSScriptRoot '..' 'module' 'RKSolutions.psd1') -Force
}

Describe 'ConvertTo-RKCoreGraphPath' {
    It 'Splits <Uri>' -ForEach @(
        @{ Uri = 'beta/deviceManagement/managedDevices'; Version = 'beta'; Path = 'deviceManagement/managedDevices' }
        @{ Uri = '/v1.0/groups/1/transitiveMembers'; Version = 'v1.0'; Path = 'groups/1/transitiveMembers' }
        @{ Uri = 'https://graph.microsoft.com/v1.0/devices?$select=id'; Version = 'v1.0'; Path = 'devices?$select=id' }
    ) {
        $r = InModuleScope RKSolutions -Parameters @{ Uri = $Uri } { param($Uri) ConvertTo-RKCoreGraphPath -Uri $Uri }
        $r.Version | Should -Be $Version
        $r.Path | Should -Be $Path
    }

    It 'Refuses a URI without a version' {
        { InModuleScope RKSolutions { ConvertTo-RKCoreGraphPath -Uri 'deviceManagement/managedDevices' } } | Should -Throw '*v1.0/ or beta/*'
    }
}

Describe 'Get-RKCoreGraphErrorStatus' {
    It 'Reads Response.StatusCode' {
        InModuleScope RKSolutions {
            $e = [System.Exception]::new('x')
            $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 404 })
            Get-RKCoreGraphErrorStatus -ErrorRecord ([System.Management.Automation.ErrorRecord]::new($e, 'id', 'NotSpecified', $null)) | Should -Be 404
        }
    }
    It 'Reads StatusCode from an HttpRequestException' {
        InModuleScope RKSolutions {
            $e = [System.Net.Http.HttpRequestException]::new('x', $null, [System.Net.HttpStatusCode]::Forbidden)
            Get-RKCoreGraphErrorStatus -ErrorRecord ([System.Management.Automation.ErrorRecord]::new($e, 'id', 'NotSpecified', $null)) | Should -Be 403
        }
    }
    It 'Returns 0 without a status' {
        InModuleScope RKSolutions {
            $e = [System.Exception]::new('x')
            Get-RKCoreGraphErrorStatus -ErrorRecord ([System.Management.Automation.ErrorRecord]::new($e, 'id', 'NotSpecified', $null)) | Should -Be 0
        }
    }
}

Describe 'Get-RKCoreBatchBody' {
    It 'Returns the body for 200' {
        InModuleScope RKSolutions {
            $w = [System.Collections.Generic.List[object]]::new()
            (Get-RKCoreBatchBody -Response ([pscustomobject]@{ Id = 'a'; Status = 200; Body = @{ value = 1 } }) -What 'x' -Area 'BitLocker' -Warnings $w).value | Should -Be 1
            $w.Count | Should -Be 0
        }
    }
    It 'Returns $null and warns for 404' {
        InModuleScope RKSolutions {
            $w = [System.Collections.Generic.List[object]]::new()
            Get-RKCoreBatchBody -Response ([pscustomobject]@{ Id = 'a'; Status = 404; Body = $null }) -What 'group g1' -Area 'Laps' -Warnings $w | Should -BeNullOrEmpty
            $w[0].Area | Should -Be 'Laps'
            $w[0].Message | Should -BeLike '*group g1*'
        }
    }
    It 'Warns once per area for 404' {
        InModuleScope RKSolutions {
            $w = [System.Collections.Generic.List[object]]::new()
            Get-RKCoreBatchBody -Response ([pscustomobject]@{ Id = 'a'; Status = 404; Body = $null }) -What 'group g1' -Area 'BitLocker', 'Laps' -Warnings $w | Should -BeNullOrEmpty
            @($w.Area) | Should -Be @('BitLocker', 'Laps')
            @($w | Where-Object { $_.Message -like '*group g1*' }).Count | Should -Be 2
        }
    }
    It 'Throws for any other status' {
        { InModuleScope RKSolutions {
            $w = [System.Collections.Generic.List[object]]::new()
            Get-RKCoreBatchBody -Response ([pscustomobject]@{ Id = 'a'; Status = 403; Body = $null }) -What 'policy p1' -Area 'BitLocker' -Warnings $w
        } } | Should -Throw '*policy p1*403*'
    }
}

Describe 'Module adapters' {
    It 'Invoke-RKCoreGraphRequest -All follows nextLink' {
        InModuleScope RKSolutions {
            Mock Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*skip=2' } { [pscustomobject]@{ value = @(3) } }
            Mock Invoke-MgGraphRequest -ParameterFilter { $Uri -notlike '*skip=2' } { [pscustomobject]@{ value = @(1, 2); '@odata.nextLink' = 'https://graph.microsoft.com/beta/x?skip=2' } }
            @(Invoke-RKCoreGraphRequest -Uri 'beta/x' -All) | Should -Be @(1, 2, 3)
        }
    }
    It 'Invoke-RKCoreGraphRequest -All throws when a nextLink repeats' {
        InModuleScope RKSolutions {
            Mock Invoke-MgGraphRequest { [pscustomobject]@{ value = @(1); '@odata.nextLink' = 'https://graph.microsoft.com/beta/x?skip=1' } }
            { Invoke-RKCoreGraphRequest -Uri 'beta/x' -All } | Should -Throw '*same @odata.nextLink*'
        }
    }
    It 'Invoke-RKCoreGraphRequest without -All returns the raw body' {
        InModuleScope RKSolutions {
            Mock Invoke-MgGraphRequest { [pscustomobject]@{ value = @(1); '@odata.nextLink' = 'n' } }
            (Invoke-RKCoreGraphRequest -Uri 'beta/x').'@odata.nextLink' | Should -Be 'n'
        }
    }
    It 'Invoke-RKCoreGraphRequest retries a 502 once and returns the next result' {
        InModuleScope RKSolutions {
            $script:calls = 0
            Mock Start-Sleep { }
            Mock Invoke-MgGraphRequest {
                $script:calls++
                if ($script:calls -eq 1) {
                    $e = [System.Exception]::new('bad gateway')
                    $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 502 })
                    throw $e
                }
                [pscustomobject]@{ value = @(1) }
            }
            @(Invoke-RKCoreGraphRequest -Uri 'beta/x' -All) | Should -Be @(1)
        }
    }
    It 'Invoke-RKCoreGraphRequest does not retry a 403' {
        InModuleScope RKSolutions {
            Mock Start-Sleep { }
            Mock Invoke-MgGraphRequest {
                $e = [System.Exception]::new('forbidden')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 403 })
                throw $e
            }
            { Invoke-RKCoreGraphRequest -Uri 'beta/x' -All } | Should -Throw
            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        }
    }
    It 'Invoke-RKCoreGraphBatch groups requests by version' {
        InModuleScope RKSolutions {
            Mock Invoke-RKGraphBatch { foreach ($r in $Requests) { [pscustomobject]@{ Id = $r.Id; Status = 200; Body = "$GraphVersion$($r.Url)" } } }
            $out = @(Invoke-RKCoreGraphBatch -Requests @(
                    [pscustomobject]@{ Id = 'a'; Url = 'beta/x' }
                    [pscustomobject]@{ Id = 'b'; Url = 'v1.0/y' }
                ))
            ($out | Where-Object Id -EQ 'a').Body | Should -Be 'beta/x'
            ($out | Where-Object Id -EQ 'b').Body | Should -Be 'v1.0/y'
        }
    }
}

Describe 'Get-DetailedPolicyAssignments -ThrowOnError' {
    It 'returns an empty result for a 404' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest {
                $e = [System.Exception]::new('http 404')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 404 })
                throw $e
            }
            @(Get-DetailedPolicyAssignments -EntityType 'configurationPolicies' -EntityId 'p1' -PolicyName 'P' -ThrowOnError).Count | Should -Be 0
        }
    }
    It 'throws for a 500' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest {
                $e = [System.Exception]::new('http 500')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 500 })
                throw $e
            }
            { Get-DetailedPolicyAssignments -EntityType 'configurationPolicies' -EntityId 'p1' -PolicyName 'P' -ThrowOnError } | Should -Throw
        }
    }
    It 'returns an empty result for a 500 without the switch' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest {
                $e = [System.Exception]::new('http 500')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 500 })
                throw $e
            }
            @(Get-DetailedPolicyAssignments -EntityType 'configurationPolicies' -EntityId 'p1' -PolicyName 'P').Count | Should -Be 0
        }
    }
}

Describe 'Collector error kinds' {
    It 'Get-BitLockerLapsAssignmentContext marks the BitLocker keys unreadable on 403 and still returns the context' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -like 'v1.0/informationProtection/bitlocker/recoveryKeys*' } {
                $e = [System.Exception]::new('http 403')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 403 })
                throw $e
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $ctx = Get-BitLockerLapsAssignmentContext -Warnings $w
            $ctx.BitLockerKeysReadable | Should -BeFalse
            $ctx.BitLockerKeysError | Should -Be 'http 403'
            $ctx.LapsCredentialsReadable | Should -BeTrue
        }
    }
    It 'Get-BitLockerLapsAssignmentContext throws when the configurationPolicies list answers 500' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -eq 'beta/deviceManagement/configurationPolicies' } {
                $e = [System.Exception]::new('http 500')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 500 })
                throw $e
            }
            $w = [System.Collections.Generic.List[object]]::new()
            { Get-BitLockerLapsAssignmentContext -Warnings $w } | Should -Throw 'http 500'
        }
    }
    It 'Get-BitLockerLapsAssignmentContext warns, and does not throw, when a group is not found (404)' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -eq 'beta/deviceManagement/deviceConfigurations' } {
                [pscustomobject]@{ id = 'c1'; displayName = 'BitLocker'; '@odata.type' = '#microsoft.graph.windows10EndpointProtectionConfiguration' }
            }
            Mock Get-DetailedPolicyAssignments { [pscustomobject]@{ GroupId = 'g1' } }
            Mock Invoke-RKCoreGraphBatch -ParameterFilter { @($Requests.Url) -like '*/transitiveMembers/*' } {
                [pscustomobject]@{ Id = 'gd:g1'; Status = 404; Body = $null }
                [pscustomobject]@{ Id = 'gu:g1'; Status = 200; Body = [pscustomobject]@{ value = @([pscustomobject]@{ id = 'u1'; userPrincipalName = 'User@contoso.com' }) } }
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $ctx = Get-BitLockerLapsAssignmentContext -Warnings $w
            @($w | Where-Object { $_.Message -like '*group g1*' }).Area | Should -Be @('BitLocker', 'Laps')
            $ctx.GroupUserUpns['g1'].Contains('user@contoso.com') | Should -BeTrue
        }
    }
    It 'Get-BitLockerLapsAssignmentContext throws when recoveryKeys answers 500' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -like 'v1.0/informationProtection/bitlocker/recoveryKeys*' } {
                $e = [System.Exception]::new('http 500')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 500 })
                throw $e
            }
            $w = [System.Collections.Generic.List[object]]::new()
            { Get-BitLockerLapsAssignmentContext -Warnings $w } | Should -Throw 'http 500'
        }
    }
}

Describe 'Collector logic' {
    It 'Get-AllDeviceData returns nothing, and throws nothing, for a tenant without devices' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            $w = [System.Collections.Generic.List[object]]::new()
            @(Get-AllDeviceData -Users @() -TenantName '' -Warnings $w).Count | Should -Be 0
            $w.Count | Should -Be 0
        }
    }
    It 'Get-AllDeviceData asks once for the settingStates of a policy state listed twice' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -like 'beta/deviceManagement/managedDevices?*' } {
                [pscustomobject]@{ id = 'd1'; deviceName = 'PC1'; serialNumber = 'S1'; isEncrypted = $true; complianceState = 'noncompliant'; managementAgent = 'mdm' }
            }
            Mock Invoke-RKCoreGraphBatch -ParameterFilter { @($Requests.Url) -like '*/deviceCompliancePolicyStates' } {
                $state = [pscustomobject]@{ id = 'ps1'; state = 'nonCompliant' }
                [pscustomobject]@{ Id = 'ps:d1'; Status = 200; Body = [pscustomobject]@{ value = @($state, $state) } }
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $null = Get-AllDeviceData -Users @() -TenantName 'T' -Warnings $w
            Should -Invoke Invoke-RKCoreGraphBatch -Times 1 -Exactly -ParameterFilter {
                (@($Requests.Url) -join '|') -eq 'beta/deviceManagement/managedDevices/d1/deviceCompliancePolicyStates/ps1/settingStates'
            }
        }
    }
    It 'Get-AllDeviceData skips the settingStates of a device that fails more than 10 policies, and counts it' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -like 'beta/deviceManagement/managedDevices?*' } {
                [pscustomobject]@{ id = 'd2'; deviceName = 'PC2'; serialNumber = 'S2'; isEncrypted = $true; complianceState = 'noncompliant'; managementAgent = 'mdm' }
            }
            Mock Invoke-RKCoreGraphBatch -ParameterFilter { @($Requests.Url) -like '*/deviceCompliancePolicyStates' } {
                [pscustomobject]@{ Id = 'ps:d2'; Status = 200; Body = [pscustomobject]@{ value = @(1..11 | ForEach-Object { [pscustomobject]@{ id = "s$_"; state = 'nonCompliant' } }) } }
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $null = Get-AllDeviceData -Users @() -TenantName 'T' -Warnings $w
            Should -Invoke Invoke-RKCoreGraphBatch -Times 0 -Exactly -ParameterFilter { @($Requests.Url) -like '*/settingStates' }
            @($w | Where-Object { $_.Area -eq 'Compliance' -and $_.Message -like 'The compliance reasons of 1 device(s)*' }).Count | Should -Be 1
        }
    }
    It 'Get-BitLockerLapsAssignmentContext keeps Entra-backed LAPS policies (_1), not on-premises ones (_2), and BitLocker policies' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Get-DetailedPolicyAssignments { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -eq 'beta/deviceManagement/configurationPolicies' } {
                [pscustomobject]@{ id = 'p1'; name = 'LAPS Entra' }
                [pscustomobject]@{ id = 'p2'; name = 'LAPS AD' }
                [pscustomobject]@{ id = 'p3'; name = 'BitLocker' }
            }
            Mock Invoke-RKCoreGraphBatch -ParameterFilter { @($Requests.Url) -like '*/settings?*' } {
                foreach ($p in @{ p1 = 1; p2 = 2 }.GetEnumerator()) {
                    $choice = [pscustomobject]@{ value = "device_vendor_msft_laps_policies_backupdirectory_$($p.Value)" }
                    $setting = [pscustomobject]@{ settingInstance = [pscustomobject]@{ settingDefinitionId = 'device_vendor_msft_laps_policies_backupdirectory'; choiceSettingValue = $choice } }
                    [pscustomobject]@{ Id = "cps:$($p.Key)"; Status = 200; Body = [pscustomobject]@{ value = @($setting) } }
                }
                $bitLocker = [pscustomobject]@{ settingInstance = [pscustomobject]@{ settingDefinitionId = 'device_vendor_msft_bitlocker_requiredeviceencryption' } }
                [pscustomobject]@{ Id = 'cps:p3'; Status = 200; Body = [pscustomobject]@{ value = @($bitLocker) } }
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $ctx = Get-BitLockerLapsAssignmentContext -Warnings $w
            @($ctx.LapsPolicies.Id) | Should -Be @('p1')
            @($ctx.BitLockerPolicies.Id) | Should -Be @('p3')
        }
    }
    It 'Get-BitLockerLapsAssignmentContext follows the nextLink of a group-members page once and adds its members' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -eq 'beta/deviceManagement/deviceConfigurations' } {
                [pscustomobject]@{ id = 'c1'; displayName = 'BitLocker'; '@odata.type' = '#microsoft.graph.windows10EndpointProtectionConfiguration' }
            }
            Mock Get-DetailedPolicyAssignments { [pscustomobject]@{ GroupId = 'g1' } }
            Mock Invoke-RKCoreGraphBatch -ParameterFilter { @($Requests.Url) -like '*/transitiveMembers/*' } {
                $next = 'https://graph.microsoft.com/v1.0/groups/g1/transitiveMembers/microsoft.graph.device?$skiptoken=x'
                [pscustomobject]@{ Id = 'gd:g1'; Status = 200; Body = [pscustomobject]@{ value = @([pscustomobject]@{ id = 'm1'; deviceId = 'dev1' }); '@odata.nextLink' = $next } }
                [pscustomobject]@{ Id = 'gu:g1'; Status = 200; Body = [pscustomobject]@{ value = @() } }
            }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -like '*skiptoken=x' } { [pscustomobject]@{ id = 'm2'; deviceId = 'dev2' } }
            $w = [System.Collections.Generic.List[object]]::new()
            $ctx = Get-BitLockerLapsAssignmentContext -Warnings $w
            Should -Invoke Invoke-RKCoreGraphRequest -Times 1 -Exactly -ParameterFilter { $Uri -like '*skiptoken=x' -and $All }
            @($ctx.GroupDeviceAadIds['g1'] | Sort-Object) | Should -Be @('dev1', 'dev2')
        }
    }
    It 'Get-BitLockerLapsAssignmentContext marks the LAPS credentials unreadable on 403' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Invoke-RKCoreGraphBatch { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -like 'v1.0/directory/deviceLocalCredentials*' } {
                $e = [System.Exception]::new('http 403')
                $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 403 })
                throw $e
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $ctx = Get-BitLockerLapsAssignmentContext -Warnings $w
            $ctx.LapsCredentialsReadable | Should -BeFalse
            $ctx.LapsCredentialsError | Should -Be 'http 403'
            $ctx.BitLockerKeysReadable | Should -BeTrue
        }
    }
}

Describe 'Resolvers' {
    BeforeAll {
        $script:data = InModuleScope RKSolutions {
            $old = (Get-Date).AddDays(-200)
            [pscustomobject]@{
                TenantName = 'T'
                Devices = @(
                    [pscustomobject]@{ Customer = 'T'; DeviceName = 'D1'; PrimaryUser = 'off@t'; Serialnumber = 'S1'; DeviceManufacturer = 'M'; DeviceModel = 'X'; LastContact = $old; DeviceHashUploaded = $false; DeviceAssociated = $false; IsCloudPC = $false; usersLoggedOnCount = 2; usersLoggedOnIds = 'a, b'; EnrollmentProfile = 'Single'; ComplianceStatus = 'noncompliant'; NoncompliantBasedOn = 'Require BitLocker, Real-time protection (turned off)'; NoncompliantAlert = 'Require BitLocker' }
                    [pscustomobject]@{ Customer = 'T'; DeviceName = 'D2'; PrimaryUser = 'on@t'; Serialnumber = 'S2'; DeviceManufacturer = 'M'; DeviceModel = 'X'; LastContact = (Get-Date); DeviceHashUploaded = $true; DeviceAssociated = $false; IsCloudPC = $false; usersLoggedOnCount = 1; usersLoggedOnIds = 'a'; EnrollmentProfile = 'Single'; ComplianceStatus = 'compliant'; NoncompliantBasedOn = ''; NoncompliantAlert = '' }
                    [pscustomobject]@{ Customer = 'T'; DeviceName = 'CPC'; PrimaryUser = 'on@t'; Serialnumber = 'S3'; DeviceManufacturer = 'M'; DeviceModel = 'X'; LastContact = (Get-Date); DeviceHashUploaded = $false; DeviceAssociated = $false; IsCloudPC = $true; usersLoggedOnCount = 1; usersLoggedOnIds = 'a'; EnrollmentProfile = ''; ComplianceStatus = 'compliant'; NoncompliantBasedOn = ''; NoncompliantAlert = '' }
                )
                Users = @(
                    [pscustomobject]@{ id = '1'; userPrincipalName = 'off@t'; accountEnabled = $false }
                    [pscustomobject]@{ id = '2'; userPrincipalName = 'on@t'; accountEnabled = $true }
                )
                AutopilotProfiles = @([pscustomobject]@{ displayName = 'Single'; outOfBoxExperienceSettings = [pscustomobject]@{ deviceUsageType = 'SingleUser' } })
                SecurityKeyContext = [pscustomobject]@{ ConfigurationPolicies = @() }
                Warnings = @([pscustomobject]@{ Area = 'Compliance'; Message = 'w1' })
            }
        }
    }

    It 'Inactive: devices older than -Days' {
        $r = InModuleScope RKSolutions -Parameters @{ D = $script:data } { param($D) Resolve-IntuneInactiveDevices -Data $D -Days 90 }
        @($r.Rows).DeviceName | Should -Be @('D1')
    }
    It 'Disabled primary user' {
        $r = InModuleScope RKSolutions -Parameters @{ D = $script:data } { param($D) Resolve-IntuneDisabledPrimaryUsers -Data $D }
        @($r.Rows).DeviceName | Should -Be @('D1')
    }
    It 'Missing Autopilot leaves out Cloud PCs' {
        $r = InModuleScope RKSolutions -Parameters @{ D = $script:data } { param($D) Resolve-IntuneMissingAutopilotDevices -Data $D }
        @($r.Rows).DeviceName | Should -Be @('D1')
    }
    It 'Multi-user only on single-user Autopilot profiles' {
        $r = InModuleScope RKSolutions -Parameters @{ D = $script:data } { param($D) Resolve-IntuneMultiUserDevices -Data $D }
        @($r.Rows).DeviceName | Should -Be @('D1')
    }
    It 'Noncompliant: one row per reason, with its Compliance warnings' {
        $r = InModuleScope RKSolutions -Parameters @{ D = $script:data } { param($D) Resolve-IntuneNoncompliantDevices -Data $D }
        @($r.Rows).Count | Should -Be 2
        @($r.Warnings).Message | Should -Be @('w1')
    }
    It 'Disabled primary user matches the UPN case-insensitively' {
        $r = InModuleScope RKSolutions {
            $d = [pscustomobject]@{
                Devices = @([pscustomobject]@{ Customer = 'T'; DeviceName = 'U1'; PrimaryUser = 'OFF@T'; Serialnumber = 'S'; DeviceManufacturer = 'M'; DeviceModel = 'X' })
                Users   = @([pscustomobject]@{ id = '1'; userPrincipalName = 'off@t'; accountEnabled = $false })
            }
            Resolve-IntuneDisabledPrimaryUsers -Data $d
        }
        @($r.Rows).DeviceName | Should -Be @('U1')
    }
    It 'Noncompliant without a reason gives one row with Unknown' {
        $r = InModuleScope RKSolutions {
            $d = [pscustomobject]@{ Devices = @([pscustomobject]@{ Customer = 'T'; DeviceName = 'N1'; ComplianceStatus = 'noncompliant'; NoncompliantBasedOn = '' }); Warnings = @() }
            Resolve-IntuneNoncompliantDevices -Data $d
        }
        @($r.Rows).Count | Should -Be 1
        @($r.Rows)[0].NoncompliantBasedOn | Should -Be 'Unknown'
    }
    It 'Inactive: a device without LastContact counts as inactive' {
        $r = InModuleScope RKSolutions {
            $d = [pscustomobject]@{ Devices = @([pscustomobject]@{ Customer = 'T'; DeviceName = 'Z1'; LastContact = $null }) }
            Resolve-IntuneInactiveDevices -Data $d -Days 90
        }
        @($r.Rows).DeviceName | Should -Be @('Z1')
    }
    It 'Resolvers work on a Data object without Warnings' {
        $r = InModuleScope RKSolutions {
            $d = [pscustomobject]@{ Devices = @(); Users = @(); AutopilotProfiles = @() }
            @(
                (Resolve-IntuneNoncompliantDevices -Data $d)
                (Resolve-IntuneMissingAutopilotDevices -Data $d)
            )
        }
        foreach ($x in $r) { @($x.Warnings).Count | Should -Be 0 }
    }
    It 'Excluded devices come back as Info rows (BitLocker and LAPS)' {
        $r = InModuleScope RKSolutions {
            $policy = [pscustomobject]@{ DisplayName = 'P'; Source = 'Catalog'; Assignments = @(
                    [pscustomobject]@{ AssignmentType = 'All Devices' }
                    [pscustomobject]@{ AssignmentType = 'Group (Exclude)'; GroupId = 'g1' }) }
            $ctx = [pscustomobject]@{
                EntraDeviceIdByName        = @{}
                GroupDeviceAadIds          = @{ g1 = [System.Collections.Generic.HashSet[string]]::new([string[]]@('aad1')) }
                GroupUserUpns              = @{}
                Filters                    = @{}
                BitLockerPolicies          = @($policy)
                BitLockerKeyDeviceIds      = [System.Collections.Generic.HashSet[string]]::new()
                LapsPolicies               = @($policy)
                LapsCredentialByDeviceId   = @{}
                LapsCredentialByDeviceName = @{}
            }
            $d = [pscustomobject]@{
                TenantName         = 'T'
                Devices            = @([pscustomobject]@{ DeviceName = 'E1'; AzureAdDeviceId = 'aad1'; PrimaryUser = 'u@t'; Serialnumber = 'S'; Encrypted = $false; IsCloudPC = $false })
                SecurityKeyContext = $ctx
                Warnings           = @()
            }
            @{ BitLocker = (Resolve-IntuneBitLockerAnomalies -Data $d); Laps = (Resolve-IntuneLapsAnomalies -Data $d) }
        }
        foreach ($k in 'BitLocker', 'Laps') {
            @($r[$k].Rows).Count | Should -Be 1
            @($r[$k].Rows)[0].Severity | Should -Be 'Info'
            @($r[$k].Rows)[0].Status | Should -BeLike 'Intentionally excluded*'
        }
    }
    It 'BitLocker and LAPS resolvers work when the tenant has no such policy (empty lists)' {
        $r = InModuleScope RKSolutions {
            $ctx = [pscustomobject]@{
                EntraDeviceIdByName        = @{}
                GroupDeviceAadIds          = @{}
                GroupUserUpns              = @{}
                Filters                    = @{}
                BitLockerPolicies          = [System.Collections.Generic.List[object]]::new()
                BitLockerKeyDeviceIds      = [System.Collections.Generic.HashSet[string]]::new()
                LapsPolicies               = [System.Collections.Generic.List[object]]::new()
                LapsCredentialByDeviceId   = @{}
                LapsCredentialByDeviceName = @{}
            }
            $d = [pscustomobject]@{
                TenantName         = 'T'
                Devices            = @([pscustomobject]@{ DeviceName = 'N1'; AzureAdDeviceId = 'aad1'; PrimaryUser = 'u@t'; Serialnumber = 'S'; Encrypted = $false; IsCloudPC = $false })
                SecurityKeyContext = $ctx
                Warnings           = @()
            }
            @{ BitLocker = (Resolve-IntuneBitLockerAnomalies -Data $d); Laps = (Resolve-IntuneLapsAnomalies -Data $d) }
        }
        @($r.BitLocker.Rows)[0].Severity | Should -Be 'Critical'
        @($r.BitLocker.Rows)[0].Status | Should -Be 'Device not encrypted and no BitLocker policy assigned'
        @($r.Laps.Rows)[0].Severity | Should -Be 'Warning'
        @($r.Laps.Rows)[0].Status | Should -Be 'No Entra-backed LAPS policy assigned'
    }
    It 'Deprecated: reads display names from the expanded setting definitions' {
        $r = InModuleScope RKSolutions {
            $policy = [pscustomobject]@{ id = 'p1'; name = 'Pol'; platforms = 'windows10'; technologies = 'mdm'; settings = @(
                    [pscustomobject]@{
                        settingInstance    = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance'; settingDefinitionId = 'def_a'; choiceSettingValue = [pscustomobject]@{ value = 'def_a_1' } }
                        settingDefinitions = @([pscustomobject]@{ id = 'def_a'; displayName = 'Old thing (Deprecated)' })
                    }) }
            $d = [pscustomobject]@{ TenantName = 'T'; SecurityKeyContext = [pscustomobject]@{ ConfigurationPolicies = @($policy) }; Warnings = @() }
            Resolve-IntuneDeprecatedAnomalies -Data $d
        }
        @($r.Rows).SettingDisplayName | Should -Be @('Old thing (Deprecated)')
        @($r.Rows)[0].DetectionSource | Should -Be 'CatalogDisplayName'
    }
}

Describe 'Unreadable keys, settings fallback and null children' {
    It 'BitLocker gives Unknown, and no missing-key anomaly, when the keys cannot be read' {
        $r = InModuleScope RKSolutions {
            $policy = [pscustomobject]@{ DisplayName = 'P'; Source = 'Catalog'; Assignments = @([pscustomobject]@{ AssignmentType = 'All Devices' }) }
            $ctx = [pscustomobject]@{
                EntraDeviceIdByName = @{}; GroupDeviceAadIds = @{}; GroupUserUpns = @{}; Filters = @{}
                BitLockerPolicies = @($policy); BitLockerKeyDeviceIds = [System.Collections.Generic.HashSet[string]]::new()
            }
            $devices = @(
                [pscustomobject]@{ DeviceName = 'Enc'; AzureAdDeviceId = 'a1'; PrimaryUser = 'u@t'; Serialnumber = 'S1'; Encrypted = $true; IsCloudPC = $false }
                [pscustomobject]@{ DeviceName = 'Plain'; AzureAdDeviceId = 'a2'; PrimaryUser = 'u@t'; Serialnumber = 'S2'; Encrypted = $false; IsCloudPC = $false })
            $d = [pscustomobject]@{ TenantName = 'T'; Devices = $devices; SecurityKeyContext = $ctx; Warnings = @(); BitLockerKeysReadable = $false }
            $d2 = $d.PSObject.Copy(); $d2.BitLockerKeysReadable = $true
            @{ Unreadable = (Resolve-IntuneBitLockerAnomalies -Data $d); Readable = (Resolve-IntuneBitLockerAnomalies -Data $d2) }
        }
        @($r.Readable.Rows | Where-Object DeviceName -eq 'Enc').Status | Should -Be 'Encrypted but no OS-volume key in Entra'
        @($r.Unreadable.Rows | Where-Object DeviceName -eq 'Enc') | Should -BeNullOrEmpty
        $plain = @($r.Unreadable.Rows | Where-Object DeviceName -eq 'Plain')
        $plain.Status | Should -Be 'BitLocker policy assigned but device not encrypted'
        $plain.KeyEscrowed | Should -Be 'Unknown'
    }
    It 'LAPS does not flag a missing or stale backup when the credentials cannot be read' {
        $r = InModuleScope RKSolutions {
            $policy = [pscustomobject]@{ DisplayName = 'P'; Source = 'Catalog'; Assignments = @([pscustomobject]@{ AssignmentType = 'All Devices' }) }
            $ctx = [pscustomobject]@{
                EntraDeviceIdByName = @{}; GroupDeviceAadIds = @{}; GroupUserUpns = @{}; Filters = @{}
                LapsPolicies = @($policy); LapsCredentialByDeviceId = @{}; LapsCredentialByDeviceName = @{}
            }
            $d = [pscustomobject]@{ TenantName = 'T'; Devices = @([pscustomobject]@{ DeviceName = 'N1'; AzureAdDeviceId = 'a1'; PrimaryUser = 'u@t'; Serialnumber = 'S' }); SecurityKeyContext = $ctx; Warnings = @(); LapsCredentialsReadable = $false }
            $d2 = $d.PSObject.Copy(); $d2.LapsCredentialsReadable = $true
            @{ Unreadable = (Resolve-IntuneLapsAnomalies -Data $d); Readable = (Resolve-IntuneLapsAnomalies -Data $d2) }
        }
        @($r.Readable.Rows)[0].Status | Should -BeLike '*no credential backed up*'
        @($r.Unreadable.Rows) | Should -BeNullOrEmpty
    }
    It 'Get-BitLockerLapsAssignmentContext reads the settings of a policy directly when the batch answers 400' {
        InModuleScope RKSolutions {
            Mock Invoke-RKCoreGraphRequest { }
            Mock Get-DetailedPolicyAssignments { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -eq 'beta/deviceManagement/configurationPolicies' } { [pscustomobject]@{ id = 'p1'; name = 'BitLocker' } }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -like '*configurationPolicies/p1/settings*' -and $All } {
                [pscustomobject]@{ settingInstance = [pscustomobject]@{ settingDefinitionId = 'device_vendor_msft_bitlocker_requiredeviceencryption' } }
            }
            Mock Invoke-RKCoreGraphBatch -ParameterFilter { @($Requests.Url) -like '*/settings?*' } {
                [pscustomobject]@{ Id = 'cps:p1'; Status = 400; Body = $null }
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $ctx = Get-BitLockerLapsAssignmentContext -Warnings $w
            @($ctx.ConfigurationPolicies[0].settings).Count | Should -Be 1
            @($ctx.BitLockerPolicies.Id) | Should -Be @('p1')
        }
    }
    It 'Get-RKCoreBatchBody puts the Graph error message in the throw' {
        { InModuleScope RKSolutions {
            $w = [System.Collections.Generic.List[object]]::new()
            Get-RKCoreBatchBody -Response ([pscustomobject]@{ Id = 'a'; Status = 400; Body = [pscustomobject]@{ error = [pscustomobject]@{ message = 'Resource not valid' } } }) -What 'policy p1' -Area 'BitLocker' -Warnings $w
        } } | Should -Throw '*400*Resource not valid*'
    }
    It 'Deprecated: a group setting whose children are null does not throw' {
        $r = InModuleScope RKSolutions {
            $si = '{"@odata.type":"#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance","settingDefinitionId":"x_group","groupSettingCollectionValue":[{"children":null}]}' | ConvertFrom-Json
            $si2 = '{"@odata.type":"#microsoft.graph.deviceManagementConfigurationGroupSettingInstance","settingDefinitionId":"y_group","groupSettingValue":{"children":null}}' | ConvertFrom-Json
            $policy = [pscustomobject]@{ id = 'p'; name = 'P'; settings = @([pscustomobject]@{ settingInstance = $si }, [pscustomobject]@{ settingInstance = $si2 }) }
            $d = [pscustomobject]@{ TenantName = 'T'; Warnings = @(); SecurityKeyContext = [pscustomobject]@{ ConfigurationPolicies = @($policy) } }
            Resolve-IntuneDeprecatedAnomalies -Data $d
        }
        @($r.Rows).Count | Should -Be 0
    }
}

Describe 'General warnings' {
    It 'Every device check carries the General warnings, the deprecated check does not' {
        $r = InModuleScope RKSolutions {
            $d = [pscustomobject]@{ TenantName = 'T'; Devices = @(); Users = @(); AutopilotProfiles = @(); Warnings = @([pscustomobject]@{ Area = 'General'; Message = 'Device X could not be processed' }, [pscustomobject]@{ Area = 'Zzz'; Message = 'other' })
                SecurityKeyContext = [pscustomobject]@{ ConfigurationPolicies = @(); BitLockerPolicies = @(); LapsPolicies = @(); BitLockerKeyDeviceIds = [System.Collections.Generic.HashSet[string]]::new(); LapsCredentialByDeviceId = @{}; LapsCredentialByDeviceName = @{}; EntraDeviceIdByName = @{}; GroupDeviceAadIds = @{}; GroupUserUpns = @{}; Filters = @{} } }
            $out = @{}
            foreach ($f in 'BitLocker', 'Laps', 'Noncompliant', 'Inactive', 'DisabledPrimaryUsers', 'MissingAutopilotDevices', 'MultiUserDevices', 'Deprecated') {
                $cmd = (Get-Command "Resolve-Intune*$f*" | Select-Object -First 1).Name
                $out[$f] = @((& $cmd -Data $d).Warnings.Message)
            }
            $out
        }
        foreach ($f in 'BitLocker', 'Laps', 'Noncompliant', 'Inactive', 'DisabledPrimaryUsers', 'MissingAutopilotDevices', 'MultiUserDevices') { $r[$f] | Should -Be @('Device X could not be processed') -Because $f }
        foreach ($f in 'Deprecated') { $r[$f] | Should -BeNullOrEmpty -Because $f }
    }
}

Describe 'Dropped connections' {
    It 'Invoke-RKCoreGraphRequest retries a failure without a status (dropped connection) once' {
        InModuleScope RKSolutions {
            $script:calls = 0
            Mock Start-Sleep { }
            Mock Invoke-MgGraphRequest {
                $script:calls++
                if ($script:calls -eq 1) { throw [System.Exception]::new('connection reset') }
                [pscustomobject]@{ value = @(1) }
            }
            @(Invoke-RKCoreGraphRequest -Uri 'beta/x' -All) | Should -Be @(1)
            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly
        }
    }
}

Describe 'Get-RKIntuneAnomaly' {
    BeforeAll {
        $script:anomalyData = {
            param($Keys = $true, $Laps = $true)
            $policy = [pscustomobject]@{ DisplayName = 'P'; Source = 'Catalog'; Assignments = @(
                    [pscustomobject]@{ AssignmentType = 'All Devices' }
                    [pscustomobject]@{ AssignmentType = 'Group (Exclude)'; GroupId = 'g1' }) }
            $ctx = [pscustomobject]@{
                EntraDeviceIdByName = @{}; Filters = @{}; GroupUserUpns = @{}
                GroupDeviceAadIds = @{ g1 = [System.Collections.Generic.HashSet[string]]::new([string[]]@('aad1')) }
                BitLockerPolicies = @($policy); BitLockerKeyDeviceIds = [System.Collections.Generic.HashSet[string]]::new()
                LapsPolicies = @($policy); LapsCredentialByDeviceId = @{}; LapsCredentialByDeviceName = @{}
                ConfigurationPolicies = @()
            }
            [pscustomobject]@{
                TenantName = 'T'; CollectedAt = [datetime]'2026-10-25'
                Devices = @([pscustomobject]@{ Customer = 'T'; DeviceName = 'E1'; AzureAdDeviceId = 'aad1'; Encrypted = $false; PrimaryUser = 'u@t'; Serialnumber = 'S'; IsCloudPC = $false; LastContact = Get-Date; DeviceHashUploaded = $true; DeviceAssociated = $false; usersLoggedOnCount = 1; OperatingSystemEdition = 'Pro'; OSFriendlyname = 'W11' })
                Users = @(); AutopilotProfiles = @(); SecurityKeyContext = $ctx
                Warnings = @([pscustomobject]@{ Area = 'BitLocker'; Message = 'same' }, [pscustomobject]@{ Area = 'BitLocker'; Message = 'same' })
                BitLockerKeysReadable = $Keys; LapsCredentialsReadable = $Laps
            }
        }
    }
    BeforeEach {
        Mock -ModuleName RKSolutions Get-MgContext { [pscustomobject]@{ TenantId = 't' } }
        Mock -ModuleName RKSolutions Get-RKIntuneAnomalyData { & $script:anomalyData $false $true }
        Mock -ModuleName RKSolutions Get-ApplicationFailures { }
    }
    It 'returns Info rows only with -ShowExcludedDevices' {
        @((Get-RKIntuneAnomaly).BitLocker) | Should -BeNullOrEmpty
        @((Get-RKIntuneAnomaly).Laps) | Should -BeNullOrEmpty
        @((Get-RKIntuneAnomaly -ShowExcludedDevices).BitLocker).Severity | Should -Be 'Info'
        @((Get-RKIntuneAnomaly -ShowExcludedDevices).Laps).Severity | Should -Be 'Info'
    }
    It 'returns an empty category as an empty array, not $null' {
        $a = Get-RKIntuneAnomaly
        foreach ($p in 'Noncompliant', 'MissingAutopilot', 'MultiUser', 'AppInstallFailures', 'DeprecatedSettings', 'DisabledPrimaryUser') {
            ($null -ne $a.$p) | Should -BeTrue -Because $p
            $a.$p -is [array] | Should -BeTrue -Because $p
            $a.$p.Count | Should -Be 0
        }
    }
    It 'removes duplicate warnings' {
        @((Get-RKIntuneAnomaly).Warnings).Count | Should -Be 1
    }
    It 'passes the Readable flags through' {
        $a = Get-RKIntuneAnomaly
        $a.BitLockerKeysReadable | Should -BeFalse
        $a.LapsCredentialsReadable | Should -BeTrue
        $a.AppsReadable | Should -BeTrue
    }
    It 'takes AppInstallFailures from Get-ApplicationFailures' {
        Mock -ModuleName RKSolutions Get-ApplicationFailures { [pscustomobject]@{ Application = 'A'; FailedDeviceCount = 3 } }
        $a = Get-RKIntuneAnomaly
        @($a.AppInstallFailures).Application | Should -Be @('A')
        Should -Invoke Get-ApplicationFailures -ModuleName RKSolutions -ParameterFilter { $TenantName -eq 'T' }
    }
    It 'gives AppsReadable $false and an empty array when the app report answers 403' {
        Mock -ModuleName RKSolutions Get-ApplicationFailures {
            $e = [System.Exception]::new('http 403')
            $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 403 })
            throw $e
        }
        $a = Get-RKIntuneAnomaly
        $a.AppsReadable | Should -BeFalse
        $a.AppInstallFailures -is [array] | Should -BeTrue
        $a.AppInstallFailures.Count | Should -Be 0
    }
    It 'throws when the app report fails with another error' {
        Mock -ModuleName RKSolutions Get-ApplicationFailures { throw 'http 500' }
        { Get-RKIntuneAnomaly } | Should -Throw '*http 500*'
    }
    It 'throws without a Graph context' {
        Mock -ModuleName RKSolutions Get-MgContext { $null }
        { Get-RKIntuneAnomaly } | Should -Throw '*Connect-RKGraph*'
    }
}

Describe 'Cached policies and failed batch requests' {
    It 'Get-BitLockerLapsAssignmentContext leaves the policy objects the adapter returned untouched' {
        InModuleScope RKSolutions {
            $script:cached = [pscustomobject]@{ id = 'p1'; name = 'BitLocker' }
            Mock Invoke-RKCoreGraphRequest { }
            Mock Get-DetailedPolicyAssignments { }
            Mock Invoke-RKCoreGraphRequest -ParameterFilter { $Uri -eq 'beta/deviceManagement/configurationPolicies' } { $script:cached }
            Mock Invoke-RKCoreGraphBatch -ParameterFilter { @($Requests.Url) -like '*/settings?*' } {
                $s = [pscustomobject]@{ settingInstance = [pscustomobject]@{ settingDefinitionId = 'device_vendor_msft_bitlocker_requiredeviceencryption' } }
                [pscustomobject]@{ Id = 'cps:p1'; Status = 200; Body = [pscustomobject]@{ value = @($s) } }
            }
            $w = [System.Collections.Generic.List[object]]::new()
            $ctx = Get-BitLockerLapsAssignmentContext -Warnings $w
            $script:cached.PSObject.Properties['settings'] | Should -BeNullOrEmpty
            @($ctx.ConfigurationPolicies[0].settings).Count | Should -Be 1
            @($ctx.BitLockerPolicies.Id) | Should -Be @('p1')
        }
    }
    It 'Invoke-RKCoreGraphBatch carries the error text of a failed request into Body.error.message' {
        InModuleScope RKSolutions {
            Mock Invoke-RKGraphBatch { [pscustomobject]@{ Id = 'a'; Status = 0; Body = $null; Headers = $null; Error = 'connection reset' } }
            $r = @(Invoke-RKCoreGraphBatch -Requests @([pscustomobject]@{ Id = 'a'; Url = 'beta/x' }))
            $r[0].Body.error.message | Should -Be 'connection reset'
        }
    }
}
