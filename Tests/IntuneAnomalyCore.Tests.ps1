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
