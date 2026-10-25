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
    It 'Invoke-RKCoreGraphRequest without -All returns the raw body' {
        InModuleScope RKSolutions {
            Mock Invoke-MgGraphRequest { [pscustomobject]@{ value = @(1); '@odata.nextLink' = 'n' } }
            (Invoke-RKCoreGraphRequest -Uri 'beta/x').'@odata.nextLink' | Should -Be 'n'
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
