# Consistency.Tests.ps1
# Validates that the script module can be imported and that expected cmdlets are exported.
# Run from repo root: Invoke-Pester ./Tests/Consistency.Tests.ps1
# These tests do NOT connect to Microsoft Graph so they pass in CI (GitHub Actions).

$ErrorActionPreference = 'Stop'

# Resolve path to RKSolutions.psd1. Defined in global scope so Pester BeforeAll blocks can call it.
function global:Get-RKSolutionsManifestPath {
    if ($script:manifestPathCache -and (Test-Path -LiteralPath $script:manifestPathCache)) { return $script:manifestPathCache }
    $here = $PSScriptRoot
    if (-not $here -and $PSCommandPath) { $here = Split-Path -Parent $PSCommandPath }
    if ($here) {
        $tryPath = Join-Path (Join-Path (Join-Path $here '..') 'module') 'RKSolutions.psd1'
        if (Test-Path -LiteralPath $tryPath) { $script:manifestPathCache = $tryPath; return $tryPath }
    }
    $root = Get-Location
    $tryPath = Join-Path (Join-Path $root 'module') 'RKSolutions.psd1'
    if (-not (Test-Path -LiteralPath $tryPath)) {
        throw "Module manifest not found. Run from repo root. Tried: $tryPath"
    }
    $script:manifestPathCache = $tryPath
    $tryPath
}

Describe 'Module import and exports' {
    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        $script:manifestPath = Get-RKSolutionsManifestPath
        $script:manifestData = Import-PowerShellDataFile -Path $script:manifestPath
        $script:expectedCmdlets = @($script:manifestData.FunctionsToExport)
    }

    It 'Can import the module without error' {
        { Import-Module $script:manifestPath -Force -ErrorAction Stop } | Should -Not -Throw
        $m = Get-Module -Name 'RKSolutions'
        $m | Should -Not -BeNullOrEmpty -Because 'the module should be loaded after Import-Module'
    }

    It 'Exports all cmdlets listed in the manifest (FunctionsToExport)' {
        Import-Module $script:manifestPath -Force -ErrorAction Stop
        $exported = @((Get-Module -Name 'RKSolutions').ExportedCommands.Keys)
        foreach ($name in $script:expectedCmdlets) {
            $exported | Should -Contain $name -Because "cmdlet '$name' is in FunctionsToExport but was not exported"
        }
        $exported.Count | Should -Be $script:expectedCmdlets.Count -Because 'exported count should match FunctionsToExport count'
    }
}

Describe 'Consistency contract' {

    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        $path = Get-RKSolutionsManifestPath
        $manifestData = Import-PowerShellDataFile -Path $path
        $script:expectedNames = @($manifestData.FunctionsToExport)
        $script:expectedCount = $script:expectedNames.Count
        Import-Module $path -Force
        $script:exported = (Get-Module -Name 'RKSolutions').ExportedCommands.Keys
        # Key parameters (subset) for a few cmdlets to validate binding
        $script:expectedParameters = @{
            'Connect-RKGraph'                 = @('RequiredScopes', 'TenantId', 'ClientId')
            'Get-IntuneEnrollmentFlowsReport' = @('AssignmentOverviewOnly', 'OutputPath')
            'Get-IntuneAnomaliesReport'       = @('ExportPath')
            'Get-EntraAdminRolesReport'       = @('ExportPath')
            'Get-M365LicenseAssignmentReport' = @('ExportPath')
            'Get-CustomSecurityAttributesReport' = @('AttributeSet', 'ExportPath')
        }
    }

    It 'Each cmdlet has expected parameters (subset check) where defined' {
        foreach ($name in $script:expectedNames) {
            $expectedParams = $script:expectedParameters[$name]
            if ($null -eq $expectedParams -or $expectedParams.Count -eq 0) { continue }
            $cmd = Get-Command -Name $name -ErrorAction Stop
            $paramNames = $cmd.Parameters.Keys
            foreach ($p in $expectedParams) {
                $paramNames | Should -Contain $p
            }
        }
    }

    It 'Get-Help is filled for every exported cmdlet' {
        foreach ($name in $script:expectedNames) {
            $help = Get-Help -Name $name -ErrorAction Stop
            $help | Should -Not -BeNullOrEmpty -Because "Get-Help $name should return help"
            $help.Synopsis | Should -Not -BeNullOrEmpty -Because "cmdlet $name must have .SYNOPSIS filled"
        }
    }
}

Describe 'No-silent-failure contract' {
    # Only tests that do NOT trigger Graph connection (safe for CI).
    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        Import-Module (Get-RKSolutionsManifestPath) -Force
    }

    It 'Disconnect-RKGraph runs when not connected' {
        { Disconnect-RKGraph } | Should -Not -Throw
    }
}

Describe 'Public report cmdlet shape contracts' {
    # AST-level checks: every public report cmdlet must follow the standard shape
    # (CmdletBinding, the five contract parameters, a connection check).

    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        Import-Module (Get-RKSolutionsManifestPath) -Force
        $script:reportCmdlets = @(
            'Get-IntuneEnrollmentFlowsReport',
            'Get-IntuneAnomaliesReport',
            'Get-EntraAdminRolesReport',
            'Get-M365LicenseAssignmentReport',
            'Get-CustomSecurityAttributesReport'
        )
    }

    It 'Every report cmdlet declares an output-path parameter (ExportPath or OutputPath)' {
        # Reports don't share a uniform Send/From/Recipient surface (Get-IntuneEnrollmentFlowsReport
        # uses its own ExportToCsv/ExportFolder model). What IS uniform: every report must let the
        # caller direct the rendered HTML somewhere, via either -ExportPath or -OutputPath.
        foreach ($name in $script:reportCmdlets) {
            $cmd = Get-Command -Name $name -ErrorAction Stop
            $paramNames = $cmd.Parameters.Keys
            $hasOutputParam = ($paramNames -contains 'ExportPath') -or ($paramNames -contains 'OutputPath')
            $hasOutputParam | Should -BeTrue -Because "$name must let the caller direct the output via -ExportPath or -OutputPath"
        }
    }

    It 'Every report cmdlet uses [CmdletBinding()] so -Verbose works' {
        foreach ($name in $script:reportCmdlets) {
            $cmd = Get-Command -Name $name -ErrorAction Stop
            $cmd.CmdletBinding | Should -BeTrue -Because "$name must declare [CmdletBinding()] to honor -Verbose"
        }
    }

    It 'Every report cmdlet checks Get-MgContext before doing work' {
        # Source-level check: each public cmdlet file must reference Get-MgContext (the connection guard).
        $publicDir = Join-Path (Split-Path (Get-RKSolutionsManifestPath) -Parent) 'Public'
        foreach ($name in $script:reportCmdlets) {
            $file = Join-Path $publicDir "$name.ps1"
            (Test-Path $file) | Should -BeTrue -Because "expected $file to exist"
            $content = Get-Content -Raw -Path $file
            $content | Should -Match 'Get-MgContext' -Because "$name must guard with Get-MgContext before any Graph call"
        }
    }
}

Describe 'Compliance setting friendly names' {
    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        Import-Module (Get-RKSolutionsManifestPath) -Force
    }

    It 'No friendly name contains ", " (the Noncompliant tab splits reasons on it)' {
        $bad = InModuleScope RKSolutions { $script:ComplianceSettingFriendlyNames.Values | Where-Object { $_ -like '*, *' } }
        $bad | Should -BeNullOrEmpty
    }

    It 'Translates known rules and passes unknown ones through' {
        InModuleScope RKSolutions {
            ConvertTo-ComplianceSettingFriendlyName 'Windows10CompliancePolicy.BitLockerEnabled' | Should -Be 'Require BitLocker'
            ConvertTo-ComplianceSettingFriendlyName './Vendor/MSFT/Custom/Thing' | Should -Be './Vendor/MSFT/Custom/Thing'
        }
    }
}

Describe 'BitLocker / LAPS anomaly resolver output contracts' {
    # The HTML template emits <td> cells for each property in a specific order;
    # if these resolvers ever stop emitting one of these columns, the rendered
    # report will silently lose a value. These tests run with synthetic input
    # (no Microsoft Graph) and assert the exact column set + order.

    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        Import-Module (Get-RKSolutionsManifestPath) -Force
    }

    It 'Resolve-IntuneBitLockerAnomalies emits the documented column set in order' {
        $row = InModuleScope RKSolutions {
            $ctx = [PSCustomObject]@{
                BitLockerPolicies     = [System.Collections.Generic.List[object]]::new()
                LapsPolicies          = [System.Collections.Generic.List[object]]::new()
                BitLockerKeyDeviceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                GroupDeviceAadIds     = @{}
                GroupUserUpns         = @{}
                GroupUserIds          = @{}
                Filters               = @{}
                EntraDeviceIdByName   = @{}
            }
            $d = [PSCustomObject]@{
                DeviceName         = 'TEST-DEVICE-1'
                AzureAdDeviceId    = $null   # Force the "no aadId, not encrypted" branch -> produces a row
                Encrypted          = $false
                PrimaryUser        = 'user@test.local'
                Serialnumber       = 'TEST-SN-1'
                DeviceManufacturer = 'TestCorp'
                DeviceModel        = 'TestModel'
            }
            $rows = Resolve-IntuneBitLockerAnomalies -Devices @($d) -Context $ctx -TenantName 'TestTenant'
            $rows[0]
        }
        $row | Should -Not -BeNullOrEmpty -Because 'an unencrypted device with no aadId must surface as an anomaly row'
        $actualColumns = @($row.PSObject.Properties.Name)
        $expectedColumns = @('Customer','DeviceName','PrimaryUser','Serialnumber','IsEncrypted','AppliedPolicies','KeyEscrowed','Severity','Status')
        $actualColumns | Should -Be $expectedColumns -Because 'BitLocker anomaly rows must match the HTML template <td> order (Severity column comes before Status)'
    }

    It 'Resolve-IntuneLapsAnomalies emits the documented column set in order' {
        $row = InModuleScope RKSolutions {
            # Create a LAPS policy that matches via "All Devices" so the synthetic device hits the
            # "policy assigned but no credential" branch (which produces an anomaly row).
            $policy = [PSCustomObject]@{
                Id              = 'fake-policy-id'
                DisplayName     = 'Fake LAPS Policy'
                Source          = 'SettingsCatalog'
                BackupDirectory = 1
                Assignments     = @(
                    [PSCustomObject]@{
                        AssignmentType = 'All Devices'
                        GroupId        = $null
                        FilterId       = $null
                        FilterType     = 'None'
                    }
                )
            }
            $lapsPolicies = [System.Collections.Generic.List[object]]::new()
            $lapsPolicies.Add($policy)
            $ctx = [PSCustomObject]@{
                BitLockerPolicies          = [System.Collections.Generic.List[object]]::new()
                LapsPolicies               = $lapsPolicies
                BitLockerKeyDeviceIds      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                GroupDeviceAadIds          = @{}
                GroupUserUpns              = @{}
                GroupUserIds               = @{}
                Filters                    = @{}
                EntraDeviceIdByName        = @{}
                LapsCredentialByDeviceId   = @{}
                LapsCredentialByDeviceName = @{}
            }
            # Use the REAL field names produced by Get-AllDeviceData so a typo'd
            # resolver lookup (e.g. $d.OwnerType when the device exposes DeviceOwnership)
            # is caught by this test instead of being silently hidden behind a synthetic
            # field. See FINDINGS: "synthetic-fixture-hides-field-name-bug".
            $d = [PSCustomObject]@{
                DeviceName         = 'TEST-DEVICE-1'
                AzureAdDeviceId    = '11111111-2222-3333-4444-555555555555'
                Encrypted          = $true
                PrimaryUser        = 'user@test.local'
                Serialnumber       = 'TEST-SN-1'
                DeviceManufacturer = 'TestCorp'
                DeviceModel        = 'TestModel'
                DeviceOwnership    = 'company'   # matches Get-AllDeviceData's output property
            }
            $rows = Resolve-IntuneLapsAnomalies -Devices @($d) -Context $ctx -TenantName 'TestTenant'
            $rows[0]
        }
        $row | Should -Not -BeNullOrEmpty -Because 'a device covered by a LAPS policy with no credential must surface'
        $actualColumns = @($row.PSObject.Properties.Name)
        $expectedColumns = @('Customer','DeviceName','PrimaryUser','Serialnumber','OwnerType','AppliedPolicies','LastBackupDateTime','BackupAgeDays','Severity','Status')
        $actualColumns | Should -Be $expectedColumns -Because 'LAPS anomaly rows must match the HTML template <td> order (Severity column comes before Status)'
        $row.OwnerType | Should -Be 'company' -Because 'the LAPS resolver must read the device ownership from the same field Get-AllDeviceData emits (DeviceOwnership), otherwise the Ownership column will be blank on every real device'
    }
}

Describe 'Connect-RKGraph default scopes match docs/PERMISSIONS.md' {
    # Detects drift between Connect-RKGraph's hard-coded default -RequiredScopes
    # and the documented permissions table. Either both are right, or both are wrong,
    # but they must not silently diverge.

    BeforeAll {
        $modulePath = Split-Path (Get-RKSolutionsManifestPath) -Parent
        $repoRoot   = Split-Path $modulePath -Parent
        $script:connectSource = Get-Content -Raw -Path (Join-Path $modulePath 'Public/Connect-RKGraph.ps1')
        $script:permissionsDoc = Get-Content -Raw -Path (Join-Path $repoRoot 'docs/PERMISSIONS.md')
    }

    It 'Every scope mentioned in Connect-RKGraph appears in docs/PERMISSIONS.md' {
        $scopeMatches = [regex]::Matches($script:connectSource, "'([A-Za-z0-9]+\.[A-Za-z]+(?:\.[A-Za-z]+)?(?:\.All)?)'")
        $scopes = @($scopeMatches | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique | Where-Object { $_ -match '\.' -and $_ -notmatch '^Microsoft\.' })
        $scopes.Count | Should -BeGreaterThan 0 -Because 'Connect-RKGraph should declare some default Graph scopes'
        foreach ($scope in $scopes) {
            $script:permissionsDoc | Should -Match ([regex]::Escape($scope)) -Because "scope '$scope' is granted by Connect-RKGraph but is not documented in docs/PERMISSIONS.md"
        }
    }
}
