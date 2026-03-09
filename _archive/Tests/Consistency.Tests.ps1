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
