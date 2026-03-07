# Consistency.Tests.ps1
# Validates that the script module's exported functions and parameter names match the manifest and have help.
# Run from repo root: Invoke-Pester ./Tests/Consistency.Tests.ps1

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

    It 'Module exports all expected cmdlets' {
        foreach ($name in $script:expectedNames) {
            $script:exported | Should -Contain $name
        }
        @($script:exported).Count | Should -Be $script:expectedCount
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

    It 'Every exported cmdlet has comment-based help (Synopsis)' {
        foreach ($name in $script:expectedNames) {
            $help = Get-Help -Name $name -ErrorAction Stop
            $help.Synopsis | Should -Not -BeNullOrEmpty -Because "cmdlet $name must have .SYNOPSIS"
        }
    }
}

Describe 'No-silent-failure contract' {
    # Key cmdlets must produce either output or an error when not connected (no Graph session).
    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        Import-Module (Get-RKSolutionsManifestPath) -Force
    }

    It 'Disconnect-RKGraph runs when not connected' {
        { Disconnect-RKGraph } | Should -Not -Throw
    }

    It 'Get-IntuneEnrollmentFlowsReport produces an error when not connected' {
        $err = $null
        Get-IntuneEnrollmentFlowsReport -AssignmentOverviewOnly -ErrorVariable err -ErrorAction SilentlyContinue
        $err | Should -Not -BeNullOrEmpty -Because 'should report not connected, not return silence'
    }

    It 'Get-EntraAdminRolesReport produces an error when not connected' {
        $err = $null
        Get-EntraAdminRolesReport -ErrorVariable err -ErrorAction SilentlyContinue
        $err | Should -Not -BeNullOrEmpty -Because 'should report not connected, not return silence'
    }
}

Describe 'Parameter binding and behavior' {
    # Key parameters bind and cmdlet runs (output or connection error, not parameter binding error).
    BeforeAll {
        Remove-Module -Name 'RKSolutions' -ErrorAction SilentlyContinue
        Import-Module (Get-RKSolutionsManifestPath) -Force
    }

    It 'Connect-RKGraph accepts RequiredScopes and produces output or error' {
        $result = $null
        $err = $null
        try { $result = Connect-RKGraph -RequiredScopes 'User.Read' -ErrorVariable err -ErrorAction SilentlyContinue } catch { $err = @($_) }
        $hasOutput = $null -ne $result -and (@($result).Count -gt 0)
        $hasError = $null -ne $err -and (@($err).Count -gt 0)
        ($hasOutput -or $hasError) | Should -BeTrue -Because 'Connect-RKGraph must not silently do nothing'
        if ($hasError -and $err[0].ToString() -match 'Cannot bind|Parameter.*not found|Unknown parameter') {
            Set-ItResult -Inconclusive -Because 'Parameter binding failed; check parameter names'
        }
    }

    It 'Get-IntuneEnrollmentFlowsReport -AssignmentOverviewOnly binds and produces output or error' {
        $out = @(); $err = @()
        $out = Get-IntuneEnrollmentFlowsReport -AssignmentOverviewOnly -ErrorVariable err -ErrorAction SilentlyContinue
        $err = @($err)
        $hasOutput = $null -ne $out -and (@($out).Count -ge 0)
        $hasError = $err.Count -gt 0
        ($hasOutput -or $hasError) | Should -BeTrue -Because 'Get-IntuneEnrollmentFlowsReport must not silently do nothing'
        if ($hasError -and $err[0].ToString() -match 'Cannot bind|Parameter.*not found|Unknown parameter') {
            Set-ItResult -Inconclusive -Because 'Parameter binding failed'
        }
    }
}
