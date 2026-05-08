# RKSolutions.psm1 - Dot-source Private then Public; export only public functions
# Private scripts (e.g. IntuneEnrollmentFlows.ps1) may define helper functions (e.g. Get-DeviceEvaluationContext,
# Get-CloudPCProvisioningPolicyGroupInfo) that are used internally by Public cmdlets but are NOT exported as cmdlets.

# Require PowerShell 7.0 or higher
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host ""
    Write-Host "RKSolutions module requires PowerShell 7.0 or higher." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    Write-Host ""
    Write-Host "You are running Windows PowerShell (powershell.exe)." -ForegroundColor Yellow
    Write-Host "Please use PowerShell 7 (pwsh.exe) instead." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Install PowerShell 7: " -ForegroundColor White -NoNewline
    Write-Host "https://aka.ms/powershell" -ForegroundColor Cyan
    Write-Host "After installing, run: " -ForegroundColor White -NoNewline
    Write-Host "pwsh" -ForegroundColor Green
    Write-Host ""
    exit 1
}

$moduleRoot = $PSScriptRoot

# Load Private scripts first (shared helpers, then report-specific)
$privatePath = Join-Path $moduleRoot 'Private'
if (Test-Path $privatePath) {
    $sharedOrder = @(
        'Install-Requirements.ps1',
        'Export-Results.ps1',
        'Connect-ToMgGraph.ps1',
        'Invoke-RKSolutionsWithConnection.ps1',
        'Invoke-GraphRequestWithPaging.ps1',
        'Send-EmailWithAttachment.ps1',
        'ConvertTo-DateString.ps1',
        'Get-RKSolutionsReportTemplate.ps1'
    )
    foreach ($name in $sharedOrder) {
        $fp = Join-Path $privatePath $name
        if (Test-Path $fp) { . $fp }
    }
    # Report-specific private scripts (order matters if they depend on each other)
    $domainOrder = @('IntuneEnrollmentFlows.ps1', 'IntuneAnomalies.ps1', 'EntraAdminRoles.ps1', 'M365License.ps1', 'CustomSecurityAttributes.ps1')
    foreach ($name in $domainOrder) {
        $fp = Join-Path $privatePath $name
        if (Test-Path $fp) { . $fp }
    }
}

# Load Public scripts (exported cmdlets)
$publicPath = Join-Path $moduleRoot 'Public'
if (Test-Path $publicPath) {
    Get-ChildItem -Path $publicPath -Filter '*.ps1' -File | ForEach-Object { . $_.FullName }
}

# Export public cmdlets (Connect, Disconnect, report cmdlets only; helpers stay private)
Export-ModuleMember -Function @(
    'Connect-RKGraph',
    'Disconnect-RKGraph',
    'Get-IntuneEnrollmentFlowsReport',
    'Get-IntuneAnomaliesReport',
    'Get-EntraAdminRolesReport',
    'Get-M365LicenseAssignmentReport',
    'Get-CustomSecurityAttributesReport'
)
