# Private: Windows apps with failed installs, from the Intune app install status report (top 50 by failed devices).
function Get-ApplicationFailures {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [AllowEmptyString()] [string] $TenantName)

    # Cross-platform temporary file path
    # Detect OS and set appropriate temp path
    $detectedWindows = $false

    # Check if automatic variables exist (PowerShell Core 6.0+)
    if (Get-Variable -Name "IsWindows" -ErrorAction SilentlyContinue) {
        $detectedWindows = $IsWindows
    }
    # Fallback for older PowerShell versions
    else {
        $osInfo = [System.Environment]::OSVersion.Platform
        switch ($osInfo) {
            "Win32NT" { $detectedWindows = $true }
            "Unix" {
                # macOS/Linux detected but not used in this context
            }
            default {
                try {
                    if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
                        $detectedWindows = $true
                    }
                    else {
                        # Non-Windows OS detected
                    }
                } catch {
                    $detectedWindows = $true
                }
            }
        }
    }

    $apps = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=(isof(%27microsoft.graph.win32CatalogApp%27)%20or%20isof(%27microsoft.graph.windowsStoreApp%27)%20or%20isof(%27microsoft.graph.microsoftStoreForBusinessApp%27)%20or%20isof(%27microsoft.graph.officeSuiteApp%27)%20or%20(isof(%27microsoft.graph.win32LobApp%27)%20and%20not(isof(%27microsoft.graph.win32CatalogApp%27)))%20or%20isof(%27microsoft.graph.windowsMicrosoftEdgeApp%27)%20or%20isof(%27microsoft.graph.windowsPhone81AppX%27)%20or%20isof(%27microsoft.graph.windowsPhone81StoreApp%27)%20or%20isof(%27microsoft.graph.windowsPhoneXAP%27)%20or%20isof(%27microsoft.graph.windowsAppX%27)%20or%20isof(%27microsoft.graph.windowsMobileMSI%27)%20or%20isof(%27microsoft.graph.windowsUniversalAppX%27)%20or%20isof(%27microsoft.graph.webApp%27)%20or%20isof(%27microsoft.graph.windowsWebApp%27)%20or%20isof(%27microsoft.graph.winGetApp%27))%20and%20(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&`$orderby=displayName&").value

    # Use a unique temporary file to avoid race conditions. Created after the app list, so a missing permission
    # (which fails that first call) leaves no file behind.
    $Data = [System.IO.Path]::GetTempFileName()

    $params = @{
        Select  = @(
            "DisplayName"
            "Publisher"
            "Platform"
            "AppVersion"
            "FailedDevicePercentage"
            "FailedDeviceCount"
            "FailedUserCount"
            "ApplicationId"
        )
        Skip    = 0
        Top     = 50
        Filter  = "(FailedDeviceCount gt '0')"
        OrderBy = @(
            "FailedDeviceCount desc"
        )
    }
    Invoke-MgGraphRequest -Body $params -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/getAppsInstallSummaryReport" -Method POST -OutputFilePath $Data

    $DataFile = Get-Content $Data
    # Fix char encoding to UTF-8
    $Response = [system.Text.Encoding]::UTF8.GetString(($DataFile).ToCharArray()) | ConvertFrom-Json

    # Pre-build app lookup (id -> app) so the per-row joins are O(1) instead of O(N) Where-Object scans.
    $appLookup = @{}
    foreach ($a in $apps) { if ($a.Id) { $appLookup[[string]$a.Id] = $a } }

    # First pass: materialise the schema-flattened rows that have a matching app,
    # collecting unique ApplicationIds so we can batch-fetch their assignments.
    $rows = [System.Collections.Generic.List[hashtable]]::new()
    $expandIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($value in $Response.Values) {
        $LineObject = @{ }
        foreach ($prop in $Response.Schema) {
            $LineObject[$prop.Column] = $value[$Response.Schema.IndexOf($prop)]
        }
        $appId = [string]$LineObject.ApplicationId
        if ($appLookup.ContainsKey($appId)) {
            $rows.Add($LineObject)
            [void]$expandIds.Add($appId)
        }
    }

    # Batch-fetch ?$expand=assignments for every failed app in parallel sub-requests.
    $assignmentStatusByAppId = @{}
    if ($expandIds.Count -gt 0) {
        $expandRequests = foreach ($id in $expandIds) {
            [PSCustomObject]@{
                Id  = "app:$id"
                Url = "/deviceAppManagement/mobileApps/$id/?`$expand=assignments"
            }
        }
        $expandResponses = Invoke-RKGraphBatch -Requests @($expandRequests) -Activity 'App assignments'
        foreach ($resp in $expandResponses) {
            if ($resp.Status -ne 200 -or -not $resp.Body) { continue }
            $appId = $resp.Id -replace '^app:', ''
            $hasAssignments = $resp.Body.assignments -and (@($resp.Body.assignments).Count -gt 0)
            $assignmentStatusByAppId[$appId] = [bool]$hasAssignments
        }
    }

    # Build result array from the prepared rows using O(1) lookups.
    $ReturnObject = New-Object System.Collections.ArrayList
    foreach ($LineObject in $rows) {
        $appId = [string]$LineObject.ApplicationId
        $PlatformName = if ($LineObject.Platform_loc) { $LineObject.Platform_loc } else { $LineObject.Platform }
        $AssignmentStatus = $false
        if ($assignmentStatusByAppId.ContainsKey($appId)) { $AssignmentStatus = $assignmentStatusByAppId[$appId] }

        $ReturnObject.Add([PSCustomObject][ordered]@{
                Customer               = $TenantName
                Application            = $appLookup[$appId].displayName
                Platform               = $PlatformName
                Version                = $LineObject.AppVersion
                AssignmentStatus       = $AssignmentStatus
                FailedUserCount        = $LineObject.FailedUserCount
                FailedDeviceCount      = $LineObject.FailedDeviceCount
                FailedDevicePercentage = [math]::Round([double]$LineObject.FailedDevicePercentage, 2)
            }) | Out-Null
    }

    # Clean up temporary data file
    if (Test-Path -Path $Data) {
        Remove-Item -Path $Data -Force
    }

    return $ReturnObject | Sort-Object -Property FailedDeviceCount -Descending
}
