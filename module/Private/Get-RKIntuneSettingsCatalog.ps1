# Private: load the Intune Settings Catalog data into an in-memory lookup keyed by
# settingDefinitionId. Used by IntuneAnomalies' deprecation walker to resolve a
# settingDefinitionId to its Microsoft DisplayName (Microsoft flags retired
# settings in the display name itself, e.g. "...(deprecated)").
#
# Source: https://github.com/royklo/IntuneSettingsCatalogData
# Cache : $LOCALAPPDATA/RKSolutions/settings-catalog.json on Windows,
#         ~/.rksolutions/settings-catalog.json elsewhere
# TTL   : 24 hours (based on file LastWriteTime; no separate metadata file)
#
# On any failure (offline, parse error, missing release) the function returns an
# empty hashtable. Callers must handle the empty case gracefully - the
# deprecation regex check still works against raw setting name/value strings, so
# detection only degrades to "less complete" not "broken".

function Get-RKIntuneSettingsCatalog {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $false)]
        [switch] $Force
    )

    if (-not $Force -and $null -ne $script:RKIntuneSettingsCatalog) {
        return $script:RKIntuneSettingsCatalog
    }

    # Per-OS cache directory. Falls back to $HOME on platforms without LOCALAPPDATA.
    $cacheDir = if ($IsWindows -and $env:LOCALAPPDATA) {
        Join-Path $env:LOCALAPPDATA 'RKSolutions'
    } else {
        Join-Path $HOME '.rksolutions'
    }
    $cacheFile = Join-Path $cacheDir 'settings-catalog.json'
    $ttl = New-TimeSpan -Hours 24
    $sourceUrl = 'https://github.com/royklo/IntuneSettingsCatalogData/releases/latest/download/settings.json'

    $needDownload = $true
    if (Test-Path -LiteralPath $cacheFile) {
        $age = (Get-Date) - (Get-Item -LiteralPath $cacheFile).LastWriteTime
        if ($age -lt $ttl) {
            $needDownload = $false
            Write-Verbose ("Settings catalog cache is fresh (age {0:N1}h)" -f $age.TotalHours)
        }
    }

    if ($needDownload) {
        if (-not (Test-Path -LiteralPath $cacheDir)) {
            New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
        }
        Write-Host '  Refreshing Intune Settings Catalog data from GitHub...' -ForegroundColor Cyan
        $savedProgress = $ProgressPreference
        try {
            $ProgressPreference = 'SilentlyContinue'
            $tmpFile = "$cacheFile.tmp"
            Invoke-WebRequest -Uri $sourceUrl -OutFile $tmpFile -UseBasicParsing -ErrorAction Stop
            Move-Item -Path $tmpFile -Destination $cacheFile -Force
        }
        catch {
            Remove-Item -Path "$cacheFile.tmp" -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $cacheFile) {
                Write-Warning "Could not refresh Intune Settings Catalog data ($($_.Exception.Message)); using stale cache."
            } else {
                Write-Warning "Could not download Intune Settings Catalog data ($($_.Exception.Message)). Deprecation detection will fall back to name/value regex only."
                $script:RKIntuneSettingsCatalog = @{}
                return $script:RKIntuneSettingsCatalog
            }
        }
        finally {
            $ProgressPreference = $savedProgress
        }
    }

    try {
        $raw = Get-Content -Path $cacheFile -Raw -Encoding UTF8
        $entries = $raw | ConvertFrom-Json -AsHashtable -Depth 100
        $lookup = @{}
        foreach ($e in $entries) {
            $id = $e['id']
            if ([string]::IsNullOrEmpty($id)) { continue }
            $lookup[$id] = @{
                DisplayName = $e['displayName']
                Description = $e['description']
            }
        }
        $script:RKIntuneSettingsCatalog = $lookup
        Write-Host "    Loaded $($lookup.Count) catalog entries." -ForegroundColor Gray
    }
    catch {
        Write-Warning "Failed to parse Intune Settings Catalog data: $($_.Exception.Message)"
        $script:RKIntuneSettingsCatalog = @{}
    }

    return $script:RKIntuneSettingsCatalog
}
