# Private: Unified export to file (csv, html, pdf, json, xml, txt) with cross-platform paths
function Export-Results {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [ValidateNotNull()] $Results,
        [Parameter(Mandatory = $false)] [string] $FileName = 'Report',
        [Parameter(Mandatory = $true)] [ValidateSet('csv', 'html', 'pdf', 'json', 'xml', 'txt')] [string] $Extension,
        [Parameter(Mandatory = $false)] [string] $OutputFolder = '',
        [Parameter(Mandatory = $false)] [bool] $IncludeTimestamp = $true,
        [Parameter(Mandatory = $false)] [switch] $DebugMode,
        [Parameter(Mandatory = $false)] [switch] $ShowOSDetection
    )

    try {
        if ($ShowOSDetection) {
            Write-Host '=== PowerShell OS Detection Test ===' -ForegroundColor Cyan
            Write-Host "`nPowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
            if (Get-Variable -Name 'IsWindows' -ErrorAction SilentlyContinue) {
                Write-Host "IsWindows: $IsWindows, IsMacOS: $IsMacOS, IsLinux: $IsLinux" -ForegroundColor Green
            }
            $osInfo = [System.Environment]::OSVersion
            Write-Host "OSVersion: $($osInfo.Platform), $($osInfo.Version)" -ForegroundColor Green
        }

        $documentsPath = if ($OutputFolder -and (Test-Path $OutputFolder)) {
            $OutputFolder
        } else {
            (Get-Location).Path
        }

        if (-not (Test-Path $documentsPath)) { New-Item -Path $documentsPath -ItemType Directory -Force | Out-Null }

        $finalFileName = if ($IncludeTimestamp) { "$FileName`_$(Get-Date -Format 'yyyyMMdd_HHmmss').$Extension" } else { "$FileName.$Extension" }
        $filePath = Join-Path $documentsPath $finalFileName

        switch ($Extension.ToLower()) {
            'csv' { $Results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 }
            'html' {
                $htmlContent = $Results | ConvertTo-Html -Title $FileName -PreContent "<h1>$FileName</h1><p>Generated: $(Get-Date)</p>"
                $htmlContent | Out-File -FilePath $filePath -Encoding UTF8
            }
            'json' { $Results | ConvertTo-Json -Depth 10 | Out-File -FilePath $filePath -Encoding UTF8 }
            'xml' { $Results | Export-Clixml -Path $filePath -Encoding UTF8 }
            'txt' { $Results | Out-String | Out-File -FilePath $filePath -Encoding UTF8 }
            'pdf' {
                Write-Warning 'PDF export not implemented; exporting as HTML. Convert to PDF externally if needed.'
                $htmlPath = [System.IO.Path]::ChangeExtension($filePath, 'html')
                $htmlContent = $Results | ConvertTo-Html -Title $FileName -PreContent "<h1>$FileName</h1><p>Generated: $(Get-Date)</p>"
                $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
                $filePath = $htmlPath
            }
        }

        if ($DebugMode) { Write-Host "Export completed: $filePath" -ForegroundColor Green }
        return $filePath
    } catch {
        Write-Error "Error exporting to $Extension`: $($_.Exception.Message)"
        throw
    }
}
