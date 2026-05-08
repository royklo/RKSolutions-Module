# Private: Graph API request with paging and retry
function Invoke-GraphRequestWithPaging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $Uri,
        [Parameter(Mandatory = $false)] [string] $Method = 'GET',
        [Parameter(Mandatory = $false)] [int] $MaxRetries = 3,
        [Parameter(Mandatory = $false)] [switch] $DebugMode
    )
    $results = [System.Collections.Generic.List[object]]::new()
    $currentUri = $Uri
    do {
        $retryCount = 0
        $success = $false
        do {
            try {
                $response = Invoke-MgGraphRequest -Uri $currentUri -Method $Method -OutputType PSObject -ErrorAction Stop
                $success = $true
                if ($response -and $response.PSObject.Properties['value']) {
                    if ($response.value -and $response.value.Count -gt 0) { $results.AddRange($response.value) }
                    $currentUri = $response.'@odata.nextLink'
                } else {
                    $currentUri = $null
                }
            } catch {
                $statusCode = $null
                if ($_.Exception.Response) { $statusCode = $_.Exception.Response.StatusCode.value__ }
                elseif ($_.Exception.InnerException -and $_.Exception.InnerException.Response) { $statusCode = $_.Exception.InnerException.Response.StatusCode.value__ }
                if ($statusCode -eq 400) { return $results.ToArray() }
                $retryCount++
                if ($retryCount -ge $MaxRetries) { return $results.ToArray() }
                Start-Sleep -Seconds (2 * $retryCount)
            }
        } while (-not $success -and $retryCount -lt $MaxRetries)
        if (-not $success) { break }
        if ($results.Count -gt 10000) {
            Write-Warning "Invoke-GraphRequestWithPaging: Results truncated at 10,000 items for URI: $Uri"
            break
        }
    } while ($currentUri)
    $results.ToArray()
}
