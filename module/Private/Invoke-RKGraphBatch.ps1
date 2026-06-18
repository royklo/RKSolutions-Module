# Private: Microsoft Graph $batch helper.
# Sends up to 20 sub-requests per HTTP round trip, handles 429/5xx retry
# (respecting Retry-After), and returns one response object per input request
# correlated by Id so callers can map results back without re-issuing requests.
function Invoke-RKGraphBatch {
    [CmdletBinding()]
    param(
        # Each input request is a hashtable / PSObject with:
        #   Id      (string, optional - auto-generated when missing)
        #   Url     (string, required - absolute Graph URL or path)
        #   Method  (string, optional - defaults to GET)
        #   Body    (object, optional - serialised to JSON; sets Content-Type)
        #   Headers (hashtable, optional)
        [Parameter(Mandatory)]
        [object[]] $Requests,

        [ValidateSet('beta', 'v1.0')]
        [string] $GraphVersion = 'beta',

        [ValidateRange(1, 20)]
        [int] $BatchSize = 20,

        [ValidateRange(0, 10)]
        [int] $MaxRetries = 5,

        [string] $Activity,

        [switch] $DebugMode
    )

    if (-not $Requests -or $Requests.Count -eq 0) { return @() }

    $batchUri = "https://graph.microsoft.com/$GraphVersion/`$batch"
    $allResponses = [System.Collections.Generic.List[object]]::new()

    # Normalise inputs into the JSON shape the $batch endpoint expects.
    $normalised = [System.Collections.Generic.List[hashtable]]::new()
    $counter = 0
    foreach ($req in $Requests) {
        $counter++
        $idValue = $null
        if ($req.PSObject.Properties['Id']) { $idValue = $req.Id }
        $id = if ($idValue) { [string]$idValue } else { "req-$counter" }

        $url = [string]$req.Url
        if ([string]::IsNullOrWhiteSpace($url)) {
            Write-Verbose "Invoke-RKGraphBatch: skipping request '$id' with empty Url"
            continue
        }
        # Strip absolute host so the batch sub-request URL is a path.
        $url = $url -replace '^https?://graph\.microsoft\.com/(beta|v1\.0)', ''
        if (-not $url.StartsWith('/')) { $url = "/$url" }

        $entry = [ordered]@{
            id     = $id
            method = if ($req.Method) { [string]$req.Method } else { 'GET' }
            url    = $url
        }
        if ($req.PSObject.Properties['Body'] -and $null -ne $req.Body) {
            $entry.body = $req.Body
            $entry.headers = @{ 'Content-Type' = 'application/json' }
        }
        if ($req.PSObject.Properties['Headers'] -and $req.Headers) {
            if (-not $entry.Contains('headers')) { $entry.headers = @{} }
            foreach ($k in $req.Headers.Keys) { $entry.headers[$k] = $req.Headers[$k] }
        }
        $normalised.Add([hashtable]$entry)
    }

    if ($normalised.Count -eq 0) { return @() }

    $chunkCount = [Math]::Ceiling($normalised.Count / [double]$BatchSize)
    $chunkIndex = 0

    for ($offset = 0; $offset -lt $normalised.Count; $offset += $BatchSize) {
        $chunkIndex++
        $take = [Math]::Min($BatchSize, $normalised.Count - $offset)
        $chunk = $normalised.GetRange($offset, $take)
        $pending = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($c in $chunk) { $pending.Add($c) }

        if ($Activity) {
            $pct = [int](($chunkIndex / [double]$chunkCount) * 100)
            Write-Progress -Activity $Activity -Status "Batch $chunkIndex of $chunkCount" -PercentComplete $pct
        }

        $attempt = 0
        while ($pending.Count -gt 0) {
            $payload = @{ requests = @($pending) } | ConvertTo-Json -Depth 20
            $response = $null
            try {
                $response = Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $payload -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
            } catch {
                # Inspect HTTP status. 4xx (except 408 / 429) is permanent: retrying wastes
                # time. Only 408, 429, 5xx, and unknown (network / parse error) are retried.
                $statusCode = 0
                if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                $errorBody = $null
                if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $errorBody = $_.ErrorDetails.Message }

                $isTransient = ($statusCode -eq 0 -or $statusCode -eq 408 -or $statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600))

                if (-not $isTransient) {
                    Write-Verbose "Invoke-RKGraphBatch: batch POST failed with non-transient HTTP $statusCode (not retried): $($_.Exception.Message)"
                    if ($errorBody) { Write-Verbose "Invoke-RKGraphBatch: response body: $errorBody" }
                    foreach ($p in $pending) {
                        $allResponses.Add([PSCustomObject]@{ Id = $p.id; Status = $statusCode; Body = $null; Headers = $null; Error = $_.Exception.Message })
                    }
                    break
                }

                $attempt++
                if ($attempt -gt $MaxRetries) {
                    Write-Verbose "Invoke-RKGraphBatch: batch POST failed after $MaxRetries retries (HTTP $statusCode): $($_.Exception.Message)"
                    if ($errorBody) { Write-Verbose "Invoke-RKGraphBatch: response body: $errorBody" }
                    foreach ($p in $pending) {
                        $allResponses.Add([PSCustomObject]@{ Id = $p.id; Status = $statusCode; Body = $null; Headers = $null; Error = $_.Exception.Message })
                    }
                    break
                }
                Start-Sleep -Seconds ([Math]::Min(60, [Math]::Pow(2, $attempt)))
                continue
            }

            $retryList = [System.Collections.Generic.List[hashtable]]::new()
            $maxRetryAfter = 0
            foreach ($r in $response.responses) {
                $status = 0
                if ($r.PSObject.Properties['status']) { $status = [int]$r.status }

                $isRetryable = ($status -eq 429 -or ($status -ge 500 -and $status -lt 600))
                if ($isRetryable -and $attempt -lt $MaxRetries) {
                    if ($r.headers) {
                        $retryAfterProp = $r.headers.PSObject.Properties | Where-Object { $_.Name -ieq 'Retry-After' } | Select-Object -First 1
                        if ($retryAfterProp) {
                            $parsed = 0
                            if ([int]::TryParse([string]$retryAfterProp.Value, [ref]$parsed) -and $parsed -gt $maxRetryAfter) {
                                $maxRetryAfter = $parsed
                            }
                        }
                    }
                    $match = $pending | Where-Object { $_.id -eq $r.id } | Select-Object -First 1
                    if ($match) { $retryList.Add($match) }
                } else {
                    $allResponses.Add([PSCustomObject]@{
                            Id      = $r.id
                            Status  = $status
                            Body    = $r.body
                            Headers = $r.headers
                            Error   = $null
                        })
                }
            }

            if ($retryList.Count -eq 0) { break }
            $attempt++
            if ($attempt -gt $MaxRetries) {
                foreach ($p in $retryList) {
                    $allResponses.Add([PSCustomObject]@{ Id = $p.id; Status = 429; Body = $null; Headers = $null; Error = 'Max retries exceeded' })
                }
                break
            }
            $pending = $retryList
            $sleepSec = if ($maxRetryAfter -gt 0) { [Math]::Min(60, $maxRetryAfter) } else { [Math]::Min(30, [Math]::Pow(2, $attempt)) }
            if ($DebugMode) { Write-Verbose "Invoke-RKGraphBatch: retrying $($pending.Count) sub-request(s) after $sleepSec s" }
            Start-Sleep -Seconds $sleepSec
        }
    }

    if ($Activity) { Write-Progress -Activity $Activity -Completed }
    return $allResponses.ToArray()
}
