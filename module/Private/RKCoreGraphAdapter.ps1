# Private: the module's implementation of the two Graph adapters module/Core calls.
# The RKSolutions Maester pack has its own implementation (helpers/GraphAdapter.ps1); this file is not synced.

function Invoke-RKCoreGraphGet {
    # One GET. The Graph SDK retries 429/503/504 itself; retry 500, 502 and a dropped connection or timeout (no status) here like Maester's own Graph helper.
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Uri)

    $delays = @(2, 4)
    for ($attempt = 0; ; $attempt++) {
        try { return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop }
        catch {
            if ($attempt -ge $delays.Count -or (Get-RKCoreGraphErrorStatus -ErrorRecord $_) -notin 0, 500, 502) { throw }
            Start-Sleep -Seconds $delays[$attempt]
        }
    }
}

function Invoke-RKCoreGraphRequest {
    <#
    .SYNOPSIS
        One Graph GET. Without -All: the raw body. With -All: every item of value, following @odata.nextLink.
        Throws on an HTTP error.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [switch] $All
    )

    $target = ConvertTo-RKCoreGraphPath -Uri $Uri
    $next = "https://graph.microsoft.com/$($target.Version)/$($target.Path)"
    if (-not $All) { return Invoke-RKCoreGraphGet -Uri $next }

    $items = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    while ($next) {
        if (-not $seen.Add($next)) { throw "Graph returned the same @odata.nextLink twice for $Uri." }
        $page = Invoke-RKCoreGraphGet -Uri $next
        if ($page.PSObject.Properties['value']) { foreach ($item in @($page.value | Where-Object { $null -ne $_ })) { $items.Add($item) } }
        else { $items.Add($page) }
        $next = $page.'@odata.nextLink'
    }
    $items.ToArray()
}

function Invoke-RKCoreGraphBatch {
    <#
    .SYNOPSIS
        Many Graph GETs through $batch. Returns Id, Status and Body (the raw first page) per request; never throws
        for a single request.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Requests)

    # Invoke-RKGraphBatch takes one version per call, so group the requests by the version in their URL.
    $byVersion = @{}
    foreach ($request in $Requests) {
        $target = ConvertTo-RKCoreGraphPath -Uri $request.Url
        if (-not $byVersion.ContainsKey($target.Version)) { $byVersion[$target.Version] = [System.Collections.Generic.List[object]]::new() }
        $byVersion[$target.Version].Add([pscustomobject]@{ Id = [string]$request.Id; Url = "/$($target.Path)" })
    }
    foreach ($version in $byVersion.Keys) {
        foreach ($response in (Invoke-RKGraphBatch -Requests $byVersion[$version].ToArray() -GraphVersion $version -Activity 'Reading Intune data from Microsoft Graph')) {
            # A request that failed without an answer from Graph has no Body but an Error: give it the shape Get-RKCoreBatchBody reads.
            $body = $response.Body
            if ($null -eq $body -and $response.PSObject.Properties['Error'] -and $response.Error) { $body = [pscustomobject]@{ error = [pscustomobject]@{ message = [string]$response.Error } } }
            [pscustomobject]@{ Id = [string]$response.Id; Status = [int]$response.Status; Body = $body }
        }
    }
}
