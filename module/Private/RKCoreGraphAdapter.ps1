# Private: the module's implementation of the two Graph adapters module/Core calls.
# The RKSolutions Maester pack has its own implementation (helpers/GraphAdapter.ps1); this file is not synced.

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
    if (-not $All) { return Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop }

    $items = [System.Collections.Generic.List[object]]::new()
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
        if ($page.PSObject.Properties['value']) { foreach ($item in @($page.value)) { $items.Add($item) } }
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
        foreach ($response in (Invoke-RKGraphBatch -Requests $byVersion[$version].ToArray() -GraphVersion $version)) {
            [pscustomobject]@{ Id = [string]$response.Id; Status = [int]$response.Status; Body = $response.Body }
        }
    }
}
