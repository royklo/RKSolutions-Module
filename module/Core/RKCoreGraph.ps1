function ConvertTo-RKCoreGraphPath {
    <#
    .SYNOPSIS
        Splits a Graph URL ('beta/x', '/v1.0/x' or 'https://graph.microsoft.com/beta/x') into Version and Path.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)] [string] $Uri)

    $u = ($Uri -replace '^https?://graph\.microsoft\.com', '') -replace '^/+', ''
    if ($u -notmatch '^(v1\.0|beta)/(.+)$') { throw "The Graph URI '$Uri' must start with v1.0/ or beta/." }
    [pscustomobject]@{ Version = $Matches[1]; Path = $Matches[2] }
}

function Get-RKCoreGraphErrorStatus {
    <#
    .SYNOPSIS
        The HTTP status of a failed Graph request, from the exception or its inner exception; 0 when there is none.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)] [System.Management.Automation.ErrorRecord] $ErrorRecord)

    foreach ($ex in @($ErrorRecord.Exception, $ErrorRecord.Exception.InnerException)) {
        if (-not $ex) { continue }
        if ($ex.PSObject.Properties['Response'] -and $ex.Response -and $ex.Response.StatusCode) { return [int]$ex.Response.StatusCode }
        if ($ex.PSObject.Properties['StatusCode'] -and $ex.StatusCode) { return [int]$ex.StatusCode }
    }
    0
}

function Get-RKCoreBatchBody {
    <#
    .SYNOPSIS
        The body of a required batch response. 404 (an object deleted or referenced by a stale assignment) is
        empty with a warning; any other status other than 200 throws.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [pscustomobject] $Response,
        [Parameter(Mandatory)] [string] $What,
        [Parameter(Mandatory)] [string] $Area,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]] $Warnings
    )

    if ($Response.Status -eq 200) { return $Response.Body }
    if ($Response.Status -eq 404) {
        $Warnings.Add([pscustomobject]@{ Area = $Area; Message = "$What was not found (404) and is skipped." })
        return $null
    }
    throw "Reading $What from Microsoft Graph failed with status $($Response.Status)."
}
