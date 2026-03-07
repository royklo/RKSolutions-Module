<#
.SYNOPSIS
    Disconnects from Microsoft Graph.
.DESCRIPTION
    Clears the current Microsoft Graph session. Use after running report cmdlets when you no longer need the connection.
    If you used Connect-RKGraph to connect, run Disconnect-RKGraph when done to disconnect.
#>
function Disconnect-RKGraph {
    [CmdletBinding()]
    param()
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host 'Disconnected from Microsoft Graph.' -ForegroundColor Green
}
