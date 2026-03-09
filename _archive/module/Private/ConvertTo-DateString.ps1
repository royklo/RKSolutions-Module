# Private: Robust date parsing with fallback (used by M365 License report)
function ConvertTo-DateString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Value,
        [string]$Format = 'dd-MM-yyyy HH:mm',
        [string]$Fallback = 'No sign-in activity'
    )
    if (-not $Value) { return $Fallback }
    try {
        return (Get-Date ([DateTime]$Value) -Format $Format)
    } catch {
        return 'Invalid date value'
    }
}
