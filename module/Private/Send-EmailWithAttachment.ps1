# Private: Send email with HTML attachment via Microsoft Graph (Mail.Send). Caller provides Subject and BodyHtml.
function Send-EmailWithAttachment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string[]] $Recipient,
        [Parameter(Mandatory = $true)]
        [string] $AttachmentPath,
        [Parameter(Mandatory = $false)]
        [string] $From,
        [Parameter(Mandatory = $true)]
        [string] $Subject,
        [Parameter(Mandatory = $true)]
        [string] $BodyHtml
    )

    try {
        $contextInfo = Get-MgContext -ErrorAction Stop
        if (-not $contextInfo) {
            Write-Error 'Not connected to Microsoft Graph. Connect using Connect-MgGraph with Mail.Send scope.'
            return $false
        }
    } catch {
        Write-Error "Error checking Graph connection: $_"
        return $false
    }

    if (-not (Test-Path -Path $AttachmentPath)) {
        Write-Error "Attachment file not found: $AttachmentPath"
        return $false
    }
    $fileInfo = Get-Item -Path $AttachmentPath
    if ($fileInfo.Length -gt 3MB) {
        Write-Error "Attachment is too large ($(($fileInfo.Length / 1MB).ToString('0.00')) MB). Maximum recommended size is 3 MB."
        return $false
    }

    try {
        $fileName = Split-Path -Path $AttachmentPath -Leaf
        $contentBytes = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($AttachmentPath))
    } catch {
        Write-Error "Failed to read attachment file: $_"
        return $false
    }

    $toRecipients = @()
    foreach ($email in $Recipient) {
        $toRecipients += @{ emailAddress = @{ address = $email } }
    }

    $mailRequestBody = @{
        message = @{
            subject       = $Subject
            toRecipients  = $toRecipients
            body          = @{
                contentType = 'html'
                content     = $BodyHtml
            }
            attachments   = @(
                @{
                    '@odata.type'  = '#microsoft.graph.fileAttachment'
                    name           = $fileName
                    contentType   = 'text/html'
                    contentBytes  = $contentBytes
                }
            )
        }
    }

    $sendMailUri = 'https://graph.microsoft.com/v1.0/me/sendMail'
    if ($From) {
        if ($From -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$' -and
            $From -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
            Write-Error "Invalid -From value '$From'. Must be an email address or user object ID."
            return $false
        }
        $sendMailUri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
    }

    try {
        $jsonBody = ConvertTo-Json -InputObject $mailRequestBody -Depth 20
        $null = Invoke-MgGraphRequest -Method POST -Uri $sendMailUri -Body $jsonBody -ErrorAction Stop
        Write-Host "Email sent successfully to: $($Recipient -join ', ')" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Failed to send email. Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}
