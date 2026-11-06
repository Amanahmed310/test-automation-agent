$ErrorActionPreference = 'Stop'

if ($args.Count -ne 1) {
    throw 'Missing Outlook draft payload.'
}

$payload = Get-Content -LiteralPath $args[0] -Raw | ConvertFrom-Json
try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = [string]$payload.to
    $mail.Subject = [string]$payload.subject
    $mail.Body = [string]$payload.body
    if ($payload.imagePath) {
        $mail.Attachments.Add([string]$payload.imagePath) | Out-Null
    }
    $mail.Display($false)
} catch {
    throw "Could not create an Outlook message. Classic Outlook for Windows must be installed and configured. $($_.Exception.Message)"
}
