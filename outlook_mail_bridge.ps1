<#
  outlook_mail_bridge.ps1
  Sends the TRAM app's queued emails (with the Excel attached) through the LOCAL Outlook.

  How it works:
    The Railway app queues every renewal/offboarding email in its outbox.
    This script polls the outbox every few seconds, creates the email in your Outlook
    (recipient, subject, body, Excel attachment) and sends it, then confirms to the app.

  Requirements: Outlook desktop open and signed in on this PC; this script left running.
  Usage:        .\outlook_mail_bridge.ps1 -Pin <UPLOAD_PIN>
                .\outlook_mail_bridge.ps1 -Pin <UPLOAD_PIN> -SaveOnly   # test: saves drafts, sends nothing
#>
param(
    [string]$Pin     = $env:UPLOAD_PIN,
    [string]$BaseUrl = "https://tram-renewals-production.up.railway.app",
    [int]$PollSeconds = 15,
    [switch]$SaveOnly,
    [switch]$Once
)

if (-not $Pin) { $Pin = Read-Host "Enter the UPLOAD_PIN" }
$ErrorActionPreference = "Stop"

function Log($m) { Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m) }

function Send-OutlookMail($item) {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)   # olMailItem
    $mail.To = $item.to
    $mail.Subject = $item.subject
    if ($item.body_type -eq "HTML") { $mail.HTMLBody = $item.body } else { $mail.Body = $item.body }
    $tmp = $null
    if ($item.has_attachment -and $item.attachment_b64) {
        $tmp = Join-Path $env:TEMP ("tram_" + $item.id + "_" + $item.attachment_filename)
        [IO.File]::WriteAllBytes($tmp, [Convert]::FromBase64String($item.attachment_b64))
        [void]$mail.Attachments.Add($tmp)
    }
    if ($SaveOnly) { $mail.Save() } else { $mail.Send() }
    if ($tmp -and (Test-Path $tmp)) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

Log "Outlook mail bridge started -> $BaseUrl  (SaveOnly=$SaveOnly)"
do {
    try {
        $resp = Invoke-WebRequest "$BaseUrl/api/outbox/next?pin=$Pin" -UseBasicParsing -TimeoutSec 30
        if ($resp.StatusCode -eq 200) {
            $item = $resp.Content | ConvertFrom-Json
            Log "Mail $($item.id): '$($item.subject)' -> $($item.to)"
            try {
                Send-OutlookMail $item
                Invoke-WebRequest "$BaseUrl/api/outbox/$($item.id)/sent?pin=$Pin" -Method POST -UseBasicParsing | Out-Null
                Log "  done ($(if ($SaveOnly) {'saved as draft'} else {'sent'}))"
            } catch {
                Log "  FAILED: $($_.Exception.Message)"
                $b = (@{ error = $_.Exception.Message } | ConvertTo-Json)
                Invoke-WebRequest "$BaseUrl/api/outbox/$($item.id)/fail?pin=$Pin" -Method POST -Body $b -ContentType "application/json" -UseBasicParsing | Out-Null
            }
            continue   # check for more immediately
        }
    } catch {
        Log "Poll error: $($_.Exception.Message)"
    }
    if (-not $Once) { Start-Sleep -Seconds $PollSeconds }
} while (-not $Once)
