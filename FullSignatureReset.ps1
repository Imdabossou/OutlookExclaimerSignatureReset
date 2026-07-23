$outlookprofiles = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
$commonMailSettings = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
$exclaimerPath = "\\Fs0131uk\o365\Cloud\Exclaimer.CloudSignatureAgent.application"

$targetKeys = @(
    "Reply-Forward Signature",
    "ReplyForwardSignature",
    "Reply Signature",
    "ReplySignature",
    "New Signature",
    "NewSignature",
    "Exclaimer.Duplicate.New Signature",
    "Exclaimer.Duplicate.NewSignature",
    "Exclaimer.Duplicate.Reply Signature",
    "Exclaimer.Duplicate.ReplySignature",
    "Exclaimer.Duplicate.Reply-Forward Signature",
    "Exclaimer.Duplicate.ReplyForwardSignature"
)

function Remove-TargetValues {
    param([string]$KeyPath)

    foreach ($targetName in $targetKeys) {
        try {
            if (Get-ItemProperty -Path $KeyPath -Name $targetName -ErrorAction SilentlyContinue) {
                Write-Host "Match found ($targetName) at: $KeyPath" -ForegroundColor Yellow
                Remove-ItemProperty -Path $KeyPath -Name $targetName -Force
                if (Get-ItemProperty -Path $KeyPath -Name $targetName -ErrorAction SilentlyContinue) {
                    Write-Host "WARNING: $targetName still present after removal attempt at $KeyPath" -ForegroundColor Red
                } else {
                    Write-Host "Successfully removed: $targetName" -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Host "Error processing key $KeyPath for target $targetName : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "Script running under user context $env:USERNAME" -ForegroundColor Cyan

Write-Host "`n--- Deleting Keys from Profiles Directory ($outlookprofiles) ---" -ForegroundColor Cyan
if (Test-Path $outlookprofiles) {
    Get-ChildItem -Path $outlookprofiles -Recurse | ForEach-Object {
        Remove-TargetValues -KeyPath $_.PSPath
    }
} else {
    Write-Host "Outlook Profiles path not found: $outlookprofiles" -ForegroundColor Yellow
}

Write-Host "`n--- Deleting Keys from Common MailSettings ($commonMailSettings) ---" -ForegroundColor Cyan
if (Test-Path $commonMailSettings) {
    Remove-TargetValues -KeyPath $commonMailSettings
} else {
    Write-Host "Common MailSettings path not found: $commonMailSettings" -ForegroundColor Yellow
}

Write-Host "`n--- Launching Exclaimer Agent ---" -ForegroundColor Cyan
try {
    Start-Process -FilePath $exclaimerPath
    Write-Host "Exclaimer Agent launched successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to launch Exclaimer Agent: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nProcess complete." -ForegroundColor Cyan
