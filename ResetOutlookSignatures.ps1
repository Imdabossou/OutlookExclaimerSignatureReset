# Define the paths and the keys to be removed
$basePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
$mailSettingsPath = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
$sigFilePath = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Signatures"
$exclaimerPath = "\\SERVER NAME HERE\SHARE NAME HERE\Exclaimer.CloudSignatureAgent.application"

$targetKeys = @(
    "Reply-Forward Signature",
    "Exclaimer.Duplicate.New Signature",
    "Exclaimer.Duplicate.Reply Signature",
    "New Signature"
)

Write-Host "Searching for registry entries matching target signature keys..." -ForegroundColor Cyan

# 1. Recursively search through the Profiles directory
Get-ChildItem -Path $basePath -Recurse | ForEach-Object {
    $currentKey = $_.Name.Replace("HKEY_CURRENT_USER", "HKCU:")
    
    foreach ($targetName in $targetKeys) {
        $val = Get-ItemProperty -Path $currentKey -Name $targetName -ErrorAction SilentlyContinue
        
        if ($null -ne $val) {
            Write-Host "Match found ($targetName) at: $currentKey" -ForegroundColor Yellow
            try {
                Remove-ItemProperty -Path $currentKey -Name $targetName -Force
                Write-Host "Successfully removed: $targetName" -ForegroundColor Green
            }
            catch {
                Write-Host "Error removing $targetName : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

# 2. Specifically target the Common MailSettings keys
Write-Host "Checking Common MailSettings..." -ForegroundColor Cyan
$mailSettingsTargets = @("NewSignature", "ReplySignature")

foreach ($msName in $mailSettingsTargets) {
    if (Get-ItemProperty -Path $mailSettingsPath -Name $msName -ErrorAction SilentlyContinue) {
        try {
            Remove-ItemProperty -Path $mailSettingsPath -Name $msName -Force
            Write-Host "Successfully removed $msName from MailSettings." -ForegroundColor Green
        }
        catch {
            Write-Host "Error removing $msName from MailSettings: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# 3. Delete all files in the Signatures directory
Write-Host "Cleaning up signature files for user: $env:USERNAME" -ForegroundColor Cyan
if (Test-Path $sigFilePath) {
    try {
        Get-ChildItem -Path $sigFilePath | Remove-Item -Recurse -Force
        Write-Host "Successfully cleared all files from: $sigFilePath" -ForegroundColor Green
    }
    catch {
        Write-Host "Error clearing Signatures folder: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "Signatures directory not found at $sigFilePath, skipping file cleanup." -ForegroundColor Yellow
}

# 4. Launch Exclaimer Cloud Signature Agent independently
Write-Host "Launching Exclaimer Agent..." -ForegroundColor Cyan
try {
    Start-Process -FilePath $exclaimerPath
    Write-Host "Exclaimer Agent launched successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to launch Exclaimer Agent: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Process complete." -ForegroundColor Cyan
