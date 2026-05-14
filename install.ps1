# Requires -RunAsAdministrator

# 1. Define paths
$TargetDir = "C:\Program Files\doi-handler"
$VbsFile = "doi_resolver.vbs"
$TargetPath = Join-Path $TargetDir $VbsFile

# 2. Ensure target directory exists
if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    Write-Host "Created directory: $TargetDir" -ForegroundColor Green
}

# 3. Copy the VBScript handler to the target directory
if (Test-Path $VbsFile) {
    Copy-Item -Path $VbsFile -Destination $TargetPath -Force
    Write-Host "Copied $VbsFile to $TargetDir" -ForegroundColor Green
} else {
    Write-Error "Could not find $VbsFile in the current directory."
    Exit
}

# 4. Create Windows Registry structure using the universal Registry provider path
$RegistryPath = "Registry::HKEY_CLASSES_ROOT\doi"
$CommandPath = "$RegistryPath\shell\open\command"

# Build main key
if (-not (Test-Path $RegistryPath)) { New-Item -Path $RegistryPath -Force | Out-Null }
New-ItemProperty -Path $RegistryPath -Name "URL Protocol" -Value "" -PropertyType String -Force | Out-Null

# Build command subkeys and set default path execution string
if (-not (Test-Path $CommandPath)) { New-Item -Path $CommandPath -Force | Out-Null }
$ExecString = "wscript.exe `"$TargetPath`" `"%1`""
Set-Item -Path $CommandPath -Value $ExecString

Write-Host "Successfully registered 'doi:' protocol in Windows Registry!" -ForegroundColor Green
Write-Host "Configuration complete." -ForegroundColor Cyan

