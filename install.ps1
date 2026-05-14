# Requires -RunAsAdministrator

$TargetDir = "C:\Program Files\doi-handler"
$VbsFile = "doi_resolver.vbs"
# Force script to look in the exact folder it was extracted to
$CurrentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$SourcePath = Join-Path $CurrentDir $VbsFile
$TargetPath = Join-Path $TargetDir $VbsFile

if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
}

if (Test-Path $SourcePath) {
    Copy-Item -Path $SourcePath -Destination $TargetPath -Force
} else {
    Exit
}

$RegistryPath = "Registry::HKEY_CLASSES_ROOT\doi"
$CommandPath = "$RegistryPath\shell\open\command"

if (-not (Test-Path $RegistryPath)) { New-Item -Path $RegistryPath -Force | Out-Null }
New-ItemProperty -Path $RegistryPath -Name "URL Protocol" -Value "" -PropertyType String -Force | Out-Null

if (-not (Test-Path $CommandPath)) { New-Item -Path $CommandPath -Force | Out-Null }
$ExecString = "wscript.exe `"$TargetPath`" `"%1`""
Set-Item -Path $CommandPath -Value $ExecString


