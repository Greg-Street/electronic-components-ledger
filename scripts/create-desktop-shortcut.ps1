$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$launcherPath = Join-Path $projectRoot 'scripts\start-app.ps1'
$desktopPath = [Environment]::GetFolderPath('Desktop')
$shortcutName = [string]::Concat([char]0x7535, [char]0x5B50, [char]0x5143, [char]0x4EF6, [char]0x5E93, [char]0x5B58, [char]0x53F0, [char]0x8D26, '.lnk')
$shortcutPath = Join-Path $desktopPath $shortcutName

if (-not (Test-Path -LiteralPath $launcherPath)) {
  throw "App launcher not found: $launcherPath"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$quotedLauncherPath = [char]34 + $launcherPath + [char]34
$shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ' + $quotedLauncherPath
$shortcut.WorkingDirectory = $projectRoot
$shortcut.Description = 'Open the Electronic Components Ledger app window; closing it stops the local service'
$shortcut.IconLocation = (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe') + ',0'
$shortcut.Save()

Write-Output "Desktop app shortcut created: $shortcutPath"
