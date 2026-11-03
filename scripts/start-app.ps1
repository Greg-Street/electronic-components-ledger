$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$appHtml = Join-Path $projectRoot ([string]::Concat([char]0x7535, [char]0x5B50, [char]0x5143, [char]0x4EF6, [char]0x5E93, [char]0x5B58, [char]0x53F0, [char]0x8D26, '.html'))
$browserProcess = $null
$browserPath = $null
$profileRoot = Join-Path $env:LOCALAPPDATA 'ElectronicComponentsLedger\browser-profile'

function Find-Browser {
  $candidates = @(
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
    (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
    (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
    (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
  )
  return $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
}

function Stop-AppBrowserProcesses {
  param([string]$BrowserName, [string]$ProfilePath)
  Get-CimInstance Win32_Process -Filter "Name='$BrowserName'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -and $_.CommandLine.Contains($ProfilePath) } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}

try {
  if (-not (Test-Path -LiteralPath $appHtml)) {
    throw "Offline app was not found: $appHtml. Run npm run build first."
  }

  $browserPath = Find-Browser
  if (-not $browserPath) {
    throw 'Microsoft Edge or Google Chrome was not found.'
  }

  New-Item -ItemType Directory -Force -Path $profileRoot | Out-Null
  $appUrl = ([System.Uri]$appHtml).AbsoluteUri
  $profileArgument = '--user-data-dir=' + [char]34 + $profileRoot + [char]34
  $browserProcess = Start-Process -FilePath $browserPath -ArgumentList @(
    "--app=$appUrl",
    $profileArgument,
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-background-mode'
  ) -WorkingDirectory $projectRoot -PassThru

  Start-Sleep -Milliseconds 500
  if ($browserProcess.HasExited) {
    throw 'The application window did not open.'
  }

  $browserProcess.WaitForExit()
} finally {
  if ($browserProcess -and -not $browserProcess.HasExited) {
    Stop-Process -Id $browserProcess.Id -Force -ErrorAction SilentlyContinue
  }
  if ($browserPath) {
    Stop-AppBrowserProcesses -BrowserName (Split-Path -Leaf $browserPath) -ProfilePath $profileRoot
  }
}
