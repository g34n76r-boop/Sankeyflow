# Installs (or with -Uninstall, removes) the SankeyFlow add-in for desktop Excel
# on Windows, for the current user. No admin rights needed.
#
# Office loads add-ins listed under HKCU\...\Office\16.0\Wef\Developer: each
# value is named after an add-in's ID and holds the path of its manifest. The
# manifest only points Excel at the website, so updates to the add-in arrive
# without reinstalling.
#
# Keep this file ASCII: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.

param([switch]$Uninstall)

$ErrorActionPreference = "Stop"

$AddinId = "b57d1c29-484e-4154-a980-a248f0feaf54"
$ManifestUrl = "https://g34n76r-boop.github.io/SankeyFlow/manifest.xml"
$InstallDir = Join-Path $env:LOCALAPPDATA "SankeyFlow"
$ManifestPath = Join-Path $InstallDir "manifest.xml"
$RegistryKey = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"

function Wait-ExcelClosed {
  while (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "Excel is open. Close all Excel windows, then press Enter (or Ctrl+C to cancel)." -ForegroundColor Yellow
    [void](Read-Host)
  }
}

function Install-SankeyFlow {
  Wait-ExcelClosed

  Write-Host "Downloading the SankeyFlow manifest..."
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
  try {
    $content = (Invoke-WebRequest -Uri $ManifestUrl -UseBasicParsing).Content
  } catch {
    throw "Couldn't download $ManifestUrl. Check the internet connection and try again. ($($_.Exception.Message))"
  }
  if ($content -is [byte[]]) { $content = [Text.Encoding]::UTF8.GetString($content) }

  # Make sure it's the right add-in and the published (not development) manifest.
  try { $xml = [xml]$content } catch { throw "The downloaded manifest isn't valid XML." }
  if ($xml.OfficeApp.Id -ne $AddinId) { throw "The downloaded manifest isn't the SankeyFlow add-in." }
  if ($content -match "localhost") { throw "The downloaded manifest points at localhost, not the published add-in." }

  New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
  [IO.File]::WriteAllText($ManifestPath, $content, (New-Object Text.UTF8Encoding $false))

  if (-not (Test-Path $RegistryKey)) { New-Item -Path $RegistryKey -Force | Out-Null }
  New-ItemProperty -Path $RegistryKey -Name $AddinId -Value $ManifestPath -PropertyType String -Force | Out-Null

  Write-Host ""
  Write-Host "SankeyFlow is installed." -ForegroundColor Green
  Write-Host "Open Excel: the SankeyFlow button is on the Home tab."
}

function Uninstall-SankeyFlow {
  Wait-ExcelClosed

  if (Test-Path $RegistryKey) {
    Remove-ItemProperty -Path $RegistryKey -Name $AddinId -ErrorAction SilentlyContinue
  }
  if (Test-Path $InstallDir) { Remove-Item -Recurse -Force $InstallDir }

  Write-Host ""
  Write-Host "SankeyFlow has been removed." -ForegroundColor Green
}

try {
  if ($Uninstall) { Uninstall-SankeyFlow } else { Install-SankeyFlow }
} catch {
  Write-Host ""
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}
