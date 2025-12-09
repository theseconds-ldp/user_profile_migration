<#
  Rewritten clean version (ASCII only) to eliminate hidden characters causing parse errors.
  Provides backup of browser data to OneDrive with optional scheduled task creation.
  NOTE: param MUST be the first executable statement; Set-StrictMode moved below it.
#>

param(
  [string]$OneDrivePath = $env:OneDrive,
  [switch]$Chrome,
  [switch]$Edge,
  [switch]$Firefox,
  [switch]$InternetExplorer,
  [switch]$All,
  [switch]$IncludeOutlook = $true,
  [switch]$SkipKFM,            # Skip attempt to enable Known Folder Move (KFM)
  [switch]$VerboseList
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($All) { $Chrome = $Edge = $Firefox = $InternetExplorer = $true }

if (-not ($Chrome -or $Edge -or $Firefox -or $InternetExplorer -or $IncludeOutlook)) {
  Write-Host 'Browser/Outlook Data Sync to OneDrive' -ForegroundColor Green
  Write-Host 'Usage: .\Sync-BrowserData-OneDrive-Clean.ps1 [-Chrome] [-Edge] [-Firefox] [-InternetExplorer] [-All] [-IncludeOutlook] [-SkipKFM]' -ForegroundColor Yellow
  return
}

if (-not $OneDrivePath -or -not (Test-Path $OneDrivePath)) {
  Write-Host "ERROR: OneDrive path not found: $OneDrivePath" -ForegroundColor Red
  return
}

$backupRoot = Join-Path $OneDrivePath ("Browser-Sync\" + $env:COMPUTERNAME)
New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

function Copy-ItemSafe {
  param(
    [Parameter(Mandatory)] [string]$Source,
    [Parameter(Mandatory)] [string]$Destination,
    [switch]$Quiet
  )
  if (Test-Path $Source) {
    $destDir = Split-Path $Destination -Parent
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    Copy-Item -LiteralPath $Source -Destination $Destination -Force
    if (-not $Quiet) { Write-Host "  [OK] $([IO.Path]::GetFileName($Source))" -ForegroundColor Green }
    return $true
  } else {
    if (-not $Quiet) { Write-Host "  - Missing: $Source" -ForegroundColor DarkGray }
    return $false
  }
}

$total = 0

# region KFM (Desktop/Documents/Pictures) status & optional enable
function Get-KFMStatus {
  $shellFoldersPath = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders'
  $result = @{}
  foreach ($k in 'Desktop','Personal','My Pictures') {
    $val = (Get-ItemProperty -Path $shellFoldersPath -ErrorAction SilentlyContinue).$k
    if ($val) { $expanded = [Environment]::ExpandEnvironmentVariables($val); $result[$k] = $expanded } else { $result[$k] = $null }
  }
  return [pscustomobject]@{
    Desktop   = $result.Desktop
    Documents = $result.Personal
    Pictures  = $result.'My Pictures'
    DesktopRedirected   = ($result.Desktop   -like '*OneDrive*')
    DocumentsRedirected = ($result.Personal  -like '*OneDrive*')
    PicturesRedirected  = ($result.'My Pictures' -like '*OneDrive*')
  }
}

function Enable-KFMAttempt {
  Write-Host 'Attempting to enable OneDrive Known Folder Move...' -ForegroundColor Yellow
  $base = 'HKCU:\\Software\\Policies\\Microsoft\\OneDrive'
  $kfm = Join-Path $base 'KnownFolderMove'
  if (-not (Test-Path $kfm)) { New-Item -Path $kfm -Force | Out-Null }
  try {
    New-ItemProperty -Path $kfm -Name 'OptInToKFM' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $kfm -Name 'KFMSilentOptIn' -Value '' -PropertyType String -Force | Out-Null
    Write-Host 'KFM registry opt-in set. A OneDrive restart may be required.' -ForegroundColor Green
  } catch { Write-Host "[ERR] Failed to set KFM registry keys: $($_.Exception.Message)" -ForegroundColor Red }
}

$kfmStatusBefore = Get-KFMStatus
if (-not $SkipKFM -and -not ($kfmStatusBefore.DesktopRedirected -and $kfmStatusBefore.DocumentsRedirected -and $kfmStatusBefore.PicturesRedirected)) {
  Enable-KFMAttempt
  $kfmStatusAfter = Get-KFMStatus
} else { $kfmStatusAfter = $kfmStatusBefore }
# endregion

Write-Host '=== Browser Data Sync to OneDrive ===' -ForegroundColor Cyan
Write-Host "Backup root: $backupRoot" -ForegroundColor DarkCyan

if ($Chrome) {
  Write-Host '\n[Chrome]' -ForegroundColor Yellow
  $profile = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\Default'
  if (Test-Path $profile) {
    $dest = Join-Path $backupRoot 'Chrome'
    $files = 'Bookmarks','Preferences','Login Data','Web Data','Favicons'
    foreach ($f in $files) { if (Copy-ItemSafe -Source (Join-Path $profile $f) -Destination (Join-Path $dest $f)) { $total++ } }
  } else { Write-Host '  Not found' -ForegroundColor DarkGray }
}

if ($Edge) {
  Write-Host '\n[Edge]' -ForegroundColor Yellow
  $profile = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data\Default'
  if (Test-Path $profile) {
    $dest = Join-Path $backupRoot 'Edge'
    $files = 'Bookmarks','Preferences','Login Data','Web Data','Favicons'
    foreach ($f in $files) { if (Copy-ItemSafe -Source (Join-Path $profile $f) -Destination (Join-Path $dest $f)) { $total++ } }
  } else { Write-Host '  Not found' -ForegroundColor DarkGray }
}

if ($Firefox) {
  Write-Host '\n[Firefox]' -ForegroundColor Yellow
  $profilesDir = Join-Path $env:APPDATA 'Mozilla\\Firefox\\Profiles'
  if (Test-Path $profilesDir) {
    # Prefer default-release, then default, else newest
    $candidates = Get-ChildItem -Path $profilesDir -Directory | Sort-Object LastWriteTime -Descending
    $default = $candidates | Where-Object { $_.Name -like '*.default-release*' } | Select-Object -First 1
    if (-not $default) { $default = $candidates | Where-Object { $_.Name -like '*.default*' } | Select-Object -First 1 }
    if (-not $default) { $default = $candidates | Select-Object -First 1 }
    if ($default) {
      Write-Host "  Using profile: $($default.Name)" -ForegroundColor DarkCyan
      $dest = Join-Path $backupRoot 'Firefox'
      $files = 'places.sqlite','prefs.js','key4.db','logins.json','favicons.sqlite','addonStartup.json.lz4'
      foreach ($f in $files) { if (Copy-ItemSafe -Source (Join-Path $default.FullName $f) -Destination (Join-Path $dest $f)) { $total++ } }
    } else { Write-Host '  No profile directories found' -ForegroundColor DarkGray }
  } else { Write-Host '  Profiles directory not found' -ForegroundColor DarkGray }
}

if ($InternetExplorer) {
  Write-Host '\n[Internet Explorer / Legacy Edge Favorites]' -ForegroundColor Yellow
  $src = Join-Path $env:USERPROFILE 'Favorites'
  if (Test-Path $src) {
    $dest = Join-Path $backupRoot 'InternetExplorer'
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    $rc = & robocopy $src $dest /MIR /R:2 /W:3 /NFL /NDL /NJH /NJS /NP
    if ($LASTEXITCODE -le 7) { Write-Host '  [OK] Favorites synced' -ForegroundColor Green; $total++ } else { Write-Host "  [ERR] Robocopy exit code $LASTEXITCODE" -ForegroundColor Red }
  } else { Write-Host '  Favorites folder not found' -ForegroundColor DarkGray }
}

if ($IncludeOutlook) {
  Write-Host '\n[Outlook]' -ForegroundColor Yellow
  #$sigSrc = Join-Path $env:APPDATA 'Microsoft\\Signatures'
  $tmplSrc = Join-Path $env:APPDATA 'Microsoft\\Templates'
  $rulesSrc = Join-Path $env:APPDATA 'Microsoft\\Outlook'
  $roamCache = Join-Path $env:LOCALAPPDATA 'Microsoft\\Outlook\\RoamCache'
  $outDest = Join-Path $backupRoot 'Outlook'
  New-Item -ItemType Directory -Force -Path $outDest | Out-Null

  #if (Test-Path $sigSrc) { $rc = & robocopy $sigSrc (Join-Path $outDest 'Signatures') /MIR /R:2 /W:3 /NFL /NDL /NJH /NJS /NP; if ($LASTEXITCODE -le 7) { Write-Host '  [OK] Signatures' -ForegroundColor Green } else { Write-Host '  [ERR] Signatures copy' -ForegroundColor Red } }
  #else { Write-Host '  - No Signatures folder' -ForegroundColor DarkGray }

  if (Test-Path $tmplSrc) {
    $tmplDest = Join-Path $outDest 'Templates'
    New-Item -ItemType Directory -Force -Path $tmplDest | Out-Null
    Get-ChildItem $tmplSrc -Filter *.oft -File -ErrorAction SilentlyContinue | ForEach-Object {
      Copy-Item $_.FullName (Join-Path $tmplDest $_.Name) -Force; $total++
    }
    Write-Host '  [OK] Templates (*.oft) copied' -ForegroundColor Green
  } else { Write-Host '  - No Templates folder' -ForegroundColor DarkGray }

  if (Test-Path $rulesSrc) {
    $ruleFiles = Get-ChildItem $rulesSrc -Filter *.rwz -File -ErrorAction SilentlyContinue
    if ($ruleFiles) {
      $rulesDest = Join-Path $outDest 'Rules'
      New-Item -ItemType Directory -Force -Path $rulesDest | Out-Null
      foreach ($rf in $ruleFiles) { Copy-Item $rf.FullName (Join-Path $rulesDest $rf.Name) -Force; $total++ }
      Write-Host '  [OK] Legacy rule files (*.rwz) copied' -ForegroundColor Green
    } else { Write-Host '  - No legacy rule files (*.rwz) found (modern Outlook stores rules server-side)' -ForegroundColor DarkGray }
  }

  if (Test-Path $roamCache) {
    $roamDest = Join-Path $outDest 'RoamCache'
    New-Item -ItemType Directory -Force -Path $roamDest | Out-Null
    Get-ChildItem $roamCache -Filter 'Stream_Autocomplete*' -File -ErrorAction SilentlyContinue | ForEach-Object {
      Copy-Item $_.FullName (Join-Path $roamDest $_.Name) -Force; $total++ }
    Write-Host '  [OK] Autocomplete cache copied (Stream_Autocomplete*)' -ForegroundColor Green
  }
}

Write-Host "\n=== Summary ===" -ForegroundColor Green
Write-Host "Items (files/folders) backed up (counted individual file copies + category markers): $total" -ForegroundColor Green
Write-Host "Backup path: $backupRoot" -ForegroundColor Green
Write-Host '\nKnown Folder Move (Desktop/Documents/Pictures) Status:' -ForegroundColor Cyan
Write-Host ("  Before -> Desktop:{0} Documents:{1} Pictures:{2}" -f $kfmStatusBefore.DesktopRedirected,$kfmStatusBefore.DocumentsRedirected,$kfmStatusBefore.PicturesRedirected) -ForegroundColor DarkGray
Write-Host ("  After  -> Desktop:{0} Documents:{1} Pictures:{2}" -f $kfmStatusAfter.DesktopRedirected,$kfmStatusAfter.DocumentsRedirected,$kfmStatusAfter.PicturesRedirected) -ForegroundColor DarkGray
if (-not $SkipKFM -and ($kfmStatusAfter.DesktopRedirected -or $kfmStatusAfter.DocumentsRedirected -or $kfmStatusAfter.PicturesRedirected) -and -not ($kfmStatusBefore.DesktopRedirected -or $kfmStatusBefore.DocumentsRedirected -or $kfmStatusBefore.PicturesRedirected)) {
  Write-Host 'KFM enable attempt applied. If folders are still not redirected, open OneDrive settings > Backup > Manage Backup and confirm.' -ForegroundColor Yellow
}
Write-Host '\nDone.' -ForegroundColor Cyan

# SIG # Begin signature block
# MIIIkwYJKoZIhvcNAQcCoIIIhDCCCIACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUQU3UF6JI39aO8ZqQ30Z0JxsD
# gTqgggXrMIIF5zCCBM+gAwIBAgITWgAAAYQgvD3TcWSXFQAAAAABhDANBgkqhkiG
# 9w0BAQ0FADBaMRIwEAYKCZImiZPyLGQBGRYCbnoxEzARBgoJkiaJk/IsZAEZFgNu
# ZXQxFzAVBgoJkiaJk/IsZAEZFgdjZXJub256MRYwFAYDVQQDEw1NQ0wtUFJELUNB
# UzAxMB4XDTI1MDcyODAzNDc1NFoXDTI2MDcyODAzNDc1NFowFDESMBAGA1UEAxMJ
# SmltbXkgTGFtMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAoQJCf4rn
# N5AZ83pUo4LR5t7O2lz9/OZgtMvT+YhYv5Zw8xwLLla+6qShxCEWdPi9hl4JWQe3
# RCyKDF/f9xES+Svcooze6AsHI3eA+41JtcvqeILENPEiERu6Kha4UULguCW2Edf2
# P8WVB531tJ4nTQ6fawFCbC69EeOb6KHXyGGc356mV2bD71itW0c4AcRVSWJLJOVr
# aI2qTPAeITo9TFPPbfBtBZBk7UhAzzX1XB/3x/p23BhwPWWnjst9kGyTYVFjeQR3
# cQQfuNMdi4TUFFZWCX5XoFra1Ooll8vJMgByqsdnUlzPwDGyGmXHuytxuh80449b
# uA1TkKdd5HBXKQIDAQABo4IC6jCCAuYwOwYJKwYBBAGCNxUHBC4wLAYkKwYBBAGC
# NxUIha2BAej1JIbNnTmBrOIHt+53T4Tpz2uH9KF3AgFkAgELMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIHgDAbBgkrBgEEAYI3FQoEDjAMMAoGCCsG
# AQUFBwMDMB0GA1UdDgQWBBRrKX4Pkjv3hIYDD7H3t3i4RusnfTAfBgNVHSMEGDAW
# gBSu1SniGFK39ZxihlCrNiyrgqnk/zCB1wYDVR0fBIHPMIHMMIHJoIHGoIHDhoHA
# bGRhcDovLy9DTj1NQ0wtUFJELUNBUzAxLENOPU1DTC1QUkQtQ0FTMDEsQ049Q0RQ
# LENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZp
# Z3VyYXRpb24sREM9Y2Vybm9ueixEQz1uZXQsREM9bno/Y2VydGlmaWNhdGVSZXZv
# Y2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50
# MIHFBggrBgEFBQcBAQSBuDCBtTCBsgYIKwYBBQUHMAKGgaVsZGFwOi8vL0NOPU1D
# TC1QUkQtQ0FTMDEsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENO
# PVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9Y2Vybm9ueixEQz1uZXQsREM9
# bno/Y0FDZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRpZmljYXRpb25B
# dXRob3JpdHkwMwYDVR0RBCwwKqAoBgorBgEEAYI3FAIDoBoMGGppbW15LmxhbUBt
# Y2xhcmVucy5jby5uejBOBgkrBgEEAYI3GQIEQTA/oD0GCisGAQQBgjcZAgGgLwQt
# Uy0xLTUtMjEtMzIxMjQ4NDY5Ni0zODY2NDU3MjUtMjM2NzQyNzY5My02MDM3MA0G
# CSqGSIb3DQEBDQUAA4IBAQChPnXwuSRXDU3VZMG5W60bMpx4vb84Hfl2em5cCxGY
# YP6KmibbXDU0IJ1JIrxfnJh7UgvVaKsQZ/4/gT9VDfUyi8ytMVKCzvMGJKLixha7
# 7qp1BuRYuicCSup3po8934LlDnqJfNVDRvtP1eajBWxe+UF5T9U1+aJMISbqK5EM
# UcInor81VDB17MLyYbowLqEGJxTiXJxPFv0xr5xi3L2jA9j3ccJZXZfi3Rq4z1sY
# IxElVkXjMxl/DX9ZNNUzXJ6oKe5Nlmr8x4s+XIpsq3WBdJz+w7Q8valxt1+A1wtK
# t/fGMJphTv1aP1W7sNQWtKEzzkpZOKW4FHZEO/rbk3oCMYICEjCCAg4CAQEwcTBa
# MRIwEAYKCZImiZPyLGQBGRYCbnoxEzARBgoJkiaJk/IsZAEZFgNuZXQxFzAVBgoJ
# kiaJk/IsZAEZFgdjZXJub256MRYwFAYDVQQDEw1NQ0wtUFJELUNBUzAxAhNaAAAB
# hCC8PdNxZJcVAAAAAAGEMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBRZFoy4nRH2LzTFjTFBJyAz
# C5JYJDANBgkqhkiG9w0BAQEFAASCAQBfUV4O4bx9US1rw4NRDK2hAtp8b/qgl+mW
# FIdv9B3FNOuYFmACVLvSh2x/gI/Y+K6PQap1DMom1F02aDF95EWmGUvAG5p/TMc/
# Cur9aUwOWFwQ1mAYEWEQ/wwIYCl6quitsjrZFhb355W0SLqebuZAzBNVQX16N06C
# d27SzhUdFdlyzQIA5tYLd7GXpUGCOiw8rA8tbO52/KNkj6oFZMqeRYIrLadWwR5u
# 4ge8MNCSNoz8WaNfBL9yXxexT3zESFW2EVwQuoYKxQSJn9/VcLC08zefhcnbLfps
# yaZsJqlT+T9Nvs6x2K8wYCYC5AEhuwU+Se1P9npUUiDPikKql0LQ
# SIG # End signature block
