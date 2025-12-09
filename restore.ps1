<#
 Restore-BrowserOutlook-OneDrive.ps1
 Restores browser data (Chrome, Edge, Firefox, IE Favorites) and Outlook artifacts
 from a backup created by Sync-BrowserData-OneDrive-Clean.ps1.

 NOTES / DISCLAIMERS:
  - CLOSE all target applications (Outlook, Chrome, Edge, Firefox, IE/Explorer windows) before restoring.
  - Browser "Login Data" (saved passwords) may NOT decrypt on another Windows profile or machine due to DPAPI binding.
  - Outlook rules: modern (Exchange/365) rules are server-side; only legacy *.rwz files are restored.
  - Outlook autocomplete cache (Stream_Autocomplete*.dat) may require Outlook restart; only works for same profile.
  - Firefox profile choice: tries to match existing profile or offers to copy into first / create new.

 PARAMETERS:
  -BackupRoot      Path to the backup root (default: autodetect under OneDrive Browser-Sync/<COMPUTERNAME> or prompt)
  -Chrome|-Edge|-Firefox|-InternetExplorer|-Outlook or -All to select components.
  -Force           Overwrite existing destination files without prompting.
  -WhatIf          Simulate actions only.
  -SelectFirefox   Interactively choose a Firefox profile to restore into (if more than one).
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$BackupRoot,
  [switch]$Chrome,
  [switch]$Edge,
  [switch]$Firefox,
  [switch]$InternetExplorer,
  [switch]$Outlook,
  [switch]$All,
  [switch]$Force,           # Aggressively replace (remove / reset attributes before copy)
  [switch]$SkipExisting,    # Skip existing destination files instead of overwriting
  [switch]$SelectFirefox,
  [switch]$NoProcessKill,   # Do NOT attempt to close running target apps
  [switch]$VerboseProcess   # Show detailed process info when killing
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($All) { $Chrome = $Edge = $Firefox = $InternetExplorer = $Outlook = $true }

if (-not ($Chrome -or $Edge -or $Firefox -or $InternetExplorer -or $Outlook)) {
  Write-Host 'Usage: .\Restore-BrowserOutlook-OneDrive.ps1 -All | -Chrome -Edge -Firefox -InternetExplorer -Outlook [-BackupRoot <path>] [-Force] [-SkipExisting] [-WhatIf]' -ForegroundColor Yellow
  Write-Host 'Default: existing files are overwritten. Use -SkipExisting to prevent that, or -Force to aggressively replace locked/read-only files.' -ForegroundColor DarkGray
  return
}

# --- Resolve Backup Root ---
if (-not $BackupRoot) {
  $oneDrive = $env:OneDrive
  if ($oneDrive -and (Test-Path $oneDrive)) {
    # Find most recent machine folder if multiple
    $candidateRoot = Join-Path $oneDrive 'Browser-Sync'
    if (Test-Path $candidateRoot) {
      $machineFolders = Get-ChildItem -Path $candidateRoot -Directory | Sort-Object LastWriteTime -Descending
      if ($machineFolders) {
        $BackupRoot = $machineFolders[0].FullName  # default to newest
        Write-Host "Auto-detected backup root: $BackupRoot" -ForegroundColor Cyan
      }
    }
  }
}

if (-not $BackupRoot -or -not (Test-Path $BackupRoot)) {
  throw "Backup root not found. Specify -BackupRoot explicitly."
}

Write-Host "Using backup root: $BackupRoot" -ForegroundColor Green

function Copy-Back {
  param(
    [Parameter(Mandatory)] [string]$Source,
    [Parameter(Mandatory)] [string]$Destination,
    [switch]$Dir
  )
  if (-not (Test-Path $Source)) { Write-Host "  - Missing source: $Source" -ForegroundColor DarkGray; return }
  $destParent = Split-Path $Destination -Parent
  if ($Dir) { $destParent = $Destination }
  if ($destParent -and -not (Test-Path $destParent)) { New-Item -ItemType Directory -Force -Path $destParent | Out-Null }

  if ($Dir) {
    if ($PSCmdlet.ShouldProcess($Destination, "Mirror directory from $Source")) {
      # Always mirror; /MIR overwrites differences. Add /FFT for cross-FS tolerance? Not needed here.
      $rc = & robocopy $Source $Destination /MIR /R:2 /W:3 /NFL /NDL /NJH /NJS /NP
      if ($LASTEXITCODE -le 7) {
        Write-Host "  [OK] $(Split-Path $Source -Leaf) (dir)" -ForegroundColor Green
      } else {
        Write-Host "  [ERR] Robocopy code $LASTEXITCODE for $Source" -ForegroundColor Red
        $script:ErrorsEncountered++
      }
    }
  } else {
    if (Test-Path $Destination) {
      if ($SkipExisting) {
        Write-Host "  [SKIP] $(Split-Path $Destination -Leaf) exists" -ForegroundColor Yellow
        $script:Skipped++
        return
      }
      if ($Force) {
        try {
          # Remove read-only or locked (best effort)
          Attrib -R $Destination 2>$null | Out-Null
          Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        } catch { }
      }
    }
    if ($PSCmdlet.ShouldProcess($Destination, "Copy file from $Source")) {
      try {
        Copy-Item -LiteralPath $Source -Destination $Destination -Force
        Write-Host "  [OK] $(Split-Path $Destination -Leaf)" -ForegroundColor Green
        $script:Copied++
      } catch {
        Write-Host "  [ERR] Failed copy: $(Split-Path $Destination -Leaf) -> $($_.Exception.Message)" -ForegroundColor Red
        $script:ErrorsEncountered++
      }
    }
  }
}

# Ensure apps closed suggestion
Write-Host 'Ensure Chrome / Edge / Firefox / Outlook are CLOSED before proceeding.' -ForegroundColor Yellow
Write-Host ("Mode: Overwrite={( -not $SkipExisting)} Force={$Force}") -ForegroundColor DarkGray

$script:Copied = 0
$script:Skipped = 0
$script:ErrorsEncountered = 0

# --- Kill running related processes (optional) ---
function Stop-TargetProcesses {
  param(
    [string[]]$Names,
    [switch]$Detailed,
    [switch]$ForceKill
  )
  $found = @()
  foreach ($n in $Names) {
    $procs = Get-Process -Name $n -ErrorAction SilentlyContinue
    if ($procs) { $found += $procs }
  }
  if (-not $found) { return }
  Write-Host "Detected running processes that may lock files:" -ForegroundColor Yellow
  if ($Detailed) {
    $found | Select-Object Name,Id,CPU,StartTime | Format-Table | Out-String | Write-Host
  } else {
    ($found | Select-Object -ExpandProperty Name -Unique) -join ', ' | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
  }
  $proclist = ($found | Select-Object -ExpandProperty Name -Unique) -join ', '
  if (-not $ForceKill) {
    $ans = Read-Host "Terminate these processes now? [$proclist] (y/N)"
    if ($ans -notin @('y','Y')) { Write-Host 'Skipping process termination.' -ForegroundColor DarkGray; return }
  }
  foreach ($p in ($found | Sort-Object -Property Id -Unique)) {
    try {
      Stop-Process -Id $p.Id -Force -ErrorAction Stop
      Write-Host "  [KILLED] $($p.Name) ($($p.Id))" -ForegroundColor Magenta
    } catch {
      Write-Host "  [WARN] Could not kill $($p.Name) ($($p.Id)): $($_.Exception.Message)" -ForegroundColor Yellow
    }
  }
  Start-Sleep -Milliseconds 500
}

if (-not $NoProcessKill) {
  Stop-TargetProcesses -Names @('chrome','msedge','firefox','outlook','iexplore','MicrosoftEdge','MicrosoftEdgeCP','msedgewebview2') -Detailed:$VerboseProcess -ForceKill:$Force
}

# --- Chrome ---
if ($Chrome) {
  $srcDir = Join-Path $BackupRoot 'Chrome'
  Write-Host '\n[Restore Chrome]' -ForegroundColor Cyan
  if (Test-Path $srcDir) {
    $destProfile = Join-Path $env:LOCALAPPDATA 'Google\\Chrome\\User Data\\Default'
    New-Item -ItemType Directory -Force -Path $destProfile | Out-Null
    foreach ($f in 'Bookmarks','Preferences','Login Data','Web Data','Favicons') {
      Copy-Back -Source (Join-Path $srcDir $f) -Destination (Join-Path $destProfile $f)
    }
    Write-Host '  NOTE: Saved passwords may not work if DPAPI context differs.' -ForegroundColor DarkGray
  } else { Write-Host '  No backup found for Chrome' -ForegroundColor DarkGray }
}

# --- Edge ---
if ($Edge) {
  $srcDir = Join-Path $BackupRoot 'Edge'
  Write-Host '\n[Restore Edge]' -ForegroundColor Cyan
  if (Test-Path $srcDir) {
    $destProfile = Join-Path $env:LOCALAPPDATA 'Microsoft\\Edge\\User Data\\Default'
    New-Item -ItemType Directory -Force -Path $destProfile | Out-Null
    foreach ($f in 'Bookmarks','Preferences','Login Data','Web Data','Favicons') {
      Copy-Back -Source (Join-Path $srcDir $f) -Destination (Join-Path $destProfile $f)
    }
    Write-Host '  NOTE: Saved passwords may not work if DPAPI context differs.' -ForegroundColor DarkGray
  } else { Write-Host '  No backup found for Edge' -ForegroundColor DarkGray }
}

# --- Firefox ---
if ($Firefox) {
  $srcDir = Join-Path $BackupRoot 'Firefox'
  Write-Host '\n[Restore Firefox]' -ForegroundColor Cyan
  if (Test-Path $srcDir) {
    $profilesDir = Join-Path $env:APPDATA 'Mozilla\\Firefox\\Profiles'
    if (-not (Test-Path $profilesDir)) { New-Item -ItemType Directory -Force -Path $profilesDir | Out-Null }
    $profiles = Get-ChildItem -Path $profilesDir -Directory
    $targetProfile = $profiles | Where-Object { $_.Name -like '*.default-release*' } | Select-Object -First 1
    if (-not $targetProfile) { $targetProfile = $profiles | Where-Object { $_.Name -like '*.default*' } | Select-Object -First 1 }
    if (-not $targetProfile -and $SelectFirefox) {
      $i = 0
      $profiles | ForEach-Object { Write-Host "  [$i] $($_.Name)"; $i++ }
      $choice = Read-Host 'Select profile index (blank to create new)'
      if ($choice -match '^[0-9]+$' -and [int]$choice -lt $profiles.Count) { $targetProfile = $profiles[[int]$choice] }
    }
    if (-not $targetProfile) {
      $newName = (Get-Random -Minimum 10000 -Maximum 99999).ToString() + '.manual'
      $targetProfile = New-Item -ItemType Directory -Path (Join-Path $profilesDir $newName)
      Write-Host "  Created new profile folder: $newName" -ForegroundColor DarkGray
    }
    $destProfile = $targetProfile.FullName
    foreach ($f in 'places.sqlite','prefs.js','key4.db','logins.json','favicons.sqlite','addonStartup.json.lz4') {
      Copy-Back -Source (Join-Path $srcDir $f) -Destination (Join-Path $destProfile $f)
    }
    Write-Host '  Firefox files restored. Launch Firefox to verify.' -ForegroundColor Green
  } else { Write-Host '  No backup found for Firefox' -ForegroundColor DarkGray }
}

# --- Internet Explorer / Favorites ---
if ($InternetExplorer) {
  $srcDir = Join-Path $BackupRoot 'InternetExplorer'
  Write-Host '\n[Restore IE/Legacy Favorites]' -ForegroundColor Cyan
  if (Test-Path $srcDir) {
    $dest = Join-Path $env:USERPROFILE 'Favorites'
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Back -Source $srcDir -Destination $dest -Dir
  } else { Write-Host '  No backup found for Favorites' -ForegroundColor DarkGray }
}

# --- Outlook ---
if ($Outlook) {
  Write-Host '\n[Restore Outlook]' -ForegroundColor Cyan
  $outRoot = Join-Path $BackupRoot 'Outlook'
  if (Test-Path $outRoot) {
    #$sigSrc = Join-Path $outRoot 'Signatures'
    #$sigDest = Join-Path $env:APPDATA 'Microsoft\\Signatures'
    #if (Test-Path $sigSrc) { Copy-Back -Source $sigSrc -Destination $sigDest -Dir } else { Write-Host '  - No Signatures backup' -ForegroundColor DarkGray }

    $tmplSrc = Join-Path $outRoot 'Templates'
    $tmplDest = Join-Path $env:APPDATA 'Microsoft\\Templates'
    if (Test-Path $tmplSrc) { Copy-Back -Source $tmplSrc -Destination $tmplDest -Dir } else { Write-Host '  - No Templates backup' -ForegroundColor DarkGray }

    $rulesSrc = Join-Path $outRoot 'Rules'
    if (Test-Path $rulesSrc) {
      $rulesDest = Join-Path $env:APPDATA 'Microsoft\\Outlook'
      New-Item -ItemType Directory -Force -Path $rulesDest | Out-Null
      Get-ChildItem $rulesSrc -Filter *.rwz -File | ForEach-Object { Copy-Back -Source $_.FullName -Destination (Join-Path $rulesDest $_.Name) }
      Write-Host '  NOTE: Import .rwz manually via Rules > Manage Rules & Alerts > Options > Import Rules.' -ForegroundColor DarkGray
    }

    $cacheSrc = Join-Path $outRoot 'RoamCache'
    if (Test-Path $cacheSrc) {
      $roamDest = Join-Path $env:LOCALAPPDATA 'Microsoft\\Outlook\\RoamCache'
      New-Item -ItemType Directory -Force -Path $roamDest | Out-Null
      Get-ChildItem $cacheSrc -Filter 'Stream_Autocomplete*' -File | ForEach-Object { Copy-Back -Source $_.FullName -Destination (Join-Path $roamDest $_.Name) }
      Write-Host '  NOTE: Autocomplete restored; if not visible, clear auto-complete cache and retry.' -ForegroundColor DarkGray
    }
  } else { Write-Host '  No Outlook backup folder found.' -ForegroundColor DarkGray }
}

Write-Host '\nRestore operation complete.' -ForegroundColor Green
Write-Host ("Summary: Copied={0} Skipped={1} Errors={2}" -f $script:Copied,$script:Skipped,$script:ErrorsEncountered) -ForegroundColor Cyan
if ($script:ErrorsEncountered -gt 0) { Write-Host 'Some items failed to restore. Review errors above.' -ForegroundColor Yellow }

# SIG # Begin signature block
# MIIIkwYJKoZIhvcNAQcCoIIIhDCCCIACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYannJJaRLYoenCyKEvUhNYPV
# NNOgggXrMIIF5zCCBM+gAwIBAgITWgAAAYQgvD3TcWSXFQAAAAABhDANBgkqhkiG
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
# DjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBRrdQkmmKj7RUHjBO+HtR4J
# PykpqzANBgkqhkiG9w0BAQEFAASCAQAI8zImAmkNLLLbIVXtTzZe5h6BF0SdLVtz
# xqElh45YQiwBLr29KiAXzz8n61hGy4WSo3btGREapwap4A11zYLlTjnuuWlRabH0
# 3g13gSMufLXd6hs49sqc/dalfgZRi2kMh0bwc1ffPzAk0NR2RfujyLkx/6wCL+hl
# tizuP4JTVNYyfg+FWoLOp6fzjfD8m6x1H0npGsAuc6V3C/bdtev4EuaTS1h72J5c
# vn/AbH69ZZMRposdQH7YvDav9YFXNc4LPar8CWc8qEctGDqdMTmngjJHa1sf3bL3
# 7nUUgakwwpLh/uhtuhyBVNYB/iENqLU224ZEvIcxSKIouP3bIzK5
# SIG # End signature block
