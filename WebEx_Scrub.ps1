<#
.SYNOPSIS
    Cleans up leftover Webex and Cisco Spark files, shortcuts, and registry keys for all user profiles.
#>

#region Configuration
$UserProfilesRoot   = 'C:\Users'

# Data folders to remove under each profile
$RelativeDataPaths  = @(
    'AppData\Local\Webex',
    'AppData\Roaming\Webex',
    'AppData\Local\Cisco\Spark',
    'AppData\Roaming\Cisco Spark'
)

# Shortcut filename patterns
$ShortcutPatterns   = @('*Webex*.lnk', '*Cisco Spark*.lnk')

# Machine‑wide registry keys to delete
$MachineRegKeys = @(
    'HKLM:\SOFTWARE\Cisco\Webex',
    'HKLM:\SOFTWARE\Webex',
    'HKLM:\SOFTWARE\Wow6432Node\Cisco\Webex',
    'HKLM:\SOFTWARE\Wow6432Node\Webex',
    'HKLM:\SOFTWARE\Cisco\Spark',
    'HKLM:\SOFTWARE\Cisco Spark',
    'HKLM:\SOFTWARE\Wow6432Node\Cisco\Spark',
    'HKLM:\SOFTWARE\Wow6432Node\Cisco Spark',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Cisco Webex',
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Cisco Webex',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Cisco Spark',
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Cisco Spark'
)

# Per‑user registry subkeys under HKCU to delete
$UserRegSubKeys = @(
    'Software\Cisco\Webex',
    'Software\Webex',
    'Software\Cisco\Spark',
    'Software\Cisco Spark'
)

# Public shortcuts locations
$PublicDesktop   = Join-Path $UserProfilesRoot 'Public\Desktop'
$PublicStartMenu = Join-Path $Env:ProgramData 'Microsoft\Windows\Start Menu\Programs'
#endregion

function Remove-Paths {
    param([string[]] $Paths)
    foreach ($p in $Paths) {
        if (Test-Path $p) {
            Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "✔ Removed: $p"
        }
        else {
            Write-Host "ℹ Not found: $p"
        }
    }
}

function Remove-Shortcuts {
    param([string] $BaseFolder, [string[]] $Patterns)
    foreach ($pattern in $Patterns) {
        $matches = Get-ChildItem -Path $BaseFolder -Filter $pattern -Recurse -ErrorAction SilentlyContinue
        if ($matches) {
            foreach ($lnk in $matches) {
                Remove-Item -LiteralPath $lnk.FullName -Force -ErrorAction SilentlyContinue
                Write-Host "✔ Removed shortcut: $($lnk.FullName)"
            }
        }
        else {
            Write-Host "ℹ No shortcuts matching '$pattern' in $BaseFolder"
        }
    }
}

function Remove-MachineRegKeys {
    foreach ($key in $MachineRegKeys) {
        if (Test-Path $key) {
            Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "✔ Removed registry key: $key"
        }
        else {
            Write-Host "ℹ Registry key not found: $key"
        }
    }
}

function Remove-UserRegKeys {
    param([string] $NtUserDatPath)

    $hiveName = 'TempHive_' + [Guid]::NewGuid().ToString('N')
    # Load hive
    reg.exe load "HKU\$hiveName" "$NtUserDatPath" 2>$null
    foreach ($subKey in $UserRegSubKeys) {
        $fullKey = "Registry::HKEY_USERS\$hiveName\$subKey"
        if (Test-Path $fullKey) {
            Remove-Item -Path $fullKey -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "✔ Removed user-registry: HKU:\$hiveName\$subKey"
        }
        else {
            Write-Host "ℹ User-registry not found: HKU:\$hiveName\$subKey"
        }
    }
    # Unload hive
    reg.exe unload "HKU\$hiveName" 2>$null
}

# 1) Remove machine‑wide registry keys
Write-Host "`n=== Removing machine‑wide registry entries ==="
Remove-MachineRegKeys

# 2) Process each real user profile
$skip = 'Default','Default User','Public','All Users'
$profiles = Get-ChildItem -Directory -Path $UserProfilesRoot |
            Where-Object { $skip -notcontains $_.Name }

foreach ($prof in $profiles) {
    Write-Host "`n=== Profile: $($prof.Name) ==="

    # a) Delete AppData folders
    $paths = $RelativeDataPaths | ForEach-Object { Join-Path $prof.FullName $_ }
    Remove-Paths -Paths $paths

    # b) Delete shortcuts
    Remove-Shortcuts -BaseFolder (Join-Path $prof.FullName 'Desktop') -Patterns $ShortcutPatterns
    Remove-Shortcuts -BaseFolder (Join-Path $prof.FullName 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs') -Patterns $ShortcutPatterns

    # c) Delete per-user registry
    $ntuser = Join-Path $prof.FullName 'NTUSER.DAT'
    if (Test-Path $ntuser) {
        Remove-UserRegKeys -NtUserDatPath $ntuser
    }
    else {
        Write-Warning "NTUSER.DAT not found for profile $($prof.Name)"
    }
}

# 3) Clean Public shortcuts only
Write-Host "`n=== Cleaning Public shortcuts ==="
Remove-Shortcuts -BaseFolder $PublicDesktop   -Patterns $ShortcutPatterns
Remove-Shortcuts -BaseFolder $PublicStartMenu -Patterns $ShortcutPatterns

 