<#
.SYNOPSIS
    Cleans up leftover Webex and Cisco Spark files, shortcuts, registry keys, and uninstalls MSI packages for all user profiles.
#>

#region Configuration
$UserProfilesRoot   = 'C:\Users'

# MSI product name patterns to uninstall
$MsiNamePatterns   = @('*Webex*', '*Cisco Spark*')

# Data folders to remove under each profile
$RelativeDataPaths  = @(
    'AppData\Local\Webex',
    'AppData\Roaming\Webex',
    'AppData\Local\Cisco\Spark',
    'AppData\Roaming\Cisco Spark',
    'AppData\Local\Programs\Cisco Spark\',
    'AppData\Local\CiscoSpark',
    'AppData\Local\CiscoSparkLauncher',
    'C:\Program Files\Cisco Spark\',
    'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Webex'
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

function Uninstall-MSIProducts {
    Write-Host "`n=== Uninstalling MSI‑based Webex/Spark products ==="
    # Use Get-WmiObject Win32_Product to find installed MSI products
    $installed = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue |
                 Where-Object { 
                     foreach ($pattern in $MsiNamePatterns) {
                         if ($_.Name -like $pattern) { return $true }
                     }
                     return $false
                 }

    if (-not $installed) {
        Write-Host "ℹ No MSI products matching Webex/Spark found."
        return
    }

    foreach ($pkg in $installed) {
        Write-Host "→ Uninstalling: $($pkg.Name) (ProductCode: $($pkg.IdentifyingNumber))"
        try {
            $exit = $pkg.Uninstall()
            if ($exit.ReturnValue -eq 0) {
                Write-Host "✔ Successfully uninstalled $($pkg.Name)"
            }
            else {
                Write-Warning "⚠ Failed to uninstall $($pkg.Name) (ReturnValue: $($exit.ReturnValue))"
            }
        }
        catch {
            Write-Warning "⚠ Exception uninstalling $($pkg.Name): $_"
        }
    }
}

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

    reg.exe unload "HKU\$hiveName" 2>$null
}

# === MAIN ===

# 1) Uninstall any MSI‑based Webex/Spark products
Uninstall-MSIProducts

# 2) Remove machine‑wide registry keys
Write-Host "`n=== Removing machine‑wide registry entries ==="
Remove-MachineRegKeys

# 3) Process each real user profile
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

# 4) Clean Public shortcuts only
Write-Host "`n=== Cleaning Public shortcuts ==="
Remove-Shortcuts -BaseFolder $PublicDesktop   -Patterns $ShortcutPatterns
Remove-Shortcuts -BaseFolder $PublicStartMenu -Patterns $ShortcutPatterns
