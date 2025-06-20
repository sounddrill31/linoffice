# Quick Access Management Script
# This script unpins all folders from Quick Access and pins \\tsclient\home
# Avoids operations requiring SYSTEM access to avoid "access denied" prompts

# Function to unpin all folders from Quick Access
function Remove-AllQuickAccessPins {
    Write-Host "Removing all Quick Access pins..." -ForegroundColor Yellow

    try {
        # Get the Quick Access object
        $qa = New-Object -ComObject shell.application
        $quickAccess = $qa.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")

        # Get all pinned items
        $pinnedItems = $quickAccess.Items()

        # Remove each pinned item
        foreach ($item in $pinnedItems) {
            try {
                $item.InvokeVerb("unpinfromhome")
                Write-Host "Unpinned: $($item.Name)" -ForegroundColor Green
            }
            catch {
                Write-Host "Could not unpin: $($item.Name)" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "Error accessing Quick Access: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to clear recent folders from Quick Access
function Clear-QuickAccessRecentFolders {
    Write-Host "Clearing recent folders from Quick Access..." -ForegroundColor Yellow

    try {
        # Clear recent folders registry entries
        $recentFoldersPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs\.folder"
        if (Test-Path $recentFoldersPath) {
            Remove-Item -Path $recentFoldersPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Cleared recent folders registry" -ForegroundColor Green
        }

        # Clear Quick Access recent items
        $quickAccessRecentPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        if (Test-Path $quickAccessRecentPath) {
            Set-ItemProperty -Path $quickAccessRecentPath -Name "ShowFrequent" -Value 0 -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $quickAccessRecentPath -Name "ShowRecent" -Value 0 -ErrorAction SilentlyContinue
            # Re-enable them but they'll be cleared
            Set-ItemProperty -Path $quickAccessRecentPath -Name "ShowFrequent" -Value 1 -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $quickAccessRecentPath -Name "ShowRecent" -Value 1 -ErrorAction SilentlyContinue
            Write-Host "Reset recent folders display settings" -ForegroundColor Green
        }

        Write-Host "Cleared recent folders from Quick Access" -ForegroundColor Green
    }
    catch {
        Write-Host "Error clearing recent folders: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to remove Gallery and OneDrive from File Explorer
function Remove-ExplorerPins {
    Write-Host "Removing Gallery and OneDrive from File Explorer..." -ForegroundColor Yellow

    try {
        # Remove Gallery from File Explorer sidebar
        $galleryRegPath = "HKCU:\Software\Classes\CLSID\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c}"
        if (Test-Path $galleryRegPath) {
            Set-ItemProperty -Path $galleryRegPath -Name "System.IsPinnedToNameSpaceTree" -Value 0 -ErrorAction SilentlyContinue
            Write-Host "Removed Gallery from File Explorer" -ForegroundColor Green
        } else {
            # Create the key if it doesn't exist and disable it
            New-Item -Path $galleryRegPath -Force | Out-Null
            Set-ItemProperty -Path $galleryRegPath -Name "System.IsPinnedToNameSpaceTree" -Value 0
            Write-Host "Created registry entry and removed Gallery from File Explorer" -ForegroundColor Green
        }

        # Remove OneDrive from File Explorer sidebar
        $oneDriveRegPaths = @(
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
            "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
        )

        foreach ($path in $oneDriveRegPaths) {
            if (Test-Path $path) {
                try {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "Removed OneDrive registry entry: $path" -ForegroundColor Green
                }
                catch {
                    Write-Host "Could not remove OneDrive registry entry: $path" -ForegroundColor Yellow
                }
            }
        }

        # Additional OneDrive removal - disable via policy
        $oneDrivePolicyPath = "HKCU:\Software\Microsoft\OneDrive"
        if (Test-Path $oneDrivePolicyPath) {
            Set-ItemProperty -Path $oneDrivePolicyPath -Name "DisableFileSyncNGSC" -Value 1 -ErrorAction SilentlyContinue
        } else {
            New-Item -Path $oneDrivePolicyPath -Force | Out-Null
            Set-ItemProperty -Path $oneDrivePolicyPath -Name "DisableFileSyncNGSC" -Value 1
        }

        Write-Host "Completed Gallery and OneDrive removal from File Explorer" -ForegroundColor Green
    }
    catch {
        Write-Host "Error removing Explorer pins: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to pin network folder to Quick Access
function Is-RDPSession {
    return [bool](Get-CimInstance -ClassName Win32_TerminalServiceSetting -Namespace root\cimv2\TerminalServices)
}

function Add-NetworkFolderToQuickAccess {
    param(
        [string]$NetworkPath = "\\tsclient\home"
    )

    Write-Host "Pinning $NetworkPath to Quick Access..." -ForegroundColor Yellow

    if (Is-RDPSession -and (Test-Path $NetworkPath)) {
        try {
            $shell = New-Object -ComObject shell.application
            $folder = $shell.Namespace($NetworkPath)

            if ($folder) {
                $folder.Self.InvokeVerb("pintohome")
                Write-Host "Successfully pinned $NetworkPath to Quick Access using Shell" -ForegroundColor Green
            }
            else {
                Write-Host "Could not access folder object for $NetworkPath" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Error pinning network folder via Shell: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Falling back to registry method..." -ForegroundColor Yellow
            Add-NetworkFolderViaRegistry -NetworkPath $NetworkPath
        }
    }
    else {
        Write-Host "Not in RDP session or path inaccessible, using registry method..." -ForegroundColor Yellow
        Add-NetworkFolderViaRegistry -NetworkPath $NetworkPath
    }
}

# Alternative method to add network folder via registry
function Add-NetworkFolderViaRegistry {
    param(
        [string]$NetworkPath = "\\tsclient\home",
        [string]$DisplayName = "Home"
    )

    try {
        # Normalize the network path for comparison
        $normalizedPath = $NetworkPath -replace '\\\\', '\'

        # Check if the path is already pinned
        $quickAccessPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace"
        $existingPin = Get-ChildItem -Path $quickAccessPath | Where-Object {
            (Get-ItemProperty -Path $_.PSPath -Name "(Default)" -ErrorAction SilentlyContinue).'(Default)' -eq $normalizedPath
        }

        if ($existingPin) {
            Write-Host "Network path $NetworkPath is already pinned to Quick Access" -ForegroundColor Yellow
            return
        }

        # Create a unique GUID for the network location
        $guid = [System.Guid]::NewGuid().ToString()

        # Registry path for Quick Access items
        $regPath = "$quickAccessPath\{$guid}"

        # Create the registry key
        New-Item -Path $regPath -Force | Out-Null

        # Set properties for the pinned item
        Set-ItemProperty -Path $regPath -Name "(Default)" -Value $NetworkPath
        Set-ItemProperty -Path $regPath -Name "LocalizedResourceName" -Value $DisplayName
        # Indicate it's a network location
        Set-ItemProperty -Path $regPath -Name "System.IsPinnedToNameSpaceTree" -Value 1 -Type DWord

        Write-Host "Added network folder to Quick Access via registry: $NetworkPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Error adding network folder via registry: $($_.Exception.Message)" -ForegroundColor Red
    }
}


# Function to restart Explorer to apply changes
function Restart-Explorer {
    Write-Host "Restarting Windows Explorer to apply changes..." -ForegroundColor Yellow

    try {
        # Kill Explorer process
        Get-Process explorer | Stop-Process -Force

        # Wait a moment
        Start-Sleep -Seconds 2

        # Start Explorer again
        Start-Process explorer.exe

        Write-Host "Explorer restarted successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Error restarting Explorer: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Main execution
Write-Host "=== Quick Access Management Script ===" -ForegroundColor Cyan
Write-Host "This script will:" -ForegroundColor Cyan
Write-Host "- Unpin all folders from Quick Access" -ForegroundColor Cyan
Write-Host "- Clear recent folders from Quick Access" -ForegroundColor Cyan
Write-Host "- Remove Gallery and OneDrive from File Explorer" -ForegroundColor Cyan
Write-Host "- Pin \\tsclient\home to Quick Access" -ForegroundColor Cyan
Write-Host "- Restart Windows Explorer" -ForegroundColor Cyan
Write-Host ""

# Step 1: Remove all current pins
Remove-AllQuickAccessPins

# Step 2: Clear recent folders
Clear-QuickAccessRecentFolders

# Step 3: Remove Gallery and OneDrive from File Explorer
Remove-ExplorerPins

# Step 4: Add network folder
Add-NetworkFolderToQuickAccess

# Step 5: Automatically restart Explorer
Write-Host "Restarting Windows Explorer..." -ForegroundColor Yellow
Restart-Explorer

Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Cyan
Write-Host "File Explorer has been restarted with your new settings." -ForegroundColor Green
Write-Host ""
Write-Host "If the network folder didn't pin automatically, you may need to:" -ForegroundColor Yellow
Write-Host "1. Ensure you're in an RDP session" -ForegroundColor Yellow
Write-Host "2. Navigate to \\tsclient\home in Explorer" -ForegroundColor Yellow
Write-Host "3. Right-click and select 'Pin to Quick access'" -ForegroundColor Yellow
