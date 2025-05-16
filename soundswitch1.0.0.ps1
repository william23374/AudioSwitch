#Requires -Module AudioDeviceCmdlets

<#
.SYNOPSIS
    Manages system audio devices, allowing listing and setting of default devices.
.DESCRIPTION
    This script provides a menu-driven interface to:
    - List all available playback AND recording audio devices.
    - Display the current default playback and recording devices.
    - Set a new default playback device by ID.
    - Set a new default recording device by ID.
    - Manage audio profiles (combinations of playback/recording devices) via Advanced Settings.
    The audio profiles JSON file is stored in the same directory as the script.
    It requires the 'AudioDeviceCmdlets' module to be installed.
.NOTES
    Version: 1.0.0
    Creation Date: 20250516
    Author: William Liu
    Email: william23374@163.com
    Internal Script Version: 2.1.1 (Profile JSON stored in script's directory)
    Requires: PowerShell 5.1 or higher, AudioDeviceCmdlets module.
    To install AudioDeviceCmdlets: Install-Module -Name AudioDeviceCmdlets -Scope CurrentUser -Force
#>

# --- Configuration ---
# MODIFIED: Determine profile path based on script location
if ($PSScriptRoot) {
    $Global:AudioProfilesPath = $PSScriptRoot
} else {
    # Fallback to current working directory if script is not run from a file (e.g., pasted in console)
    $Global:AudioProfilesPath = $PWD.Path
    Write-Warning "PSScriptRoot not found. Profiles will be saved in current working directory: $($Global:AudioProfilesPath)"
}
$Global:AudioProfilesFile = Join-Path $Global:AudioProfilesPath "audioprofiles.json"
Write-Verbose "Audio profiles will be stored in: $($Global:AudioProfilesFile)"


# Ensure the module is available
if (-not (Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
    Write-Warning "The 'AudioDeviceCmdlets' module is not found."
    Write-Host "Please install it by running: Install-Module -Name AudioDeviceCmdlets -Scope CurrentUser -Force"
    Read-Host "Press Enter to exit."
    Exit 1
}

try {
    Import-Module AudioDeviceCmdlets -ErrorAction Stop
    $Global:ModuleVersion = (Get-Module AudioDeviceCmdlets).Version
    Write-Host "AudioDeviceCmdlets module version $ModuleVersion imported successfully." -ForegroundColor DarkGray
}
catch {
    Write-Error "Failed to import AudioDeviceCmdlets module: $($_.Exception.Message)"
    Read-Host "Press Enter to exit."
    Exit 1
}

# --- Helper Functions for Audio Profiles (JSON) ---
function Get-AudioProfiles {
    if (-not (Test-Path $Global:AudioProfilesFile)) {
        return @()
    }
    try {
        $content = Get-Content -Path $Global:AudioProfilesFile -Raw -ErrorAction SilentlyContinue
        if ($null -eq $content -or [string]::IsNullOrWhiteSpace($content)) {
            return @()
        }

        $parsedProfiles = $content | ConvertFrom-Json -ErrorAction Stop

        if ($null -eq $parsedProfiles) {
            return @()
        }

        if ($parsedProfiles -isnot [System.Array]) {
            return @($parsedProfiles)
        } else {
            return $parsedProfiles
        }
    }
    catch {
        Write-Warning "Error reading or parsing profiles from $($Global:AudioProfilesFile): $($_.Exception.Message)"
        return @()
    }
}

function Save-AudioProfiles {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Profiles
    )
    try {
        # Directory for JSON is now $Global:AudioProfilesPath (script's dir or PWD)
        # Test-Path for the directory itself isn't strictly necessary here if $Global:AudioProfilesPath is always valid
        # but New-Item for directory would fail if path is bad.
        # For simplicity, since $Global:AudioProfilesPath should be the script's directory,
        # we assume it exists. If not, Set-Content will create the file in a valid path or fail.
        # If $Global:AudioProfilesPath was more complex (e.g., a subfolder), directory creation would be important.

        if ($Profiles.Count -eq 0) {
            Set-Content -Path $Global:AudioProfilesFile -Value "[]" -Force -ErrorAction Stop
        } else {
            $Profiles | ConvertTo-Json -Depth 5 | Set-Content -Path $Global:AudioProfilesFile -Force -ErrorAction Stop
        }
        Write-Verbose "Profiles saved to $($Global:AudioProfilesFile)"
    }
    catch {
        Write-Error "Failed to save audio profiles to $($Global:AudioProfilesFile): $($_.Exception.Message)"
    }
}

# --- Core Audio Device Functions ---
function Show-PlaybackDevicesInternal {
    Write-Host "`n--- Available Playback Devices ---" -ForegroundColor Cyan
    $playbackDevices = Get-AudioDevice -List | Where-Object {$_.Type -eq 'Playback'}
    if (-not $playbackDevices) {
        Write-Host "No playback devices found."
        return $null
    }
    $playbackDevicesArray = @($playbackDevices)
    $playbackDevicesArray | ForEach-Object -Begin {$i = 1} -Process {
        $defaultMarker = if ($_.Default) {"*"} else {" "}
        Write-Host ("{0}. [{1}] {2}" -f $i++, $defaultMarker, $_.Name)
    }
    return $playbackDevicesArray
}

function Show-RecordingDevicesInternal {
    Write-Host "`n--- Available Recording Devices ---" -ForegroundColor Cyan
    $recordingDevices = Get-AudioDevice -List | Where-Object {$_.Type -eq 'Recording'}
    if (-not $recordingDevices) {
        Write-Host "No recording devices found."
        return $null
    }
    $recordingDevicesArray = @($recordingDevices)
    $recordingDevicesArray | ForEach-Object -Begin {$i = 1} -Process {
        $defaultMarker = if ($_.Default) {"*"} else {" "}
        Write-Host ("{0}. [{1}] {2}" -f $i++, $defaultMarker, $_.Name)
    }
    return $recordingDevicesArray
}

function List-AllAudioDevices {
    Write-Host "`n====== System Audio Devices ======" -ForegroundColor White
    $null = Show-PlaybackDevicesInternal
    Write-Host
    $null = Show-RecordingDevicesInternal
    Write-Host "================================" -ForegroundColor White
}

function Get-CurrentDefaultDevices {
    Write-Host "`n--- Current Default Audio Devices ---" -ForegroundColor Yellow
    $defaultPlayback = Get-AudioDevice -Playback -ErrorAction SilentlyContinue
    $defaultRecording = Get-AudioDevice -Recording -ErrorAction SilentlyContinue

    if ($defaultPlayback) { Write-Host "Default Playback Device : $($defaultPlayback.Name)" }
    else { Write-Host "Default Playback Device : Not Set or Error retrieving." }

    if ($defaultRecording) { Write-Host "Default Recording Device: $($defaultRecording.Name)" }
    else { Write-Host "Default Recording Device: Not Set or Error retrieving." }
}

function Set-NewDefaultPlaybackDevice {
    $devices = Show-PlaybackDevicesInternal
    if (-not $devices -or $devices.Count -eq 0) { Write-Warning "No playback devices available to set."; return }

    [int]$selection = 0; $validSelection = $false
    while (-not $validSelection) {
        try {
            $input = Read-Host "Enter playback device number to set as default (0 to cancel)"
            if ($input -eq '0') { Write-Host "Operation cancelled."; return }
            $selection = [int]$input
            if ($selection -ge 1 -and $selection -le $devices.Count) { $validSelection = $true }
            else { Write-Warning "Invalid selection. Please enter a number from the list." }
        } catch { Write-Warning "Invalid input. Please enter a numeric value." }
    }
    $selectedDevice = $devices[$selection - 1]
    Write-Host "Setting '$($selectedDevice.Name)' as default playback device..." -FG Green
    try {
        Set-AudioDevice -ID $selectedDevice.ID -EA Stop | Out-Null
        Write-Host "'$($selectedDevice.Name)' is now the default playback device." -FG Green
    }
    catch { Write-Error "Failed to set default playback device: $($_.Exception.Message)" }
}

function Set-NewDefaultRecordingDevice {
    $devices = Show-RecordingDevicesInternal
    if (-not $devices -or $devices.Count -eq 0) { Write-Warning "No recording devices available to set."; return }

    [int]$selection = 0; $validSelection = $false
    while (-not $validSelection) {
        try {
            $input = Read-Host "Enter recording device number to set as default (0 to cancel)"
            if ($input -eq '0') { Write-Host "Operation cancelled."; return }
            $selection = [int]$input
            if ($selection -ge 1 -and $selection -le $devices.Count) { $validSelection = $true }
            else { Write-Warning "Invalid selection. Please enter a number from the list." }
        } catch { Write-Warning "Invalid input. Please enter a numeric value." }
    }
    $selectedDevice = $devices[$selection - 1]
    Write-Host "Setting '$($selectedDevice.Name)' as default recording device..." -FG Green
    try {
        Set-AudioDevice -ID $selectedDevice.ID -EA Stop | Out-Null
        Write-Host "'$($selectedDevice.Name)' is now the default recording device." -FG Green
    }
    catch { Write-Error "Failed to set default recording device: $($_.Exception.Message)"; if ($Global:ModuleVersion) {Write-Host "INFO: Module version: $Global:ModuleVersion"}}
}


# --- Advanced Settings Profile Management Functions ---
function Create-NewAudioProfile {
    Write-Host "`n--- Create New Audio Profile ---" -ForegroundColor Yellow

    $loadedProfiles = Get-AudioProfiles
    if ($null -eq $loadedProfiles -or $loadedProfiles.Count -eq 0) {
        $profiles = @()
    } elseif ($loadedProfiles -isnot [System.Array]) {
        $profiles = @($loadedProfiles)
    } else {
        $profiles = $loadedProfiles
    }

    $profileName = ""
    while ([string]::IsNullOrWhiteSpace($profileName)) {
        $profileName = Read-Host "Enter a name for this audio profile (e.g., Work, Gaming Headset)"
        if ([string]::IsNullOrWhiteSpace($profileName)) {
            Write-Warning "Profile name cannot be empty. Please enter a valid name."
        }
    }

    $existingProfileIndex = -1
    $tempProfilesForIteration = @($profiles)
    for ($i = 0; $i -lt $tempProfilesForIteration.Count; $i++) {
        if ($tempProfilesForIteration[$i].ProfileName -eq $profileName) {
            $existingProfileIndex = $i
            break
        }
    }

    if ($existingProfileIndex -ne -1) {
        Write-Warning "A profile named '$profileName' already exists."
        $overwrite = Read-Host "Do you want to overwrite the existing profile? (Y/N)"
        if ($overwrite.ToUpper() -ne 'Y') {
            Write-Host "Profile creation cancelled."
            return
        }
        $newProfilesList = @()
        for ($i = 0; $i -lt $tempProfilesForIteration.Count; $i++) {
            if ($i -ne $existingProfileIndex) {
                $newProfilesList += $tempProfilesForIteration[$i]
            }
        }
        $profiles = $newProfilesList
    }

    $playbackDevices = Show-PlaybackDevicesInternal
    if (-not $playbackDevices -or $playbackDevices.Count -eq 0) { Write-Warning "Cannot create profile: No playback devices found."; return }

    [int]$pbSelection = 0; $pbValid = $false
    while (-not $pbValid) {
        try {
            $pbInput = Read-Host "Select PLAYBACK device for profile '$profileName' (enter number, 0 to cancel)"
            if ($pbInput -eq '0') { Write-Host "Profile creation cancelled."; return }
            $pbSelection = [int]$pbInput
            if ($pbSelection -ge 1 -and $pbSelection -le $playbackDevices.Count) { $pbValid = $true }
            else { Write-Warning "Invalid selection. Please enter a number from the list." }
        } catch { Write-Warning "Invalid input. Please enter a numeric value." }
    }
    $selectedPlaybackDevice = $playbackDevices[$pbSelection - 1]

    $recordingDevices = Show-RecordingDevicesInternal
    if (-not $recordingDevices -or $recordingDevices.Count -eq 0) { Write-Warning "Cannot create profile: No recording devices found."; return }

    [int]$recSelection = 0; $recValid = $false
    while (-not $recValid) {
        try {
            $recInput = Read-Host "Select RECORDING device for profile '$profileName' (enter number, 0 to cancel)"
            if ($recInput -eq '0') { Write-Host "Profile creation cancelled."; return }
            $recSelection = [int]$recInput
            if ($recSelection -ge 1 -and $recSelection -le $recordingDevices.Count) { $recValid = $true }
            else { Write-Warning "Invalid selection. Please enter a number from the list." }
        } catch { Write-Warning "Invalid input. Please enter a numeric value." }
    }
    $selectedRecordingDevice = $recordingDevices[$recSelection - 1]

    $newProfile = [PSCustomObject]@{
        ProfileName         = $profileName
        PlaybackDeviceID    = $selectedPlaybackDevice.ID
        PlaybackDeviceName  = $selectedPlaybackDevice.Name
        RecordingDeviceID   = $selectedRecordingDevice.ID
        RecordingDeviceName = $selectedRecordingDevice.Name
    }

    $profiles = $profiles + $newProfile
    Save-AudioProfiles -Profiles $profiles
    Write-Host "Profile '$profileName' saved successfully." -ForegroundColor Green
    Write-Host "  Playback Device : $($selectedPlaybackDevice.Name)"
    Write-Host "  Recording Device: $($selectedRecordingDevice.Name)"
}

function List-SavedAudioProfiles {
    Write-Host "`n--- Saved Audio Profiles ---" -ForegroundColor Yellow
    $profiles = Get-AudioProfiles

    if ($null -eq $profiles -or $profiles.Count -eq 0) {
        Write-Host "No audio profiles have been saved."
        return $null
    }
    # MODIFIED: Display the location of the profile file
    Write-Host "Profiles loaded from: $Global:AudioProfilesFile"
    @($profiles) | ForEach-Object -Begin {$i = 1} -Process {
        Write-Host ("{0}. Profile Name: {1}" -f $i++, $_.ProfileName)
        Write-Host ("     Playback Device : {0}" -f $_.PlaybackDeviceName)
        Write-Host ("     Recording Device: {0}" -f $_.RecordingDeviceName)
    }
    return $profiles
}

function Apply-AudioProfile {
    Write-Host "`n--- Apply Audio Profile ---" -ForegroundColor Yellow
    $profilesFromList = List-SavedAudioProfiles

    if ($null -eq $profilesFromList) { return }

    $profiles = @($profilesFromList)
    if ($profiles.Count -eq 0) { Write-Host "No profiles available to apply."; return }

    [int]$selection = 0; $validSelection = $false
    while (-not $validSelection) {
        try {
            $input = Read-Host "Enter profile number to apply (0 to cancel)"
            if ($input -eq '0') { Write-Host "Operation cancelled."; return }
            $selection = [int]$input

            if ($selection -ge 1 -and $selection -le $profiles.Count) {
                $validSelection = $true
            } else {
                Write-Warning "Invalid selection. Please enter a number from the list."
            }
        } catch {
            Write-Warning "Invalid input. Please enter a numeric value."
        }
    }
    $selectedProfile = $profiles[$selection - 1]

    Write-Host "`nApplying profile: '$($selectedProfile.ProfileName)'..." -ForegroundColor Cyan

    $playbackDeviceExists = Get-AudioDevice -ID $selectedProfile.PlaybackDeviceID -ErrorAction SilentlyContinue
    $recordingDeviceExists = Get-AudioDevice -ID $selectedProfile.RecordingDeviceID -ErrorAction SilentlyContinue
    $pbSuccess = $false
    $recSuccess = $false

    if ($playbackDeviceExists) {
        Write-Host "Setting playback device: '$($selectedProfile.PlaybackDeviceName)'..."
        try {
            Set-AudioDevice -ID $selectedProfile.PlaybackDeviceID -ErrorAction Stop | Out-Null
            Write-Host " -> Playback device set." -FG Green
            $pbSuccess = $true
        }
        catch { Write-Error " -> Failed to set playback device: $($_.Exception.Message)" }
    } else {
        Write-Warning "The playback device '$($selectedProfile.PlaybackDeviceName)' (ID: $($selectedProfile.PlaybackDeviceID)) from profile '$($selectedProfile.ProfileName)' was not found on this system."
    }

    if ($recordingDeviceExists) {
        Write-Host "Setting recording device: '$($selectedProfile.RecordingDeviceName)'..."
        try {
            Set-AudioDevice -ID $selectedProfile.RecordingDeviceID -ErrorAction Stop | Out-Null
            Write-Host " -> Recording device set." -FG Green
            $recSuccess = $true
        }
        catch { Write-Error " -> Failed to set recording device: $($_.Exception.Message)" }
    } else {
        Write-Warning "The recording device '$($selectedProfile.RecordingDeviceName)' (ID: $($selectedProfile.RecordingDeviceID)) from profile '$($selectedProfile.ProfileName)' was not found on this system."
    }

    if ($pbSuccess -and $recSuccess) {
        Write-Host "`nProfile '$($selectedProfile.ProfileName)' applied successfully." -ForegroundColor Green
    } elseif ($pbSuccess -or $recSuccess) {
        Write-Host "`nProfile '$($selectedProfile.ProfileName)' partially applied." -ForegroundColor Yellow
    } else {
        Write-Warning "`nProfile '$($selectedProfile.ProfileName)' could not be applied as one or more devices were not found or errors occurred."
    }
}

function Delete-AudioProfile {
    Write-Host "`n--- Delete Audio Profile ---" -ForegroundColor Yellow
    $profilesFromList = List-SavedAudioProfiles

    if ($null -eq $profilesFromList) { return }

    $profiles = @($profilesFromList)
    if ($profiles.Count -eq 0) { Write-Host "No profiles available to delete."; return }

    [int]$selection = 0; $validSelection = $false
    while (-not $validSelection) {
        try {
            $input = Read-Host "Enter profile number to DELETE (0 to cancel)"
            if ($input -eq '0') { Write-Host "Operation cancelled."; return }
            $selection = [int]$input
            if ($selection -ge 1 -and $selection -le $profiles.Count) { $validSelection = $true }
            else { Write-Warning "Invalid selection. Please enter a number from the list." }
        } catch { Write-Warning "Invalid input. Please enter a numeric value." }
    }
    $profileToDelete = $profiles[$selection - 1]
    $confirm = Read-Host "Are you sure you want to permanently delete the profile '$($profileToDelete.ProfileName)'? (Y/N)"
    if ($confirm.ToUpper() -ne 'Y') {
        Write-Host "Deletion cancelled."
        return
    }

    $updatedProfiles = @()
    foreach ($p in $profiles) {
        if ($p.ProfileName -ne $profileToDelete.ProfileName) {
            $updatedProfiles += $p
        }
    }
    Save-AudioProfiles -Profiles $updatedProfiles
    Write-Host "Profile '$($profileToDelete.ProfileName)' has been deleted." -ForegroundColor Green
}

function Show-AdvancedSettingsMenu {
    $advChoice = ''
    do {
        Clear-Host
        Write-Host "`n------- Advanced Audio Settings -------" -ForegroundColor Magenta
        # MODIFIED: Display the location of the profile file in the submenu as well
        Write-Host "Profile file location: $Global:AudioProfilesFile"
        Write-Host "1. Create Audio Profile"
        Write-Host "2. List Saved Audio Profiles"
        Write-Host "3. Apply Audio Profile"
        Write-Host "4. Delete Audio Profile"
        Write-Host "R. Return to Main Menu"
        Write-Host "-----------------------------------"

        $advChoice = Read-Host "Enter selection"

        switch ($advChoice.ToUpper()) {
            '1' { Create-NewAudioProfile }
            '2' { $null = List-SavedAudioProfiles }
            '3' { Apply-AudioProfile }
            '4' { Delete-AudioProfile }
            'R' { Write-Host "Returning to Main Menu..." }
            default { Write-Warning "Invalid selection. Please try again." }
        }

        if ($advChoice.ToUpper() -ne 'R') {
            Read-Host "`nPress Enter to continue..."
        }

    } while ($advChoice.ToUpper() -ne 'R')
}

# --- Main Menu Loop ---
do {
    Clear-Host
    Write-Host "`n======= Audio Device Manager (v2.1.1) =======" -ForegroundColor White # This is the internal script functionality version
    Write-Host "1. List All Audio Devices"
    Write-Host "2. Display Current Default Devices"
    Write-Host "3. Set Default Playback Device"
    Write-Host "4. Set Default Recording Device"
    Write-Host "5. Advanced Audio Settings"
    Write-Host "Q. Quit"
    Write-Host "=========================================="

    $choice = Read-Host "Enter selection"

    switch ($choice.ToUpper()) {
        '1' { List-AllAudioDevices }
        '2' { Get-CurrentDefaultDevices }
        '3' { Set-NewDefaultPlaybackDevice }
        '4' { Set-NewDefaultRecordingDevice }
        '5' { Show-AdvancedSettingsMenu }
        'Q' { Write-Host "Exiting application..." }
        default { Write-Warning "Invalid selection. Please try again." }
    }

    if ($choice.ToUpper() -notin @('Q', '5')) {
        Read-Host "`nPress Enter to continue..."
    }

} while ($choice.ToUpper() -ne 'Q')

Write-Host "Audio Device Manager closed."