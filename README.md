# PowerShell Audio Device Manager (v2.1.1)

Manages system audio devices, allowing listing, setting of default devices, and managing audio profiles through a menu-driven interface.

## Overview

This PowerShell script provides a user-friendly command-line interface to interact with your system's audio playback and recording devices. It simplifies common tasks such as checking current defaults, switching devices, and creating/applying custom audio profiles (e.g., for "Work," "Gaming," "Streaming").

The script requires the `AudioDeviceCmdlets` module from the PowerShell Gallery.

## Features

*   **List Devices:** Display all available playback and recording audio devices with their current default status.
*   **Show Defaults:** Quickly view the currently set default playback and recording devices.
*   **Set Default Playback Device:** Choose from a list to set a new default playback device.
*   **Set Default Recording Device:** Choose from a list to set a new default recording device.
*   **Advanced Audio Profiles:**
    *   **Create Profiles:** Define and save custom audio profiles, each with a specific playback and recording device.
    *   **List Profiles:** View all saved audio profiles.
    *   **Apply Profiles:** Quickly switch both playback and recording devices by applying a saved profile.
    *   **Delete Profiles:** Remove unwanted audio profiles.
*   **Profile Storage:** Audio profiles are saved in an `audioprofiles.json` file located in the same directory as the script. If the script is run from a context where its path isn't defined (e.g., pasted directly into a console), profiles will be saved in the current working directory.

## Requirements

*   **PowerShell:** Version 5.1 or higher.
*   **AudioDeviceCmdlets Module:** This script relies on the `AudioDeviceCmdlets` module.

## Installation

1.  **Download the Script:**
    Save the script content as a `.ps1` file (e.g., `Manage-AudioDevices.ps1`) to your desired location.

2.  **Install AudioDeviceCmdlets Module:**
    If you don't have the module installed, open PowerShell (as Administrator if you want to install for all users, otherwise as a regular user for CurrentUser scope) and run:
    ```powershell
    Install-Module -Name AudioDeviceCmdlets -Scope CurrentUser -Force
    ```
    The script will check for this module on startup and prompt you if it's missing.

## Usage

1.  **Open PowerShell:** Navigate to the directory where you saved the script.
2.  **Run the Script:**
    ```powershell
    .\Manage-AudioDevices.ps1
    ```
    *Note: You might need to adjust your PowerShell execution policy if you haven't run scripts before. You can set it for the current session by running `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process` before running the script.*

3.  **Navigate the Menu:**
    The script will display a main menu with options:
    ```
    ======= Audio Device Manager (v2.1.1) =======
    1. List All Audio Devices
    2. Display Current Default Devices
    3. Set Default Playback Device
    4. Set Default Recording Device
    5. Advanced Audio Settings
    Q. Quit
    ==========================================
    Enter selection:
    ```
    Enter the number corresponding to your desired action and press Enter.

4.  **Advanced Audio Settings:**
    Selecting option `5` will take you to a submenu for managing audio profiles:
    ```
    ------- Advanced Audio Settings -------
    Profile file location: C:\Path\To\Your\Script\audioprofiles.json
    1. Create Audio Profile
    2. List Saved Audio Profiles
    3. Apply Audio Profile
    4. Delete Audio Profile
    R. Return to Main Menu
    -----------------------------------
    Enter selection:
    ```

## Audio Profiles (`audioprofiles.json`)

*   This JSON file stores your custom audio configurations.
*   It is automatically created/updated in the same directory as the script (`$PSScriptRoot`) when you save a profile.
*   If the script cannot determine its own path (e.g., if you paste the script's content directly into a console and run it), the `audioprofiles.json` file will be created/used in the current PowerShell working directory (`$PWD`).
*   **Example `audioprofiles.json` content:**
    ```json
    [
      {
        "ProfileName": "Work Headset",
        "PlaybackDeviceID": "{0.0.0.00000000}.{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}",
        "PlaybackDeviceName": "Headset (My Work Headset)",
        "RecordingDeviceID": "{0.0.1.00000000}.{yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy}",
        "RecordingDeviceName": "Microphone (My Work Headset)"
      },
      {
        "ProfileName": "Speakers and Webcam Mic",
        "PlaybackDeviceID": "{0.0.0.00000000}.{aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa}",
        "PlaybackDeviceName": "Speakers (Realtek Audio)",
        "RecordingDeviceID": "{0.0.1.00000000}.{bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb}",
        "RecordingDeviceName": "Microphone (Webcam)"
      }
    ]
    ```

## Troubleshooting

*   **"The 'AudioDeviceCmdlets' module is not found."**: Ensure you have installed the module using the command in the "Installation" section.
*   **"Failed to import AudioDeviceCmdlets module"**: Try re-installing the module. Ensure PowerShell can access the module path.
*   **Errors setting devices**:
    *   Ensure the device is properly connected and recognized by Windows.
    *   Rarely, some specific device drivers might have compatibility issues with the underlying `AudioDeviceCmdlets` module.
*   **Profile device not found when applying**: If you've changed audio hardware or drivers, the device IDs stored in the profile might no longer be valid. You may need to recreate the profile for the new/changed devices.

## Author

*   **William Liu**
*   Email: `william23374@163.com`

## Version History

*   **Script Functionality Version:** 2.1.1 (Profile JSON stored in script's directory or PWD)
*   **Documentation Version:** 1.0.0 (Initial README based on script v2.1.1)

## License

This script is provided as-is. You are free to use, modify, and distribute it. Please consider giving credit to the original author if you make significant modifications or use it as a base for your own projects.
