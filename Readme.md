# Windows 11 Audio Troubleshooting Guide

This guide provides step-by-step instructions for troubleshooting and fixing audio issues on Windows 11, with a special focus on Realtek audio drivers.

## Quick Diagnosis

Your system shows:
- Realtek High Definition Audio device with "Error" status
- Service: IntcAzAudAddService (Intel Audio service)
- Multiple audio endpoints with "Unknown" status

## Step 1: Basic Troubleshooting

- **Check physical connections**: Ensure speakers/headphones are properly connected
- **Check volume settings**: Make sure volume isn't muted and correct output device is selected
- **Run Windows troubleshooter**: Right-click sound icon > Troubleshoot sound problems

## Step 2: Fix Realtek Driver Issues

### Option A: Complete Driver Reinstallation (Recommended)

1. Open Device Manager (right-click Start > Device Manager)
2. Expand "Sound, video and game controllers"
3. Right-click on "Realtek High Definition Audio"
4. Select "Uninstall device"
5. **Important**: Check the box that says "Delete the driver software for this device"
6. Click Uninstall
7. Restart your computer
8. Windows should automatically reinstall basic drivers

### Option B: Download Fresh Drivers

Based on your system ID (SUBSYS_1462CC37), you appear to have an MSI motherboard:

1. Visit [MSI's support website](https://www.msi.com/support)
2. Enter your motherboard model
3. Download the latest audio drivers specifically for your model
4. Run the installer and follow the prompts
5. Restart your computer

## Step 3: Check for Conflicts

### Audio Service Conflicts

Your system is showing Intel Audio service (IntcAzAudAddService) with Realtek hardware:

1. Open Run dialog (Win+R)
2. Type `services.msc` and press Enter
3. Find "Windows Audio" and "Windows Audio Endpoint Builder"
4. Ensure both are set to "Automatic" and are "Running"
5. Look for any Intel or Realtek audio services and check their status

### Audio Software Conflicts

You have multiple audio devices that might be causing conflicts:

1. THX Spatial
2. Steam Streaming Microphone/Speakers
3. XSplit Stream Audio Renderer

Try temporarily disabling these in Device Manager to see if Realtek audio starts working.

## Step 4: Advanced Fixes

### Fix Windows Audio Components

Run PowerShell as Administrator and execute:

```powershell
# Stop audio services
Stop-Service -Name Audiosrv -Force
Stop-Service -Name AudioEndpointBuilder -Force

# Wait a moment
Start-Sleep -Seconds 3

# Start audio services
Start-Service -Name AudioEndpointBuilder
Start-Service -Name Audiosrv
```

### Check System Files

Run Command Prompt as Administrator and execute:

```cmd
sfc /scannow
DISM /Online /Cleanup-Image /RestoreHealth
```

## Step 5: Hardware Troubleshooting

If software fixes don't resolve the issue:

1. Check if audio works in BIOS/UEFI (indicates hardware is functional)
2. Try different audio ports (front/back)
3. Try a USB audio adapter if available

## Diagnostics Information

Your Realtek Audio Device Details:
```
Status       : Error
FriendlyName : Realtek High Definition Audio
InstanceId   : HDAUDIO\FUNC_01&VEN_10EC&DEV_1168&SUBSYS_1462CC37&REV_1001\5&2E660546&0&0001
Service      : IntcAzAudAddService
```

## Python Diagnostic Scripts

The repository includes two Python scripts for audio testing:
- `simple-audio-test.py`: Tests basic Windows audio functionality
- `realtek-audio-fix.py`: Specifically troubleshoots Realtek driver issues

### Requirements
- Python 3.6+
- Admin rights (for some operations)

### Usage
```
python simple-audio-test.py
python realtek-audio-fix.py
```

Run as administrator for full functionality.
