#!/usr/bin/env python3
"""
Simple Windows Audio Test Script
This script tests basic Windows audio functionality without requiring many dependencies.
"""

import os
import subprocess
import time
import ctypes
import platform
import sys

def print_header(text):
    """Print a formatted header."""
    # Pause before displaying a new section header
    if text != "Windows Audio Test Script":  # Don't pause before the very first header
        input("\nPress Enter to continue to the next section...")
    
    print("\n" + "=" * 60)
    print(f" {text}")
    print("=" * 60)
    
def pause_after_error(error_message):
    """Print error message and pause for user to read it."""
    print(f"ERROR: {error_message}")
    input("Press Enter to continue...")

def check_admin_privileges():
    """Check if the script is running with admin privileges."""
    print_header("Checking Admin Privileges")
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        print(f"Running with admin privileges: {is_admin}")
        if not is_admin:
            print("NOTE: Some tests may fail without admin privileges.")
            print("Consider running this script as administrator.")
        return is_admin
    except Exception as e:
        pause_after_error(f"Error checking admin privileges: {e}")
        return False

def check_windows_version():
    """Check if running on Windows 11."""
    print_header("Checking Windows Version")
    try:
        version = platform.version()
        build = platform.release()
        build_number = 0
        
        # Try to extract actual build number for better detection
        try:
            # The version string format is usually like "10.0.26100"
            # so we try to get the last part as the build number
            build_number = int(version.split('.')[-1])
        except:
            pass
            
        print(f"Windows Version: {version}")
        print(f"Build: {build}")
        print(f"Build Number: {build_number}")
        
        # Check if Windows 11 (more reliable detection)
        # Windows 11 is build 22000 or higher
        is_win11 = build_number >= 22000
        print(f"Is Windows 11: {is_win11}")
        
        if not is_win11:
            print("Warning: This script is designed for Windows 11.")
            print("However, detection might be incorrect. If you're running Windows 11, please ignore this warning.")
        
        return is_win11
    except Exception as e:
        pause_after_error(f"Error checking Windows version: {e}")
        return False

def check_audio_services():
    """Check if Windows audio services are running using sc query."""
    print_header("Checking Windows Audio Services")
    
    services_to_check = [
        "Audiosrv",          # Windows Audio
        "AudioEndpointBuilder"  # Windows Audio Endpoint Builder
    ]
    
    try:
        for service_name in services_to_check:
            cmd = f"sc query {service_name}"
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if "STATE" in result.stdout:
                print(f"Service: {service_name}")
                
                # Extract state information
                state_line = [line for line in result.stdout.split('\n') if "STATE" in line][0]
                state_value = state_line.split(':')[1].strip()
                
                print(f"  Status: {state_value}")
                
                if "RUNNING" not in state_value:
                    print(f"  WARNING: {service_name} is not running!")
                    print(f"  Consider starting it with: net start {service_name}")
            else:
                print(f"Service {service_name} not found!")
    except Exception as e:
        pause_after_error(f"Error checking audio services: {e}")

def list_audio_devices_with_windows_cmd():
    """List audio devices using PowerShell commands."""
    print_header("Listing Audio Devices")
    
    try:
        # PowerShell command to list audio devices
        cmd = 'powershell -command "Get-CimInstance -Class Win32_SoundDevice | Select-Object Name, Status"'
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            print("Found audio devices:")
            print(result.stdout)
        else:
            print("No audio devices found or command failed.")
            if result.stderr:
                print("Error:", result.stderr)
                input("Press Enter to continue...")
            
            # Alternative approach
            print("\nTrying alternative method...")
            cmd2 = 'powershell -command "Get-WmiObject -Class Win32_SoundDevice | Format-List Name, Status"'
            result2 = subprocess.run(cmd2, capture_output=True, text=True)
            
            if result2.returncode == 0 and result2.stdout.strip():
                print(result2.stdout)
            else:
                print("Alternative method also failed.")
                if result2.stderr:
                    print("Error:", result2.stderr)
                    input("Press Enter to continue...")
    except Exception as e:
        pause_after_error(f"Error listing audio devices: {e}")

def check_volume_settings_with_cmd():
    """Check system volume settings using PowerShell."""
    print_header("Checking Volume Settings")
    
    try:
        # Check if system is muted 
        cmd = 'powershell -command "[Audio]::Mute"'
        
        # Alternative approach
        cmd2 = 'powershell -command "Get-WmiObject -Class Win32_SoundDevice | Format-List StatusInfo"'
        result = subprocess.run(cmd2, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("Audio device status:")
            print(result.stdout)
        else:
            print("Could not get volume/mute status.")
            print("This requires PowerShell modules that may not be installed.")
            print("You can manually check if your system is muted by looking at the volume icon in the taskbar.")
    except Exception as e:
        pause_after_error(f"Error checking volume: {e}")

def play_test_sound():
    """Play Windows default test sound using PowerShell."""
    print_header("Playing Windows Test Sound")
    
    try:
        print("Attempting to play Windows default sound...")
        cmd = 'powershell -command "[System.Media.SystemSounds]::Beep.Play()"'
        subprocess.run(cmd, shell=True)
        time.sleep(1)
        
        print("Attempting to play Windows exclamation sound...")
        cmd = 'powershell -command "[System.Media.SystemSounds]::Exclamation.Play()"'
        subprocess.run(cmd, shell=True)
        time.sleep(1)
        
        print("Did you hear the system sounds? (y/n)")
        response = input().lower()
        if response.startswith('y'):
            print("Great! Your audio is working.")
        else:
            print("Audio test failed. No sound was heard.")
    except Exception as e:
        pause_after_error(f"Error playing system sound: {e}")

def generate_test_tone_powershell():
    """Generate and play a test tone using PowerShell."""
    print_header("Playing Test Tone")
    
    try:
        print("Attempting to play a 1000Hz test tone for 3 seconds...")
        ps_script = """
        [console]::beep(1000,3000)
        """
        cmd = f'powershell -command "{ps_script}"'
        subprocess.run(cmd, shell=True)
        
        print("Did you hear the test tone? (y/n)")
        response = input().lower()
        if response.startswith('y'):
            print("Great! Your audio is working.")
        else:
            print("Test tone failed. No sound was heard.")
    except Exception as e:
        pause_after_error(f"Error playing test tone: {e}")

def run_windows_troubleshooter():
    """Run the Windows audio troubleshooter."""
    print_header("Running Windows Audio Troubleshooter")
    
    try:
        print("Launching Windows audio troubleshooter...")
        cmd = "msdt.exe /id AudioPlaybackDiagnostic"
        subprocess.Popen(cmd, shell=True)
        print("Troubleshooter launched. Please follow the on-screen instructions.")
    except Exception as e:
        pause_after_error(f"Error launching troubleshooter: {e}")
        print("Try running it manually: Settings > System > Troubleshoot > Other troubleshooters > Playing Audio")

def check_default_device():
    """Check the default audio device."""
    print_header("Checking Default Audio Device")
    
    try:
        cmd = 'powershell -command "Get-WmiObject -Class Win32_SoundDevice | Where-Object {$_.Status -eq \'OK\'} | Select-Object -First 1 Name"'
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            print(f"Default audio device: {result.stdout.strip()}")
        else:
            print("Could not determine default audio device.")
            print("You can check this manually in Sound settings.")
    except Exception as e:
        pause_after_error(f"Error checking default device: {e}")

def main():
    """Run all audio tests."""
    print_header("Windows Audio Test Script")
    
    print("This script will test various aspects of your Windows audio system.")
    print("Some tests may require administrator privileges.")
    print("Make sure your speakers/headphones are connected and at a comfortable volume.")
    print("\nThe script will pause between sections so you can read all information.")
    print("Press Enter when prompted to continue to the next test.")
    
    # Run tests
    check_admin_privileges()
    check_windows_version()
    check_audio_services()
    list_audio_devices_with_windows_cmd()
    check_volume_settings_with_cmd()
    check_default_device()
    
    # Interactive tests
    input("\nPress Enter to continue with audio playback tests...")
    play_test_sound()
    
    input("\nPress Enter to play a test tone...")
    generate_test_tone_powershell()
    
    # Final step
    print("\nDo you want to run the Windows Audio Troubleshooter? (y/n)")
    choice = input().lower()
    if choice.startswith('y'):
        run_windows_troubleshooter()
    
    print_header("Audio Tests Completed")
    print("If you're still having audio issues, consider:")
    print("1. Updating or reinstalling audio drivers")
    print("2. Checking for Windows updates")
    print("3. Verifying hardware connections")
    print("4. Checking application-specific audio settings")

if __name__ == "__main__":
    main()
