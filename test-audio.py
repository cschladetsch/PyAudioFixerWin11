#!/usr/bin/env python3
"""
Realtek Audio Fix Script for Windows 11
This script helps diagnose and potentially fix Realtek audio driver issues.
"""

import os
import subprocess
import time
import ctypes
import platform
import sys

def print_header(text):
    """Print a formatted header."""
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
            print("WARNING: Many driver operations require admin privileges.")
            print("Consider running this script as administrator for full functionality.")
        return is_admin
    except Exception as e:
        pause_after_error(f"Error checking admin privileges: {e}")
        return False

def check_realtek_driver_status():
    """Check detailed status of Realtek audio driver."""
    print_header("Checking Realtek Audio Driver Status")
    
    try:
        # Get detailed device info
        cmd = 'powershell -command "Get-PnpDevice | Where-Object {$_.FriendlyName -like \'*Realtek*Audio*\'} | Select-Object Status, FriendlyName, InstanceId, ProblemCode, Service | Format-List"'
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            print("Realtek driver details:")
            print(result.stdout)
            
            # Extract device instance ID for potential troubleshooting
            instance_id = None
            for line in result.stdout.split('\n'):
                if 'InstanceId' in line:
                    instance_id = line.split(':')[1].strip()
                    break
                    
            if instance_id:
                print(f"Device Instance ID: {instance_id}")
                return instance_id
            else:
                print("Could not find Realtek device instance ID")
                return None
        else:
            print("Could not retrieve Realtek driver information.")
            if result.stderr:
                print("Error:", result.stderr)
                input("Press Enter to continue...")
            return None
    except Exception as e:
        pause_after_error(f"Error checking Realtek driver: {e}")
        return None

def check_driver_events(instance_id=None):
    """Check Windows event logs for audio driver events."""
    print_header("Checking Audio Driver Events")
    
    try:
        if instance_id:
            # Get events specific to this device
            cmd = f'powershell -command "Get-WinEvent -FilterHashtable @{{LogName=\'System\'; ProviderName=\'Microsoft-Windows-DriverFrameworks-UserMode\'}} -MaxEvents 10 | Where-Object {{$_.Message -like \'*{instance_id}*\'}} | Format-List TimeCreated, Message"'
        else:
            # Get general audio events
            cmd = 'powershell -command "Get-WinEvent -FilterHashtable @{LogName=\'System\'; ProviderName=@(\'Microsoft-Windows-Audio\', \'Microsoft-Windows-AudioEndpointBuilder\')} -MaxEvents 10 | Format-List TimeCreated, Message"'
            
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            print("Recent audio driver events:")
            print(result.stdout)
        else:
            print("No relevant audio driver events found.")
            if result.stderr and "No events were found" not in result.stderr:
                print("Error:", result.stderr)
                input("Press Enter to continue...")
    except Exception as e:
        pause_after_error(f"Error checking driver events: {e}")

def attempt_driver_restart():
    """Attempt to restart the Realtek audio driver."""
    print_header("Attempting to Restart Realtek Audio Driver")
    
    try:
        # Check if admin
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("ERROR: Admin privileges required to restart drivers.")
            print("Please run this script as administrator.")
            return False
            
        # Try to restart the driver
        print("Attempting to restart Realtek Audio driver...")
        cmd1 = 'powershell -command "Get-PnpDevice | Where-Object {$_.FriendlyName -like \'*Realtek*Audio*\'} | Disable-PnpDevice -Confirm:$false"'
        subprocess.run(cmd1, capture_output=True, text=True)
        
        print("Waiting 5 seconds...")
        time.sleep(5)
        
        cmd2 = 'powershell -command "Get-PnpDevice | Where-Object {$_.FriendlyName -like \'*Realtek*Audio*\'} | Enable-PnpDevice -Confirm:$false"'
        result = subprocess.run(cmd2, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("Realtek Audio driver successfully restarted.")
            return True
        else:
            print("Failed to restart Realtek Audio driver.")
            if result.stderr:
                print("Error:", result.stderr)
            return False
    except Exception as e:
        pause_after_error(f"Error restarting driver: {e}")
        return False

def download_latest_driver():
    """Provide guidance on downloading the latest Realtek driver."""
    print_header("Downloading Latest Realtek Audio Driver")
    
    print("To download the latest Realtek HD Audio driver, you have several options:")
    print("\n1. From your PC manufacturer's website (recommended):")
    print("   - Visit your PC/motherboard manufacturer's support page")
    print("   - Find your specific model and download drivers from there")
    print("   - This is the recommended method as it has drivers tested for your hardware")
    
    print("\n2. From Realtek's website:")
    print("   - Visit https://www.realtek.com/en/downloads")
    print("   - Look for 'High Definition Audio Codecs' under PC category")
    
    print("\n3. Using Device Manager:")
    print("   - Right-click Start > Device Manager")
    print("   - Expand 'Sound, video and game controllers'")
    print("   - Right-click Realtek Audio device > Update driver")
    print("   - Choose 'Search automatically for updated driver software'")
    
    print("\n4. Using Windows Update:")
    print("   - Go to Settings > Windows Update > Optional updates")
    print("   - Check for driver updates")
    
    choice = input("\nWould you like to open Device Manager to update the driver? (y/n): ")
    if choice.lower().startswith('y'):
        subprocess.Popen("devmgmt.msc", shell=True)
        print("Device Manager opened. Please navigate to:")
        print("Sound, video and game controllers > Realtek Audio > Update driver")

def reset_windows_audio():
    """Reset Windows audio components."""
    print_header("Resetting Windows Audio Components")
    
    try:
        # Check if admin
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("ERROR: Admin privileges required to reset audio components.")
            print("Please run this script as administrator.")
            return False
            
        print("This will restart all Windows audio services.")
        print("WARNING: Your system audio will be temporarily disabled.")
        choice = input("Continue? (y/n): ")
        
        if not choice.lower().startswith('y'):
            print("Operation cancelled.")
            return False
            
        services = ["Audiosrv", "AudioEndpointBuilder"]
        
        # Stop services
        print("Stopping audio services...")
        for service in services:
            cmd = f"net stop {service}"
            subprocess.run(cmd, shell=True, capture_output=True)
            
        print("Waiting 3 seconds...")
        time.sleep(3)
        
        # Start services
        print("Starting audio services...")
        # Start in reverse order
        for service in reversed(services):
            cmd = f"net start {service}"
            subprocess.run(cmd, shell=True, capture_output=True)
            
        print("Audio services have been restarted.")
        return True
    except Exception as e:
        pause_after_error(f"Error resetting audio components: {e}")
        return False

def run_audio_troubleshooter():
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

def main():
    """Main function to diagnose and fix Realtek audio issues."""
    print_header("Realtek Audio Fix Tool")
    
    print("This script will help diagnose and potentially fix Realtek audio driver issues.")
    print("Windows 11 sometimes has compatibility issues with Realtek drivers.")
    
    # Check admin status
    is_admin = check_admin_privileges()
    
    # Check current driver status
    instance_id = check_realtek_driver_status()
    
    # Check event logs
    input("\nPress Enter to check Windows event logs for audio driver issues...")
    check_driver_events(instance_id)
    
    # Offer fixes
    print_header("Possible Solutions")
    print("Based on your system, here are some possible solutions:")
    
    print("\n1. Restart Realtek audio driver")
    print("2. Download latest Realtek driver")
    print("3. Reset Windows audio components")
    print("4. Run Windows audio troubleshooter")
    print("5. Exit")
    
    while True:
        try:
            choice = int(input("\nEnter your choice (1-5): "))
            if choice == 1:
                if is_admin:
                    attempt_driver_restart()
                else:
                    print("Admin privileges required. Please restart the script as administrator.")
            elif choice == 2:
                download_latest_driver()
            elif choice == 3:
                if is_admin:
                    reset_windows_audio()
                else:
                    print("Admin privileges required. Please restart the script as administrator.")
            elif choice == 4:
                run_audio_troubleshooter()
            elif choice == 5:
                break
            else:
                print("Invalid choice. Please enter a number between 1 and 5.")
                
            input("\nPress Enter to return to the menu...")
        except ValueError:
            print("Please enter a number.")
    
    print_header("Realtek Audio Fix Tool Completed")
    print("If your audio issues persist, consider:")
    print("1. Checking your PC manufacturer's website for specific audio drivers")
    print("2. Running sfc /scannow in an admin command prompt to repair system files")
    print("3. Checking if Windows needs updating")
    print("4. Checking your audio hardware connections")

if __name__ == "__main__":
    main()