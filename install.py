import os
import sys
import time
import urllib.request
import zipfile
import shutil
import subprocess
import ctypes
from pathlib import Path
import settings

"""
SETTINGS AREA FOR INSTALL
"""

PROGRAM_ENTRYPOINT_FILE = "virtprint.py"
GHOSTSCRIPT_URL = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10060/ghostxps-10.06.0-win64.zip"
GHOSTSCRIPT_DIR = "gxps"

"""
END SETTINGS AREA
"""

scriptdir = os.path.dirname(os.path.abspath(sys.argv[0]))


print("Creating python venv ...")
try:
    os.system(f"python -m venv {scriptdir}/.venv")
    print(f"Created venv to to {scriptdir}/.venv")
except Exception:
    print("error occurred when (re)creating venv!")
    exit(1)

print("Installing packages to venv...")
time.sleep(1.5)
try:
    print(f"{scriptdir}\\.venv\\Scripts\\pip install -r " + f"{scriptdir}\\requirements.txt")
    os.system(
        f"{scriptdir}\\.venv\\Scripts\\pip install -r " + f"{scriptdir}\\requirements.txt"
    )
    print("Installed required packages to venv!")
    
    # Run pywin32 post-install to set up service support properly
    print("\nConfiguring pywin32 for Windows services...")
    python_exe = Path(scriptdir) / ".venv" / "Scripts" / "python.exe"
    pywin32_postinstall = Path(scriptdir) / ".venv" / "Scripts" / "pywin32_postinstall.py"
    
    if pywin32_postinstall.exists():
        result = subprocess.run(
            [str(python_exe), str(pywin32_postinstall), "-install"],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            print("✓ pywin32 configured successfully!")
        else:
            print("⚠ pywin32 post-install warning (service may still work)")
            if result.stderr:
                print(f"  {result.stderr.strip()}")
    else:
        print("⚠ pywin32_postinstall.py not found - running anyway...")
        
except Exception:
    print("error occurred when installing packages to venv!")
    exit(1)


print("Downloading GhostScript...")
gxps_dir = Path(scriptdir) / GHOSTSCRIPT_DIR
gxps_zip = Path(scriptdir) / "ghostxps.zip"

try:
    # Download the zip file
    print(f"Downloading from {GHOSTSCRIPT_URL}...")
    urllib.request.urlretrieve(GHOSTSCRIPT_URL, gxps_zip)
    print(f"Downloaded to {gxps_zip}")

    # Create gxps directory
    gxps_dir.mkdir(exist_ok=True)
    print(f"Created directory {gxps_dir}")

    # Extract the zip file
    print(f"Extracting to {gxps_dir}...")
    temp_extract_dir = Path(scriptdir) / "temp_gxps_extract"
    temp_extract_dir.mkdir(exist_ok=True)

    with zipfile.ZipFile(gxps_zip, 'r') as zip_ref:
        zip_ref.extractall(temp_extract_dir)

    # Find the inner directory and move its contents to gxps_dir
    inner_dirs = [d for d in temp_extract_dir.iterdir() if d.is_dir()]
    if inner_dirs:
        inner_dir = inner_dirs[0]  # Get the ghostxps-10.06.0-win64 folder
        print(f"Moving contents from {inner_dir.name} to {gxps_dir}...")

        # Move all files from inner directory to gxps_dir
        for item in inner_dir.iterdir():
            dest = gxps_dir / item.name
            if dest.exists():
                if dest.is_dir():
                    import shutil
                    shutil.rmtree(dest)
                else:
                    dest.unlink()
            item.rename(dest)

        # Remove the temporary extraction directory
        import shutil
        shutil.rmtree(temp_extract_dir)
    else:
        # If no inner directory, just move everything
        for item in temp_extract_dir.iterdir():
            item.rename(gxps_dir / item.name)
        temp_extract_dir.rmdir()

    print(f"Extracted GhostScript to {gxps_dir}")

    # Clean up the zip file
    gxps_zip.unlink()
    print("Cleaned up temporary files")
    
except Exception as e:
    print(f"Error occurred when downloading/extracting GhostScript: {e}")
    print("You may need to install GhostScript manually from:")
    print(GHOSTSCRIPT_URL)
    # Don't exit - continue with installation


print("\n" + "="*60)
print("Installation complete!")
print("="*60)

# Ask if user wants to install the Windows service
print("\n" + "="*60)
print("WINDOWS SERVICE INSTALLATION")
print("="*60)
print("Would you like to install VirtPrint as a Windows service?")
print("This will allow the virtual printer to run automatically at system startup.")
print("\nNote: This requires administrator privileges.")
response = input("Install service? (y/n): ").lower().strip()

if response == 'y' or response == 'yes':
    print("\nInstalling Windows service...")
    
    # Check if running as administrator
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False
    
    if not is_admin:
        print("\n⚠ WARNING: Not running as administrator!")
        print("The service installation may fail without admin privileges.")
        print("\nTo run as administrator:")
        print("1. Right-click on Command Prompt or PowerShell")
        print("2. Select 'Run as administrator'")
        print("3. Run this install script again")
        
        proceed = input("\nTry to install anyway? (y/n): ").lower().strip()
        if proceed != 'y' and proceed != 'yes':
            print("\nSkipping service installation.")
            print("You can install the service later by running:")
            print(f"  {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\service.py install")
            sys.exit(0)
    
    # Install the service
    try:
        python_exe = str(Path(scriptdir) / ".venv" / "Scripts" / "python.exe")
        service_script = str(Path(scriptdir) / "service.py")
        
        result = subprocess.run(
            [python_exe, service_script, "install"],
            capture_output=True,
            text=True,
            cwd=scriptdir
        )
        
        if result.returncode == 0:
            print("✓ Service installed successfully!")
            print(f"\nOutput:\n{result.stdout}")
            
            # Ask if they want to start the service now
            start_now = input("\nStart the service now? (y/n): ").lower().strip()
            if start_now == 'y' or start_now == 'yes':
                print("Starting VirtPrint service...")
                result = subprocess.run(
                    [python_exe, service_script, "start"],
                    capture_output=True,
                    text=True,
                    cwd=scriptdir
                )
                
                if result.returncode == 0:
                    print("✓ Service started successfully!")
                    print(f"\nThe virtual printer '{settings.PRINTER_NAME}' is now running.")
                else:
                    print("✗ Failed to start service")
                    print(f"Error: {result.stderr}")
                    print("\nYou can start it manually later with:")
                    print(f"  {python_exe} {service_script} start")
            else:
                print("\nService installed but not started.")
                print("You can start it later with:")
                print(f"  {python_exe} {service_script} start")
                print("Or use Windows Services (services.msc)")
        else:
            print("✗ Failed to install service")
            print(f"Error: {result.stderr}")
            print("\nPlease ensure you are running as administrator.")
            
    except Exception as e:
        print(f"✗ Error installing service: {e}")
        print("\nYou can try installing the service manually:")
        print(f"  {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\service.py install")
else:
    print("\nSkipping service installation.")
    print("You can install the service later by running:")
    print(f"  {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\service.py install")

print("\n" + "="*60)
print("INSTALLATION SUMMARY")
print("="*60)
print(f"Virtual environment: {scriptdir}\\.venv")
print(f"GhostScript directory: {scriptdir}\\{GHOSTSCRIPT_DIR}")
print(f"\nTo run manually (without service):")
print(f"  {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\{PROGRAM_ENTRYPOINT_FILE}")
print(f"\nTo manage the Windows service:")
print(f"  Install:  {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\service.py install")
print(f"  Start:    {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\service.py start")
print(f"  Stop:     {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\service.py stop")
print(f"  Remove:   {scriptdir}\\.venv\\Scripts\\python.exe {scriptdir}\\service.py remove")
print("="*60)
