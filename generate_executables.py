import subprocess
import sys
import os
import glob
import shutil

# Install pyinstaller if not already installed
def install_pyinstaller():
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    except subprocess.CalledProcessError as e:
        print(f"Error installing pyinstaller: {e}")

# Function to run pyinstaller command
def run_pyinstaller_command(script_path, name, windowed=False):
    cwd = os.getcwd()
    dist_path = os.path.join(cwd, "dist")
    mode = "--windowed" if windowed else "--console"
    
    command = [
        "pyinstaller",
        "--noconfirm",
        "--onefile",
        mode,
        "--name", name,
        "--clean",
        "--noupx",
        "--distpath", dist_path,
        "--log-level", "ERROR",  # Set log level to ERROR to disable unnecessary logs
        script_path
    ]
    
    try:
        subprocess.run(command, check=True)
        print(f"✅ Successfully created executable for '{script_path}'")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error running pyinstaller for '{script_path}': {e}")

# Function to clean up generated files and directories
def clean_up():
    cwd = os.getcwd()
    
    # Delete .spec files
    for spec_file in glob.glob(os.path.join(cwd, "*.spec")):
        os.remove(spec_file)
    
    # Delete build directory
    build_dir = os.path.join(cwd, "build")
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)

# Paths and names for your scripts
path = os.getcwd()
installer_script_path = os.path.join(path, "installer.py")
application_script_path = os.path.join(path, "Application.py")

# Install pyinstaller
print("🔧 Installing pyinstaller...")
install_pyinstaller()
print("✅ Successfully installed pyinstaller")

print("🚀 Creating executable files...")

# Run pyinstaller for installer.py
run_pyinstaller_command(installer_script_path, "Installer")

# Run pyinstaller for Application.py (use `windowed=True` if it's a GUI)
run_pyinstaller_command(application_script_path, "Application", windowed=False)

# Clean up generated files
clean_up()
print("🧹 Cleanup complete. Executables are available in the 'dist' folder.")
