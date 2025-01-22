import subprocess
import platform

def run_packages_script(script_path):
    try:
        if platform.system() == 'Windows':
            subprocess.run(["python", script_path], check=True)
        else:
            subprocess.run(["python3", script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running 'python {script_path}': {e}")

# Define the path to your packages.py script
packages_script_path = "packages.py"

# Run the packages.py script
run_packages_script(packages_script_path)
