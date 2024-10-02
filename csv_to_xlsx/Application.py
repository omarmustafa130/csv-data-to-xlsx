import subprocess
import platform

def run_packages_script(script_path):
    try:
        if platform.system() == 'Windows':
            subprocess.run(["python", script_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            subprocess.run(["python3", script_path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        print(f"Error running 'python {script_path}': {e}")

# Define the path to your packages.py script
packages_script_path = "csv_to_xlsx.py"

# Run the packages.py script without opening a console window
run_packages_script(packages_script_path)
