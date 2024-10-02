import subprocess
import sys
import platform
import os
# List of libraries to install
libraries = [
    "pandas",
    "openpyxl",
    "pillow",
    "customtkinter"
]

# Function to install libraries
def install_libraries():
    for lib in libraries:
        subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

# Install libraries
if __name__ == "__main__":
    install_libraries()
    print("All libraries installed successfully!")
