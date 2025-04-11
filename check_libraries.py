import importlib.util
import subprocess
import sys

required_packages = ['pandas', 'pyautogui']  # Add more libraries as needed

def check():
    for package in required_packages:
        if importlib.util.find_spec(package) is None:
            print(f"❌ {package} is not installed.")
            install = input(f"Would you like to install {package}? (y/n): ").strip().lower()
            if install == 'y':
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            else:
                print(f"Skipping installation of {package}.")
        else:
            print(f"✅ {package} is installed.")