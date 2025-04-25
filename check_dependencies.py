import importlib.util
import subprocess
import sys
import os

current_directory = os.getcwd()

required_packages = ['pandas', 'pyautogui']
required_folders = ['Processed Images', 'Raw Images', 'Excel Files']

def check_packages():
    for package in required_packages:
        if importlib.util.find_spec(package) is None:
            print(f"❌ {package} is not installed.")
            install = input(f"Would you like to install {package}? (y/n): ").strip().lower()
            if install == 'y':
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            else:
                print("Exiting program, please install the required package(s) and try again.")
                exit()

def check_folders():
    os.chdir(current_directory)

    folders = os.listdir(current_directory)
    processed_images_exists = False
    raw_images_exists = False
    excel_files_exists = False

    for folder in folders:
        if folder == 'Processed Images':
            processed_images_exists = True
            for subfolder in os.listdir(folder):
                if subfolder == 'Excel Files':
                    excel_files_exists = True
                    break
        elif folder == 'Raw Images':
            raw_images_exists = True

    if not raw_images_exists:
        os.mkdir('Raw Images')

    if not excel_files_exists:
        os.mkdir('Processed Images')
        os.mkdir(os.path.join('Processed Images', 'Excel Files'))
        processed_images_exists = True

    if not processed_images_exists:
        os.mkdir('Processed Images')
    
        