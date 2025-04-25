"""
Imaris Image Processing Pipeline
Author: Armaan Hajar and Chandler Asnes
Date: 4/11/25
"""

from input_validation import y_or_n, float_input
from export_to_excel import save_to_excel
from check_dependencies import check_packages, check_folders

import pyautogui as auto
import time as tm
import os
import shutil

current_directory = os.getcwd()

source_folder = os.path.join(current_directory, 'Raw Images')
destination_folder = os.path.join(current_directory, 'Processed Images')
excel_files_folder = os.path.join(destination_folder, 'Excel Files')
path = ""

from shared_data import file_names, diam_of_lar_sphere, background_threshold, adj_filter, est_larg_diam, start_point_thres, seed_points_threshold

def save_statistics(excel_file_path):
    find("statistics.png") # Press "Statistics" (mini graph)
    try_until_found("exportstatistics.png", .9) # Press "Export all Statistics to File" (several floppy disks)
    tm.sleep(0.5)
    find("thispc.png") # Press "This PC"
    tm.sleep(0.5)
    find("ddrive.png") # Open D Drive
    auto.press('enter')
    tm.sleep(0.5)
    find("searchdrive.png")
    auto.typewrite(excel_file_path)
    tm.sleep(1)
    find("folder.png")
    tm.sleep(0.5)
    auto.press('enter')
    tm.sleep(0.5)
    find("save.png")
    tm.sleep(3)
    auto.hotkey('alt', 'f4')

def launch_app():
    auto.press("win")
    tm.sleep(0.5)
    auto.typewrite("Imaris 10.1")
    tm.sleep(0.5)
    auto.press("enter")

def open_next_file():
    find("openfolder.png")
    tm.sleep(0.5)
    auto.press('enter')
    try_until_found("open_iif.png", .9)
    auto.press('tab', 9)
    auto.press('right', 2)
    auto.press('enter')

def try_until_found(looking: str, ci: float, count = 0):
    if count > 500: # Base case: stop if more than 500 attempts to find have been made
        print("Unable to Find")
        return

    path = os.path.join(current_directory, 'Button Images')
    full_path = os.path.join(path, looking)

    try:
        location = auto.locateOnScreen(full_path, confidence=ci)
        if location:
            if looking == "exportstatistics.png":
                find("exportstatistics.png")
            else:
                return
        else:
            raise Exception("Not found")
    except Exception:
        try_until_found(looking, ci, count + 1)

def find(button_name: str, ci = 1.0):
    '''Recursive function to find a button on the screen at the highest confidence interval'''
    if ci < 0.7:  # Base case: stop if confidence interval is too low
        print("Button not found")
        return

    path = os.path.join(current_directory, 'Button Images', button_name)

    try:
        location = auto.locateOnScreen(path, confidence=ci)
        if location:
            if button_name == 'thinnestdiameter.png': # Triple Click "Thinnest Diameter"
                x_ = location.left - 20
                y_ = location.top + 10
                auto.tripleClick(x_, y_)
            else:
                x_ = location.left + 10
                y_ = location.top + 10
                auto.click(x_, y_)
                if button_name == 'voxelinside.png': # Selects "Set voxel intensity inside surface to" Text Box
                    tm.sleep(.5)
                    auto.tripleClick(location.left + location.width - 30, y_)
        else:
            raise Exception("Button not found")
    except Exception:
        find(button_name, ci - 0.01)

def open_vscode():
    path = os.path.join(current_directory, 'Button Images')
    os.chdir(path)

    try:
        location = auto.locateOnScreen("vscode1.png", confidence=.98)
    except:
        location = auto.locateOnScreen("vscode2.png", confidence=.98)

    center = auto.center(location)
    auto.click(center)

def processing(excel_file_path):
    path = os.path.join(current_directory, 'Button Images')
    os.chdir(path)

    satisfied = False

    tm.sleep(2)
    try:
        location = auto.locateOnScreen("w1tritc.png", confidence=.9)
        if location:
            auto.hotkey('ctrl', 'd')
        else:
            raise Exception
    except Exception:
        pass
    find("addnewsurfaces.png")  # Presses "Add new surfaces"
    tm.sleep(0.5)
    find("classifysurfaces.png")  # Presses "Classify Surfaces"
    tm.sleep(0.5)
    find("objectobjectstats.png")  # Presses "Object-Object Statistics"
    tm.sleep(0.5)
    find("nextbutton.png")  # Presses Blue next page button
    tm.sleep(0.5)
    find("backgroundsubtraction.png")  # Selects "Background Subtraction (Local Contrast)" Text Box
    tm.sleep(0.5)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Select the Diameter of Largest Sphere Which Fits into the Object")
    diam_of_lar_sphere.append(float_input("What Did You Input?: "))
    # maybe add section where it automatically types in the number so you don't need to type it twice
    tm.sleep(2)
    find("nextbutton.png")  # Presses Blue next page button
    tm.sleep(2)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Select the Threshold (Background Subtraction)")
    background_threshold.append(float_input("What Did You Input?: "))
    tm.sleep(2)
    find("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Adjust the Filter")
    adj_filter.append(str((float_input("What Did You Input For the Minimum?: "),
                           float_input("What Did You Input For the Maximum? (if no change, input 0): "))))
    tm.sleep(2)
    find("greennextbutton.png")  # Presses Green next page button
    tm.sleep(2)
    find("editpencil.png")  # Presses edit pencil
    tm.sleep(1)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Remove the Unwanted Artifacts")
    input("Hit Enter Once You've Remove the Unwanted Artifacts ")
    tm.sleep(0.5)
    find("maskall.png")  # Presses "Mask All"
    tm.sleep(0.5)
    find("voxelinside.png")  # Presses "Set voxel intensity inside surface to"
    tm.sleep(0.5)
    auto.typewrite("10000")
    tm.sleep(0.5)
    auto.press("tab")
    tm.sleep(0.5)
    auto.press("enter")
    tm.sleep(3.5)
    auto.hotkey("ctrl", "d") # Open "Show Display Adjustment"
    tm.sleep(0.5)
    find("w1tritc.png") # Uncheck "W1 - TRITC"
    tm.sleep(1)
    auto.hotkey("ctrl", "d") # Close "Show Display Adjustment"
    tm.sleep(0.5)
    find("addnewfilaments.png") # Press "Add new Filaments" (little green leaf)
    tm.sleep(1.5)
    find("objectobjectstats.png") # Uncheck "Object-Object Statistics"
    tm.sleep(0.5)
    find("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    find("selectsourcechannel.png") # Select "Select Source Channel" Dropdown
    tm.sleep(0.5)
    find("channel2.png") # Change Source Channel to "Channel 2 - Masked"
    tm.sleep(0.5)
    find("slicerrendering.png") # Uncheck "Turn on/off slicer rendering for selected object" (yellow box)
    tm.sleep(0.5)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Input The Estimated Largest Diameter")
    est_larg_diam.append(float_input("What Did You Input For the Estimated Largest Diameter?: "))
    tm.sleep(2)
    find("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(2)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Set the Starting Points Threshold")
    start_point_thres.append(str((float_input("What Did You Input For the Minimum?: "),
                                  float_input("What Did You Input For the Maximum?: "))))
    tm.sleep(2)
    find("calculatesomamodel.png") # Uncheck Calculate Soma Model
    tm.sleep(0.5)
    find("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    find("thinnestdiameter.png") # Triple Click "Thinnest Diameter"
    auto.typewrite("10")
    tm.sleep(0.5)
    find("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    find("classifyseedpoints.png") # Uncheck "Classify Seed Points"
    tm.sleep(0.5)
    find("classifysegments.png") # Uncheck "Classify Segments"
    tm.sleep(0.5)
    print("----------------------------------------------------------------")
    print("Adjust the Seed Points Threshold")
    while satisfied == False:
        open_vscode()
        temp_spt = float_input("What Did You Input For the Seed Points Threshold?: ")
        tm.sleep(0.5)
        find("nextbutton.png")  # Presses Blue next page button    
        tm.sleep(4)
        open_vscode()
        print("----------------------------------------------------------------")
        if y_or_n("Are You Satitsfied with the Results?"):
            find("backbutton.png") # Presses back button
            print("Readjust the Seed Points Threshold")
        else:
            satisfied = True
            # auto.hotkey("ctrl", s)
            seed_points_threshold.append(temp_spt)
            find("greennextbutton.png")  # Presses Green next page button
    tm.sleep(0.5)
    save_statistics(excel_file_path)

def done_with_file(current_file):
    old_dest = os.path.join(source_folder, current_file)
    new_dest = os.path.join(destination_folder, current_file)
    shutil.move(old_dest, new_dest)

def batch(excel_file_path):
    folder = os.listdir(source_folder)
    input("Please Open The First File You Would Like to Process Then Hit Enter ")

    for i in range(len(folder)):
        file = folder[i]
        file_name, file_extension = os.path.splitext(file)
        if file_extension != ".ims":
            pass
        else:
            print(f"Now Working on {file}")
            file_names.append(file_name)
            processing(excel_file_path)
            open_vscode()
            if i == len(folder) - 1:
                print("Thank you")
                # save file
                # close Imaris
                save_to_excel(excel_file_path)
                exit()
            else:
                if y_or_n("Want to Continue to the Next File?"):
                    open_next_file()
                    tm.sleep(1)
                    done_with_file(file)
                else:
                    print("Thank you")
                    save_to_excel(excel_file_path)
                    exit()
    print("Thank you")
    save_to_excel(excel_file_path)

def main():
    print("-------------------------------------------------------------")
    print("          Imaris Pipeline (Created by Armaan Hajar)          ")
    print("-------------------------------------------------------------")

    check_packages()  # Check if required libraries are installed
    check_folders()  # Check if required folders exist

    if not y_or_n("Is Imaris Running and Full Screen?"):
        print("Please Open and Full Screen the Imaris Application")
        exit()

    if not y_or_n("Are Your Unprocessed Images in the 'Raw Images' Folder?"):
        print("Please Put the Unprocessed Images in the 'Raw Images' Folder")
        exit()

    if not y_or_n("Have You Selected Your Preferred Statistics?"):
        print("Please Selected Your Preferred Statistics And Restart the Script")
        exit()
        
    batch(excel_files_folder)

if __name__ == "__main__":
    main()
