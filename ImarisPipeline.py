"""
Imaris Image Processing Pipeline
Author: Armaan Hajar and Chandler Asnes
Date: 2/5/23
"""

import pyautogui as auto
import time as tm
import os
import shutil
from xlwt import Workbook

source_folder = r'D:\Chandler\Raw Imaris Files'
destination_folder = r'D:\Chandler\Processed Imaris Files'
dest_excel_files = r'D:\Chandler\Exported Excel Data'
path = ""
file_names = [] # Names of current file pipeline is working on
diam_of_lar_sphere = [] # Diameter of largest Sphere which fits into the Object
background_threshold = [] # Threshold (Background Subtraction)
adj_filter = [] # Adjust Filter
est_larg_diam = [] # Estimated Largest Diameter
start_point_thres = [] # Starting Points Threshold
seed_points_threshold = [] # Seed Points Threshold

def launch_app():
    auto.press("win")
    tm.sleep(0.5)
    auto.typewrite("Imaris 10.1")
    tm.sleep(0.5)
    auto.press("enter")

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
    tm.sleep(2)
    auto.hotkey('alt', 'f4')

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

    path = r'D:\Chandler\Imairs Pipeline\Button Images'
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

    path = r'D:\Chandler\Imairs Pipeline\Button Images'
    full_path = os.path.join(path, button_name)

    try:
        location = auto.locateOnScreen(full_path, confidence=ci)
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
    path = r'D:\Chandler\Imairs Pipeline\Button Images'
    os.chdir(path)

    try:
        location = auto.locateOnScreen("vscode1.png", confidence=.98)
    except:
        location = auto.locateOnScreen("vscode2.png", confidence=.98)

    center = auto.center(location)
    auto.click(center)

def processing(excel_file_path):
    path = r'D:\Chandler\Imairs Pipeline\Button Images'
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
    diam_of_lar_sphere.append(float(input("What Did You Input?: ")))
    tm.sleep(2)
    find("nextbutton.png")  # Presses Blue next page button
    tm.sleep(2)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Select the Threshold (Background Subtraction)")
    background_threshold.append(float(input("What Did You Input?: ")))
    tm.sleep(2)
    find("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Adjust the Filter")
    adj_filter.append(str((float(input("What Did You Input For the Minimum?: ")),
                           float(input("What Did You Input For the Maximum? (if no change, input 0): ")))))
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
    est_larg_diam.append(float(input("What Did You Input For the Estimated Largest Diameter?: ")))
    tm.sleep(2)
    find("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(2)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Set the Starting Points Threshold")
    start_point_thres.append(str((float(input("What Did You Input For the Minimum?: ")),
                                  float(input("What Did You Input For the Maximum?: ")))))
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
        temp_spt = float(input("What Did You Input For the Seed Points Threshold?: "))
        tm.sleep(0.5)
        find("nextbutton.png")  # Presses Blue next page button    
        tm.sleep(4)
        open_vscode()
        print("----------------------------------------------------------------")
        if input("Are You Satitsfied with the Results? (y/n): ") == 'n':
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
    old_dest = source_folder + f"\{current_file}"
    new_dest = destination_folder + f"\{current_file}"
    shutil.move(old_dest, new_dest)

def save_to_excel(excel_file_path):
    path = f"D:\Chandler\Exported Excel Data\{excel_file_path}"
    os.chdir(path)

    wb = Workbook()
    user_input_sheet = wb.add_sheet('Data')
    names = ['File Name', 'Diameter of Largest Sphere', 'Threshold (Background Subtraction)', 'Adjust Filter',
            'Estimated Largest Diameter', 'Starting Points Threshold', 'Seed Points Threshold']
    i_hate_myself = [file_names, diam_of_lar_sphere, background_threshold, adj_filter,
                     est_larg_diam, start_point_thres, seed_points_threshold]
    
    for i in range(len(names)):
        user_input_sheet.write(0, i, names[i])
    for i in range(len(i_hate_myself)):
        for j in range(len(file_names)):
            user_input_sheet.write(j+1, i, i_hate_myself[i][j])

    wb.save(f"{excel_file_path}.xls")

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
            if i == len(folder):
                print("Thank you")
                save_to_excel(excel_file_path)
                exit()
            else:
                if input("Want to Continue to the Next File? (y/n): ") == 'y':
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
    print("----------------------------------------------------------------")
    if input("Is Imaris Running? (y/n): ") == 'n':
        launch_app()
        print("Please Full Screen the Imaris Application")

    if input("Is Imaris Full Screen? (y/n): ") == 'n':
        print("Please Restart the Script and Full Screen Imaris")
        exit()

    if input("Have You Selected Your Preferred Statistics? (y/n): ")  == 'n':
        print("Please Selected Your Preferred Statistics And Restart the Script")
        exit()

    excel_file_path = str(input("What is the Name of the Folder You Would Like to Save the Excel Data To?: "))
    path = os.path.join(dest_excel_files, excel_file_path)
    os.mkdir(path)
    batch(excel_file_path)

if __name__ == "__main__":
    main()
