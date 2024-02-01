"""
Imaris Image Processing Pipeline
Author: Armaan Hajar and Chandler Asnes
Date: 1/26/23
"""

import pyautogui as auto
import time as tm
import os
import shutil
from xlwt import Workbook

"""
1. Convert file @@@@
2. Set statistics you want @@@@
3. Add new surface
4. Go to page 1/4 on surfaces
5. Uncheck “Classify Surfaces”
6. Uncheck “Object-Object Statistics
7. Go to page 2/4
8. Select “Background Subtraction (Local Contrast)”
9. Set “Diameter of largest Sphere which fits into the Object” @@@@
10. Go to page 3/4
11. Adjust “Threshold (Background Subtraction)“ @@@@
12. Go to page 4/4
13. Adjust Filter @@@@
14. Press green arrow
15. Remove unwanted artifacts @@@
16. Select created surface
17. Press “Edit” pencil
18. Press “Mask All”
19. Check “Set voxel intensity inside surface to”
20. Set to 10000
21. Press enter
22. Enter “Edit” menu
23. Press “Show Display Adjustment”
24. Uncheck “W1 - TRITC” (or first option)
25. Press “Add new Filaments” (little green leaf)
26. Uncheck “Object-Object Statistics”
27. Go to page 2/10
28. Change “Select Source Channel” to “Channel 2 - Masked”
29. Uncheck “Turn on/off slicer rendering for selected object” (little yellow box)
30. Set “Estimated Largest Diameter" @@@@
31. Go to page 3/10
32. Set “Starting Points Threshold" @@@@
33. Set Starting Points @@@@
34. Uncheck “Calculate Soma Model”
35. Go to page 4/9
36. Set “Thinnest Diameter” to 10
37. Go to page 5/9
38. Uncheck “Classify Seed Points”
39. Uncheck “Classify Segments”
40. Adjust “Seed Points Threshold” @@@@
41. Go to page 6/6
42. Press green arrow if you are satisfied with results @@@@
43. Press Statistics (mini graph)
44. Press "Export all Statistics to file" (several floppy disks)
45. Open "Exported Excel Data"
46. Open folder user inputted at the start of the program
47. Press Save
48. Long pause until user is done
49. Alt F4 to close excel
"""

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
    find_and_click("statistics.png") # Press "Statistics" (mini graph)
    tm.sleep(4)
    find_and_click("exportstatistics.png") # Press "Export all Statistics to File" (several floppy disks)
    tm.sleep(0.5)
    find_and_click("thispc.png") # Press "This PC"
    tm.sleep(0.5)
    find_and_click("ddrive.png") # Open D Drive
    auto.press('enter')
    tm.sleep(0.5)
    find_and_click("searchdrive.png")
    auto.typewrite(excel_file_path)
    tm.sleep(1)
    find_and_click("folder.png")
    tm.sleep(0.5)
    auto.press('enter')
    tm.sleep(0.5)
    find_and_click("save.png")
    tm.sleep(2)
    auto.hotkey('alt', 'f4')

def open_next_file():
    find_and_click("openfolder.png")
    tm.sleep(0.5)
    auto.press('enter')
    tm.sleep(11)
    auto.press('tab', 9)
    auto.press('right', 2)
    auto.press('enter')

def find_and_click(button_name):
    path = r'D:\Chandler\Imairs Pipeline\Button Images'
    os.chdir(path)

    if button_name == 'selectsourcechannel.png' or button_name == 'channel2.png':
        location = auto.locateOnScreen(button_name, confidence=.89)
    elif button_name == 'statistics.png' or button_name == "w1tritc.png":
        location = auto.locateOnScreen(button_name, confidence=.9)
    elif button_name == 'nextbutton.png':
        location = auto.locateOnScreen(button_name, confidence=.995)
    elif button_name == "x_out":
        try:
            location = auto.locateOnScreen("x_out.png", confidence=.966)
        except:
            location = auto.locateOnScreen("x_out_gray.png", confidence=.966)
    else:
        location = auto.locateOnScreen(button_name, confidence=.966)

    if button_name == 'thinnestdiameter.png':
        # Triple Click "Thinnest Diameter"
        x_ = location.left - 20
        y_ = location.top + 10
        auto.tripleClick(x_, y_)
    else:
        x_ = location.left + 10
        y_ = location.top + 10
        auto.click(x_, y_)
        if button_name == 'voxelinside.png':
            # Selects "Set voxel intensity inside surface to" Text Box
            tm.sleep(.5)
            auto.tripleClick(location.left + location.width - 30, y_)

def open_vscode():
    path = r'D:\Chandler\Imairs Pipeline\Button Images'
    os.chdir(path)

    try:
        location = auto.locateOnScreen("vscode1.png", confidence=.98)
    except :
        location = auto.locateOnScreen("vscode2.png", confidence=.98)

    center = auto.center(location)
    auto.click(center)

def processing(excel_file_path):
    satisfied = False
    tm.sleep(2)
    try:
        find_and_click("x_out")
    except:
        pass
    find_and_click("addnewsurfaces.png")  # Presses "Add new surfaces"
    tm.sleep(1)
    find_and_click("classifysurfaces.png")  # Presses "Classify Surfaces"
    tm.sleep(0.5)
    find_and_click("objectobjectstats.png")  # Presses "Object-Object Statistics"
    tm.sleep(0.5)
    find_and_click("nextbutton.png")  # Presses Blue next page button
    tm.sleep(0.5)
    find_and_click("backgroundsubtraction.png")  # Selects "Background Subtraction (Local Contrast)" Text Box
    tm.sleep(0.5)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Select the Diameter of Largest Sphere Which Fits into the Object")
    diam_of_lar_sphere.append(float(input("What Did You Input?: ")))
    tm.sleep(2)
    find_and_click("nextbutton.png")  # Presses Blue next page button
    tm.sleep(2)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Select the Threshold (Background Subtraction)")
    background_threshold.append(float(input("What Did You Input?: ")))
    tm.sleep(2)
    find_and_click("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Adjust the Filter")
    adj_filter.append((float(input("What Did You Input For the Minimum?: ")),
                       float(input("What Did You Input For the Maximum? (if no change, input 0): "))))
    tm.sleep(2)
    find_and_click("greennextbutton.png")  # Presses Green next page button
    tm.sleep(2)
    find_and_click("editpencil.png")  # Presses edit pencil
    tm.sleep(1)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Remove the Unwanted Artifacts")
    input("Hit Enter Once You've Remove the Unwanted Artifacts ")
    tm.sleep(0.5)
    find_and_click("maskall.png")  # Presses "Mask All"
    tm.sleep(0.5)
    find_and_click("voxelinside.png")  # Presses "Set voxel intensity inside surface to"
    tm.sleep(0.5)
    auto.typewrite("10000")
    tm.sleep(0.5)
    auto.press("tab")
    tm.sleep(0.5)
    auto.press("enter")
    tm.sleep(3.5)
    find_and_click("editmenu.png") # Enter Edit Menu
    tm.sleep(1)
    auto.press('tab', 5)
    tm.sleep(1)
    auto.press('enter') # Press "Show Display Adjustment"
    tm.sleep(1)
    find_and_click("w1tritc.png") # Uncheck "W1 - TRITC"
    tm.sleep(1)
    find_and_click("x_out") # Closes "Display Adjustment"
    tm.sleep(0.5)
    find_and_click("addnewfilaments.png") # Press "Add new Filaments" (little green leaf)
    tm.sleep(1.5)
    find_and_click("objectobjectstats.png") # Uncheck "Object-Object Statistics"
    tm.sleep(0.5)
    find_and_click("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    find_and_click("selectsourcechannel.png") # Select "Select Source Channel" Dropdown
    tm.sleep(0.5)
    find_and_click("channel2.png") # Change Source Channel to "Channel 2 - Masked"
    tm.sleep(0.5)
    find_and_click("slicerrendering.png") # Uncheck "Turn on/off slicer rendering for selected object" (yellow box)
    tm.sleep(0.5)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Input The Estimated Largest Diameter")
    est_larg_diam.append(float(input("What Did You Input For the Estimated Largest Diameter?: ")))
    tm.sleep(2)
    find_and_click("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(2)
    open_vscode()
    print("----------------------------------------------------------------")
    print("Set the Starting Points Threshold")
    start_point_thres.append((float(input("What Did You Input For the Minimum?: ")),
                              float(input("What Did You Input For the Maximum?: "))))
    tm.sleep(2)
    find_and_click("calculatesomamodel.png") # Uncheck Calculate Soma Model
    tm.sleep(0.5)
    find_and_click("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    find_and_click("thinnestdiameter.png") # Triple Click "Thinnest Diameter"
    auto.typewrite("10")
    tm.sleep(0.5)
    find_and_click("nextbutton.png")  # Presses Blue next page button    
    tm.sleep(0.5)
    find_and_click("classifyseedpoints.png") # Uncheck "Classify Seed Points"
    tm.sleep(0.5)
    find_and_click("classifysegments.png") # Uncheck "Classify Segments"
    tm.sleep(0.5)
    print("----------------------------------------------------------------")
    print("Adjust the Seed Points Threshold")
    while satisfied == False:
        open_vscode()
        temp_spt = float(input("What Did You Input For the Seed Points Threshold?: "))
        tm.sleep(0.5)
        find_and_click("nextbutton.png")  # Presses Blue next page button    
        tm.sleep(4)
        open_vscode()
        print("----------------------------------------------------------------")
        if input("Are You Satitsfied with the Results? (y/n): ") == 'n':
            find_and_click("backbutton.png") # Presses back button
            print("Readjust the Seed Points Threshold")
        else:
            satisfied = True
            seed_points_threshold.append(temp_spt)
            find_and_click("greennextbutton.png")  # Presses Green next page button
    tm.sleep(0.5)
    save_statistics(excel_file_path)

def done_with_file(current_file):
    old_dest = source_folder + f"\{current_file}"
    new_dest = destination_folder + f"\{current_file}"
    shutil.move(old_dest, new_dest)

def save_to_excel():
    wb = Workbook()
    user_input_sheet = wb.add_sheet('Data')
    names = ['File Name', 'Diameter of Largest Sphere', 'Threshold (Background Subtraction)', 'Adjust Filter',
            'Estimated Largest Diameter', 'Starting Points Threshold', 'Seed Points Threshold']
    i_hate_myself = (names, file_names, diam_of_lar_sphere, background_threshold, adj_filter,
                     est_larg_diam, start_point_thres, seed_points_threshold)

    for i, li in enumerate(i_hate_myself):
        for j, item in enumerate(li):
            user_input_sheet.write(i, j, item)

    wb.save('User Inputs.xls')

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
            if input("Want to Continue to the Next File? (y/n): ") == 'y':
                open_next_file()
                tm.sleep(1)
                done_with_file(file)
            else:
                print("Thank you")
                save_to_excel()
                exit()

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
