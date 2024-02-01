"""
Imaris Image Processing Pipeline
Author: Armaan Hajar and Chandler Asnes
Date: 1/18/23
"""

import pyautogui as auto
import time as tm
import os
import shutil
from xlwt import Workbook
from tkinter import *
from tkinter import ttk

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
excel_file_path = ""
file_names = [] # Names of current file pipeline is working on
diam_of_lar_sphere = [] # Diameter of largest Sphere which fits into the Object
background_threshold = [] # Threshold (Background Subtraction)
adj_filter = [] # Adjust Filter
est_larg_diam = [] # Estimated Largest Diameter
start_point_thres = [] # Starting Points Threshold
seed_points_threshold = [] # Seed Points Threshold

def launch_app():
    auto.press("win")
    tm.sleep(1)
    auto.typewrite("Imaris 10.1")
    tm.sleep(0.5)
    auto.press("enter")

def pop_up(output: str):
    pass
    '''
    win = Tk()
    win.lift()
    Label(win, text=output, font=('Helvetica 14 bold'), wraplength=300, justify=CENTER).pack(pady=10, padx=20)
    ttk.Button(win, text= "Okay", command=win.destroy).pack(pady=10, padx=20)
    win.lift()
    win.mainloop()
    '''

def open_folder():
    auto.click(x=221, y=66) # Presses "Observe Folder"
    for i in range(4):
        auto.press("tab")
    for i in range(2):
        auto.press('right')
    tm.sleep(0.5)
    auto.press("enter")
    tm.sleep(0.5)
    auto.press("tab")
    tm.sleep(0.5)
    auto.typewrite("Raw Imaris Files")
    tm.sleep(0.5)
    auto.moveTo(x=1007, y=413)
    tm.sleep(1.5)
    auto.click(clicks=2) # Click Folder
    tm.sleep(0.5)
    auto.press("enter")
    tm.sleep(0.5)

def save_statistics():
    auto.click(x=184, y=1103) # Press "Statistics" (mini graph)
    tm.sleep(2.5)
    auto.click(x=351, y=1512) # Press "Export all Statistics to File" (several floppy disks)
    tm.sleep(0.5)
    for i in range(6):
        auto.press('tab')
    tm.sleep(0.5)
    auto.press('right')
    tm.sleep(0.5)
    auto.press('enter')
    tm.sleep(0.5)
    auto.press('tab')
    tm.sleep(0.5)
    auto.typewrite(excel_file_path)
    auto.click(x=184, y=1103) # Click Folder
    tm.sleep(0.5)
    auto.press('enter')
    tm.sleep(2)
    auto.hotkey('alt', 'f4')

def open_next_file():
    pass

def processing():
    satisfied = False # line 239
    tm.sleep(3)
    auto.click(x=99, y=138)  # Presses "Add new surfaces"
    tm.sleep(0.5)
    auto.click(x=19, y=1338)  # Presses "Classify Surfaces"
    tm.sleep(0.5)
    auto.click(x=21, y=1360)  # Presses "Object-Object Statistics"
    tm.sleep(0.5)
    auto.click(x=311, y=1509)  # Presses Blue next page button
    tm.sleep(0.5)
    auto.click(x=21, y=1318)  # Selects "Background Subtraction (Local Contrast)" Text Box
    tm.sleep(0.5)
    pop_up("Select the Diameter of Largest Sphere Which Fits into the Object")
    print("----------------------------------------------------------------")
    print("Select the Diameter of Largest Sphere Which Fits into the Object")
    diam_of_lar_sphere.append(float(input("What Did You Input?: ")))
    tm.sleep(2)
    auto.click(x=305, y=1514)  # Presses next page
    tm.sleep(2)
    pop_up("Select the Threshold (Background Subtraction)")
    print("----------------------------------------------------------------")
    print("Select the Threshold (Background Subtraction)")
    background_threshold.append(float(input("What Did You Input?: ")))
    tm.sleep(2)
    auto.click(x=310, y=1506)  # Presses next page
    tm.sleep(0.5)
    pop_up("Adjust the Filter")
    print("----------------------------------------------------------------")
    print("Adjust the Filter")
    adj_filter.append((float(input("What Did You Input For the Minimum?: ")),
                       float(input("What Did You Input For the Maximum? (if no change, input 0): "))))
    tm.sleep(2)
    auto.click(x=331, y=1507)  # Presses green arrow
    tm.sleep(2)
    auto.click(x=99, y=1109)  # Presses edit pencil
    tm.sleep(1)
    pop_up("Remove the Unwanted Artifacts")
    print("----------------------------------------------------------------")
    print("Remove the Unwanted Artifacts")
    if input("Did You Remove the Unwanted Artifacts? (y/n): ") == 'n':
        print("im sowwy")
    tm.sleep(0.5)
    auto.click(x=195, y=1318)  # Presses "Mask All"
    tm.sleep(0.5)
    auto.click(x=1086, y=834)  # Presses "Set voxel intensity inside surface to"
    tm.sleep(0.5)
    auto.tripleClick(x=1317, y=833)  # Selects "Set voxel intensity inside surface to" Text Box
    tm.sleep(0.5)
    auto.typewrite("10000")
    tm.sleep(0.5)
    auto.press("tab")
    tm.sleep(0.5)
    auto.press("enter")
    tm.sleep(0.5)
    auto.click(x=44, y=26) # Enter Edit Menu
    tm.sleep(0.5)
    auto.click(x=44, y=26)
    tm.sleep(0.5)
    for i in range(5):
        auto.press('tab')
    tm.sleep(0.5)
    auto.press('enter') # Press "Show Display Adjustment"
    tm.sleep(0.5)
    auto.click(x=822, y=1306) # Uncheck "W1 - TRITC"
    tm.sleep(0.5)
    auto.click(x=136, y=136) # Press "Add new Filaments" (little green leaf)
    tm.sleep(0.5)
    auto.click(x=20, y=1415) # Uncheck "Object-Object Statistics"
    tm.sleep(0.5)
    auto.click(x=309, y=1514) # Next Page
    tm.sleep(0.5)
    auto.click(x=147, y=1090) # Select "Select Source Channel" Dropdown
    tm.sleep(0.5)
    auto.click(x=22, y=1123) # Change Source Channel to "Channel 2 - Masked"
    tm.sleep(0.5)
    auto.click(x=391, y=1045) # Uncheck "Turn on/off slicer rendering for selected object" (yellow box)
    tm.sleep(0.5)
    pop_up("Input The Estimated Largest Diameter")
    print("----------------------------------------------------------------")
    print("Input The Estimated Largest Diameter")
    est_larg_diam.append(float(input("What Did You Input For the Estimated Largest Diameter?: ")))
    tm.sleep(2)
    auto.click(x=308, y=1508) # Next Page
    tm.sleep(2)
    pop_up("Set the Starting Points Threshold")
    print("----------------------------------------------------------------")
    print("Set the Starting Points Threshold")
    start_point_thres.append((float(input("What Did You Input For the Minimum?: ")),
                              float(input("What Did You Input For the Maximum?: "))))
    tm.sleep(2)
    auto.click(x=13, y=1260) # Uncheck Calculate Soma Model
    tm.sleep(0.5)
    auto.click(x=309, y=1510) # Next Page
    tm.sleep(0.5)
    auto.tripleClick(x=258, y=1088) # Triple Click "Thinnest Diameter"
    auto.typewrite("10")
    tm.sleep(0.5)
    auto.click(x=310, y=1509) # Next Page
    tm.sleep(0.5)
    auto.click(x=13, y=1240) # Uncheck "Classify Seed Points"
    tm.sleep(0.5)
    auto.click(x=13, y=1269) # Uncheck "Classify Segments"
    tm.sleep(0.5)
    pop_up("Adjust the Seed Points Threshold")
    print("----------------------------------------------------------------")
    print("Adjust the Seed Points Threshold")
    while satisfied == False:
        temp_spt = float(input("What Did You Input For the Seed Points Threshold?: "))
        tm.sleep(2)
        auto.click(x=314, y=1512) # Next Page
        tm.sleep(0.5)
        pop_up("Are You Satitsfied with the Results?")
        print("----------------------------------------------------------------")
        if input("Are You Satitsfied with the Results? (y/n): ") == 'n':
            auto.click(x=287, y=1511)
            print("Readjust the Seed Points Threshold")
        else:
            satisfied = True
    seed_points_threshold.append(temp_spt)
    auto.click(x=338, y=1504) # Press Green Arrow
    tm.sleep(0.5)
    save_statistics()

# call this function in the open_next_file function right after pipeline presses
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

def batch():
    folder = os.listdir(source_folder)
    open_folder()

    for i in range(len(folder)):
        file = folder[i]
        file_name, file_extension = os.path.splitext(file)
        if file_extension != ".ims":
            pass
        else:
            print(f"Now Working on {file}")
            file_names.append(file_name)
            auto.doubleClick(x=325, y=163)
            processing()
            if input("Want to Continue to the Next File? (y/n): ") == 'y':
                open_next_file()
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
    excel_file_path = os.path.join(dest_excel_files, excel_file_path)
    os.mkdir(excel_file_path)
    batch()

if __name__ == "__main__":
    main()