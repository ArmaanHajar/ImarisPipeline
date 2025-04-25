import pandas as pd
import datetime as dt
import os

from shared_data import file_names, diam_of_lar_sphere, background_threshold, adj_filter, est_larg_diam, start_point_thres, seed_points_threshold

current_directory = os.getcwd()

current_datetime = dt.datetime.now()
datetime_string = current_datetime.strftime("%Y-%m-%d %H:%M:%S")

def save_to_excel(excel_file_path):
    path = os.path.join(current_directory, excel_file_path)
    os.chdir(path)

    names = ['File Name', 'Diameter of Largest Sphere', 'Threshold (Background Subtraction)', 'Adjust Filter',
             'Estimated Largest Diameter', 'Starting Points Threshold', 'Seed Points Threshold']
    
    variables = [file_names, diam_of_lar_sphere, background_threshold, adj_filter,
                 est_larg_diam, start_point_thres, seed_points_threshold]

    # Create a dictionary mapping each column name to its data
    data_dict = {name: column for name, column in zip(names, variables)}

    # Create a DataFrame
    df = pd.DataFrame(data_dict)

    # Save to Excel
    df.to_excel(f"{excel_file_path}.xlsx", index=False, sheet_name=datetime_string)
