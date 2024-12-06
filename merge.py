import os
import shutil
import pandas as pd

def merge_excel_files(folder_path, target_folder1, target_folder2):

    output_filename = 'source_file.xlsx'
    excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]
    dfs = []

    for i, file in enumerate(excel_files):
        df = pd.read_excel(os.path.join(folder_path, file), header=1)       
        dfs.append(df)

    merged_data = pd.concat(dfs, ignore_index=True)

    merged_data['Date Created'] = pd.to_datetime(merged_data['Date Created'], format='%d-%b-%Y %H:%M')

    merged_data.sort_values(by='Date Created', inplace=True)

    merged_data['Date Created'] = merged_data['Date Created'].dt.strftime('%d-%b-%Y %H:%M')

    merged_data.to_excel(output_filename, index=False)

    shutil.move(output_filename, os.path.join(target_folder1, output_filename))

    base_name, extension = os.path.splitext(output_filename)
    counter = 1
    new_filename = f"{base_name}{counter}{extension}"

    while os.path.exists(os.path.join(target_folder2, new_filename)):
        counter += 1
        new_filename = f"{base_name}{counter}{extension}"
    shutil.copy(os.path.join(target_folder1, output_filename), os.path.join(target_folder2, new_filename))

    print("Successfully merged and moved the file to both folders")
    print(f"Name of file added to the first target folder: {output_filename}")
    print(f"Name of file added to the second target folder: {new_filename}")
