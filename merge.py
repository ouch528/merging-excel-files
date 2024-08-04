import os
import shutil
import pandas as pd

def merge_excel_files(folder_path, target_folder1, target_folder2):

    output_filename = 'source_file.xlsx'
    
    # Get a list of all Excel files in the folder
    excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]

    # Initialise empty DataFrame
    dfs = []

    # Check through each Excel file and read data into a DataFrame
    for i, file in enumerate(excel_files):
        
        # Read Excel file into a DataFrame
        df = pd.read_excel(os.path.join(folder_path, file), header=1)
       
        # Append DataFrame to dfs list
        dfs.append(df)

    # Concatenate all DataFrames in dfs list without preserving index
    merged_data = pd.concat(dfs, ignore_index=True)

    # Convert 'Date Created' column to datetime object with the correct format
    merged_data['Date Created'] = pd.to_datetime(merged_data['Date Created'], format='%d-%b-%Y %H:%M')

    # Sort by 'Date Created' column in ascending order
    merged_data.sort_values(by='Date Created', inplace=True)

    # Convert 'Date Created' column back to the original format
    merged_data['Date Created'] = merged_data['Date Created'].dt.strftime('%d-%b-%Y %H:%M')

    # Write merged data to a new Excel file
    merged_data.to_excel(output_filename, index=False)

    # Move the output file to the first target folder
    shutil.move(output_filename, os.path.join(target_folder1, output_filename))

    # Generate a new filename if the file already exists in the second target folder
    base_name, extension = os.path.splitext(output_filename)
    counter = 1
    new_filename = f"{base_name}{counter}{extension}"

    while os.path.exists(os.path.join(target_folder2, new_filename)):
        counter += 1
        new_filename = f"{base_name}{counter}{extension}"

    # Copy the output file to the second target folder with the new filename
    shutil.copy(os.path.join(target_folder1, output_filename), os.path.join(target_folder2, new_filename))

    # Print the new filename that has been added
    print("Successfully merged and moved the file to both folders")
    print(f"Name of file added to the first target folder: {output_filename}")
    print(f"Name of file added to the second target folder: {new_filename}")
