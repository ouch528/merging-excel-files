## Contents inside the folder:

merge.py

to_run.py

### Note:

merge.py is an implementation (do not need to open), whereas to_run.py is the file to open & run in order for the merge.py to execute. The 2 files MUST always be in the same folder.

## First-Time Installations:



1. Install Python using this [link](https://www.python.org/ftp/python/3.11.2/python-3.11.2-amd64.exe) (Windows installer (64-bit)). Ensure that the ‘Add to Path’ is ticked too.
2. Open the downloaded file, and follow the steps to complete installation
3. Open the Command Prompt (Not the downloaded python file), search for "cmd" in the Start menu.
4. To verify that the Python has been installed, type **python --version** and press enter (the version should be 3.11.2.
5. To install the additional packages, type **pip3 install pandas openpyxl** and press enter (there should be a long list of things running) (if it doesn’t work, try pip instead of pip3)
6. Once installations are done, open IDLE (looks like this) 
7. Open the to_run.py by clicking on file -> open… -> locate the to_run.py file (might have to extract from the zip folder)



## To run the script:



1. (First-Time Only) Retrieve location of the folder that contains the excel files
2. To retrieve: go to the folder, the location will be displayed in the address bar at the top of the File Explorer window. It will typically look something like: C:\\Users\YourUsername\Documents. **(NOTE THAT THERE MUST BE 2 BACKSLASH \\ AFTER THE C:)**
3. Copy the entire address, and paste it at the source_folder_path inside the to_run.py in between the ‘ ‘.
4. Repeat Steps 1-3 for the 2 target folders (the folder that the output file will appear in), paste it at destination_folder_path1 & destination_folder_path2 respectively.
5. Save the file.
6. **(For subsequent Runs) Run the file by pressing F5.**
7. **If the running of script is successful, it will output: **

    **‘Successfully merged and moved the file to both folders**


    **Name of file added to the first target folder: source_file.xlsx**


    **Name of file added to the second target folder: source_file#.xlsx’**

8. **The output filename for the first target folder is source_file.xlsx (fixed), the filename for the second target folder is source_file#.xlsx (changes).**
