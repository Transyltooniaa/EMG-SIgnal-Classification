import os
import pandas as pd
import re

# Function to convert text files to Excel files
def txt_to_excel(input_folder, output_folder):


    # Iterate through each folder in the input directory
    for folder in os.listdir(input_folder):
        folder_path = os.path.join(input_folder, folder)


        if os.path.isdir(folder_path):

            # Create corresponding output folder
            output_subfolder = os.path.join(output_folder, folder)
            os.makedirs(output_subfolder, exist_ok=True)
            
            # Navigate to DataFor1 and DataFor2 folders
            for data_folder in ['DataFor1', 'DataFor2']:
                data_path = os.path.join(folder_path, data_folder)

                if os.path.isdir(data_path):
                    # Create output folder for DataFor1 and DataFor2
                    output_data_folder = os.path.join(output_subfolder, data_folder)
                    os.makedirs(output_data_folder, exist_ok=True)
                    
                    # Iterate through each text file in DataFor1 and DataFor2 folders
                    for i in range(7):
                        txt_file = os.path.join(data_path, f"{i}.txt")


                        if os.path.isfile(txt_file):
                            # Read text file with variable whitespace as delimiter using regex
                            with open(txt_file, 'r') as file:
                                lines = file.readlines()
                            data = []
                            for line in lines:
                                # Split line using regex to handle variable whitespace
                                row = [float(value) if re.match(r'^-?\d+(?:\.\d+)?(?:e[-+]?\d+)?$', value) else value for value in re.split(r'\s+', line.strip()) if value]  # Convert to float if possible
                                data.append(row)
                            
                            # Convert list of lists to DataFrame
                            data = pd.DataFrame(data)
                            
                            # Write data to Excel file in the respective output folder
                            output_file = os.path.join(output_data_folder, f"{i}.xlsx")
                            data.to_excel(output_file, index=False, header=False)  # Assuming no headers in text files
                            print(f"Converted {txt_file} to {output_file}")



#TODO: Change the folder path
input_folder = '/Users/ajiteshkumarsingh/Documents/Semester 4/Signal Processing/Project/Output'  # Folder containing the 36 folders
output_folder = '/Users/ajiteshkumarsingh/Documents/Semester 4/Signal Processing/Project/OutputExcel'  # Folder to store the Excel files

# Convert text files to Excel files
txt_to_excel(input_folder, output_folder)
