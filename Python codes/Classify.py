import os

# Create directory if it does not exist
def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


# Process files and sort them into different classes
def process_files(folder_path, folder_path_destination , n):
    outputMap = {
        "0": "0.txt",
        "1": "1.txt",
        "2": "2.txt",
        "3": "3.txt",
        "4": "4.txt",
        "5": "5.txt",
        "6": "6.txt",
        "7": "7.txt",
    }

    # Check if the folder path exists
    if not os.path.exists(folder_path):
        print("Folder path does not exist.")
        return

    # Create the destination folder if it does not exist
    create_directory(folder_path_destination)


    # Walk through the folder and process the files
    for root, dirs, files in os.walk(folder_path):

        for folder in dirs:

            # Check if the number of folders to sample is 0
            if(n==0):
                break

            n = n-1

            # Get the folder name and the folder path
            folder_name = os.path.join(root, folder)

            # Check if the folder name is .DS_Store
            for filename in os.listdir(folder_name):
                if filename == ".DS_Store":
                    continue
                    
                print("Processing file:", filename + " in folder:", folder)

                # Open the file and read the lines in it
                with open(os.path.join(folder_name, filename), 'r') as file:

                    # Iterate through the lines in the file and process them
                    for line in file:
                        line = line.strip()
                        class_name = line[-1]
                        output_file = ""

                        # Check if the class name is in the output map and get the output file
                        if class_name in outputMap:
                            output_file = outputMap[class_name]

                            # Check if the output file is not empty and write the line to the output file
                            create_directory(os.path.join(folder_path_destination,folder, "DataFor1"))
                            create_directory(os.path.join(folder_path_destination,folder, "DataFor2"))

    
                            output_location = ""

                            # Check if the filename starts with 1 or 2 
                            if filename.startswith("1"):
                                output_location = os.path.join(folder_path_destination, folder , "DataFor1", output_file)
                            elif filename.startswith("2"):
                                output_location = os.path.join(folder_path_destination, folder , "DataFor2", output_file)

                            # Write the line to the output file 
                            if output_location:
                                with open(output_location, 'a') as f:
                                    f.write(line + "\n")

    print("Data sorted successfully.")



# TODO: Change the folder path
folder_path = "/Users/ajiteshkumarsingh/Documents/Semester 4/Signal Processing/Project/EMG_data_for_gestures-master"
folder_path_destination = "/Users/ajiteshkumarsingh/Documents/Semester 4/Signal Processing/Project/Output"

# Get the number of folders to sample
n = int(input("Enter the number of folders you want sammple of : "))

# Process the files
process_files(folder_path, folder_path_destination,n)
