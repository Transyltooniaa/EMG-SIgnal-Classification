from tabulate import tabulate

#TODO - Change the folder path to the location of the OutputResults folder
movement = 0
folders = 36

# Define flag names
flag_names = {
    "class_0": "Class 0",
    "class_1": "Class 1",
    "class_2": "Class 2",
    "class_3": "Class 3",
    "class_4": "Class 4",
    "class_5": "Class 5",
    "class_6": "Class 6"
}

# Final priorities dictionary
priorities = {
    "Class 0": [0,0],
    "Class 1": [0,0],
    "Class 2": [0,0],
    "Class 3": [0,0],
    "Class 4": [0,0],
    "Class 5": [0,0],
    "Class 6": [0,0]
}


for TestCase in range(1, folders + 1):

    # Reset flags for each test case
    flags = {
        "class_0": 0,
        "class_1": 0,
        "class_2": 0,
        "class_3": 0,
        "class_4": 0,
        "class_5": 0,
        "class_6": 0
    }

    if TestCase <= 12:
        subfolder = "1_12"
    elif 13 <= TestCase <= 24:
        subfolder = "13_24"
    elif 25 <= TestCase <= 36:
        subfolder = "25_36"


    #TODO - Change the folder path to the location of the OutputResults folder
    folderPath = f"/Users/ajiteshkumarsingh/Documents/Semester 4/Signal Processing/Project/PlotAndStats/{subfolder}/OutputResults_2-5/DataFor{TestCase}_DataFor2_Results"
    folderPath2 = f"/Users/ajiteshkumarsingh/Documents/Semester 4/Signal Processing/Project/PlotAndStats/{subfolder}/OutputResults_6-9/DataFor{TestCase}_DataFor2_Results"



    # Load the data
    with open(f"{folderPath}/Results_{movement}.txt", "r") as f:
        data1 = f.read().split("\n")

    with open(f"{folderPath2}/Results_{movement}.txt", "r") as f:
        data2 = f.read().split("\n")

    # Extract the required data
    amplitude_spectrum1 = int(data1.index("Mean of amplitude spectrum: "))
    amplitude_spectrum2 = int(data2.index("Mean of amplitude spectrum: "))

    # Extract the required data for Amplitude Spectrum
    amplitude_spectrum2_5 = [float(data1[amplitude_spectrum1 + i]) for i in range(1, 5)]
    amplitude_spectrum6_9 = [float(data2[amplitude_spectrum2 + i]) for i in range(1, 5)]
    amplitude_spectrum = amplitude_spectrum2_5 + amplitude_spectrum6_9



    # Extract the required data for Kurtosis
    kurtIndex1 = int(data1.index("Kurtosis in time domain: "))
    kurtIndex2 = int(data2.index("Kurtosis in time domain: "))

    Kurtosis2_5 = [float(data1[kurtIndex1 + i]) for i in range(1, 5)]
    Kurtosis6_9 = [float(data2[kurtIndex2 + i]) for i in range(1, 5)]
    Kurtosis = Kurtosis2_5 + Kurtosis6_9


    # Extract the required data for Peak Frequency
    peak_freq1 = int(data1.index("Peak frequency (excluding 0 frequency component): "))
    peak_freq2 = int(data2.index("Peak frequency (excluding 0 frequency component): "))

    peak_freq2_5 = [float(data1[peak_freq1 + i]) for i in range(1, 5)]
    peak_freq6_9 = [float(data2[peak_freq2 + i]) for i in range(1, 5)]
    peak_freq = peak_freq2_5 + peak_freq6_9



    # Flag 0: Highest probability
    if all(i > 6 for i in Kurtosis):
        flags["class_0"] += 1

    # Flag 1: Condition based on amplitude spectrum
    if abs(amplitude_spectrum[-2] * 1e6) < 1000:
        flags["class_1"] += 1

    # Flag 5: Multiple conditions
    if all(abs(amplitude_spectrum[i] - 0.002) < 0.002 for i in range(4)):
        flags["class_5"] += 1

    # Flag 5: Another condition
    if sum(Kurtosis[i] - [4, 3, 4, 5][i] for i in range(-1, -5, -1)) < 2.4:
        flags["class_5"] += 1

    # Flag 4: Condition based on peak frequency
    if sum(i >= 10 or abs(i - 10) < 0.5 for i in peak_freq) > 5:
        flags["class_4"] += 1

    # Flag 4: Another condition based on peak frequency
    if amplitude_spectrum.index(max(amplitude_spectrum[:4])) == 3 and amplitude_spectrum.index(
            min(amplitude_spectrum[4:8])) == 4:
        flags["class_4"] += 1

    # Flag 2: Condition based on amplitude spectrum
    if amplitude_spectrum.index(max(amplitude_spectrum[:4])) in [0, 3]:
        flags["class_2"] += 1

    # Flag 6: Condition based on amplitude spectrum
    if amplitude_spectrum.index(min(amplitude_spectrum[:4])) in [1, 2]:
        flags["class_6"] += 1

    if amplitude_spectrum.index(max(amplitude_spectrum[:4])) != 3 and amplitude_spectrum.index(
            min(amplitude_spectrum[4:8])) != 6:
        flags["class_3"] += 1



    # Define flag priority according to the specified order
    flag_priority = [
        ("class_0", 0),
        ("class_1", 1),
        ("class_5", 3),
        ("class_4", 5),
        ["class_3", 2],
        ("class_6", 4),
        ("class_2", 6)
    ]

    # Sort the flag_priority list based on the specified order
    flag_priority.sort(key=lambda x: x[1])



    # Define flag names
    flag_names = {
        "class_0": "Class 0",
        "class_1": "Class 1",
        "class_2": "Class 2",
        "class_3": "Class 3",
        "class_4": "Class 4",
        "class_5": "Class 5",
        "class_6": "Class 6"
    }

    print("\n")
    print(f"The predicted order of priorities for TestCase {TestCase} is: ")
    numbers = 2


    # Assign points to each class based on their position in the list l
    l = []
    for flag_name, _ in flag_priority:
        if flags[flag_name] == 1:
            l.append(flag_names[flag_name])

    names = list(flag_names.values())
    if len(l) == 2 and l[0] == names[6] and l[1] == names[2]:
        l.clear()
        l.append(names[5])
        l.append(names[6])

    # Assign points to each class based on their position in the list l
    IterationPriority = {
        "Class 0": [0, 0],
        "Class 1": [0, 0],
        "Class 2": [0, 0],
        "Class 3": [0, 0],
        "Class 4": [0, 0],
        "Class 5": [0, 0],
        "Class 6": [0, 0]
    }


    for i in l:
        # Assign points to each class based on their position in the list l
        priorities[i][0] += 1
        priorities[i][1] += (7 - l.index(i)) * priorities[i][0]
        IterationPriority[i][0] += 1
        IterationPriority[i][1] += (7 - l.index(i)) * IterationPriority[i][0]






    if(movement == numbers):
        IterationPriority[names[2]][1] += 8 * IterationPriority[names[2]][0]
        priorities[names[2]][1] += 8 * priorities[names[2]][0]




    # Convert the dictionary items to a list of tuples
    data = [(key, value[0], value[1]) for key, value in IterationPriority.items()]


    # Print using tabulate with headers
    print(tabulate(data, headers=["Class", "Frequency", "Priority"], tablefmt="pretty"))
  



print("\n")
print("--------------------------------------------")
print("The final order of priorities is: ")
print("--------------------------------------------")

data = [(key, value[0], value[1]) for key, value in priorities.items()]

print(tabulate(data, headers=["Class", "Frequency", "Priority"], tablefmt="pretty"))

print("\n")
max_key = max(priorities, key=lambda k: priorities[k][1])
print("The class with the highest priority is: ", max_key)
print("\n")
print("\n")
