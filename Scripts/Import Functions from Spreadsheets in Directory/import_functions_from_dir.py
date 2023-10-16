import os
import csv
import tkinter as tk
from tkinter import filedialog
import FemapAPI

app = FemapAPI.App()

def get_csv_files_from_directory(directory_path):
    return [f for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f)) and f.endswith('.csv')]

def read_csv_file(file_path):
    with open(file_path, 'r') as f:
        reader = csv.reader(f)
        data = list(reader)
    return data

def main():
    # Check for import_list.csv
    script_dir = os.path.dirname(os.path.realpath(__file__))
    import_list_path = os.path.join(script_dir, "import_list.csv")

    import_all = True
    functions_to_import = set()

    if os.path.exists(import_list_path):
        with open(import_list_path, 'r') as f:
            functions_to_import = set(line.strip() for line in f)
        import_all = False

    # Create a Tkinter root window (used for file dialog)
    root = tk.Tk()
    root.withdraw()

    # Prompt the user to select a directory
    directory_path = filedialog.askdirectory()
    #directory_path = input("Please enter the directory containing the CSV files: ")
    csv_files = get_csv_files_from_directory(directory_path)

    print(f"Found {len(csv_files)} CSV files in the directory.")

    for file in csv_files:
        print(f"Processing file: {file}")
        data = read_csv_file(os.path.join(directory_path, file))
        X = [float(row[0]) for row in data[1:]]
        print(f"Extracted {len(X)} X values.")
        
        for col in range(1, len(data[0])):  # Starting from 1 to skip the X values column
            F_Title = data[0][col]

            if not import_all and F_Title not in functions_to_import:
                print(f"Skipping function titled '{F_Title}' as it's not in the import list.")
                continue

            y = [float(row[col]) for row in data[1:]]
            print(f"Processing function titled '{F_Title}' with {len(y)} Y values.")
            
            try:
                rc = app.createFunction(F_Title, X, y)
                print(f"Function '{F_Title}' added with ID: {rc}")
            except Exception as e:
                print(f"Error while creating function for {F_Title}: {e}")
    app.rebuild()

if __name__ == "__main__":
    main()