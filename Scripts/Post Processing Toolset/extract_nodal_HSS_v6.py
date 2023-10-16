import os
import FemapAPI
import numpy as np
import time
import pandas as pd
import scipy.spatial
import sys
import tkinter as tk
from tkinter import filedialog, simpledialog, ttk
import FLife

import tkinter as tk
from tkinter import messagebox


##GUI Functions

def prompt_continue_or_exit_gui(message):
    def on_yes():
        root.destroy()

    def on_no():
        messagebox.showinfo("Exit", "Exiting the script...")
        root.destroy()

    root = tk.Tk()
    root.title("Continue or Exit")

    label = tk.Label(root, text=message)
    label.pack(pady=20)

    yes_button = tk.Button(root, text="Yes (y)", command=on_yes)
    yes_button.pack(side="left", padx=20)

    no_button = tk.Button(root, text="No (n)", command=on_no)
    no_button.pack(side="right", padx=20)

    root.mainloop()
    
def get_directory_path(message=None):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    if not message:
        message="Select a Directory"
    directory_path = filedialog.askdirectory(title=message)

    return directory_path

def get_file_path(message=None):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    if not message:
        message="Select a File"
    
    file_path = filedialog.askopenfilename(title=message)

    return file_path

def get_user_input(prompt, initialvalue, datatype):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    if datatype == "float":
        return simpledialog.askfloat("Input", prompt, initialvalue=initialvalue)
    elif datatype == "string":
        return simpledialog.askstring("Input", prompt, initialvalue=initialvalue)

class MultiInputDialog(simpledialog.Dialog):
    def __init__(self, parent, prompts, initialvalues, datatypes):
        self.entries = []
        self.datatypes = datatypes
        self.results = []
        self.prompts = prompts
        self.initialvalues = initialvalues
        super().__init__(parent)

    def body(self, master):
        for i, prompt in enumerate(self.prompts):
            ttk.Label(master, text=prompt).grid(row=i, column=0)
            entry = ttk.Entry(master)
            entry.insert(0, str(self.initialvalues[i]))
            entry.grid(row=i, column=1)
            self.entries.append(entry)
        return self.entries[0]  # set initial focus

    def apply(self):
        for i, entry in enumerate(self.entries):
            if self.datatypes[i] == "float":
                self.results.append(float(entry.get()))
            else:
                self.results.append(entry.get())

def get_user_inputs(prompts, initialvalues, datatypes):
    root = tk.Tk()
    root.withdraw()
    dialog = MultiInputDialog(root, prompts, initialvalues, datatypes)
    # No need for root.mainloop() since the Dialog class already manages this.
    root.destroy()  # Ensure the root window is destroyed after the dialog is closed.
    return dialog.results


##CLI Prompt Dialogues

def should_continue():
    while True:
        choice = input("Do you want to continue? (y/n): ").lower()
        if choice == 'y':
            return True
        elif choice == 'n':
            return False
        else:
            print("Invalid input. Please enter 'y' or 'n'.")

def prompt_continue_or_exit():
    while True:
        choice = input("(y/n)").lower()
        if choice == 'y':
            return
        elif choice == 'n':
            sys.exit("Exiting the script...")
        else:
            print("Invalid input. Please enter 'y' or 'n'.")


##Time progressed functions

def convert_timer_to_readable(secs):
    mins, s = divmod(secs, 60)
    h, mins = divmod(mins, 60)

    h_str = f'{int(h)}h ' if h > 0 else ''
    mins_str = f'{int(mins)}m ' if mins > 0 else ''
    s_str = f'{round(s, 2)}s'

    return h_str + mins_str + s_str

def get_elapsed_time_readable(start_time):
    elapsed_time = time.time() - start_time
    return convert_timer_to_readable(elapsed_time)



##File Reading and Saving

def read_file(filepath):
    if filepath.endswith('.xlsx') or filepath.endswith('.xls'):
        return pd.read_excel(filepath)
    elif filepath.endswith('.csv'):
        return pd.read_csv(filepath)
    else:
        raise ValueError(f"Unsupported file type for: {filepath}")

def save_dataframe(df, filepath, **kwargs):
    if filepath.endswith('.xlsx') or filepath.endswith('.xls'):
        df.to_excel(filepath, **kwargs)
    elif filepath.endswith('.csv'):
        df.to_csv(filepath, **kwargs)
    else:
        raise ValueError(f"Unsupported file type for: {filepath}")

## FEA Post Processing

def combined_locs(inner_nodes, outer_nodes):
    
    tree = scipy.spatial.KDTree(inner_nodes[:, 1:4])
    
    combined_data = []
    distances = []

    used_inner_node_ids = set()  # Keep track of used inner nodes

    for outer_node in outer_nodes:
        dist, inner_node_id = tree.query(outer_node[1:4])

        # Skip if this inner_node_id has already been used
        if inner_node_id in used_inner_node_ids:
            continue

        used_inner_node_ids.add(inner_node_id)  # Mark this inner node as used

        _inner_match = inner_nodes[inner_node_id, :]
        combined_data.append(np.hstack((_inner_match, outer_node)))
        distances.append(dist)

    combined_data_df = pd.DataFrame(combined_data)
    combined_data_df.columns = ['Inner Node ID', 'Inner X', 'Inner Y', 'Inner Z', 'Outer Node ID', 'Outer X', 'Outer Y', 'Outer Z']

    return combined_data_df, distances

def extract_MPS_HSS(app, start_set, end_set, matched_nodes, name=None):

        
    num_nodes = len(matched_nodes)
    total_iterations = int(end_set) - int(start_set)

    inner_stress_array = np.zeros((num_nodes, total_iterations))
    outer_stress_array = np.zeros((num_nodes, total_iterations))

    # Create a Femap set for inner and outer nodes in matched_nodes

    # Create a Femap set for the nodes
    inner_node_set = app.create_set_of_nodes(node_list=list(matched_nodes['Inner Node ID']))
    outer_node_set = app.create_set_of_nodes(node_list=list(matched_nodes['Outer Node ID']))

    # Create a Femap set for the test output sets
    output_sets = app.create_set_of_outputs(output_ids=range(int(start_set), int(end_set)))

    # Fetch results for this vector
    vector_id = 24000000

    if not name:
        inner_node_name=None
        outer_node_name=None
    else:
        inner_node_name=f"{name} inner node "
        outer_node_name=f"{name} outer node "

    print(f"Fetching {inner_node_name}results from Femap:")

    inner_stress_array_fetched = app.get_node_results(output_sets, inner_node_set, [vector_id])

    print(f"Fetching {outer_node_name}results from Femap:")
    outer_stress_array_fetched = app.get_node_results(output_sets, outer_node_set, [vector_id])

    # After fetching results, print some diagnostics:
    #print(f"Inner stress array shape: {inner_stress_array.shape}")
    #print(f"Outer stress array shape: {outer_stress_array.shape}")
    #print(f"Inner stress array min, max, mean values: {np.min(inner_stress_array)}, {np.max(inner_stress_array)}, {np.mean(inner_stress_array)}")
    #print(f"Outer stress array min, max, mean values: {np.min(outer_stress_array)}, {np.max(outer_stress_array)}, {np.mean(outer_stress_array)}")
    

    # Convert to NumPy arrays if they are DataFrames
    if isinstance(inner_stress_array_fetched, pd.DataFrame):
        inner_stress_array_fetched = inner_stress_array_fetched.to_numpy()
    if isinstance(outer_stress_array_fetched, pd.DataFrame):
        outer_stress_array_fetched = outer_stress_array_fetched.to_numpy()

    # Create a dictionary mapping node IDs to their index positions in the original DataFrame
    inner_node_id_to_index = {node_id: index for index, node_id in enumerate(matched_nodes['Inner Node ID'])}
    outer_node_id_to_index = {node_id: index for index, node_id in enumerate(matched_nodes['Outer Node ID'])}

    # Sort node IDs in descending order (the order in fetched arrays)
    sorted_inner_node_ids = sorted(matched_nodes['Inner Node ID'], reverse=True)
    sorted_outer_node_ids = sorted(matched_nodes['Outer Node ID'], reverse=True)

    # Create index arrays for reordering
    inner_index_mapping = [inner_node_id_to_index[node_id] for node_id in sorted_inner_node_ids]
    outer_index_mapping = [outer_node_id_to_index[node_id] for node_id in sorted_outer_node_ids]

    # Add debug prints to understand the sizes and shapes
    print("Shape of inner_stress_array_fetched:", inner_stress_array_fetched.shape)
    print("Shape of outer_stress_array_fetched:", outer_stress_array_fetched.shape)
    print("Length of matched_nodes['Inner Node ID']:", len(matched_nodes['Inner Node ID']))
    print("Length of matched_nodes['Outer Node ID']:", len(matched_nodes['Outer Node ID']))
    print("Max index in inner_index_mapping:", max(inner_index_mapping))
    print("Max index in outer_index_mapping:", max(outer_index_mapping))

    # Then proceed with the reordering
    try:
        inner_stress_array = inner_stress_array_fetched[np.array(inner_index_mapping), :]
    except IndexError as e:
        print("IndexError encountered:", e)

    outer_stress_array = outer_stress_array_fetched[np.array(outer_index_mapping), :]

    # Compute HSS and print diagnostics:
    hss_array = 1.67 * inner_stress_array - 0.67 * outer_stress_array
    print(f"HSS array shape: {hss_array.shape}")
    print(f"HSS array min, max, mean values: {np.min(hss_array)}, {np.max(hss_array)}, {np.mean(hss_array)}")
    
    return inner_stress_array, outer_stress_array, hss_array

def main_process(time_step, length, start_set, weld_node_filepath, save_directory, constants):
       
    start_time = time.time()
    app = FemapAPI.App()
    print(f"Connected to Femap. Elapsed time: {get_elapsed_time_readable(start_time)}")
    
    no_sets = length / time_step
    end_set = start_set + no_sets
    
    rc = app._rbo.VectorExistsV2(nSetID = start_set, nVectorID = 24000000)

    if rc == 0:
        nodal_stress_exists = False
        print(f"Nodal stress vector '24000000' not detected. Convert Max Principal Stress (vector id = 60016) to nodal stress?")
        prompt_continue_or_exit()
        print(f"Beginning Conversion of Max Principal Stress for {start_set} to {end_set}.")
        app.convert_nodal_stress(range(int(start_set), int(end_set)), 60016, 105)
    else:
        nodal_stress_exists = True
        print(f"Vector '24000000' detected. Use this vector to calculate HSS?")
        prompt_continue_or_exit()
    

    df = read_file(weld_node_filepath)
    unique_welds = df["Weld"].unique()
    results = {f"{weld}": None for weld in unique_welds}        
    
    for weld in unique_welds:
        #print(f"Processing Weld: {weld}. Elapsed time: {get_elapsed_time_readable(start_time)}")
        df_weld = df[df["Weld"] == weld]

        print(f"Beginning extraction of {weld} inner nodes. Elapsed time: {get_elapsed_time_readable(start_time)}")
        inner_nodes = df_weld[df_weld["Extrapolation Point"] == "0.4t"].drop(columns=['CSys ID', 'Weld', 'Extrapolation Point']).dropna().to_numpy()
        print(f"Extracted {weld} inner nodes. Elapsed time: {get_elapsed_time_readable(start_time)}")

        print(f"Beginning extraction of {weld} outer nodes. Elapsed time: {get_elapsed_time_readable(start_time)}")
        outer_nodes = df_weld[df_weld["Extrapolation Point"] == "1t"].drop(columns=['CSys ID', 'Weld', 'Extrapolation Point']).dropna().to_numpy()
        print(f"Extracted {weld} outer nodes. Elapsed time: {get_elapsed_time_readable(start_time)}")
        
        # Match nodes before extracting stresses
        print(f"Beginning {weld} node pairing for {len(inner_nodes)} inner nodes and {len(outer_nodes)} outer nodes. Elapsed time: {get_elapsed_time_readable(start_time)}")
        matched_nodes, dist_array = combined_locs(inner_nodes, outer_nodes)

        print(f"Created {len(matched_nodes['Inner Node ID'])} node pairs for {weld}. Elapsed time: {get_elapsed_time_readable(start_time)}")

        print(f"Begin extraction of {weld} nodal hotspot stresses. Elapsed time: {get_elapsed_time_readable(start_time)}")
        inner_stress_array, outer_stress_array, hss_array = extract_MPS_HSS(app, start_set, end_set, matched_nodes, name = f"{weld}")
        print(f"Completed extraction of {weld} nodal hotspot stresses. Elapsed time: {get_elapsed_time_readable(start_time)}")

        hss_df = pd.DataFrame(hss_array)
        hss_df = hss_df.drop(hss_df.columns[0], axis=1)
        inner_stress_df = pd.DataFrame(inner_stress_array)
        outer_stress_df = pd.DataFrame(outer_stress_array)

        max_hss_values = hss_df.max(axis=1)
        min_hss_values = hss_df.min(axis=1)
        mean_hss_values = hss_df.mean(axis=1)
        median_hss_values = hss_df.median(axis=1)
        hss_summary_df = pd.DataFrame({
            'Max HSS': max_hss_values,
            'Min HSS': min_hss_values,
            'Mean HSS': mean_hss_values,
            'Median HSS': median_hss_values
        })

        overall_max_hss = hss_df.max().max()
        results[f"{weld}"] = overall_max_hss * 1E-6

        hss_columns = list(hss_df.columns)
        inner_stress_columns = list(inner_stress_df.columns)
        outer_stress_columns = list(outer_stress_df.columns)
        
        num_decimal_places = len(str(time_step).split('.')[-1])
        #time_array = np.linspace(0, no_sets*time_step, hss_df.shape[1]) 
        time_array = np.arange(0, no_sets*time_step, time_step)

        inner_stress_columns[0] = f"Inner Node ID"
        outer_stress_columns[0] = f"Outer Node ID"

        for i in range(0, len(hss_columns)):
            hss_columns[i] = f"MPS (HSS); Time {time_array[i]:.{num_decimal_places}f}"
            inner_stress_columns[i+1] = f"MPS (Inner Stress); Time {time_array[i]:.{num_decimal_places}f}"
            outer_stress_columns[i+1] = f"MPS (Outer Stress); Time {time_array[i]:.{num_decimal_places}f}"
        
        hss_df.columns = hss_columns
        inner_stress_df.columns = inner_stress_columns
        outer_stress_df.columns = outer_stress_columns

        distances_df = pd.DataFrame(dist_array,columns=["Distance"])

        hss_temp_df1 = distances_df.join(hss_df)
        hss_temp_df2 = outer_stress_df[["Outer Node ID"]].join(hss_temp_df1)
        hss_df = inner_stress_df[["Inner Node ID"]].join(hss_temp_df2)

        hss_summary_temp_df1 = distances_df.join(hss_summary_df)
        hss_summary_temp_df2 = outer_stress_df[["Outer Node ID"]].join(hss_summary_temp_df1)
        hss_summary_df = inner_stress_df[["Inner Node ID"]].join(hss_summary_temp_df2)        
        savename = f"{weld}_HSS.xlsx"
        weld_save_filepath = os.path.join(save_directory, savename)

        # Save both DataFrames to the same Excel file but in different worksheets
        try:
            with pd.ExcelWriter(weld_save_filepath) as writer:
                hss_summary_df.to_excel(writer, sheet_name="HSS Summary", index=False)
                hss_df.to_excel(writer, sheet_name="HSS Array", index=False)
                inner_stress_df.to_excel(writer, sheet_name="0.4t Stress Array", index=False)
                outer_stress_df.to_excel(writer, sheet_name="1t Stress Array", index=False)
                matched_nodes.to_excel(writer, sheet_name="matched_nodes", index=False)
            print(f"Saved results for {weld} to {savename}. Elapsed time: {get_elapsed_time_readable(start_time)}")

        except Exception as e:
            print(f"Error saving results for {weld}: {e}. Elapsed time: {get_elapsed_time_readable(start_time)}")

    results_df = pd.DataFrame(list(results.items()), columns=['Weld', 'Max HSS (MPa)'])
    results_df['Constants'] = constants[:len(results_df)]

    results_df['Scaled HSS (MPa)'] = results_df['Max HSS (MPa)'] * results_df['Constants']
    results_df['Scaled HSS pk-pk (MPa)'] = results_df['Scaled HSS (MPa)'] * 2

    summary_save_filepath = os.path.join(save_directory, "Weld_Stress_Summary.xlsx")
    try:
        with pd.ExcelWriter(summary_save_filepath) as writer:
            results_df.to_excel(writer, sheet_name="Weld Stress Summary", index=False)
        os.system(f'start excel "{summary_save_filepath}"')
    except Exception as e:
        print(f"Error saving Weld Stress Summary: {e}. Elapsed time: {get_elapsed_time_readable(start_time)}")

if __name__ == "__main__":
    prompts = ["Timestep for analysis", "Input total time length for analysis", "Input first output set"]
    initialvalues = [0.002, 18, 69]
    datatypes = ["float", "float", "float"]

    time_step, length, start_set = get_user_inputs(prompts, initialvalues, datatypes)
    #time step in seconds
    #time_step = 0.002understanding of what I can help you with from a technica
    #time_step = get_user_input("Timestep for analysis", 0.002, "float" )
    #total time in seconds
    #length = 18
    #length = get_user_input("Input total time length for analysis", 18, "float" )
    #Femap set no of first time step
    #start_set = 9010
    #start_set = get_user_input("Input first output set", 9010, "float" )
    #constants to scale weld stress
    constants = [0.929, 0.904, 1.169, 0.796, 1.331]
    #constants_filepath = get_file_path("Select the scaling factor .xlsx file")
    #constants=read_file(constants_filepath).to_numpy
    print(f"Scaling factors imported are: {constants}")
    weld_node_filepath = get_file_path("Select the weld node .xlsx file")
    save_directory = get_directory_path("Select the save directory")

    main_process(time_step, length, start_set, weld_node_filepath, save_directory, constants)