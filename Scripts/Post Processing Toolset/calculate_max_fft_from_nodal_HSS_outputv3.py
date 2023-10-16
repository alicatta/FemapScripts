import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import butter, lfilter, find_peaks
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import tkinter as tk
from tkinter import filedialog, simpledialog
import re
import sestoolkit.array.timeseries as ts


# Global Constants
SHEET_NAME = "HSS Array"
TIME_STEP = 0.002
FS = 1 / TIME_STEP

def highpass_filter(data, cutoff, fs, order=5):
    """ High-pass filter for the data """
    nyq = 0.5 * fs
    normal_cutoff = cutoff / nyq
    b, a = butter(order, normal_cutoff, btype='high', analog=False)
    y = lfilter(b, a, data)
    return y

def perform_fft(node_data):
    """Calculate FFT for the given node data."""
    Nfft = len(node_data)
    df = FS / Nfft
    f = np.arange(0, Nfft) * df
    Nyq = f[int(Nfft/2)]
    f[f > Nyq] = f[f > Nyq] - (2 * Nyq)
    fTwoSided = f.copy()
    f = f[f >= 0]

    windowCoefficients = np.hanning(len(node_data)) * 2  # Hanning window is used
    node_data = node_data * windowCoefficients

    rawFFT = np.fft.fft(node_data)
    positive_f_indices = np.where(f >= 0)[0]
    F = rawFFT[positive_f_indices]
    F[1:int(np.ceil(Nfft/2))] = F[1:int(np.ceil(Nfft/2))] * 2

    return f, np.abs(F)

def calculate_max_fft_from_weld_output(weld_output_filepath):
    df = pd.read_excel(weld_output_filepath, sheet_name=SHEET_NAME)
    df.drop(columns=['Distance', 'Inner Node ID', 'Outer Node ID'], inplace=True)
    df /= 1e9
    
    all_fftAmp = [perform_fft(highpass_filter(row.values, 0.5, FS))[1] for _, row in df.iterrows()]
    all_fftAmp = [ts.fft(highpass_filter(row.values, 0.5, FS))[1] for _, row in df.iterrows()]
    max_fftAmp = np.max(all_fftAmp, axis=0)

    return perform_fft(df.iloc[0].values)[0], max_fftAmp  # Frequency remains the same for all rows

def calculate_max_fft_from_weld_output2(weld_output_filepath):
    df = pd.read_excel(weld_output_filepath, sheet_name="HSS Array")
    df = df.drop(columns=['Distance', 'Inner Node ID', 'Outer Node ID'])
    df = df / 1e6
    time_step = 0.002
    Fs = 1 / time_step
    
    all_fftAmp = []

    for index, row in df.iterrows():
        node_data = row.values
        node_data = highpass_filter(node_data, 0.5, Fs)

        # FFT calculation
        Nfft = len(node_data)
        df = Fs / Nfft
        f = np.arange(0, Nfft) * df
        Nyq = f[int(Nfft/2)]
        f[f > Nyq] = f[f > Nyq] - (2 * Nyq)
        fTwoSided = f.copy()
        f = f[f >= 0]

        window = "Hanning"
        if window == "Hanning":
            windowCoefficients = np.hanning(len(node_data)) * 2
        elif window == "Uniform":
            windowCoefficients = np.ones(len(node_data))

        node_data = node_data * windowCoefficients
        rawFFT = np.fft.fft(node_data)
        positive_f_indices = np.where(f >= 0)[0]
        F = rawFFT[positive_f_indices]
        F[1:int(np.ceil(Nfft/2))] = F[1:int(np.ceil(Nfft/2))] * 2
        fftAmp = np.abs(F)

        all_fftAmp.append(fftAmp)

    max_fftAmp = np.max(all_fftAmp, axis=0)
    return f, max_fftAmp

def annotate_peaks(frequencies, amplitudes, ax, all_weld_names=None):
    # Find peaks with a prominence threshold
    peaks, _ = find_peaks(amplitudes, prominence=1)  # The prominence parameter helps identify true peaks
    peaks = peaks[np.argsort(amplitudes[peaks])][-3:]  # Sort the peaks by amplitude and select the top 3

    # Annotate the peaks
    for peak in peaks:
        if all_weld_names:  # If a list of weld names is provided, extract the name for the current peak
            weld_name = all_weld_names[peak]
            annotation_text = f'{weld_name}\n({frequencies[peak]:.2f}, {amplitudes[peak]:.2f})'
        else:
            annotation_text = f'({frequencies[peak]:.2f}, {amplitudes[peak]:.2f})'
        
        plt.annotate(annotation_text, 
                    (frequencies[peak], amplitudes[peak]),
                    textcoords="offset points",
                    xytext=(0,10),
                    ha='center')   

def plot_fft(f, max_fftAmp, save_directory, file_name, freq_limit=None, y_limit=None, subtitle="", display=False, save=False):
    plt.figure(figsize=(15, 6))
    plt.plot(f, max_fftAmp)
    plt.title(f"Max Hot Spot Stress Across All Weld Nodes\n{subtitle}\n{file_name}")
    plt.xlabel("Frequency (Hz)")
    plt.ylabel("Hot Spot Stress (MPa)")
    
    annotate_peaks(f, max_fftAmp, plt.gca())  # Call the function to annotate the peaks

    if y_limit:
        plt.ylim(0, y_limit)

    if freq_limit:
        plt.xlim(freq_limit)
        major_locator = MultipleLocator(freq_limit[1]/5)
        plt.gca().xaxis.set_major_locator(major_locator)

        minor_locator = MultipleLocator(freq_limit[1]/10)
        plt.gca().xaxis.set_minor_locator(minor_locator)

        tertiary_locator = MultipleLocator(freq_limit[1]/50)
        minor_ticks = list(set(tertiary_locator.tick_values(freq_limit[0], freq_limit[1])) - set(minor_locator.tick_values(freq_limit[0], freq_limit[1])))
        plt.gca().set_xticks(minor_ticks, minor=True)

        plt.gca().xaxis.set_major_formatter(FormatStrFormatter('%.2f'))
        
    if display:
        plt.show()

    if save:
        output_file = os.path.join(save_directory, f"{file_name}_MaxFFT.png")
        plt.savefig(output_file)
        print(f"Saved data for: {file_name}.")

    plt.close()

def get_plot_settings():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    max_freq = simpledialog.askfloat("Input", "Please enter max frequency:", initialvalue=50)
    y_axis_choice = simpledialog.askstring("Y-axis Setting", "Do you want an auto y-axis max for each plot or the same y-axis max for all plots?\nChoose 'auto' or 'same':", initialvalue="auto")

    return max_freq, y_axis_choice

def get_user_input(prompt, initialvalue, datatype):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    if datatype == "float":
        return simpledialog.askfloat("Input", prompt, initialvalue=initialvalue)
    elif datatype == "string":
        return simpledialog.askstring("Input", prompt, initialvalue=initialvalue)

def get_directory_paths():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    directories = []
    subtitles = []
    
    while True:
        directory_path = filedialog.askdirectory(title="Select a Directory (cancel when done)")
        if not directory_path:  # Break the loop when the user cancels the dialog
            break
        folder_name = os.path.basename(directory_path)
        subtitle = get_subtitle(folder_name)
        
        directories.append(directory_path)
        subtitles.append(subtitle)

    return directories, subtitles

def get_subtitle(default_name):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    subtitle = simpledialog.askstring("Subtitle", "Please enter a subtitle for the plot:", initialvalue=default_name)
    return subtitle

def ensure_plots_directory(directory):
    """
    Checks if a 'plots' directory exists in the given directory. If not, it creates one.
    Returns the path to the 'plots' directory.
    """
    plots_directory = os.path.join(directory, "plots")
    if not os.path.exists(plots_directory):
        os.makedirs(plots_directory)
    return plots_directory

def process_directory(directory, subtitle, max_freq, y_axis_choice=None):

    weld_files = [file for file in os.listdir(directory) if file.endswith('_HSS.xlsx')]
    fft_data_collection = []

    # Ensure the plots directory exists and get its path
    save_directory = ensure_plots_directory(directory)
    
    all_amplitudes = []
    all_frequencies = []
    all_weld_names = []

    for index, filename in enumerate(weld_files):
        f, max_fftAmp = calculate_max_fft_from_weld_output(os.path.join(directory, filename))
        all_amplitudes.extend(max_fftAmp)
        all_frequencies.extend(f)
        label_name = os.path.splitext(filename)[0]  # Removing the .xlsx extension
        all_weld_names.extend([label_name for _ in range(len(f))])
        plot_fft(f, max_fftAmp, save_directory, label_name, freq_limit=(0, max_freq), subtitle=subtitle, save=True)
        fft_data_collection.append((f, max_fftAmp, filename))
    
    # Convert lists to numpy arrays for easier indexing
    all_amplitudes = np.array(all_amplitudes)
    all_frequencies = np.array(all_frequencies)

    # Plotting all the data on a single graph
    # Using the "Paired" colormap for colorblind-friendly colors
    colors = plt.cm.Paired(np.linspace(0,1,len(fft_data_collection)))
    plt.figure(figsize=(15, 6))
    for idx, (f, max_fftAmp, filename) in enumerate(fft_data_collection):
        label_name = os.path.splitext(filename)[0]
        plt.plot(f, max_fftAmp, label=label_name, color=colors[idx])

    annotate_peaks(all_frequencies, all_amplitudes, plt.gca(), all_weld_names=all_weld_names)  # Call the function to annotate the peaks

    plt.title(f"Summary FFT Amplitude Across All Nodes and all Welds\n{subtitle}")
    plt.xlabel("Frequency (Hz)")
    plt.ylabel("Hot Spot Stress (MPa)")
    plt.legend(loc='best')
    plt.xlim(0, max_freq)

    # Save and show the summary plot
    plt.savefig(os.path.join(save_directory, "Summary_Plot.png"))
    if len(directories) == 1:  # Only plot the summary graph if one directory is selected
        plt.show()
    else:  
        plt.close()

    print(f"All files in {directory} processed!")

if __name__ == "__main__":
    directories, subtitles = get_directory_paths()

    max_freq = get_user_input("Please enter max frequency:", 50, "float")
    #y_axis_choice = get_user_input("Do you want an auto y-axis max for each plot or the same y-axis max for all plots?\nChoose 'auto' or 'same':", "auto", "string")

    for directory, subtitle in zip(directories, subtitles):
        process_directory(directory, subtitle, max_freq)