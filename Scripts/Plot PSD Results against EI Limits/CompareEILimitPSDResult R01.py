# -*- coding: utf-8 -*-
"""
Reading 4 PSD Vibration Velocity results, integrate to Autopower and compare with EI Limits
@author: Martin
@Updated: Alastair
TODO: optimise the display layout for any number of point
"""
import FemapAPI
import numpy as np
import matplotlib.pyplot as plt

#Gets the X, Y and Z velocity PSDs for a set of given nodes, and range of output sets
def get_psd_data(app, start_set, end_set, nodes):
    freq = []
    psd_X = [[] for _ in nodes]
    psd_Y = [[] for _ in nodes]
    psd_Z = [[] for _ in nodes]

    for i in range(start_set, end_set):
        freq.append(app.getSetValue(i))
        
        for j, node in enumerate(nodes):
            psd_X[j].append(app.getEntityValue(i, 12, node))
            psd_Y[j].append(app.getEntityValue(i, 13, node))
            psd_Z[j].append(app.getEntityValue(i, 14, node))
    
    return np.array(freq), np.array(psd_X), np.array(psd_Y), np.array(psd_Z)

#Plot velocity PSDs against EI limits
#Input the X, Y, and Z PSD functions, frequency, plot name and ax setting
def plot_psd_results(ax, freq, psd_X, psd_Y, psd_Z, title):
    plt.rcParams.update({'font.size': 18})
    delta_freq = freq[1] - freq[0]

    autopowerX = np.sqrt(psd_X * delta_freq) * 1000
    autopowerY = np.sqrt(psd_Y * delta_freq) * 1000
    autopowerZ = np.sqrt(psd_Z * delta_freq) * 1000

    ax.plot(freq, autopowerX, 'b', label='x')
    ax.plot(freq, autopowerY, 'r', label='y')
    ax.plot(freq, autopowerZ, 'g', label='z')
    ax.plot(freq, 10**((np.log10(freq) + 0.48017) / 2.127612), 'y', label='Concern Limit')
    ax.plot(freq, 10**((np.log10(freq) + 1.871083) / 2.084547), 'r', label='Problem Limit')
    
    ax.legend(loc='upper right')
    ax.set_ylim((0, 60))
    ax.set_xlim((0, 50))
    ax.set_title(title)
    ax.set_xlabel('Frequency [Hz]')
    ax.set_ylabel('Velocity [mm/s]')

#Plot XYZ PSDs for four nodes against EI Limits.
#Use config .ini file to set values
def main():
    import configparser

    config = configparser.ConfigParser()
    config.read('model-config.ini')
    
    app = FemapAPI.App()
    
    output_set_start = int(config['DEFAULT']['output_set_start'])
    no_sets = int(config['DEFAULT']['no_sets'])
    model_name = config['DEFAULT']['model_name']
    nodes = [int(i) for i in config['DEFAULT']['nodes'].split(',')]
    point_name = [int(i) for i in config['DEFAULT']['point_name'].split(',')]

    output_set_end = output_set_start + no_sets

    file_name = model_name + '-PTS-' + "-".join(map(str, point_name)) + '.png'

    freq, psd_X, psd_Y, psd_Z = get_psd_data(app, output_set_start, output_set_end, nodes)

    _, axes = plt.subplots(nrows=2, ncols=2, figsize=(20, 12))  # Create a 2x2 grid of subplots

    for i, (x, y, z, ax) in enumerate(zip(psd_X, psd_Y, psd_Z, axes.ravel())):
        title = ax.set_title("Point {}".format(point_name[i]))
        plot_psd_results(ax, freq, x, y, z, title)
        
    plt.tight_layout(pad=3)
    file_name = "{}-PTS.png".format(model_name)
    plt.savefig(file_name, dpi=300, bbox_inches='tight')
    plt.close()
    
if __name__ == "__main__":
    main()