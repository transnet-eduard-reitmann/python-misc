import sys
import os
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

# Add WIM repo to path
sys.path.insert(0, r"C:\Users\ereit\GitHub\Work\wim-wim-angle-of-attack\src")

from aoa_system.utils.file_operations import WIMFileReader, FileManager

# Setup
wim_reader = WIMFileReader()
file_manager = FileManager()
test_data_path = r"C:\Users\ereit\GitHub\Work\wim-wim-angle-of-attack\test_data"

# Get WIM files
wim_files, _ = file_manager.get_files_by_extension(test_data_path, "WIM")

# Plot each file
for wim_file in wim_files:
    print(f"Plotting: {os.path.basename(wim_file)}")
    
    # Load data
    time_vector, channel_data, info = wim_reader.import_wim_data(wim_file)
    
    # Plot all channels on same plot
    plt.figure()
    
    for i in range(13):
        y = (channel_data[:, i] - np.median(channel_data[:, i]))
        y = y / np.max(y)
        plt.plot(time_vector, y + i + 1, label=info["channel_names"][i])

    for i in range(13, 26):
        y = (channel_data[:, i] - np.median(channel_data[:, i]))
        y = y / np.max(y)
        plt.plot(time_vector, -y - (i - 13) - 1, label=info["channel_names"][i])    
    
    plt.title(f'{info["filename"]}')
    plt.xlabel('Time (s)')
    plt.ylabel('Amplitude') 
    plt.show()


    ergasergerag