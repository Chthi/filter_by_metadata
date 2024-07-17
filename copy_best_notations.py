"""
This script is used to filter and copy the best-rated images from a source folder to a destination folder based on their "Notation" in metadata.

The script performs the following steps:
1. Deletes all files and folders in the destination folder.
2. Copies the files with a notation equal to or higher than the threshold to the destination folder.
Note: The metadata notation is extracted using the Windows Shell API.
Note2: This probably only works on Windows, and when setup in french (it changes the metadata key names).

Usage:
- In the user_configuration.json file, set the INPUT_PATH, OUTPUT_PATH, and NOTATION variables.
- Set the INPUT_PATH variable to the path of the source folder.
- Set the OUTPUT_PATH variable to the path of the destination folder (Warning ! It will be emptied).
- Set the NOTATION variable to the desired threshold notation (Between 0 and 5).
"""

""" START OF SCRIPT """

import win32com.client
import os
from pathlib import Path
import shutil
import tqdm
import json

with open("user_configuration.json", "r") as f:
    config = json.load(f)
INPUT_PATH = config["INPUT_PATH"]
OUTPUT_PATH = config["OUTPUT_PATH"]
NOTATION = config["NOTATION"]

# Only for french windows
txt_to_notation = {
    "Non classé": 0,
    "1 étoile": 1,
    "2 étoiles": 2,
    "3 étoiles": 3,
    "4 étoiles": 4,
    "5 étoiles": 5,
}

if OUTPUT_PATH == "" or INPUT_PATH == "":
    raise ValueError("Please set the INPUT_PATH and OUTPUT_PATH variables.")
if not (os.path.exists(INPUT_PATH) and os.path.isdir(INPUT_PATH)):
    raise ValueError(f"The INPUT_PATH folder does not exist ({INPUT_PATH}).")
if not (os.path.exists(OUTPUT_PATH) and os.path.isdir(OUTPUT_PATH)):
    raise ValueError(f"The OUTPUT_PATH folder does not exist ({OUTPUT_PATH}).")    
if NOTATION < 0 or NOTATION > 5:
    raise ValueError("The NOTATION variable must be between 0 and 5.")


# Shell object from the Windows Shell API to get metadata.
SH = win32com.client.gencache.EnsureDispatch('Shell.Application',0)

def empty_folder(folder:str):
    """
    Delete all files and folders in a folder.
    Args:
        folder (str): The path to the folder to empty.
    """
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))

def get_metadata_columns(namespace:win32com.client.CDispatch) -> list[str]:
    """
    Get the metadata columns/keys of a shell namespace (folder).
    Args:
        namespace (CDispatch): The shell namespace to get the metadata columns from.
    Returns:
        list: The list of metadata columns/keys.
    """
    colnum = 0
    columns = []
    while True:
        colname = namespace.GetDetailsOf(None, colnum)
        if not colname:
            break
        columns.append(colname)
        colnum += 1
    return columns

def get_metadata(folder:str, file:str) -> dict[str, str]:
    """
    Get the metadata of a file in a folder.
    Args:
        folder (str): The path to the folder containing the file.
        file (str): The name of the file to get the metadata from.
    Returns:
        dict: The metadata of the file.
    """
    ns = SH.NameSpace(folder)
    columns = get_metadata_columns(ns)
    item = ns.ParseName(file)
    metadata = {}
    for colnum in range(len(columns)):
        colval = ns.GetDetailsOf(item, colnum)
        if colval:
            metadata[columns[colnum]] = colval
    return metadata


print(f"Copying images with a notation equal to or higher than {NOTATION} from {INPUT_PATH} to {OUTPUT_PATH}.")
file_count_out = sum(len(files) for _, _, files in os.walk(OUTPUT_PATH))
# Add a warning and prompt to avoid accidents.
print(f"Warning: {file_count_out} files in the output folder will be cleaned. Are you sure you want to continue? (y/n)")
if input() != "y":
    exit(0)
# Try to remove the tree.
print(f"Cleaning output folder...")
empty_folder(OUTPUT_PATH)


# Get the initial number of files
file_count = sum(len(files) for _, _, files in os.walk(INPUT_PATH))
with tqdm.tqdm(total=file_count) as pbar:
    for root, dirs, files in tqdm.tqdm(os.walk(INPUT_PATH)):
        path = root.split(os.sep)
        for file in files:
            # Increment the progress bar
            pbar.update(1)
            pbar.set_description(f"Checking {os.path.join(root, file)}")
            metadata = get_metadata(root, file)
            if txt_to_notation[metadata['Notation']] >= NOTATION:
                out_folder = root.replace(INPUT_PATH, OUTPUT_PATH)
                Path(out_folder).mkdir(parents=True, exist_ok=True)
                shutil.copy2(os.path.join(root, file), out_folder)
    
