# 

This script is used to filter and copy the best-rated images from a source folder to a destination folder based on their "Notation" in metadata.

The script performs the following steps:
1. Deletes all files and folders in the destination folder.
2. Copies the files with a notation equal to or higher than the threshold to the destination folder.
Note: The metadata notation is extracted using the Windows Shell API.
Note2: This probably only works on Windows, and when setup in french (it changes the metadata key names).

Usage:
- In the user_configuration.json file, set the INPUT_PATH, OUTPUT_PATH, and NOTATION variables.
- Set the INPUT_PATH variable to the path of the source folder.
    - ex : "E:\\Images\\Mes Souvenirs"
- Set the OUTPUT_PATH variable to the path of the destination folder (Warning ! It will be emptied).
    - ex : "E:\\Images\\Mes meilleurs souvenirs"
- Set the NOTATION variable to the desired threshold notation (Between 0 and 5).

## Usage

Update `user_configuration.json`.
Install dependancies:
```
conda env create -n winauto -f requirements.txt
```
Run:
```
conda activate winauto
python copy_best_notations.py
```