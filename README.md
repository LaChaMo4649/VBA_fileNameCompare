# Excel Macro: File name comparison tool in a directory

## Overview

This Excel macro is a tool to extract a list of file names existing in two specified directories and compare them on Excel.As a result of the comparison, unique file names that exist only in one of the directories are identified and the file names are automatically colored.

## Main Features

- Writes the file names in the two specified directories as a list on an Excel sheet.
- Compares the two lists and automatically determines unique file names that exist in only one of them.
- Unique file names are automatically colored (highlighted) for easy visual identification.

## Usage

1. **Open an Excel file (.xlsm) with the macro enabled.**
2. specify the directory where the file to be compared is located in cells B3 and C3.
3. activate the Get File Names button and write out the list of file names in columns B and C.
The file name extensions of the exported file names will be all in lower case. 5.
With the File Name Compare button, color the cells with unique file names in yellow. 6.
The comparison result (unique filename) will be colored.

## Assumed sheet configuration.

- **Column BÅF** File name of directory 1
- **C column: ** File names in directory 2
- Unique file names will be colored in each column.

## Notes

- Files in subdirectories are not included (can be extended if necessary).
- File name comparisons are not case sensitive (can be changed with VBA code).
- Please enable Excel's macro functionality.

## Customize

- Coloring colors and file types (extensions) to be compared can be changed by modifying VBA code.

## Disclaimer

We are not responsible for any problems or damages caused by the use of this tool.
Please use at your own risk.
