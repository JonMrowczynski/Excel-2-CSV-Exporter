# Excel 2 CSV Exporter

## Backstory
During my data acquisition adventures, I encountered multiple Excel Workbooks with multiple Worksheets that I needed to convert into CSVs where each CSV contained data from one Worksheet. Excel's built-in CSV exporter was insufficient and impractical to accomplish this task. 

## What Does it Do?
The Excel-2-CSV-Exporter allows one to export all of the Worksheets in an Excel Workbook to CSVs in a directory that is named after the Excel Workbook. This process can also be performed over all Excel Workbooks that are contained within a directory, where the output directory is named after the input directory. The output directory will contain all of the directories where each directory corresponds to the exported data of a single Workbook.

## How 2 use Script?
Run the script in the same directory as the Excel Workbook or directory that contains the Excel Workbooks that you would like to convert to CSVs and follow the prompt. The output directory will be written to the same directory that contains the script.

Currently there is only a Windows executable, but executables for other platforms can be built using a very similar method that was used to build the Windows executable.

## How 2 Develop?
You will need:

- [PyCharm IDE](https://www.jetbrains.com/pycharm/download/) >= 2019.3.1 (recommended, but not necessary).

- [Python](https://www.python.org/downloads/) >= 3.8.1.

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) >= 3.0.3 to read Excel Workbooks.

- [PyInstaller](https://pyinstaller.readthedocs.io/en/stable/) >= 3.6 to build executable.
