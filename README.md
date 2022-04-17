# Excel 2 CSV Exporter

## Backstory
During my data acquisition adventures, I encountered multiple Excel Workbooks with multiple Worksheets that I needed to 
convert into CSVs where each CSV contained data from one Worksheet. Excel's built-in CSV exporter was insufficient and 
impractical to accomplish this task. 

## What Does it Do?
The Excel-2-CSV-Exporter allows one to export all the Worksheets in an Excel Workbook to CSVs in a directory that is 
named after the Excel Workbook. This process can also be performed on all Excel Workbooks that are contained within a 
directory. Either way, all the output directories (each corresponding to a Workbook) will be contained in a directory
called "Exports", which is placed in the same directory as the executed Python script or executable.

## How 2 use Script?
You can either...

1. Run the Python script: `python -m excel_2_csv_converter -path "path/2/workbook/or/directory"`


2. Run the executable: `.\excel_2_csv_converter.exe -path "path/2/workbook/or/directory"`

Where `"path/2/workbook/or/directory"` can be an absolute path or a path relative to the Python script or executable. 
Either way, it must point to an Excel Workbook (a file with an ".xlsx" extension) or a directory containing one or more 
Excel Workbooks.

Building a Windows executable can be done by running the build.bat script or the "build executable" configuration. 
Executables for other platforms can be built using a very similar method that was used to build the Windows executable.
However, this is not done here.

## How 2 Develop?
You will need:

- [PyCharm IDE](https://www.jetbrains.com/pycharm/download/) >= 2022.1 (recommended, but not necessary).

- [Python](https://www.python.org/downloads/) >= 3.10.

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) >= 3.0.9 to read Excel Workbooks.

- [PyInstaller](https://pyinstaller.readthedocs.io/en/stable/) >= 4.10 to build executable.
  - [UPX](https://upx.github.io/) >= 3.96 if you would like to make smaller executables.