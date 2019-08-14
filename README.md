# Excel-2-CSV-Exporter

## Backstory
During my data acquisition adventures, I encountered multiple Excel Workbooks with multiple Worksheets that I needed to convert into CSVs where each CSV contained data from one Worksheet. Excel's built-in CSV exporter was insufficient and impractical to accomplish this task. 

## What Does it Do?
The Excel-2-CSV-Exporter allows one to export all of the Worksheets in an Excel Workbook to CSVs in a directory that is named after the Excel Workbook. This process can also be performed over all Excel Workbooks that are contained within a directory.

## How 2 use Script?
Run the script in the same directory as the Excel Workbook or directory that contains the Excel Workbooks and follow the prompt. The exported CSV(s) will be written to the same directory that the script is in.

Currently there is only a Windows executable, but executables for other platforms can be built using a very similar method that was used to build the Windows executable.

## How 2 Develop?
You will need:
- PyCharm 2019.2 (recommended although not necessary). 
- Python 3.7.4 or later.
- openpyxl to read Excel Workbooks.
- PyInstaller to build executable.
