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

1.
    1. Create a venv.
    2. Install all requirements to run by entering the command `pip install -r ./requirements/run.txt`.
    3. Run the Python script: `python -m excel_2_csv_converter --input_path "path/2/workbook/or/directory"`.

2. Run the pre-built executable: `.\excel_2_csv_converter.exe --input_path "path/2/workbook/or/directory"`

Where `"path/2/workbook/or/directory"` can be an absolute path or a path relative to the Python script or executable.
Either way, it must point to an Excel Workbook (a file with an ".xlsx" extension) or a directory containing one or more
Excel Workbooks.

An optional output path can be specified using the flag `--output_path` if the default is not desired.

Building a Windows executable can be done by running the build.bat script or the "build executable" configuration.
Executables for other platforms can be built using a very similar method that was used to build the Windows executable.
However, this is not done here.

## How 2 Develop?

You will need:

- [PyCharm IDE](https://www.jetbrains.com/pycharm/download/) >= 2023.3.2 (recommended, but not necessary).

- [Python](https://www.python.org/downloads/) >= 3.12.1.

- Run `pip install -r ./requirements/deploy.txt` if you would like to build executables.

- [UPX](https://upx.github.io/) >= 4.2.2 if you would like to make smaller executables on windows.