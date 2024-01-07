@echo off
rmdir /s /q build dist\Windows
del *.spec
set year=%date:~10,4%
set month=%date:~4,2%
set day=%date:~7,2%
:: script_name should match the name of the Python file that should be built. This is most likely the only needed change.
set script_name="excel_2_csv_converter"
python -O -m PyInstaller --clean --distpath dist/Windows --onefile %script_name%.py --name %script_name%_%year%-%month%-%day%