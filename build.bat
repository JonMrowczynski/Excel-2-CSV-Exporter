@echo off
title excel_2_csv_converter build script
del /Q build, dist
set year=%date:~10,4%
set month=%date:~4,2%
set day=%date:~7,2%
set script_name="excel_2_csv_converter"
python -O -m PyInstaller --clean --onefile %script_name%.py --name %script_name%_%year%_%month%_%day%