@echo off
title excel_2_csv_converter build script
del /Q build, dist
python -O -m PyInstaller --clean --onefile excel_2_csv_converter.py