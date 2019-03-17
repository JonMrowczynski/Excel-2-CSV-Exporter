@echo off
title excel_2_csv_converter build script
python -O -m PyInstaller --onefile excel_2_csv_converter.py
exit