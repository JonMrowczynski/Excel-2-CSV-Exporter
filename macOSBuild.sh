#!/usr/bin/env bash
rm -rf ./build ./dist/macOS/* ./*.spec
# script_name should match the name of the Python file that should be built. This is most likely the only needed change.
script_name="excel_2_csv_converter"
python -O -m PyInstaller --clean --distpath ./dist/macOS --onefile "${script_name}.py" --name "${script_name}_$(date +"%Y-%m-%d")"