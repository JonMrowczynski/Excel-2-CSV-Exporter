"""
Copyright (c) 2018-2022 Jon Mrowczynski

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation the
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit
persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the
Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This script converts an Excel Workbook into one or more CSVs where each CSV file contains the contents of one Worksheet
of the Workbook. These CSVs are placed in a directory that is named after the Workbook. Each CSV is named after the
corresponding Worksheet.

This process can instead be done on a directory that contains multiple Excel Workbooks. For this case, everything is the
same except that the above process will be carried out on all of the Workbooks and the resulting directories will be
placed in a directory called root_dir_name.
"""

from argparse import ArgumentParser
from csv import writer
from os import listdir
from os.path import basename, isdir, join, splitext
from pathlib import Path
from typing import Final

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet

WORKBOOK_EXTENSION: Final = '.xlsx'
EXPORT_ROOT_DIR: Final = 'Exports'


def _load_workbook(workbook_path: str) -> dict[str: Workbook]:
    """
    Loads the Workbook at the given path into memory. If there does not exist a Workbook at the given path, then a
    FileNoteFoundError is raised. If the Workbook is open in another program, then a PermissionError is raised since
    only one program can exclusively access an Excel Workbook.

    :param workbook_path: a relative or absolute path to a Workbook.
    :return: a singleton dictionary containing the name of the Workbook as the key and the Workbook as the value.
    """
    try:
        return {splitext(basename(workbook_path))[0]: load_workbook(workbook_path, read_only=True)}
    except FileNotFoundError or PermissionError as e:
        problem_word = 'find' if isinstance(e, FileNotFoundError) else 'load'
        print(f'Could not {problem_word} {Workbook.__name__} "{workbook_path}".')


def _loads_workbooks(dir_path: str) -> dict[str: Workbook]:
    """
    Loads all the Workbooks in the given path and returns a dictionary that maps the name of the Workbook to the
    corresponding loaded Workbook.

    :param dir_path: a relative or absolute path to a directory that contains Workbooks.
    :return: a dictionary containing names of the Workbooks as the keys and the corresponding Workbook as the values.
    """
    dicts = [_load_workbook(join(dir_path, name)) for name in listdir(dir_path) if name.endswith(WORKBOOK_EXTENSION)]
    workbooks_dict = dict()
    for singleton in dicts:
        workbooks_dict |= singleton
    return workbooks_dict


def _get_input_source(path: str) -> dict:
    """
    Returns either a singleton dictionary if a path 2 a Workbook is given, a dictionary with multiple entries if a
    directory of Workbooks is given, or an empty dictionary if the given path is not valid or there are no Workbooks in
    the given dictionary.

    :param path: the path to the Workbook or directory of Workbooks that are to be converted.
    :return: a dictionary containing mappings from name(s) of the Workbook(s) to the corresponding loaded Workbook(s).
    """
    if path.endswith(WORKBOOK_EXTENSION):
        print(f'Found input source "{path}".')
        return _load_workbook(path)
    elif isdir(path):
        print(f'Found input source "{path}".')
        workbooks_map = _loads_workbooks(path)
        if not workbooks_map:
            print(f'No {Workbook.__name__}s found in "{path}".')
        return workbooks_map
    else:
        print(f'"{path}" is not a valid {Workbook.__name__} or directory.')
        return dict()


def _should_write_data(export_path: Path) -> bool:
    """
    Determines whether data should be written to the given path. By default, data will be written unless there already
    exists a file at the given export path.

    :param export_path: the path that a CSV file will possible exported to.
    :return: a boolean indicating whether data should be written to the given export path.
    """
    if export_path.exists():
        print(f'Found "{export_path}".')
        choice = input('Would you like to overwrite? (y/[n]): ').lower()
        while choice != 'y' and choice != 'n':
            print('Please enter either "y" or "n"')
            choice = input('Would you like to overwrite? (y/[n]): ').lower()
        return True if choice == 'y' else False
    return True


def _workbooks2csv(workbooks: dict) -> None:
    """
    Converts all the Worksheets in all the Workbooks into CSVs. One directory will be named after each Workbook and will
    contain all the data in the Workbook. Each CSV file will be named after each Worksheet and will contain the data
    present in the Worksheet.

    :param workbooks: a map whose keys are the names of the Workbooks and the values are the corresponding Workbooks.
    """
    print(f'Converting {Workbook.__name__}s to CSVs...')
    for workbook_name, wb in workbooks.items():
        path2workbook_export = Path(EXPORT_ROOT_DIR, workbook_name)
        path2workbook_export.mkdir(parents=True, exist_ok=True)
        for ws in wb:
            path2export = Path(join(str(path2workbook_export), ws.title + '.csv'))
            should_write_data = _should_write_data(path2export)
            if not should_write_data:
                print(f'No data was written for {Worksheet.__name__} "{ws.title}".')
                continue
            with open(path2export, 'w', encoding='utf-8', newline='') as output_file:
                writer(output_file).writerows(cell_vals for row in ws if any(cell_vals := [cell.value for cell in row]))
                print(f'Successfully saved converted data to "{path2export}".')
    print(f'Converted {Workbook.__name__}s to CSVs!')


def main():
    parser = ArgumentParser()
    parser.add_argument('-path', type=str, required=True, help=f'The name of the {Workbook.__name__} or a path to '
                                                               f'the directory containing {Workbook.__name__}s.')
    input_source = _get_input_source(parser.parse_args().path)
    if input_source:
        _workbooks2csv(input_source)


if __name__ == '__main__':
    main()
