"""
This script converts an Excel Workbook into one or more CSVs where each CSV contains the contents of one worksheet of
the Workbook. These CSVs are placed in a directory with the same name as the Workbook. Each CSV is named
after the corresponding worksheet.
"""

from csv import writer
from os import listdir, makedirs
from os.path import isdir, basename, splitext, dirname, sep, exists
from typing import Tuple, Union

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet


def _get_directory_or_workbook_name(names: tuple) -> Union[str, None]:
    """
    If there is only one name in names, then the name of the Workbook or directory is returned.

    Else the user is asked which Workbook or directory they would like to choose from and their selection is
    returned.

    :param names: of either the Workbooks or the directories that exist in the same directory as this script.
    :return: either the name of the exclusive Workbook or directory that exists in the same directory as this
             script or the name of the Workbook or directory that was chosen by the user. If names is empty or
             None, then None is returned.
    """
    if names:
        are_directory_names = isdir(names[0])
        type_singular = 'directory' if are_directory_names else 'Excel ' + Workbook.__name__
        if len(names) == 1:
            print('Found', type_singular, names[0])
            return names[0]
        type_plural = 'directories' if are_directory_names else 'Excel ' + Worksheet.__name__ + 's'
        choice = 0
        print('More than one', type_singular, 'was found.')
        print('Please choose one of the', type_plural, 'from the list or enter -1 to quit:')
        for i, name in enumerate(names, 1):
            print(str(i) + '.', name)
        while choice != -1:
            try:
                choice = int(input('Choice: '))
                return names[choice - 1] if choice != -1 else None
            except ValueError or IndexError:
                print('Please choose one of the available options.')
    else:
        print('No directory or', Workbook.__name__, 'to acquire name.')


def get_directory_or_workbook_name() -> Union[str, None]:
    """
    Locates all of the Workbooks and directories in the same directory as this script.

        If one or more directories are found and one or more Workbooks are found, then the user is asked whether
        they want to convert a single Workbook to CSVs or all of the Workbooks in a directory to CSVs.

        Else if one or more Workbooks are found, then the user is asked which Workbook they would like to
        convert to CSVs.

        Else if only one or more directories were found, then the user is asked which directory they would like to
        convert all of its containing Workbooks into CSVs.

    :return: Either the name of the Workbook or directory of choice, or None if neither an Workbook or
             directory was found in the same directory as this script.
    """
    workbook_names = tuple(name for name in listdir('.') if name.endswith('.xlsx'))
    directory_names = tuple(name for name in listdir('.') if isdir(name))
    if workbook_names and directory_names:
        choice = '-1'
        while choice != '1' or choice != '2':
            print('Would you like to...')
            print('1. Export an Excel', Workbook.__name__, 'to CSVs?')
            print('2. Export all of the Excel', Workbook.__name__ + 's in a directory to CSVs?')
            choice = input('Choice: ')
            if choice == '1':
                return _get_directory_or_workbook_name(workbook_names)
            elif choice == '2':
                return _get_directory_or_workbook_name(directory_names)
            else:
                print('Please select from one of the options')
    elif workbook_names:
        return _get_directory_or_workbook_name(workbook_names)
    elif directory_names:
        return _get_directory_or_workbook_name(directory_names)
    print('Could not find any Excel', Workbook.__name__ + 's or directories in the current directory.')
    print('Make sure that the Excel', Workbook.__name__, 'or directory is in the same directory as this script.')


def _load_workbook(name: str) -> Union[Workbook, None]:
    """
    A wrapper function for openpyxl's load_workbook method. It prints error messages based on the exception encountered
    from attempting to load the Workbook with the given name.

    :param name: of the Workbook that is to be loaded.
    :return: the loaded Workbook or None if an exception occurred.
    """
    try:
        return load_workbook(name)
    except FileNotFoundError:
        print('Could not find Excel ' + Workbook.__name__, '"' + name + '".')
    except PermissionError:
        print('Could not load ' + Workbook.__name__, '"' + name + '".')
        print('Make sure that no other program has the', Workbook.__name__, 'open.')


def get_workbooks_and_names(name: str) -> \
        Union[Tuple[Workbook, str], Tuple[Tuple[Workbook, ...], Tuple[str, ...]], None]:
    """
    If name is the name of an Workbook, then it is loaded and returned along with the Workbook's corresponding
    name.

    Else if name is the name of a directory, then all of the Workbooks in the directory are loaded and a tuple
    of tuples is returned, where the zeroth element of the tuple is another tuple of all of the loaded Workbooks, while
    the first element of the tuple is a parallel tuple that contains all of the names of those Workbooks.

    :param name: of the Workbook or directory.
    :return: A tuple that contains the Workbook and it name or a tuple of parallel tuples that contains the loaded
             workbooks along with their names. If name is empty or None, then None is returned.
    """
    if name:
        if name.endswith('.xlsx'):
            workbook = _load_workbook(name)
            print('Successfully loaded', Workbook.__name__, '"' + name + '".')
            return workbook, name
        elif isdir(name):
            print('Loading all', Workbook.__name__ + 's in directory "' + name + '".')
            workbook_names = tuple(name for name in listdir(name) if name.endswith('.xlsx'))
            workbooks = tuple(_load_workbook(name + sep + workbook_name) for workbook_name in workbook_names)
            if workbooks:
                print('Successfully loaded', Workbook.__name__ + 's in directory "' + name + '".')
            else:
                print('No', Workbook.__name__ + 's found in directory "' + name + '".')
            return workbooks, workbook_names
    print('Cannot load a', Workbook.__name__, 'or the', Workbook.__name__ + 's in a directory with no name.')


def _get_output_file_relative_path(workbook_name: str, worksheet: Worksheet, root_dir_name: str = None) -> str:
    """
    Returns the relative path to the CSV, named after the corresponding Worksheet, which will be located in
    a directory named after the Workbook.

    If a root_dir_name is given, then it will be prepended to the relative output path.

    :param workbook_name: that will be used to construct the directory name.
    :param worksheet: that will be used to construct the CSV and name.
    :param root_dir_name: that will be prepended to the relative output path if specified.
    :return: the relative path to the CSV.
    """
    workbook_name = splitext(basename(workbook_name))[0]
    relative_path = (root_dir_name + sep if root_dir_name else '') + workbook_name + sep + worksheet.title + '.csv'
    if exists(relative_path):
        print('Found', '"' + relative_path + '".')
        choice = input('Would you like to overwrite? (y/n): ').lower()
        while choice != 'y' and choice != 'n':
            print('Please enter either "y" or "n"')
            choice = input('Would you like to overwrite? (y/n): ').lower()
        return relative_path if choice == 'y' else None
    elif not exists(dirname(relative_path)):
        makedirs(dirname(relative_path))
    return relative_path


def convert_excel_to_csv(workbook: Workbook, workbook_name: str, root_dir_name: str = None) -> None:
    """
    A directory named after the Workbook is created if it does not exist already. Each Worksheet in the Workbook is
    converted into a CSV and placed inside that directory. Each CSV is named after the corresponding Worksheet.

    If a root_dir_name is given, then it will be prepended to the relative output path.

    :param workbook: that is to be converted to CSVs.
    :param workbook_name: of the Workbook that is to be converted.
    :param root_dir_name: that will be prepended to the relative output path if specified.
    """
    for worksheet in workbook:
        relative_path = _get_output_file_relative_path(workbook_name, worksheet, root_dir_name)
        if relative_path:
            with open(relative_path, 'w', encoding='utf-8', newline='') as output_file:
                csv_writer = writer(output_file)
                for row in worksheet:
                    csv_writer.writerow([cell.value for cell in row if cell.value is not None])
                if not root_dir_name:
                    print('Successfully saved converted data to', '"' + relative_path + '".')
        else:
            print('No data was written for', Worksheet.__name__, '"' + worksheet.title + '".')
    if root_dir_name:
        print('Successfully saved converted data to', '"' + dirname(relative_path) + '".')


def main():
    try:
        directory_or_file_name = get_directory_or_workbook_name()
        if directory_or_file_name:
            workbooks_and_names = get_workbooks_and_names(directory_or_file_name)
            if workbooks_and_names:
                workbooks, workbook_names = workbooks_and_names
                if isinstance(workbooks, Workbook):
                    convert_excel_to_csv(*workbooks_and_names)
                else:
                    for workbook, workbook_name in zip(workbooks, workbook_names):
                        convert_excel_to_csv(workbook, workbook_name, directory_or_file_name + '_Converted')
    except Exception as e:
        print(e)
    input('Press enter to exit...')


if __name__ == '__main__':
    main()
