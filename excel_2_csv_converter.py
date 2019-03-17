"""
This script allows one to convert an Excel Workbook into .csv files, where each .csv file corresponds to data in one
worksheet of the Excel Workbook. These .csv files are placed in a directory that has the same name as the name of the
Excel Workbook and the name of the .csv files are named after their corresponding worksheet.
"""

from os import listdir, makedirs
from os.path import isdir, basename, splitext, dirname, sep, exists

from typing import Tuple, Union

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet


def _get_directory_or_file_name(names: tuple) -> Union[str, None]:
    """
    If there is only one name in names, then the name of the Excel Workbook or directory is returned.

    Else the user is asked which Excel Workbook or directory they would like to choose from.

    :param names: of either the Excel Workbooks or the directories that exist in the current directory.
    :return: either the name of the only Excel Workbook or directory that exists in the current directory or the name of
             the Excel Workbook or directory that was chosen by the user or
    """
    if names:
        are_directory_names = isdir(names[0])
        type_singular = 'directory' if are_directory_names else 'Excel ' + Workbook.__name__
        if len(names) == 1:
            print('Found', type_singular, names[0])
            return names[0]
        type_plural = 'directories' if are_directory_names else 'Excel ' + Worksheet.__name__ + 's'
        while True:
            print('More than one', type_singular, 'was found in the current directory')
            print('Please choose one of the', type_plural, 'from the list:')
            for i, name in enumerate(names, 1):
                print(str(i) + '.', name)
            try:
                return names[int(input('Choice: ')) - 1]
            except ValueError or IndexError:
                print('Please choose one of the available options')


def get_directory_or_file_name() -> Union[str, None]:
    """
     All of the Excel Workbooks and directories in the same directory as this script are located.

        If one or more directories have been found and one or more Excel Workbooks have been found, then the user is
        asked whether they want to convert a single Excel Workbook to CSVs or all of the Excel Workbooks in a directory
        into CSVs.

        Else, if only one or more Excel Workbooks were found, then the user is asked which Excel Workbook they would
        like to convert into a CSV.

        Else, if only one or more directories were found, then the user is asked which directory they would like to
        convert all of its containing Excel Workbooks into CSVs.

        Otherwise, an error message is printed.

    :return: Either the name of the Excel Workbook or directory of choice, or None if neither an Excel Workbook or
             directory was found in the current directory.
    """
    excel_workbook_names = tuple(name for name in listdir('.') if name.endswith('.xlsx'))
    directory_names = tuple(name for name in listdir('.') if isdir(name))
    if excel_workbook_names and directory_names:
        choice = '-1'
        while choice != '1' or choice != '2':
            print('Would you like to...')
            print('1. Export an Excel', Workbook.__name__, 'to CSVs?')
            print('2. Export all of the Excel', Workbook.__name__ + 's in a directory to CSVs?')
            choice = input('Choice: ')
            if choice == '1':
                return _get_directory_or_file_name(excel_workbook_names)
            elif choice == '2':
                return _get_directory_or_file_name(directory_names)
            else:
                print('That was not one of the choices')
    elif excel_workbook_names:
        return _get_directory_or_file_name(excel_workbook_names)
    elif directory_names:
        return _get_directory_or_file_name(directory_names)
    print('Could not find any Excel', Workbook.__name__ + 's or directories in the current directory')
    print('Make sure that the Excel', Workbook.__name__, 'or directory is in the same directory as this script')


def _load_workbook(name: str) -> Workbook:
    """
    This is basically just a wrapper function for openpyxl's load_workbook method. It prints corresponding error
    messages based on the exception encountered from attempting to load the Excel Workbook with the given name.

    :param name: of the Excel Workbook that is to be loaded.
    :return: the loaded Excel Workbook.
    """
    try:
        return load_workbook(name)
    except FileNotFoundError:
        print('Could not find Excel ' + Workbook.__name__, '"' + name + '"')
    except PermissionError:
        print('Could not load ' + Workbook.__name__, '"' + name + '"')
        print('Make sure that no other program has the', Workbook.__name__, 'open')


def load_workbooks(name: str) -> Union[Tuple[Workbook, str], Tuple[Tuple[Workbook, ...], Tuple[str, ...]]]:
    """
    If name is the name of an Excel Workbook, then it is loaded and the loaded Workbook along with the Workbook's
    corresponding name is returned.

    Else if name is the name of a directory, then all of the Excel Workbooks in the directory are loaded and a tuple
    of tuples is returned, where the zeroth element of the tuple is another tuple of all of the loaded Workbooks, while
    the first element of the tuple is a parallel tuple that contains all of the names of the Workbooks.

    :param name: of the Excel Workbook or directory.
    :return: A tuple of parallel tuples that contains the loaded workbooks along with their names.
    """
    if name:
        if name.endswith('.xlsx'):
            workbook = _load_workbook(name)
            print('Successfully loaded', Workbook.__name__, '"' + name + '"')
            return workbook, name
        elif isdir(name):
            print('Loading all', Workbook.__name__ + 's in directory "' + name + '"')
            workbook_names = tuple(name for name in listdir(name) if name.endswith('.xlsx'))
            workbooks = tuple(_load_workbook(name + sep + workbook_name) for workbook_name in workbook_names)
            if workbooks:
                print('Successfully loaded', Workbook.__name__ + 's in directory "' + name + '"')
            else:
                print('No', Workbook.__name__ + 's found in directory "' + name + '"')
            return workbooks, workbook_names
    print('Cannot load a', Workbook.__name__, 'or the', Workbook.__name__ + 's in a directory with no name.')


def _get_output_file_relative_path(workbook_name: str, worksheet: Worksheet, root_directory_name: str = None) -> str:
    """
    Returns the relative path to the CSV output file which will be located in a directory named after the Workbook and
    the name of the file will be named after the name given to the Worksheet. In addition, the CSV file will also
    contain all of the information contained in the Worksheet.

    If a root_directory_name is given, then the converted data is saved in a directory named after root_directory_name.

    :param workbook_name: that will be used to construct the directory name.
    :param worksheet: that will be used to construct the CSV file data and name.
    :param root_directory_name: that will contain the converted data if specified.
    :return: the relative path to the CSV file.
    """
    workbook_name = splitext(basename(workbook_name))[0]
    output_file_relative_path = (root_directory_name + sep if root_directory_name else '') \
                                + workbook_name + sep + worksheet.title + '.csv'
    if exists(output_file_relative_path):
        print('Found', '"' + output_file_relative_path + '"')
        choice = input('Would you like to overwrite? (y/n): ').lower()
        while choice != 'y' and choice != 'n':
            print('Please enter either "y" or "n"')
            choice = input('Would you like to overwrite? (y/n): ').lower()
        return output_file_relative_path if choice == 'y' else None
    elif not exists(dirname(output_file_relative_path)):
        makedirs(dirname(output_file_relative_path))
    return output_file_relative_path


def convert_excel_to_csv(workbook: Workbook, workbook_name: str, root_directory_name: str = None) -> None:
    """
    Converts a given Excel Workbook to CSV files by outputting in a directory whose name is the name of the Excel
    Workbook a CSV file that corresponds to each Worksheet in the Excel Workbook whose file name corresponds to the
    name of the Worksheet.

    If a root_directory_name is given, then the converted data is saved in a directory named after root_directory_name.

    :param workbook: that is to be converted to .csv files.
    :param workbook_name: of the Excel Workbook that is to be converted to CSV files.
    :param root_directory_name: that will contain the converted data if specified.
    """
    for worksheet in workbook:
        output_file_relative_path = _get_output_file_relative_path(workbook_name, worksheet, root_directory_name)
        if output_file_relative_path:
            with open(output_file_relative_path, 'w', encoding='utf-8', newline='') as output_file:
                from csv import writer
                csv_writer = writer(output_file)
                for row in worksheet:
                    line = [cell.value for cell in row if cell.value is not None]
                    csv_writer.writerow(line)
                if not root_directory_name:
                    print('Successfully saved converted data to', output_file_relative_path)
        else:
            print('No data was written for', Worksheet.__name__, '"' + worksheet.title + '"')
    if root_directory_name:
        print('Successfully saved converted data to', '"' + dirname(output_file_relative_path) + '"')


def main():
    directory_or_file_name = get_directory_or_file_name()
    if directory_or_file_name:
        workbooks_and_names = load_workbooks(directory_or_file_name)
        if workbooks_and_names:
            workbooks, workbook_names = workbooks_and_names
            if isinstance(workbooks, Workbook):
                convert_excel_to_csv(*workbooks_and_names)
            else:
                for workbook, workbook_name in zip(workbooks, workbook_names):
                    convert_excel_to_csv(workbook, workbook_name, directory_or_file_name + '_Converted')
    input('Press enter to exit...')


if __name__ == '__main__':
    main()
