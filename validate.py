
import argparse
import os
import sys
from collections import OrderedDict
from itertools import zip_longest

try:
    import xlrd
except ImportError as e:
    print(e)
    print('Maybe "pip install xlrd", would help.')
    sys.exit(1)

def _cli():
    parser = argparse.ArgumentParser(
            description=__doc__,
            formatter_class=argparse.ArgumentDefaultsHelpFormatter,
            argument_default=argparse.SUPPRESS)
    parser.add_argument('excel_file', help="Excel sheet containing the records", nargs='?')
    parser.add_argument('folder', help="Path to folder containing the files", nargs='?')
    parser.add_argument('-wdir', '--working_directory', default=os.getcwd(), help="Set working directory")

    if len(sys.argv) < 3:
        parser.print_usage()
        sys.exit(1)

    args = parser.parse_args()

    return vars(args)

def first_difference(str1, str2):
    """Find first difference between two strings and return index, and character which is different"""
    for idx, (sh, fi) in enumerate(zip_longest(str1, str2)):
        if sh != fi:
            return idx
    return -1


def xls_search(sheet):
    """Searches for files in sheet and return dictionary with records
    :type sheet: xlrd sheet
    :return: dict {(row, col), value)
    """
    columns_to_search = ['Working File', 'Released File', 'Filename']

    # take row 1 containing headers
    row = sheet.row(1)
    # row where the data start
    start_row = 2

    filename_in_sheet = {}
    for idc, cell in enumerate(row):
        if cell.value in columns_to_search:
            column = sheet.col_slice(idc, start_rowx=start_row)
            # non_empty_fields = (_cell.value, idr, idc) for idc, _cell in enumerate(column) if _cell.value]
            non_empty_fields = {(idr, idc): _cell.value for idr, _cell in enumerate(column, start=start_row+1) if _cell.value}
            filename_in_sheet.update(non_empty_fields)

    return filename_in_sheet

def


def validate_sheet_duplicates(file_names):
    """Return duplicated files from sheet"""
    file_names = list(file_names)
    duplicates = [item for idx, item in enumerate(file_names) if item in file_names[:idx]]
    return duplicates


def report_differences(out_file, uniques):
    """Writes report to file"""
    with open(out_file, 'w+') as outf:
        # write header
        header = 'filename_in_sheet|filename_in_folder|corrected_value'
        print(header, file=outf)
        # write rest
        for (pos, sheet), folder in uniques:

            idx = first_difference(sheet, folder)
            line = "{sheet}|{folder}|".format(sheet=sheet,
                                              folder=folder)
            if idx >= 0:
                sign = "{spc}^{pos}".format(spc=(idx - 1) * " ",
                                            pos=pos)
            else:
                sign = "---^ names does not match ^---"
            print(line, file=outf)
            print(sign, file=outf)
        print()
        print("Report stored to:", out_file)


def report_duplicates(out_file, duplicates):

    with open(out_file, 'w+') as outf:
        print("This file contain records which are duplicated in sheet")
        for duplicate in duplicates:
            print(duplicate, file=outf)

def report_sheet_dict(uniq_in_sheet, files_in_sheet):
    """Takes set of unique names in sheet and dict files_in_sheet
    Return new dictionary containing only unique values with coordinates
    :return: dict {pos: filename}"""
    report_dict = {}
    for un in uniq_in_sheet:
        tmp_dict = {pos: file for pos, file in files_in_sheet.items() if un == file}
        report_dict.update(tmp_dict)
    return report_dict

def main(*args, **kwargs):


    wb_filename = "GAL-DN-ESA-ZZZ-X_0197_0.1_Batch 44_Sapienza.xls"
    folder_with_files = "GAL-DN-ESA-ZZZ-X_0197_0.1_Batch 44"

    # EXCEL_FILE = os.path.join("C:\\Bordel\\ipython_tmp\\", "GAL-DN-ESA-ZZZ-X_0197_0.1_Batch 44_Sapienza.xls")

    # load folder and file
    WORKING_FOLDER = os.path.abspath(kwargs['working_directory'])


    WB_FILE = os.path.abspath(os.path.join(WORKING_FOLDER, kwargs['excel_file']))  # excel file

    FOLDER_WITH_FILES = os.path.abspath(os.path.join(WORKING_FOLDER, kwargs['folder']))

    # output file handling
    report_file_name = os.path.basename(WB_FILE)
    report_file_name = "report_{xls_file}.txt".format(xls_file=report_file_name[:report_file_name.rfind('.')])
    REPORT_FILE = os.path.join(WORKING_FOLDER, report_file_name)

    duplicates_file_name = "duplicates_{xls_file}.txt".format(xls_file=report_file_name[:report_file_name.rfind('.')])
    DUPLICATES_FILE = os.path.join(WORKING_FOLDER, duplicates_file_name)

    # working with excel file
    work_book = xlrd.open_workbook(WB_FILE)
    documents_sheet = work_book.sheet_by_name('Documents')

    # getting data
    try:
        files_in_sheet = xls_search(documents_sheet)
        files_in_folder = set(os.listdir(FOLDER_WITH_FILES))
    except FileNotFoundError as ex:
        files_in_sheet = []
        files_in_folder = []
        print(ex)
        print("Have you forget file argument?")
        sys.exit(1)

    # list duplicated reports in sheet
    dups_in_sheet = validate_sheet_duplicates(files_in_sheet.values())

    # processing records
    set_files_in_sheet = set(files_in_sheet.values())
    # unique elements in folder not in sheet
    uniq_in_folder = files_in_folder - set_files_in_sheet

    # unique elements in sheet not in folder
    uniq_in_sheet = set_files_in_sheet - files_in_folder

    # compare unique records and try to find similarities
    sorted_uniq_folder = sorted(uniq_in_folder)
    report_dict = report_sheet_dict(uniq_in_sheet, files_in_sheet)
    sorted_dict = OrderedDict(sorted(report_dict.items(), key=lambda rec: rec[1]))

    to_report = filter(lambda pair: pair[0] != pair[1], zip(sorted_dict.items(), sorted_uniq_folder))

    if to_report:
        report_differences(REPORT_FILE, to_report)



    if dups_in_sheet:
        report_duplicates(DUPLICATES_FILE, dups_in_sheet)
        print()
        print("Duplicates in sheet stored in:", DUPLICATES_FILE)

if __name__ == '__main__':
    main(**_cli())