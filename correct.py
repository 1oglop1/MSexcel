"""!!!WARNING!!! Check if there are any duplicates in sheet before running this.

   Purpose of this script is to read the file('report_') produced by validate.py.
   Then change the file names of files to represent the value from the spreadsheet.

   How to use this script"""
import os
import sys
import argparse
from ast import literal_eval as make_tuple

try:
    import xlrd
    from xlutils.copy import copy as xlcopy
    import xlwt
except ImportError as e:
    print(e)
    print('Maybe "pip install xlrd xlwt xlutils", would help.')
    sys.exit(1)

def _cli():
    parser = argparse.ArgumentParser(
            description=__doc__,
            formatter_class=argparse.ArgumentDefaultsHelpFormatter,
            argument_default=argparse.SUPPRESS)
    parser.add_argument('excel_file', help="Input file - Excel sheet containing the records", nargs='?')
    parser.add_argument('-wdir', '--working_directory', default=os.getcwd(), help="Set working directory")
    parser.add_argument('-of', '--out_file', default='output.xls', help="Output file, empty will create output.xls")
    parser.add_argument('-o','--overwrite', action='store_true', default=False, help='Overwrite input excel file')

    args = parser.parse_args()
    arguments = vars(args)
    try:
        a = arguments['excel_file']
    except KeyError:
        parser.print_usage()
        sys.exit(1)

    return vars(args)


def rename_file(original_name, new_name, working_directory):
    full_original_name = os.path.abspath(os.path.join(working_directory, original_name))
    full_new_name = os.path.abspath(os.path.join(working_directory, new_name))
    os.rename(full_original_name, full_new_name)


def parse_position(string):
    """Takes empty line wint spaces and ^(13, 45) position and return position
    :return: (13, 45)"""
    string = string.strip()
    if '^(' in string.strip():
        start = string.find('^(') + 1
        stop = string.index(')') + 1
        position = make_tuple(string[start:stop])
        return position

    return False


def values_to_correct(input_file):
    """Read report_ file and return dictionary of values to be corrected
    :return: dictionary {(row, col): corrected_name}"""
    to_be_corrected = {}

    with open(input_file, 'r') as inf:
        # Check header # maybe later
        line = inf.readline()  # skip first line

        for line in inf:
            try:
                file_in_sheet, file_in_folder, corrected_value = line.strip().split('|')
            except ValueError:
                corrected_value = None

            if corrected_value:
                # if the filename has been corrected change the name of file
                line2 = inf.readline()
                pos = parse_position(line2)
                to_be_corrected[pos] = corrected_value
                #                 print("chaging:{0} => {1}".format(file_in_folder, corrected_value))

    return to_be_corrected



def main(*args, **kwargs):

    WORKING_FOLDER = os.path.abspath(kwargs['working_directory'])


    WB_FILE = os.path.abspath(os.path.join(WORKING_FOLDER, kwargs['excel_file']))  # excel file
    report_file_name = os.path.basename(WB_FILE)
    report_file_name = "report_{xls_file}.txt".format(xls_file=report_file_name[:report_file_name.rfind('.')])
    REPORT_FILE = os.path.join(WORKING_FOLDER, report_file_name)

    if kwargs['overwrite']:
        OUTPUT_FILE = WB_FILE
    else:
        OUTPUT_FILE = os.path.abspath(kwargs['out_file'])

    # open work_book to read
    read_book = xlrd.open_workbook(WB_FILE, formatting_info=True)

    # copy read only to write only
    print('Copying excel sheet in memory.')
    write_book = xlcopy(read_book)

    write_sheet = write_book.get_sheet('Documents')

    print("Trying to read report file:")
    print(REPORT_FILE)

    # get correct values from report_ file
    correct_values = values_to_correct(REPORT_FILE)

    print('Writing changes')
    for (row, col), value in correct_values.items():
        print(row - 1, col, value)
        write_sheet.write(row - 1, col, value)

    write_book.save(OUTPUT_FILE)
    print("Saved to:")
    print(OUTPUT_FILE)



if __name__ == '__main__':
    main(**_cli())