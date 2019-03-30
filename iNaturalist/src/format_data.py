'''
Created on May 14th, 2018
Author: Miles McCall
Sources:
Description: Parse the "Pollinator Plant Master List" maintained by Signe and
            generate a formatted version for uploading to the website.
'''

# External imports
import os
import sys
import string
import errno
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
    #from openpyxl import Workbook
    #from openpyxl.reader.excel import load_workbook
    #from openpyxl.compat import range
    #from openpyxl.utils import get_column_letter

# Local imports
import col_functions
import file_functions
from file_functions import test_import

#Functions
def read_xlsx(wb_name, ws_name, min_col, min_row, max_col, max_row):
    # Initialize openpyxl vars
    wb = load_workbook(str(wb_name))
    ws = wb[str(ws_name)]
    return_array = []

    # Iterate through xlsx file
    selection_str = min_col + min_row + ':' + max_col + max_row
    for row in ws[selection_str]:
        temp_array = []
        for cell in row:
            temp_array.append(str(cell.value).encode("utf8"))
        if temp_array != [None]:
            return_array.append(temp_array)

    # Return array selection
    print(return_array)
    return return_array

def letter_to_index(letter):
    """Converts a column letter, e.g. "A", "B", "AA", "BC" etc. to a zero based
    column index.

    A becomes 0, B becomes 1, Z becomes 25, AA becomes 26 etc.

    Args:
        letter (str): The column index letter.
    Returns:
        The column index as an integer.
    """
    letter = letter.upper()
    result = 0

    for index, char in enumerate(reversed(letter)):
        # Get the ASCII number of the letter and subtract 64 so that A
        # corresponds to 1.
        num = ord(char) - 64

        # Multiply the number with 26 to the power of `index` to get the correct
        # value of the letter based on it's index in the string.
        final_num = (26 ** index) * num

        result += final_num

    # Subtract 1 from the result to make it zero-based before returning.
    return result - 1

def count_rows(workbook, worksheet):
    wb = load_workbook(str(workbook))
    ws = wb[str(worksheet)]
    row_count = str(ws.max_row)
    print("count rows:",row_count)
    return row_count

def count_cols(workbook, worksheet):
    wb = load_workbook(str(workbook))
    ws = wb[str(worksheet)]
    col_count = str(ws.max_column)
    print("count cols:",col_count)
    return col_count

def genColDict(workbook, worksheet):
    # Dictionary associating output column to an input column number
    # 'Output Column Name' : 'Input Column #'
    wb = load_workbook(str(workbook))
    ws = wb[str(worksheet)]
    max_col = int(count_cols(workbook, worksheet))

    print("In column dictionary...")
    print("# of cols:\t",max_col)

    # Save header in list
    first_row = list(ws.rows)[0]
    print(first_row)

def merge_tables(observation_array, collector_array, input_wb, input_ws, num_rows=None):
    # Load XLSX file
    wb = Workbook()
    ws = wb.active
    ws.title = input_ws

    # Initialize values
    month = ['i','ii','iii','iv','v','vi','vii','viii','ix','x','xi','xii']
    header_row = [  'iNaturalist ID',
                    'Collection Day 1',
                    'Month 1',
                    'Year 1',
                    'Collection Day 2',
                    'Month 2',
                    'Year 2',
                    'Collector Name',
                    'Collection No',
                    'Sample No',
                    'State',
                    'County',
                    'City',
                    'Location',
                    'Lat',
                    'Long',
                    'Collection method',
                    'Associated plant'
                ]
    updated_header_row = [ 'Date Label Printed',
        'Date Label Sent', 'Observation No.', 'Voucher No.',
        'iNaturalist ID', 'iNaturalist login',
        'Collector - First Name', 'Collector - Last Name',
        'Collection Day 1',	'Month 1', 'Year 1', 'Time 1',
        'Collection Day 2',	'Month 2', 'Year 2', 'Time 2', 'Collection Day 2 Merge',
        'Sample ID', 'Specimen ID',
        'Country', 'State', 'County', 'Location', 'Abbreviated Location',
        'Projects', 'Dec. Lat.', 'Dec. Long.', 'Lat/Long Accuracy', 'Elevation',
        'Collection method', 'Associated plant', 'Inaturalist URL',
        'Specimen Sex/Caste', 'Sociality', 'Specimen Family', 'Specimen SubFamily',
        'Specimen Tribe', 'Specimen Genus', 'Specimen SubGenus',
        'Bee Species', 'Morphology', 'Determined By', 'Date Determined', 'Verified By',
        'Other Determiner(s)', 'Other Dets. Sci. Name(s)', 'Additional Notes'
    ]

    # Write header to first row in sheet
    ws.append(updated_header_row)

    # Translate observation list rows to template format
    print("Translating iNaturalist list...")
    for row in observation_array:
        print(row)
        # Initialize values
        sample_num_flag = 0
        result_array = []

        # Begin constructing row to append to spreadsheet, saved in result_array

        # Col 0
        # Date label printed
        result_array = eval_dateLabelPrinted(result_array, row[0])

        # Col 1
        # Specimen ID
        result_array = eval_dateLabelSent(result_array, row[1])

        # Col 1
        # iNaturalist ID
        result_array = eval_iNaturalistID(result_array, row[0])

        # Col 2, 3, 4
        # Collection Day 1
        result_array = eval_collDay1(result_array, month, row[1])

        # Col 5, 6, 7
        # Collection Day 2
        result_array = eval_collDay2(result_array, month, row[17])

    # Col 1
    # Sample ID
    # result_array = eval_sampleID(result_array, row[X])

        # Col 8
        # Collector Name
        result_array = eval_collName(result_array, collector_array, row[2])

        # Col 9
        # Collection No
        result_array = eval_collNum(result_array, row[3])

        # Col 10
        # Sample No
        result_array, sample_num_flag = eval_sampleNum(result_array, sample_num_flag, row[4])

        # Col 11
        # State
        result_array = eval_state(result_array, row[6])

        # Col 12
        # County
        result_array = eval_county(result_array, row[7])

        # Col 13
        # City
        result_array = eval_city(result_array, row[15])

        # Col 14
        # Location
        result_array = eval_location(result_array, row[8])

        # Col 15, 16
        # Lat & Long
        result_array = eval_latLong(result_array, row[9], row[10])

    # Col X
    # Positional accuracy
    # result_array = eval_positional_acc(result_array, row[X])

    # Col X
    # Elevation
    # result_array = eval_positional_acc(result_array, row[X])

        # Col 17
        # Collection Method
        result_array = eval_colMethod(result_array, row[14])

        # Col 18
        # Associated Plant
        result_array = eval_assocPlant(result_array, row[12])

        # Append to xlsx file
        if sample_num_flag == 1:
            if row[4] != None:
                temp_range = row[4]
                for x in range(0, int(temp_range), 1):
                    print(result_array)
                    ws.append(result_array)
                    result_array[9] = int(result_array[9]) + 1
        else:
            print(result_array)
            ws.append(result_array)

    print("Appending results to file...")
    wb.save(filename = input_wb)
    print("Saved.")

def parse_cmd_line():
    print("cmd line parser function")
    print("cmd arguments: " , str(sys.argv))

    # Iterate through cmd line and assign strings to input and output paths
    i = 0; in = ''; out = ''
    for arg in sys.argv:
        if arg == "--input":
            in = sys.argv[i+1]
        elif arg == "--output":
            out = sys.argv[i+1]
        i += 1

    # By now we should have a valid input path at minimum
    if in == '':
        print("\'--input\' argument must be provided. Exitting.")
        sys.exit()

    if out == '':
        print(in.split('/'))
        #out =

def create_file(file_string):
    print("create file function")

def main():
    # Variables to keep track of
    input_file = ""
    output_file = ""
    input_file_type = ""

    # Initiate / Default values
    input_folder = 'default'
    output_folder = 'default'
    min_col = 'A'
    min_row = '2'

    # Default path settings
    observation_wb = 'data/' + input_folder + '/2018_iNaturalist.xlsx'
    observation_ws = ['observations-49204', 'Oregon Bee Atlas']

    # If output/input is set in cmd line, set the output folder to mirror that
    if output_folder != 'default':
        output_wb = 'results/' + output_folder + '/Oregon_Bee_Atlas.xlsx'
    elif input_folder != 'default':
        output_wb = 'results/' + input_folder + '/Oregon_Bee_Atlas.xlsx'
    else:
        output_wb = 'results/default/Oregon_Bee_Atlas.xlsx'
    output_ws = 'Results'

    # Create directories
    if not os.path.exists(os.path.dirname(output_wb)):
        try:
            os.makedirs(os.path.dirname(output_wb))
        except OSError as exc: # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise

    file_functions.test_import()
    test_import()

    parse_cmd_line()

    # Init variables to pass into read_xlsx
        #max_col = list(string.ascii_lowercase)[ int(count_cols(observation_wb,observation_ws[0])) - 1 ].upper()
        #max_row = count_rows(observation_wb,observation_ws[0])

    # Generate dictionary connecting column names to column number
        #genColDict(observation_wb, observation_ws[0])

    # Extract values from observations table
        #observation_result = read_xlsx(observation_wb, observation_ws[0], min_col, min_row, max_col, max_row)

    # Init variables to pass into read_xlsx
        #max_col = list(string.ascii_lowercase)[ int(count_cols(observation_wb,observation_ws[1])) - 1 ].upper()
        #max_row = count_rows(observation_wb,observation_ws[1])

    # Extract values from collector ID table
        #collector_result = read_xlsx(observation_wb, observation_ws[1], min_col, min_row, max_col, max_row)

    # Generate Output Sheet
        #merge_tables(observation_result, collector_result, output_wb, output_ws)

if __name__ == '__main__':
    main()
