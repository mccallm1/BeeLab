'''
Created on May 14th, 2018
Author: Miles McCall
Sources:
Description: Parse the "Pollinator Plant Master List" maintained by Signe and
            generate a formatted version for uploading to the website.
'''

# External Imports
import os
import sys
import string
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter

#Functions
def read_xlsx(wb_name, ws_name, min_col, min_row, max_col, max_row):
    # Initialize openpyxl vars
    wb = load_workbook(filename = str(wb_name))
    ws = wb[str(ws_name)]
    return_array = []

    # Iterate through xlsx file
    selection_str = min_col + min_row + ':' + max_col + max_row
    for row in ws.iter_rows(selection_str):
        temp_array = []
        for cell in row:
            temp_array.append(cell.value)
        return_array.append(temp_array)
        #print(return_array)

    # Return array selection
    return return_array

def count_rows(workbook, worksheet):
    wb = load_workbook(filename = str(workbook))
    ws = wb[str(worksheet)]
    row_count = ws.max_row
    print("count rows:",row_count)
    return row_count

def count_cols(workbook, worksheet):
    wb = load_workbook(filename = str(workbook))
    ws = wb[str(worksheet)]
    col_count = ws.max_column
    print("count cols:",col_count)
    return col_count

def merge_tables(observation_array, collector_array, input_wb, input_ws, num_rows=None):
    # Load XLSX file
    wb = Workbook()
    ws = wb.active
    ws.title = input_ws
    #wb = load_workbook(filename = str(input_wb))
    #ws = wb.create_sheet(str(input_ws))

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
    ws.append(header_row)

    # Translate observation list rows to template format
    print("Translating iNaturalist list...")

    for row in observation_array:
        print(row)
        # Initialize values
        sample_num_flag = 0
        result_array = []

        # iNaturalist ID
        result_array.append(str(row[0]))

        # Collection Day 1
        temp_date = str(row[1]).split(" ")
        temp_date = temp_date[0].split("-")
        result_array.append(str(temp_date[2])) # Day
        result_array.append(str(month[int(temp_date[1])])) # Month
        result_array.append(str(temp_date[0])) # Year

        # Collection Day 2
        if row[17] == None:
            result_array.append("-")
            result_array.append("-")
            result_array.append("-")
        elif len(str(row[17])) == 25:
            temp_date = str(row[17])[:10]
            temp_date = temp_date.split("-")
            result_array.append(str(temp_date[2])) # Day
            result_array.append(str(month[int(temp_date[1])])) # Month
            result_array.append(str(temp_date[0])) # Year
        else:
            result_array.append("manual fill")
            result_array.append("manual fill")
            result_array.append("manual fill")

        # Collector Name
            # Convert the collector code to a matching Name from table
        match_flag = 0
        for c_row in collector_array:
            if c_row[0].lower() == str(row[2]).lower():
                match_flag = 1
                result_array.append(str(c_row[1] + " " + c_row[2]))
                break
        if match_flag == 0:
            result_array.append("-")

        # Collection No
        result_array.append(str(row[3]))

        # Sample No
        if row[4] != None:
            if int(row[4]) > 1:
                # We want to create a row for each sample collected
                # So we set the flag now to reference later
                sample_num_flag = 1
                result_array.append(1)
            else:
                result_array.append(row[4])
        else:
            # -1 represents an error or empty value
            result_array.append(-1)

        # State
        result_array.append(str(row[6]))

        # County
        result_array.append(str(row[7]))

        # City
        if row[15] == None:
            result_array.append("-")
        else:
            result_array.append(str(row[15]))

        # Location
        result_array.append(str(row[8]))

        # Lat
        if row[9] == None:
            result_array.append("-")
        else:
            result_array.append(str(round(row[9],4)))

        # Long
        if row[10] == None:
            result_array.append("-")
        else:
            result_array.append(str(round(row[10],4)))

        # Collection Method
        result_array.append(str(row[14]))

        # Associated Plant
        result_array.append(str(row[12]))

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

    print("Appending results to input file...")
    wb.save(filename = input_wb)
    print("Saved.")

def main():
    # Initiate / Default values
    input_folder = 'default'
    output_folder = 'default'
    min_col = 'A'
    min_row = '2'

    # Command line args
    print("cmd arguments: " , str(sys.argv))
    i = 0
    for arg in sys.argv:
        if arg == "--output":
            output_folder = sys.argv[i+1]
        elif arg == "--input":
            input_folder = sys.argv[i+1]
        i += 1

    # Default path settings
    observation_wb = 'data/' + input_folder + '/2018_iNaturalist.xlsx'
    observation_ws = ['observations-34528', 'Oregon Bee Atlas']

    if output_folder != 'default':
        output_wb = 'results/' + output_folder + '/Oregon_Bee_Atlas.xlsx'
    elif input_folder != 'default':
        output_wb = 'results/' + input_folder + '/Oregon_Bee_Atlas.xlsx'
    else:
        output_wb = 'results/default/Oregon_Bee_Atlas.xlsx'
    output_ws = 'Results'

    # Extract values from tables
    max_col = count_cols(observation_wb,observation_ws[0])
    max_row = string.ascii_lowercase[count_rows(observation_wb,observation_ws[0])]
    print(min_col + " | " + min_row + " | " + max_col + " | " + max_row)
    observation_result = read_xlsx(observation_wb, observation_ws[0], min_col, min_row, max_col, max_row)
    #print("observations: " + str(observation_result))

    max_col = count_cols(observation_wb,observation_ws[1])
    max_row = string.ascii_lowercase[count_rows(observation_wb,observation_ws[1])]
    print(min_col + " | " + min_row + " | " + max_col + " | " + max_row)
    collector_result = read_xlsx(observation_wb, observation_ws[1], min_col, min_row, max_col, max_row)
    #print("collectors: " + str(collector_result))

    # Generate Output Sheet
    #merge_tables(observation_result, collector_result, output_wb, output_ws)

if __name__ == '__main__':
    main()
