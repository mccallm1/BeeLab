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
from openpyxl import Workbook

# Local Imports

#Functions
def read_xlsx(wb_name, ws_name, min_col, min_row, max_col, max_row):
    # Initialize openpyxl vars
    from openpyxl import load_workbook
    wb = load_workbook(filename = str(wb_name))
    sheet_ranges = wb[str(ws_name)]

    return sheet_ranges.get_squared_range(min_col, min_row, max_col, max_row)
    #print(sheet_ranges['H11'].value)

def main():
    # Establish vars
    master = 'data/master_list.xlsx'
    m_ws = 'Master table'
    min_col = 1
    min_row = 3
    max_col = 37
    max_row = 282

    template = 'data/plants_sample_data.xlsx'
    t_ws = 'Sample spreadsheet for the plan'

    master_array = []
    master_res = read_xlsx(master, m_ws, min_col, min_row, max_col, max_row)
    for row in master_res:
        for col in row:
            master_array.append(col.value)
            #print(col.value)
        #print("\n")

    print(master_array)

if __name__ == '__main__':
    main()
