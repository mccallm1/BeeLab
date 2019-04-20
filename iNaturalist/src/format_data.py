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
import csv
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
    #from openpyxl import Workbook
    #from openpyxl.reader.excel import load_workbook
    #from openpyxl.compat import range
    #from openpyxl.utils import get_column_letter
import elevation
# Local imports
import col_functions
import file_functions
from file_functions import test_import

#Functions
def read_xlsx_orig(wb_name, ws_name, min_col, min_row, max_col, max_row):
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

def merge_tables(observation_array, collector_array, input_wb, input_ws, num_rows=None):
    # Load XLSX file
    wb = Workbook()
    ws = wb.active
    ws.title = input_ws

    # Initialize values
    month = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XIII']
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

#####################################




def parse_cmd_line():
    # Parse the command line arguments:
    # Determine an input path, output path, and input file type
    print("Parsing command line arguments...")

    # Init vars
    i = 0
    in_file = ''
    out_file = ''
    in_type = ''

    # Iterate through cmd line and assign strings to input and output paths
    for arg in sys.argv:
        if arg == "--input":
            in_file = sys.argv[i+1]
        elif arg == "--output":
            out_file = sys.argv[i+1]
        i += 1

    # Input path:
        # data/folder_name/file_name

    # Input file is required
    if in_file == '':
        print("\'--input\' argument must be provided. Exitting.")
        sys.exit()

    # The input file must be kept in data dir
    if in_file.split('/')[0] != "data":
        print("\'--input\' file must be saved inside the data directory. Exitting.")
        sys.exit()

    # Check input file type
    in_type = in_file.split('/')[len(in_file.split('/')) - 1]
    in_type = in_type.split(".")[len(in_type.split('.')) - 1]

    # Output path:
        # results/folder_name/file_name

    # If the output file does exist it must be kept in the results dir
    if out_file != '' and out_file.split('/')[0] != "results":
        print("\'--output\' file must be saved inside the results directory. Exitting.")
        sys.exit()

    # If output was not specified, use the input folder name
    if out_file == '':
        # We will use the split file path components
        out_file += "results/" + in_file.split('/')[1] + "/results.csv"

    # Its necessary to append a '/' to the output folder so its treated as a dir
    out_folder = "results/" + in_file.split('/')[1] + "/"

    # Create output folder
        #print("creating output folder & files...\t",out_folder)
        #print("out path:\t",os.path.dirname(out_folder))
        #print("path exists?:\t",str(os.path.exists(os.path.dirname(out_folder))))
    if not os.path.exists(os.path.dirname(out_folder)):
        try:
            os.makedirs(os.path.dirname(out_folder))
        except OSError as exc: # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise

    # Create output file
    f = open(out_file, "w")

    # Return vars
    return in_file, out_file, in_type.lower()

def read_csv_header(file_string):
    file = open(file_string, "r") # Open CSV file
    return file.readline() # Return only the first line of the CSV

def read_xlsx_header(wb_name, ws_name):
    wb = load_workbook(wb_name)
    ws = wb[str(ws_name)]

    max_col = list(string.ascii_lowercase)[ws.max_column - 1].upper()
    print(max_col)
    print(ws.max_column)

    selection_str = "A1:" + max_col + "1"
    return ws[selection_str].ecnode("utf8")

# Two functions defined: One for reading CSV and one for Xlsx
def read_csv(file_string):
    print("\tReading from CSV...")

    # Open CSV file
    file = open(file_string, "r")

    # Read the first line into header var
    header_row = file.readline()

    # Iterate through rest of file, saving in array
    file_rows = []
    for line in file:
        file_rows.append(line)

    # Returns the header row and line array
    return header_row, file_rows

def read_xlsx(file_string):
    print("\tReading from Excel spreadsheet...")

def read_data(file_string, file_type):
    print("Reading data from input source...")
    # Variables to capture the header row and following data
    header = ''
    file_rows = []

    # Check which file type to read from
    if file_type == "csv":
        header, file_rows = read_csv(file_string)
    elif file_type == "xlsx":
        header, file_rows = read_xlsx(file_string)
    else:
        print("Invalid input file type. Exitting.")
        sys.exit()

    # Strip header of extra characters and converts individuals chars into words
    header = header.strip().split(',')
    return header, file_rows

def search_header(header, search_str):
    # Locate the index of a string in the header row
    print("\tLocating \'",search_str,"\'... ",end="")
    for index, col in enumerate(header):
        if col == search_str:
            print("found at [",index,"].")
            return index
    print("no match found.")

def print_out_header(line_to_print, csv_file):
    with open(csv_file,'w') as file:
        for index, col in enumerate(line_to_print):
            # If the current col has a value, print it
            if col != '':
                file.write(str(col))
            # If the current col isn't the last, print a comma
            # This can be separated because both
                #blank and full cols need a comma
            if index < len(line_to_print):
                file.write(',')
        # Append newline as the last step to start the next row
        file.write("\n")


def print_out_row(line_to_print, csv_file):
    with open(csv_file,'a') as file:
        # Append generated row to output file
        print(line_to_print)
        for index, col in enumerate(line_to_print):
            # If the current col has a value, print it
            if col != '':
                file.write(str(col))
            # If the current col isn't the last, print a comma
            # This can be separated because both
                #blank and full cols need a comma
            if index < len(line_to_print):
                file.write(',')
        # Append newline as the last step to start the next row
        file.write("\n")

def gen_output(out_header, out_file, in_header, in_data):
    # Create rows of formatted data and append to output csv
    print("Generating output data...")

    # Print header row
    print_out_header(out_header,out_file)

    # Parse input rows
    i = 0
    for in_row in csv.reader(in_data, skipinitialspace=True):
        i += 1
        print("\n\t",i)

        # Init the output row
        out_row = []

        # 0 Date Label Printed
        # 1 Date Label Sent
        # 2 Observation No.
        # 3 Voucher No.
        out_row.append(" ")
        out_row.append(" ")
        out_row.append(" ")
        out_row.append(" ")

        # 4 iNaturalist ID
        id = in_row[search_header(in_header,"user_id")]
        out_row.append(id)

        # 5 Collector - First Name
        # 6 Collector - First Initial
        # 7 Collector - Last Name
        u_name = in_row[search_header(in_header,"user_login")]
        f_name, f_initial, l_name = col_functions.collector_name("data/usernames.csv",u_name)
        out_row.append(f_name)
        out_row.append(f_initial)
        out_row.append(l_name)

        # 8 Collection Day 1
        # 9 Month 1
        # 10 Year 1
        # 11 Time 1
        date1 = in_row[search_header(in_header,"observed_on")]
        day1, month1, year1 = col_functions.date_1(date1)
        time1 = col_functions.time_1(in_row[search_header(in_header,"time_observed_at")])
        out_row.append(day1)
        out_row.append(month1)
        out_row.append(year1)
        out_row.append(time1)

        # 12 Collection Day 2
        # 13 Moth 2
        # 14 Year 2
        # 15 Day 2 merge
        # 16 Time 2
        date2 = in_row[search_header(in_header,"field:trap removed")]
        day2, month2, year2, merge2 = col_functions.date_2(date2)
        time2 = col_functions.time_2(in_row[search_header(in_header,"field:trap removed")])
        out_row.append(day2)
        out_row.append(month2)
        out_row.append(year2)
        out_row.append(merge2)
        out_row.append(time2)

        # 17 Sample ID
        sampleid = in_row[search_header(in_header,"field:sample id")]
        out_row.append(sampleid)

        # 18 Specimen ID
        specimenid = in_row[search_header(in_header,"field:number of bees collected")]
        out_row.append(specimenid)

        # 19 Country
        country = "USA"
        out_row.append(country)

        # 20 State
        state = "OR"
        if in_row[search_header(in_header,"place_state_name")] != "Oregon":
            state = in_row[search_header(in_header,"place_state_name")]
        out_row.append(state)

        # 21 County
        county = in_row[search_header(in_header,"place_county_name")]
        out_row.append(county)

        # 22 Location
        # 23 Abbreviated Location
        location = col_functions.location_guess(in_row[search_header(in_header,"place_guess")],"data/OR_cities.csv")
        abbreviated_location = ''
        out_row.append(location)
        out_row.append(abbreviated_location)

        # 24 Dec. Lat.
        # 25 Dec. Long.
        lat = col_functions.round_coord(in_row[search_header(in_header,"latitude")])
        long = col_functions.round_coord(in_row[search_header(in_header,"longitude")])
        out_row.append(lat)
        out_row.append(long)

        # 26 Pos Accuracy
        pos_acc = in_row[search_header(in_header,"positional_accuracy")]
        out_row.append(pos_acc)

        # 27 Elevation
        elevation = ""
        out_row.append(elevation)

        # 28 Collection method
        collection_method = in_row[search_header(in_header,"field:oba collection method")]
        out_row.append(collection_method)

        # 29 Associated plant - family
        # 30 Associated plant - species
        # 31 Associated plant - iNaturalist url
        family = in_row[search_header(in_header,"taxon_family_name")]
        species = in_row[search_header(in_header,"scientific_name")]
        url = in_row[search_header(in_header,"url")]
        out_row.append(family)
        out_row.append(species)
        out_row.append(url)

        # End of appending to output row

        # Append generated row to output file
        # If the row has multiple bees collected, expand by that many
        if specimenid.isdigit() and int(specimenid) > 1:
            print("multiple bees, print multiple times...",specimenid)
            for i in range(1, int(specimenid)):
                print("i:",i)
                out_row[search_header(out_header,"Specimen ID")] = i
                print_out_row(out_row,out_file)
                #out_row[search_header(out_header,"Specimen ID")] = int(out_row[search_header(out_header,"Specimen ID")]) + 1
        else:
            print_out_row(out_row,out_file)

        # Break for only 1 loop
        #break


def main():
    # Intro
    print("iNaturalist Pipeline -----------------------")

    # Variables to keep track of
    input_file = ""
    output_file = ""
    input_file_type = ""

    # Parse command line arguments
    input_file, output_file, input_file_type = parse_cmd_line()

    # Pipeline Description
    print("\tInput path:\t",input_file)
    print("\tOutput path:\t",output_file)
    print("\tInput type:\t",input_file_type)
    print()

    # Choose which file reading function to call
    input_header, input_rows = read_data(input_file, input_file_type)
    print()

    # Sort columns before writing output
    output_header = "Date Label Printed,Date Label Sent,Observation No.,Voucher No.,iNaturalist ID,Collector - First Name,Collector - First Name Initial,Collector - Last Name,Collection Day 1,Month 1,Year 1,Time 1,Collection Day 2,Month 2,Year 2,Collection Day 2 Merge,Time 2,Sample ID,Specimen ID,Country,State,County,Location,Abbreviated Location,Dec. Lat.,Dec. Long.,Lat/Long Accuracy,Elevation,Collection method,Associated plant - family,Associated plant - species,Associated plant - Inaturalist URL".split(",")
        # Revisit
        #output_header2 = read_xlsx_header("data/4_16_19/Output_from_Script.xlsx","Sheet1")
        #print(output_header2)

    # Create output data
    gen_output(output_header, output_file, input_header, input_rows)
    print("writing to",output_file)
    # Test elevation
    ###################################
    import requests
    import pandas as pd

    # script for returning elevation from lat, long, based on open elevation data
    # which in turn is based on SRTM
    def get_elevation(lat, long):
        query = ('https://api.open-elevation.com/api/v1/lookup'f'?locations={lat},{long}')
        r = requests.get(query).json()  # json object, various ways you can extract value
        # one approach is to use pandas json functionality:
        elevation = pd.io.json.json_normalize(r, 'results')['elevation'].values[0]
        return elevation
    #get_elevation(44.5993, -123.3157)

        #col_functions.elevation_from_coords(1,2)

        #import geocoder
        #g = geocoder.elevation([44.5993, -123.3157])
        #print ("elevation meters:",g.meters)

    ###################################

    print()


if __name__ == '__main__':
    main()
