from __future__ import print_function
import os
import sys
import csv
import json
import requests

def test_import():
    print("col_functions.py test")

# 'Eval' functions are called from the merge_tables function to evaluate the
# contents of each column for the output spreadsheet. For each row in the
# original table, each eval function is called sequentially to construct the
# desired output row.

# I chose to encapsulate each column value in its own eval function to make
# the interpretation of input data as modular as possible. Each eval function
# uses potentially different logic to generate properly formatted output values,
# and each function can be modified independently.

def collector_name(in_file, user_name):
    # Open usernames CSV
    with open(in_file,'r') as file:
        for row in file:
            # Remove trailing characters and split line into array
            row = row.rstrip("\r\n")
            row = row.split(',')
            # Found a match...
            if row[1] == user_name:
                # First name: 1st word of column 1
                first_name = row[0].split(' ')[0]
                # First letter of the first name
                first_initial = first_name[0] + '.'
                # Last name: 2nd word of column 1
                last_name = row[0].split(' ')[1]
                #Done
                return first_name, first_initial, last_name
    return "","",""

def date_1(in_date):
    month_numeral = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XIII']
    # Check input
    if in_date == '':
        return '','',''

    date = in_date.split('/')

    if len(date) > 1:
        day = date[1]
        month = month_numeral[int(date[0]) - 1]
        year = date[2]
    else:
        date = in_date.split('-')
        day = date[2]
        month = month_numeral[int(date[1]) - 1]
        year = date[0]

    return day, month, year

def time_1(in_time):
    # Check input
    if in_time == '':
        return ''
    # Split full time string to remove date (1st word)
    in_time = in_time.split(' ')
    # Split time word by : to separate hours, mins, secs
    return_time = in_time[1].split(':')
    # Convert from UTC to PST by -7
    if (int(return_time[0]) - 7) < 6:
        return_time[0] = str(int(return_time[0]) + 24 - 7)
    else:
        return_time[0] = str(int(return_time[0]) - 7)
    # Reattach the hours and minutes, leaving out seconds
    return_time = return_time[0] + ":" + return_time[1]
    return return_time

def date_2(in_date):
    month_numeral = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XIII']
    months = ['January','February','March','April','May','June','July','August','September','October','November','December']

    # Check input
    if in_date == '' or in_date is None:
        return '','','',''

    # If the string contains 'T' it is a differently formatted date time value
    if 'T' in in_date:
        print("Found different format:",in_date)


    # Check if date was parsed into different format
    for index, curr_month in enumerate(months):
        if curr_month in in_date:
            #print("Found written format:",curr_month," | ",in_date)
            date = in_date.split(' ')
            print("split date:",date)

            day = date[0]
            month = month_numeral[index]
            year = date[2]
            merge = "-" + day + month

            return day, month, year, merge

    # If the string can be split in two, it is formatted: 'mm/dd/yy hh:mm'
    # date[0] is the date, date[1] is the time
    date = in_date.split(' ')
    print("split date:",date)

    if len(date) == 2:
        # reset date var to just include date info
        date = date[0].split('/')
        day = date[1]
        month = month_numeral[int(date[0]) - 1]
        year = date[2]
    else:
        date = in_date.split('-')
        day = date[2]
        month = month_numeral[int(date[1]) - 1]
        year = date[0]

    merge = "-" + day + month

    return day, month, year, merge

def date_2_orig(in_date):
    # Check input
    print("in_date:",in_date)
    if in_date == '':
        return '','','',''
    # Init vars
    month_numeral = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XIII']

    in_date = in_date.split('T')
    #print(in_date)
    # There are several input formats for this col
    # If the format can be split with T continue:
    if len(in_date) == 2:
        #print("Splitting with T...")
        in_date = in_date[0].split('-')
        # Parse values from full date string
        day = in_date[2]
        # Reference numeral array
        month = month_numeral[int(in_date[1]) - 1]
        # Year is straight forward
        year = in_date[0]
        # Calculate merge string
        merge = "-" + day + month
    else:
        #print("Splitting without T...")
        in_date = in_date[0].split(' ')
        in_date = in_date[0].split('/')

        #Check if month is in first or second location
        if int(in_date[1]) > 12:
            day = in_date[1]
            month = month_numeral[int(in_date[0]) - 1]
        else:
            day = in_date[0]
            month = month_numeral[int(in_date[1]) - 1]
        year = in_date[2]
        merge = "-" + day + month

    return day, month, year, merge

def time_2(in_time):
    # Check input
    if in_time == '' or in_time is None:
        return ''

    # If the string can be split in two, it is formatted: 'mm/dd/yy hh:mm'
    # in_time[0] is the date, in_time[1] is the time
    in_time = in_time.split(' ')

    if len(in_time) == 2:
        in_time = in_time[1]
        in_time = in_time[0] + ":" + in_time[1]
    #else:
    #    in_time = in_time[0].split(' ')
    #    in_time = in_time[1]

    return in_time

def time_2_orig(in_time):
    # Check input
    if in_time == '':
        return ''

    # Split full time string to remove date (1st word)
    in_time = in_time.split('T')

    if len(in_time) == 2:
        #print("Splitting with T...")
        in_time = in_time[1].split(':')
        in_time = in_time[0] + ":" + in_time[1]
    else:
        #print("Splitting without T...")
        in_time = in_time[0].split(' ')
        in_time = in_time[1]

    return in_time

def location_guess(address, cities_file):
    #print("raw:",address)

    address = address.split(", ")
    #print("split on comma:",address)

    # If 'normal' address, should split in 4
    if len(address) == 4:
        #print("guess:",address[1])
        return address[1]
    else:
        #print("unusual format...")
        return ""

def specimen_id(specimen_id_str):
    if specimen_id_str is None:
        return " "
    if specimen_id_str != '':
        if specimen_id_str[0].isalpha():
            return "NOT INT"
        else:
            return specimen_id_str

def round_coord(coord):
    temp = '%.4f'%(float(coord))
    if len(temp.split('.')[1]) < 4:
        print("coordinate didn't have 4 digits:",temp)
        temp = float(str(temp) + "0")
        print("fixed(?):",temp)
    return temp

def write_elevation_res(elevation_file, lat, long, elevation):
    lat_rounded = '%.2f'%(float(lat))
    long_rounded = '%.2f'%(float(long))
    line_to_print = str(lat_rounded),str(long_rounded),str(elevation)

    # New System:
    # Check if file has been created: Elevations/input_file_elevation.csv
    # Write file or Append file based on this
    if not os.path.exists(elevation_file):
        with open(elevation_file, 'w') as file:
            file.write(line_to_print)
    else:
        with open(elevation_file, 'a') as file:
            writer = csv.writer(file)
            writer.writerow(line_to_print)

    # Old System:
    # Always append results to same master file
    #with open(elevation_file, 'a') as f:
    #    writer = csv.writer(f)
    #    writer.writerow(line_to_print)

def read_elevation_csv(elevation_file, lat, long):
    # Create more matches by simplifying coordinates
    #print(lat,long)
    lat_rounded = '%.2f'%(float(lat))
    long_rounded = '%.2f'%(float(long))

    with open(elevation_file) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        for row in csv_reader:
            if len(row) > 1 and str(row[0]) == lat_rounded and str(row[1]) == long_rounded:
                return row[2]
    return ''

def create_elevation_csv(elevation_file):
    # If the elevation CSV hasn't been created, generate it.
    if not os.path.exists(elevation_file):
        f = open(elevation_file,"w+")
        f.close()

def elevation(input_file_name, lat, long):
    # Goal with elevation calculation:
        # Download GIS data on OR's elevation/topography
        # Don't need elevation CSV, simply reference the GIS map with coordinates

    # Current elevation calculation:
        # 1: See if coordinates have already been calculated before
            # > Parse matching elevation csv if present in output
        # 2: Then call API to request elevation of coordinates

    # Establish Variables & Files
    elevation_file_name = "elevations/" + input_file_name + "_elevations.csv"
    create_elevation_csv(elevation_file_name)

    #check if the current lat and long have already been calculated
    csv_result = read_elevation_csv(elevation_file_name, lat, long)

    #print(csv_result)
    if csv_result != '':
        # A matching set of coordinates was found in results
        return int(float(csv_result))
    else:
        # Couldn't find a matching coordinate, use the API
        apikey = "AIzaSyBoc369wPHoc2R3fKHBSiIh4iwIY4qk7P4"
        url = "https://maps.googleapis.com/maps/api/elevation/json"
        request = requests.get(url+"?locations="+str(lat)+","+str(long)+"&key="+apikey)
        try:
            results = json.loads(request.text).get('results')
            if 0 < len(results):
                elev = results[0].get('elevation')
                #resolution = results[0].get('resolution') # for RESOLUTION
                # ELEVATION
                write_elevation_res("data/elevations.csv", lat, long, int(float(elev)))
                return int(float(elev))
            else:
                print('HTTP GET Request failed.')
        except ValueError as e:
            print('JSON decode failed: '+str(request) + str(e))

        return 0
