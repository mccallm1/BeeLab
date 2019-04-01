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

def eval_iNaturalistID(result_array, id_string):
    result_array.append(str(id_string))
    return result_array

def eval_collDay1(result_array, month, date_string):
    temp_date = str(date_string).split(" ")
    temp_date = temp_date[0].split("-")

    result_array.append(str(temp_date[2])) # Day
    result_array.append(str(month[int(temp_date[1])])) # Month
    result_array.append(str(temp_date[0])) # Year

    return result_array

def eval_collDay2(result_array, month, date_string):
    if date_string == None:
        result_array.append("-")
        result_array.append("-")
        result_array.append("-")

    elif len(str(date_string)) == 25:
        temp_date = str(date_string)[:10]
        temp_date = temp_date.split("-")

        result_array.append(str(temp_date[2])) # Day
        result_array.append(str(month[int(temp_date[1])])) # Month
        result_array.append(str(temp_date[0])) # Year

    else:
        result_array.append("manual fill")
        result_array.append("manual fill")
        result_array.append("manual fill")

    return result_array

def eval_collName(result_array, collector_array, code_string):
    # Convert the collector code to a matching Name from table
    match_flag = 0
    for c_row in collector_array:
        if c_row[0].lower() == str(code_string).lower():
            match_flag = 1
            result_array.append(str(c_row[1] + " " + c_row[2]))
            break

    if match_flag == 0:
        result_array.append("-")

    return result_array

def eval_collNum(result_array, num_string):
    result_array.append(str(num_string))
    return result_array

def eval_sampleNum(result_array, sample_num_flag, sample_num):
    if sample_num != None:
        if int(sample_num) > 1:
            # We want to create a row for each sample collected,
            # so we set the flag now to reference later
            sample_num_flag = 1
        result_array.append(1)
    else:
        # -1 represents an error or empty value
        result_array.append(-1)

    return result_array, sample_num_flag

def eval_state(result_array, num_string):
    result_array.append(str(num_string))
    return result_array

def eval_county(result_array, county_string):
    result_array.append(str(county_string))
    return result_array

def eval_city(result_array, city_string):
    if city_string == None:
        result_array.append("-")
    else:
        result_array.append(str(city_string))
    return result_array

def eval_location(result_array, loc_string):
    result_array.append(str(loc_string))
    return result_array

def eval_latLong(result_array, lat, long):
    # Lat
    if lat == None:
        result_array.append("-")
    else:
        result_array.append(str(round(lat,4)))
    # Long
    if long == None:
        result_array.append("-")
    else:
        result_array.append(str(round(long,4)))

    return result_array

def eval_colMethod(result_array, method_string):
    result_array.append(str(method_string))
    return result_array

def eval_assocPlant(result_array, plant_string):
    result_array.append(str(plant_string))
    return result_array

def eval_elevation(result_array, lat_string, long_string):
    # Use lat and long to look up elevation data
    elevation = 0;
    result_array.append(str(elevation))
    return result_array

def eval_positional_acc(result_array, accuracy_string):
    result_array.append(str(accuracy_string))
    return result_array

def eval_specimenID(result_array, id_string):
    result_array.append(str(id_string))
    return result_array

def eval_sampleID(result_array, id_string):
    result_array.append(str(id_string))
    return result_array

def eval_collectorID(result_array, id_string):
    result_array.append(str(id_string))
    return result_array

def eval_specimenURL(result_array, url_string):
    result_array.append(str(url_string))
    return result_array

def eval_dateLabelPrinted(result_array, date_string):
    result_array.append(str(date_string))
    return result_array

def eval_dateLabelSent(result_array, date_string):
    result_array.append(str(date_string))
    return result_array
