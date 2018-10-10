'''
Created on Sep 28th, 2018
Author: Miles McCall
'''

# Library Imports
import os
import sys
import argparse
import random
#from random import uniform
#from random import choice
#from random import seed
#from random import random

#Functions
#def gen_row():
#    row_string =  str(uniform(20.0, 60.0)) + ","
#    row_string = row_string + str(uniform(100.0, 140.0)) + ","
#    row_string = row_string + str(choice(['Apis mellifera','Lasioglossum','Seladonia','Osmia']))
#    return row_string

def write_csv(num_rows):
    print("Writing randomized entries to output csv...")

    f = open("results/test_results.csv", "a")
    for i in range(num_rows):
        row_string = str(i) + "," # ID number
        row_string = row_string + str(random.randint(20,60)) + "," # Latitude
        row_string = row_string + str(random.randint(100,140)) + "," # Longitude
        row_string = row_string + str(random.choice(['Apis mellifera','Lasioglossum','Seladonia','Osmia'])) # Species Name
        row_string = row_string + "\n"
        f.write(row_string)

    print("\t--> Done.")

def main():
    write_csv(1000)

if __name__ == '__main__':
    main()
