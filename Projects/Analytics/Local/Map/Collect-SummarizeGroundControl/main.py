# This script combs through delimited text files to find and extract ground control points

# import dependencies
import csv







# open/read text files
with open(control_file_path, "r") as control_file:
    control_points = control_file.read()
    #append to control point collection
