# This script combs through delimited text files to find and extract ground control points

# import dependencies
import os, csv


keywords = ["control", "ctrl"]

directory = "B:\"

result = {}  # dict store our results

for filename in os.listdir(directory):
    for keyword in keywords:
        if keyword in filename:
			# open/read text files
			with open(control_file_path, "r") as control_file:
				control_points = control_file.read()
				#sort into ID,X,Y,Z,DESCRIPTION
				#control_points_sorted = control_points.sort blah blah blah
	
				#append to control point collection (csv, postGIS database, las, etc.)
				all_control = open("all_control.csv","a")
				all_control.write(control_points_sorted)
				all_control.close
