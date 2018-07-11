# #PyBank

# Create Python script for analyzing financial records of company.
# Two sets of revenue data (`budget_data_1.csv` & `budget_data_2.csv`).
# Each dataset has two columns: `Date` and `Revenue`.
# (The records are simple.)

# Calculate the following:

# * The total number of months included in the dataset (unique months?)

# * The total amount of revenue gained over the entire period

# * The average change in revenue between months over the entire period

# * The greatest increase in revenue (date and amount) over the entire period

# * The greatest decrease in revenue (date and amount) over the entire period

# As an example, your analysis should look similar to the one below:

# ```
# Financial Analysis
# ----------------------------
# Total Months: 25
# Total Revenue: $1241412
# Average Revenue Change: $216825
# Greatest Increase in Revenue: Sep-16 ($815531)
# Greatest Decrease in Revenue: Aug-12 ($-652794)
# ```

# Your final script must be able to handle any such similarly structured dataset in the future
# (your boss is going to give you more of these -- so your script has to work for the ones to come).
# In addition, your final script should both print the analysis to the terminal and export a text file with the results.

# import dependencies
import os
import csv

# make data file paths cross-platform ready
budget_data_1_path = os.path.join("data", "budget_data_1.csv")

# open/read text files & print contents/type
with open(budget_data_1_path, "r") as budget_data_1_file:
    budget_data_1 = budget_data_1_file.read()
    print(budget_data_1)
    print(type(budget_data_1))