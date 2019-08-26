# Read the csv file
import csv
codeddata = []
with open('DataBefore.csv', newline='') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=',')
    for row in spamreader:
        codeddata.append(row)
        
# We can look at the first few rows 
codeddata[:5]


# Clean up data:

merged = []
for row in codeddata:
    if len(row[0]) == 0 and len(row[1]) == 0: # Delete empty rows (if any data coded included text across paragraph 
        pass                                  # breaks you may generate empty rows when you unmerge cells in Excel)
    else:
        if len(row[1]) == 0:                  # If this data does not have an assigned code then add data to the
            merged[-1][0] += " " + row[0]    # previous row (where it belongs!)
        else:
            merged.append(row)
                
merged[:5]


# Produce a csv with the clean data:

with open('DataAfter.csv', 'w', newline='') as csvfile:
    spamwriter = csv.writer(csvfile, delimiter=';',   # If your data includes commas, choose characters to use as delimiters (e.g. %)
                            quotechar='"', quoting=csv.QUOTE_MINIMAL)
    for row in merged:
        spamwriter.writerow(row)