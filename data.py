# Download all the dependencies of the code.
import xlrd
import csv
import xlsxwriter

# Mention the path of your excel sheet
path_of_excelsheet = "C:\\Users\\DELL\Downloads\\new_mcg_246230_combined.xlsx"
# Step to open your excel workbook to memory
workbook = xlrd.open_workbook(path_of_excelsheet)
# Get the worksheet by sheet name\index number etc.
worksheet = workbook.sheet_by_index(0)
# Create a new csv file by mentionoing name for your csv file of your choice.
my_csv_file = open("my_csv_file.csv", 'w')


# Write into csv file one row at a time.
writer = csv.writer(my_csv_file, quoting=csv.QUOTE_ALL)

for rownum in range(worksheet.nrows):
    writer.writerow(list(x.encode('utf-8').decode('utf-8') if type(x) == type(u'') else x for x in worksheet.row_values(rownum)))

# Don't forget to close the csv file. 
my_csv_file.close()










