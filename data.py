import xlrd
import csv
import xlsxwriter


path_of_excelsheet = "C:\\Users\\DELL\Downloads\\new_mcg_246230_combined.xlsx"
workbook = xlrd.open_workbook(path_of_excelsheet)
worksheet = workbook.sheet_by_index(0)
my_csv_file = open("my_csv_file.csv", 'w')
writer = csv.writer(my_csv_file, quoting=csv.QUOTE_ALL)

for rownum in range(worksheet.nrows):
    writer.writerow(list(x.encode('utf-8').decode('utf-8') if type(x) == type(u'') else x for x in worksheet.row_values(rownum)))

my_csv_file.close()










