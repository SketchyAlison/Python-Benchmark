# From cmd use: py -m pip install xlsxwriter

import xlsxwriter
import datetime
import os.path

auto_incr_test_num = False
test_name = "PythonBenchmarkDAQmxAI"
test_num = 1

now = datetime.datetime.now()
file_name = test_name + "_" + str(test_num) + ".xlsx"

if auto_incr_test_num:
    while os.path.isfile(file_name):
        print('File exists, auto-increasing test number...')
        test_num += 1
        file_name = test_name + "_" + str(test_num) + ".xlsx"
        print(file_name)

try:
    xlsxwriter.Workbook(file_name).close()
except PermissionError as perm_error:
    print(perm_error)
    print("Check whether file is open in Excel.")

workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
expenses = [
    ['Rent', 2],
    ['Gas',   100],
    ['Food',  300],
    ['Gym',    50],
]

# Start from the first cell. Rows and columns are zero indexed.
test_descrip_row = 0
test_results_row = test_descrip_row + 1
data_label_row = test_results_row + 2
col = 0

worksheet.set_column(0, 4, 25)

worksheet.write(test_descrip_row, 0, "Test Name")
worksheet.write(test_descrip_row, 1, test_name)
worksheet.write(test_descrip_row, 2, "Date and Time")
worksheet.write(test_descrip_row, 3, str(now))

worksheet.write(data_label_row, 0, "Trial")
worksheet.write(data_label_row, 1, "Time(s)")
row = data_label_row + 1

# Iterate over the data and write it out row by row.
for item, cost in (expenses):
    worksheet.write(row, col,     item)
    worksheet.write(row, col + 1, cost)
    row += 1

worksheet.write(test_results_row, 0, "Average Time (s)")
worksheet.write(test_results_row, 1, "AVG GOES HERE")
worksheet.write(test_results_row, 2, "Num of Iterations")
worksheet.write(test_results_row, 3, row-(data_label_row+1)) # Perhaps revise this.

workbook.close()