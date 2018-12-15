from daqmx_class import DaqmxSession
import time
import xlsxwriter
import datetime
import os.path


def single_method_benchmark(session, method_to_test, trials):
    session.open()
    results = []
    trial = 1
    for trial in range(trial, trials+1):
        start_time = time.perf_counter()
        method_to_test()
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        print(elapsed_time)
        print(trial)
        results.append([trial, elapsed_time])
        trial += 1
    session.close()
    return results

def open_benchmark(session, trials):
    start_time = time.perf_counter()
    session.open()
    end_time = time.perf_counter()
    session.close()
    elapsed_time = end_time - start_time
    return elapsed_time

def close_benchmark(session, trials):
    session.open()
    start_time = time.perf_counter()
    session.close()
    end_time = time.perf_counter()
    elapsed_time = end_time - start_time
    return elapsed_time

# ------------- LOGGING SETUP --------------

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
    quit()

workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()

# Rows and columns are zero indexed.
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

# ----------------- DRIVER TEST --------------

session = DaqmxSession()
trials = 2

results = single_method_benchmark(session, session.read_std, trials)
print(session.data)
print(len(session.data))
print(results)

# ------------- LOG TO FILE --------------

for x, y in (results):
    worksheet.write(row, col,     x)
    worksheet.write(row, col + 1, y)
    row += 1

average = 0
for x in results:
    average = average + x[1]
average = average / len(results)

worksheet.write(test_results_row, 0, "Average Time (s)")
worksheet.write(test_results_row, 1, average)
worksheet.write(test_results_row, 2, "Num of Iterations")
worksheet.write(test_results_row, 3, row-(data_label_row+1)) # Perhaps revise this.

workbook.close()