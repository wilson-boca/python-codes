import pandas as pd
import time
import subprocess
import openpyxl
import csv
import xlrd

start_time = time.time()
data = pd.read_excel("/home/rodrigo/Downloads/Relatório consumo_SCS.xlsx", read_only=True)
data.to_csv("/home/rodrigo/Downloads/ripped.csv")
print("--- pandas->: %s seconds ---" % (time.time() - start_time))

start_time = time.time()
wb = openpyxl.load_workbook(filename="/home/rodrigo/Downloads/Relatório consumo_SCS.xlsx", read_only=True)
sh = wb.get_active_sheet()
with open('/home/rodrigo/Downloads/test.csv', 'w', newline="") as f:
    c = csv.writer(f)
    for r in sh.rows:
        c.writerow([cell.value for cell in r])
print("--- openpyxl ->%s seconds ---" % (time.time() - start_time))

start_time = time.time()
book = xlrd.open_workbook("/home/rodrigo/Downloads/Relatório consumo_SCS.xlsx")
print("--- xlrd->%s seconds ---" % (time.time() - start_time))

start_time = time.time()
call = ["xlsx2csv", "/home/rodrigo/Downloads/Relatório consumo_SCS.xlsx", "/home/rodrigo/Downloads/converted.csv"]
try:
    subprocess.call(call)  # On Windows use shell=True
except Exception as ex:
    print('Failed...{}'.format(str(ex)))
print("--- xlsx2csv->%s seconds ---" % (time.time() - start_time))

start_time = time.time()
print("--- %s seconds ---" % (time.time() - start_time))

start_time = time.time()
data = pd.read_csv("/home/rodrigo/Downloads/converted.csv", usecols=['CATEGORIA'], low_memory=False)
print("--- pandas import a csv file->%s seconds ---" % (time.time() - start_time))

start_time = time.time()
data.drop_duplicates(subset="CATEGORIA", keep='first', inplace=True)
print("--- pandas drop duplicate->%s seconds ---" % (time.time() - start_time))

start_time = time.time()
data.sort_values("Empresa", inplace=True)
print("--- pandas sort->%s seconds ---" % (time.time() - start_time))

for value in data.values:
    print(value)