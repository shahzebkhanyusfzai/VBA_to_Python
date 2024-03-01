import openpyxl
import xlwings as xw
import pandas as pd
import csv
import random
import time

# @echo off
# "C:\Users\shahz\AppData\Local\Programs\Python\Python311\python.exe" "C:\path\to\your\excelV3.py"
# pause


def projection():
    # Load the Excel workbook and sheets
    wb = xw.Book("C:/Users/shahz/Desktop/excel225/StochasticScenarios.xlsm")  #CLIENT EDIT : PASTE YOUR Main Macro Enabled Workbook Directory here
    sht_control = wb.sheets['Control']
    sht_inforce = wb.sheets['Inforce']
    sht_assumptions = wb.sheets['Assumptions']
    
    # Load ranges from Assumptions tab
    lapse_range = sht_assumptions.range('LapseRange').value
    mortality_range = sht_assumptions.range('MortalityRange').value

    # Read values from Control tab
    run_number = int(sht_control.range("NumScen").value)
    o_file = sht_control.range("file").value
    pandemic_incidence = sht_assumptions.range("PandemicIncidence").value
    pandemic_severity = sht_assumptions.range("PandemicSeverity").value

    # Loop through stochastic scenarios
    with open(o_file, 'w', newline='') as file:
        writer = csv.writer(file)
        for i in range(run_number):
            total_fee = [0] * 600
            pandemic_year = [0] * 50
            pandemic_factor = [0] * 600
            
            # Determine if a pandemic happens in a year
            for j in range(50):
                if j == 0:
                    pandemic_year[j] = 1 if random.random() < pandemic_incidence else 0
                else:
                    pandemic_year[j] = 0.5 if pandemic_year[j - 1] == 1 else (1 if random.random() < pandemic_incidence else 0)

            
            # Determine the pandemic factors by month
            for j in range(1,601):
                # pandemic_factor[j] = pandemic_year[int((j - 0.1 / 12)+1)] * pandemic_severity
                pandemic_factor[j-1] = pandemic_year[int((j-0.1)/ 12)] * pandemic_severity

            # In VBA we called the last row as 100000, but in python its best to dynamically calculate the last row,
            last_row = sht_inforce.range('A' + str(sht_inforce.cells.last_cell.row)).end('up').row
            # Load the Inforce data
            df_inforce = sht_inforce.range(f'A1:J{last_row}').options(pd.DataFrame, header=1, index=False).value

            # Loop through policies
            for _, row in df_inforce.iterrows():
                if row['Current Age (Months)'] > 0:
                    age = int(row['Current Age (Months)'])
                    duration = int(row['Policy Duration (Months)'])
                    mortality_table = int(row['Mortality Table'])
                    lapse_table = int(row['Lapse Table'])
                    fee = float(row['Monthly Fee (in dollars)'])
                    fee_mode = int(row['Premium Mode'])

                    survival_rate = 1 
                
                for k in range(600):

                    
                    if k == 0: # k = 0, because of 0 based indexing
                        lapse_rate = 0
                        mortality_rate = 0
                    else:

                        if mortality_range[age - 1][mortality_table - 1] is None:
                            mortality_rate = 0
                        else:
                            mortality_rate = 1 - ((1 - mortality_range[age - 1][mortality_table - 1] * pandemic_factor[k]) ** (1/12))
                        
                        lapse_rate = lapse_range[duration - 1][lapse_table - 1] # Considering 0-based indexing

                        if lapse_rate < random.random():
                            lapse_rate = 0
                        else:
                            lapse_rate = 1

                        if mortality_rate < random.random():
                            mortality_rate = 0
                        else:
                            mortality_rate = 1

                        
                        survival_rate = survival_rate * (1 - lapse_rate) * (1 - mortality_rate)
                    
                    total_fee[k] += survival_rate * fee
                    duration += 1
                    age += 1
                    
                    if survival_rate == 0:
                        break
            
            # Write to output csv file
            writer.writerow(total_fee)
    
            print(f'Progress {(i/100)*100}%')

if __name__ == '__main__':
    start_time = time.time()
    projection()
    end_time = time.time() 
    elapsed_time = end_time - start_time
    print(f"Program executed in {elapsed_time} seconds.")
    

