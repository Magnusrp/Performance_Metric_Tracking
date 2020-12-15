# -*- coding: utf-8 -*-

import numpy as np
import xlsxwriter
"""
The information that I used for the original program are proprietary and I cannot upload them
here. The purpose of this file is simply to produce some dummy data for the program while
replicating the formatting of the actual xlsx files that I am given at work.
"""
# hyperparameters
num_employees = 50
weeks = 6

column_labels = ['Employee',None,None,None,None,'LW']
employees = []
for i in range(num_employees):
    employees.append('Employee '+str(i+1))
    
# defines the employees AVERAGE performance    
avg_perf = np.random.uniform(20,30)
std_dev_perf = np.random.uniform(4,5)
emp_avg = np.random.normal(avg_perf, std_dev_perf, num_employees)

# creates an array of all the employees' perf metrics for every week
total_perf = np.empty((num_employees,weeks))
for i in range(len(emp_avg)):
    ith_employee = np.random.normal(emp_avg[i], 6, weeks)
    total_perf[i] = ith_employee

for i in range(weeks):
    metrics = total_perf[:,i]
    
    workbook = xlsxwriter.Workbook('spreadsheets/FW'+str(i+1)+'.xlsx')
    wks1=workbook.add_worksheet('Performance')
    
    wks1.write(5,6,np.average(metrics))      # true average of the week's metrics
    for col_num, data in enumerate(column_labels):
        wks1.write(7, col_num+1, data)
    for row_num, data in enumerate(employees):
        wks1.write(row_num+8, 1, data)
    for row_num, data in enumerate(metrics):
        wks1.write(row_num+8, 6, data)
    
    workbook.close()
