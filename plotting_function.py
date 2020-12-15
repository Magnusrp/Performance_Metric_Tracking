# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import glob
import io
import xlsxwriter
import seaborn as sns
sns.set()

data = []
store_avg = []
for i in glob.glob("spreadsheets/*.xlsx"):
    df = pd.read_excel(i)
    df.columns = df.iloc[6]
    store_avg.append(df.iloc[4,6])
    df = df[7:]
    df = df[['Employee','LW']]
    df = df.set_index('Employee')
    df.columns = [i[13:-5]]          # renames the column to FWi, where i is the week
    df = df.dropna(axis=0)           # removes employees that have no data for any week
    data.append(df)

# joins all the dataframes in order to create a master array of all the employees for every week
perf_list_ytd = data[0]
for i in np.arange(1,len(data)):
    perf_list_ytd = perf_list_ytd.join(data[i])
perf_list_ytd = perf_list_ytd[~perf_list_ytd.index.duplicated(keep='first')]

# in case there are fewer than 4 weeks of data available
if len(glob.glob("spreadsheets/*xlsx")) < 4: num_weeks = len(glob.glob("spreadsheets/*xlsx"))
else: num_weeks = 4

workbook = xlsxwriter.Workbook('Plots.xlsx')
wks1=workbook.add_worksheet('Performance')

for i in np.arange(len(perf_list_ytd.index)):
    employee = perf_list_ytd.index[i]
    employee_perf = perf_list_ytd.loc[employee][-1*num_weeks:]
    fig,ax = plt.subplots()
    # plots and labels every partner's IPMs alongside the store average
    plt.plot(np.arange(num_weeks),store_avg[-1*num_weeks:],'-r.',employee_perf,'-bo')
    plt.xlabel('LW:  '+str(round(employee_perf[-1],2)),size=14)
    plt.ylabel('Performance Metric',size=16)
    plt.title(employee_perf.name,size=20)
    # writes the plot into the Plots.xlsx spreadsheet
    wks1.write(25*i,0,employee)
    imgdata=io.BytesIO()
    plt.savefig(imgdata, format='png')
    wks1.insert_image(25*i,1, '', {'image_data': imgdata})

workbook.close()