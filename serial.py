import matplotlib.pyplot as plt
import serial
import numpy as np
from drawnow import *
import csv

from datetime import datetime
import time

import xlwt
from xlwt import Workbook

wb = Workbook()
# Workbook is created
wb = Workbook()



# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, "data1")
sheet1.write(0, 1, "data2")
sheet1.write(0, 2, "time")



arrayData = []
arrayData1 = []



dataserial = serial.Serial('COM12', 2400, timeout=0)

print("Communication  has started")

count = 0
row = 0


def plotValues():
    plt.subplot(2, 1, 1)
    plt.title('Serial Data')
    plt.grid(True)
    # plt.ylabel('Values')
    plt.plot(arrayData, 'rx-', label='data1')
    plt.legend(loc='upper left')
    plt.subplot(2, 1, 2)
    plt.grid(True)
    plt.plot(arrayData1, 'gx-', label='data2')
    plt.legend(loc='upper left')




for i in range(0, 24):
    arrayData.append(0)
    arrayData1.append(0)

plt.figure()

time.sleep(1)

while 1:
    dataserial.flush()
    mastisData = 0
    time.sleep(0.1)
    data = dataserial.readline()
    data = data[0:8]
    print(data)

    data1 = 0
    data11 = 0
    data12 = 0

    data2 = 0
    data21 = 0
    data22 = 0




    if len(data) == 8:
        count = count + 1

        data = int(data)
        dt11 = int(data / 1000000)
        dt12 = int((data - dt11 * 1000000) / 10000)
        dt21 = int((data - (dt11 * 1000000) - (dt12 * 10000)) / 100)
        dt22 = int(data % 100)
       

        data1 = dt11 * 100 + dt12
        data2 = dt21 * 100 + dt22

        print("data1", data1, "data2", data2)

        if count > 25:
            count = 25

        dt = datetime.now()
        time1 = str(dt.hour) + ":" + str(dt.minute) + ":" + str(dt.second) + "-" + str(dt.day) + "/" + str(
            dt.month) + "/" + str(dt.year)
        row = row + 1
        sheet1.write(row, 0, data1)
        sheet1.write(row, 1, data2)
        sheet1.write(row, 2, time1)
        wb.save('workbook.xls')


        arrayData.append(data1)
        arrayData1.append(data2)
        
        if count == 25:
            arrayData.pop(0)
            arrayData1.pop(0)

        drawnow(plotValues)

        dataserial.flush()
