import win32com.client
import re
import sys
import os
import shutil

print(win32com.__gen_path__)

excel1 = win32com.client.Dispatch("Excel.Application")
wb1 = excel1.Workbooks.Open(os.path.join(sys.path[0], 'debugFilter.xlsx')) # give here

ws1 = wb1.Worksheets('1')

excel1.Visible = True
ws1.Range("A1:E500").ClearContents()

i = 0 

with open(os.path.join(sys.path[0], 'DEBUG.log'), "r") as input:
           
    for line in input:
        if (line.find('venv') != -1) or (line.find('Python37') != -1):
##            print("YAH : ",  line)
            pass
        else:
            if (line.find('File') != -1):
                
                lnS1 = line.split('File ')
                lnS2 = lnS1[1].split(' ')
                lnF = lnS2[0].strip()

                i = i + 1
                ws1.Cells(i,1).Value = lnF
                dirN = os.path.dirname(lnF)
                fbaseN = os.path.basename(lnF)
                ws1.Cells(i,3).Value = dirN
                ws1.Cells(i,4).Value = fbaseN
