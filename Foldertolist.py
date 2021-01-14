import os
import datetime
import time
from pathlib import Path
import glob
from stat import S_ISREG, ST_CTIME, ST_MODE
import sys
import xlwings as xw

pathwb = r'‪F:\\Documents\\AI PROJECTS\\exceltest.xlsx'

dir_path = r'F:\\Documents\\test folder'
# get all entries in the directory
entries = (os.path.join(dir_path, file_name)
           for file_name in os.listdir(dir_path))
# Get their stats
entries = ((os.stat(path), path) for path in entries)
# leave only regular files, insert creation date
entries = ((stat[ST_CTIME], path)
           for stat, path in entries if S_ISREG(stat[ST_MODE]))
# seperate date from the list
datesc = []
filenamesc = []
for x in entries:

    datesc.append(x[0])
    filenamesc.append(x[1])

# excel
wbstatus = xw.books('status.xlsx')
ws1 = wbstatus.sheets['tab1']
ws2 = wbstatus.sheets['tab2']
ws3 = wbstatus.sheets['tab3']
ws1.range('A2').options(transpose=True).value = datesc
ws1.range('F2').options(transpose=True).value = filenamesc
