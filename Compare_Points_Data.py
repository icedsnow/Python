#! python3

import openpyxl, os, re, fnmatch, csv, sys
from timeit import default_timer as timer

start = timer()

cPath = os.getcwd()
wDir = r'C:\Users\alogan\Desktop\TEMP_GIS\python\tpy\Testing\Compare Testing'

if cPath != wDir:
    os.chdir(wDir)
    cPath = os.getcwd()

pathFolders = os.listdir()
pathFoldersL = len(pathFolders)

u_pathFolders = []

for i in range(0, pathFoldersL):
    if os.path.isdir(pathFolders[i]):
        u_pathFolders.append(pathFolders[i])
        
u_pathFoldersL = len(u_pathFolders)

dir_pathList = []

for i in range(0, u_pathFoldersL):
    newpath = os.path.join(cPath, u_pathFolders[i])
    dir_pathList.append(newpath)

dir_pathListL = len(dir_pathList)

#Paths established
masterPath = dir_pathList[1]
comparePath = dir_pathList[0]
resultsPath = dir_pathList[2]

#Grab Master data
os.chdir(masterPath)

#Find master file
for file in os.listdir('.'):
        if fnmatch.fnmatch(file, '*.xlsx'):
            masterFile = file
masterFile = os.path.join(masterPath, masterFile)

#Find Compare File
os.chdir(comparePath)
for file in os.listdir('.'):
        if fnmatch.fnmatch(file, '*.xlsx'):
            compareFile = file
compareFile = os.path.join(comparePath, compareFile)
#Master workbook/sheet
m_wb = openpyxl.load_workbook(masterFile)
m_ws = m_wb.active
#Compare workbook/sheet
c_wb = openpyxl.load_workbook(compareFile)
c_ws = c_wb.active

m_columnHeader = [m_ws.cell(row=1,column=i).value for i in range(1,28)]
m_columnHeader.append(0)
c_columnHeader = [c_ws.cell(row=1,column=i).value for i in range(1,28)]
m_chLen = len(m_columnHeader)
c_chLen = len(c_columnHeader)

#List of headers to compare
headerList = ['Target_ID', 'Date', 'Team' ,'Ch2_QC_R1', 'Anomaly_Ty', 'Instrument', 'Depth_in',
              'Offset_in', 'Offset_dir', 'Seed_Type', 'Seed_ID', 'Count', 'Weight_lb']


#Change headerList to lower case to avoid case error, search compare must be lowercase too
#THIS LINE CAUSING INDEXERROR, LOOPS BELOW WORK WITHOUT.
headerList = [x.lower() for x in headerList]
m_columnHeader = [x.lower() for x in headerList]

headerListLen = len(headerList)

#colNum starts at 1, increment with i
colNum = 1
#Loop through all column headers
for i in range(0, m_chLen):
    #Enter col i -> Loop through list of Header tags
    for k in range(0, headerListLen):
        if m_columnHeader[i] == headerList[k]:
            print(str(colNum) + ' : ' + str(m_columnHeader[i]))
    colNum = colNum + 1
#Reset colNum to 1
colNum = 1

for i in range(0, c_chLen):
    #Enter col i -> Loop through list of Header tags
    for k in range(0, headerListLen):
        if c_columnHeader[i] == headerList[k]:
            print(str(colNum) + ' : ' + str(c_columnHeader[i]))
    colNum = colNum + 1






"""
#old column header loop statements
#    if m_columnHeader[i] == h_targetid:
#        print(str(i) + ' : ' + str(m_columnHeader[i]))
#    elif m_columnHeader[i] == h_date:
#        print(str(i) + ' : ' + str(m_columnHeader[i]))

"""







end = timer()
print(end - start) 
