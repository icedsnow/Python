#! python3

import openpyxl, os, re, fnmatch, csv

print('Opening workbook...')
currentPath = os.getcwd()
workingDir = 'C:\\Users\\alogan\\Desktop\\TEMP_GIS\\python\\tpy\\Testing\\SSF\\20170426'

if currentPath != workingDir:
    os.chdir(workingDir)
    currentPath = os.getcwd()

gridFile = 'K-8'
gridFilePath = 'C:\\Users\\alogan\\Desktop\\TEMP_GIS\\python\\tpy\\Testing\\SSF\\20170426\\K-8.xlsx'
wb = openpyxl.load_workbook(gridFilePath)
ws = wb.active

print(gridFilePath + '\n' + gridFile + ' is now open...')

def delete_column(ws, delete_column):
    if isinstance(delete_column, str):
        delete_column = openpyxl.cell.column_index_from_string(delete_column)
    assert delete_column >= 1, "Column numbers must be 1 or greater"

    for column in range(delete_column, ws.max_column + 1):
        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=column).value = \
                    ws.cell(row=row, column=column+1).value

newcount = 0
columnHeader = [ws.cell(row=1,column=i).value for i in range(1,28)]
#To line up columnHeader with actual column # do columnHeader + 1
key_list = ['ID', 'Ch2', 'Time', 'Take', 'Photo']
un_key_1 = 'Target_ID'
un_key_2 = 'Seed_ID'
colmax = ws.max_column
#TODO encapsulate in Try Error block (NoneType)
#Try Delete photo columns first?
"""
try:
    delete_column(ws, 25)
    for i in range(20, 28):
        header = ws.cell(row=1,column=i).value
        if 'Field' in header:
            i += 1
        else:
            k = i
            print('PDeleting Column ' + str(k))
            delete_column(ws, k)

except TypeError:
    print('TypeError encountered: Breaking')
"""
#Start of iterator loop to delete columns
for i in range(1,5):
    for i in range(1,28):
        header = ws.cell(row=1,column=i).value
        k = i
        try:
            for i in range(0,5):
                    if key_list[i] in header:
                            if un_key_1 not in header and un_key_2 not in header:
                                if header is not None:
    #                                print('Empty column found: ' +str(k))
                                    print('Deleting column ' + str(k))
                                    delete_column(ws, k)
        except TypeError:
                print('TypeError encountered: Breaking')
                break
                        
                        
        columnHeader2 = [ws.cell(row=1,column=i).value for i in range(1,colmax)]

#delete_column(ws, k)
wb.close()
wb.save('K-8_t.xlsx')
#
#
