#! python3

import openpyxl, os, re, fnmatch, csv
print('Opening workbook...')

#Find our directory, if its not the working path change to it
currentPath = os.getcwd()
workingDir = 'C:\\Users\\alogan\\Desktop\\TEMP_GIS\\python\\tpy\\Testing\\SSF'

if currentPath != workingDir:
    os.chdir(workingDir)
    currentPath = os.getcwd()

#os.path.getsize(path)
#os.listdir(path)

pathFileList = os.listdir(currentPath)
#pathFileDateRegex = re.compile(r'\d\d\d\d\d\d\d\d')

#set Regular Expression to get ready to look for '########' (8 digit number format)
pathFileDateRegex = re.compile(r'\'\d\d\d\d\d\d\d\d\'')
fileListLen = len(pathFileList)
#def fileDateSearch:
#    pathFileDateRegex = re.search(r'\d\d\d\d\d\d\d\d')
#    return pathFileDateRegex

#Loop through the length of the file list
#for i in range(fileListLen):

#    filename = pathFileList[i]
#    filename = filename.rsplit('.', 1)[0]
#    print(filename)
    
    #run regex on file list
    #captures all 8 digit dates, TODO - sort out .files before this line executes

#findall occurences of compiled Regex within the current directory path
updatedPathFileList = pathFileDateRegex.findall(str(pathFileList))
updatedFileListLen = len(updatedPathFileList)
    #pathFileDateRegex.fullmatch(str(pathFileList))
    #print(pathFileList[i])

#updatedPathFileDate = pathFileDateRegex.findall(str(pathFileList))
#pathFileDate list includes "''" chars, write function to sort list and remove '' chars


#Look at pathFileDate list, load entries into directory change string
#os.chdir(datedFolder)
#ex.: os.chdir('C:\\Users\\alogan\\Desktop\\TEMP_GIS\\python\\tpy\\Testing\\SSF\\20170426')
#os.listdir(currentPath)
#updatePathFiles with .csv, .xlsx, .xls extensions to grab with correct Regex filtered names
#^New system doesnt use Regex, capture files using fnmatch below

#if updatedPathFileList != []:
curListLen = updatedFileListLen
curListPos = 0
csvFileList = []
#TODO refactor into function
while curListLen > 0 :
    chFolder = updatedPathFileList[curListPos]
    chFolder = chFolder.strip('\'')
    #do something
    #os.chdir(workingDir)
    #newDir = workingDirectory.join('\\' + chFolder)
    #tmpDirLen = len(workingDir) - 1
    #newDir = workingDir + 'test'
    newDir = str(workingDir)
    newDir = str(newDir + '\\' + chFolder)
    os.chdir(newDir)
    currentPath = os.getcwd()
    pathFileList = os.listdir(currentPath)
    fileListLen = len(pathFileList)
    #Loop through files in new dir, sort csv's and select one to open
    #
    #fnmatch
    #'.' =actual current path
    #Finds csvs and puts into a list
    #TODO refactor into function
    for file in os.listdir('.'):
        if fnmatch.fnmatch(file, '*.csv'):
            csvFileList.append(file)
    fileLen = len(csvFileList)
    #Build Full File pathing
    fileDir = newDir + '\\' + csvFileList[0]
    fileDir_xlsx = fileDir.strip('.csv')
    #Strip .csv and append .xlsx for new save file name
    fileDir_xlsx = fileDir_xlsx + '.xlsx'
    print(str(fileLen) + ' .CSV files found')
    print('Converting X .CSV files to .xlsx')
    
    wb = openpyxl.Workbook()
    ws = wb.active

    #csvFile = open(csvFileList[0])
    #reader = csv.reader(csvFile, delimiter=',')
    #reader = csv.reader(open(fileDir, "rU")) #"rb" bytes mode for NUL - check csv is a real csv
    #reader = csv.reader(fileDir, "rU")
    #convert csv to xlsx
    csvFile = open(fileDir)
    csvReader = csv.reader(csvFile, delimiter=',')
    for row in csvReader:
        ws.append(row)
    csvFile.close()

    wb.save(fileDir_xlsx)

    
    print('Opening workbook...')
    nwb = openpyxl.load_workbook(fileDir_xlsx)
    nws = nwb.active
    #TODO - only one csv can be opened for ease of testing, need to extend
    #

    #update list index at end of loop
    curListPos += 1
    #update list length to end loop at 0 ---- inelegant but works?
    curListLen -= 1
