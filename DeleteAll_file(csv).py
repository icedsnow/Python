#! python3

##############################################################
#
#       DELETES ALL CSV FILES IN PATH TREE
#
#
##############################################################


import os, fnmatch

cPath = os.getcwd()
wDir = r'C:\Users\alogan\Desktop\TEMP_GIS\python\tpy\Testing\SSF'

if cPath != wDir:
    os.chdir(wDir)
    cPath = os.getcwd()

pathFiles = os.listdir()
pathFilesL = len(pathFiles)

u_pathFiles = []

for i in range(0, pathFilesL):
    if os.path.isdir(pathFiles[i]):
        u_pathFiles.append(pathFiles[i])

u_pathFiles.remove('Compiled')
u_pathFilesL = len(u_pathFiles)

dir_pathList = []

for i in range(0, u_pathFilesL):
    newpath = os.path.join(cPath, u_pathFiles[i])
    dir_pathList.append(newpath)

dir_pathListL = len(dir_pathList)

for i in range(0, dir_pathListL):
    os.chdir(dir_pathList[i])
    cPath = os.getcwd()
#    u_pathFiles = os.listdir()
#    u_pathFilesL = len(u_pathFiles)

    #delete files from here
    for file in os.listdir('.'):
        if fnmatch.fnmatch(file, '*.csv'):
            filepath = os.path.join(cPath, file)
            print('Deleting... ' + str(filepath))
            os.unlink(file)
