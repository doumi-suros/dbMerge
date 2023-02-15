import os    # for rename and logs
import shutil    # for copy files
import pandas    # for reading excel format records

excelFile = 'C:\\Users\\chesh\\Desktop\\Net2\\Net2.xlsx'    # excel record file name
srcPrePath = 'C:\\Users\\chesh\\Desktop\\Net2\\'    # prefix path of source folder, copy from
dstPrePath = 'C:\\Users\\chesh\\Desktop\\dbData_SNP230215\\'    # prefix path of target folder, move to
reptList = pandas.read_excel(excelFile, sheet_name=0, usecols=[1])    # set excel record reading columns
reptHead = '報告檔案夾'    # set select columns' header
#print ('check folder list : \n', os.listdir (srcPrePath), '\n', os.listdir (dstPrePath), '\n')

for m in range (0,93):    # reading range per data counts of excel record, count from "0", excluded end number
    folderName = reptList[reptHead].values[m].strip()    # get folder name from excel record

    srcDirPath = srcPrePath + folderName    # source folder path
    if not os.path.exists(srcDirPath):    # verify if company folder existed
        print (str(m) + ' : ' + folderName + ' not existed, add-in log \n')
        logFile = open ('noSourceFolderLog.txt', 'a+')
        logFile.write(str(m+1) + ', ' + folderName + '\n')    # add-in log file if company folder not existed
        m += 1
        continue
    print (str(m) + ' : copy source: ' + folderName + ' file list \n', os.listdir (srcDirPath), '\n')

    dstDirPath = dstPrePath + folderName    # target folder path
    if not os.path.exists(dstDirPath):    # verify if target folder existed
        print ('target folder not existed, run copytree : \n')
        shutil.copytree(srcDirPath, dstDirPath)    # copy all files from source folder to target folder
        print ('copytree ', folderName, '\n', os.listdir (dstDirPath), '\n')
        m += 1
        continue
    
    # target folder existed, copy source folder files to target folder
    print ('target folder existed, run copyfile, \n', os.listdir(dstDirPath), '\n')
    srcFileList = os.listdir (srcDirPath)    # get source folder file list
    dstFileList = os.listdir (dstDirPath)    # get target folder file list
    for i in range (0, len(srcFileList)):
        if srcFileList[i] in dstFileList:    # source files existed in target folder
            print (srcFileList[i], ' file existed, pass \n')
            i += 1
            continue
        
        srcFilePath = srcDirPath + '\\' + srcFileList[i]    # get source file path
        dstFilePath = dstDirPath + '\\' + srcFileList[i]    # get target file path
        shutil.copyfile(srcFilePath, dstFilePath)    # copy all files from source folder to target folder
        i += 1
    print ('copyfile ', folderName, '\n', os.listdir (dstDirPath), '\n')
    
    m += 1

print ('done ' + str(m) + ' , renew target folder list : \n', os.listdir (dstPrePath))

