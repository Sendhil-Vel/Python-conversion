import os
import openpyxl as xl
import win32com.client as win32
import pandas as pd
import time
from os import walk
from datetime import datetime
from process import logging
from process import env
from process import pdfProcess
from process import xlsProcess
from process import file_process
from win32com.client import Dispatch


def convertToXLS(filename) :
    return
    time.sleep(3)
    ffile = str(filename).replace(".xlsx",".xls")
    # print("filename : " +filename)
    # print("ffile : " + ffile)
    # workbook = xl.load_workbook(filename)
    # workbook.save(ffile)
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Add(filename)
    wb.SaveAs(ffile, FileFormat=56)
    xl.Quit()
    time.sleep(1)
    
def convertFromXLS() :
    sourcefolder = env.SourceXLSFolder
    files = GetFileListing(sourcefolder)
    for file in files:
        time.sleep(3)
        ffile = os.path.abspath(os.getcwd()) + "\\" + env.SourceXLSFolder + "\\" + file
        fnfile = os.path.abspath(os.getcwd()) + "\\" + env.SourceFolder + "\\" + str.lower(file) + "x"
        df = pd.read_excel(ffile, header=None)
        df.to_excel(fnfile, index=False, header=False)
        # try:
        #     os.system('TASKKILL /F /IM excel.exe')
        # except Exception:
        #     print("Closeing error")
        time.sleep(3)
        print("Converted file " + ffile + " at " + str(datetime.now()))
        df = None
        ffile = None
        fnfile = None
        file_process.processedFileXLSMoving(str(file).replace("source_folder","").replace("\\",""))
        
def checkForFiles(checkingFolder) :
    print("Processing file converstion : " + str(datetime.now()))
    convertFromXLS()
    print("Processed file conversion : " + str(datetime.now()))
    time.sleep(3)
    files = GetFileListing(checkingFolder)
    for file in files:
        time.sleep(3)
        fname,ext = GetFileExtension(file)
        if (ext != ".xlsx" and ext != ".pdf") :
            logging.applicationLogProcess("New file " + file + " has invalid extension : " + fname + " " + ext)
            print("!!!!!!!!!!!!!!!!New file " + file + " has invalid extension : " + fname + " " + ext)
            file_process.processedFileMoving(file)
        else :
            logging.applicationLogProcess("Processing File " + env.SourceFolder+"\\"+file)
            print("Processing File " + env.SourceFolder+"\\"+file)
            now = str(datetime.now()).replace(" ","_").replace(":","_").replace(".","_").replace("-","_")
            if (ext == ".pdf") :
                file = str.lower(file)
                # pdfProcess.processPDFFiles(env.SourceFolder+"\\"+file,env.GeneratedFolder + "\\" + now + "_" + str.replace(file,".pdf",".xlsx"))
                if os.path.exists("..\\"+env.GeneratedFolder+"\\"+now+"_"+str.replace(file,".pdf",".xls")):
                    os.remove("..\\"+env.GeneratedFolder+"\\"+now+"_"+str.replace(file,".pdf",".xlsx"))
                try :
                    xlsxfilename = env.GeneratedFolder + "\\xlsx\\" + now + "_" + str.replace(file,".pdf",".xlsx")
                    pdfProcess.processPDFFiles(env.SourceFolder+"\\"+file,xlsxfilename)
                    convertToXLS(xlsxfilename)
                    logging.applicationLogProcess("Process completed File " + env.SourceFolder+"\\"+file)
                    print("**********************Process completed File " + env.SourceFolder+"\\"+file)
                    # processedFileMoving(file)
                except :
                    print("!!!!!!!!!!!!!!!!Processed File with error : " + file)
                    logging.applicationLogProcess("!!!!!!!!!!!!!!!!Processed File with error : " + file)
                    # corruptedFileMoving(file)
            else :
                file = str.lower(file)
                # xlsProcess.processXlsFiles(env.SourceFolder+"\\"+file,env.GeneratedFolder+"\\"+now+"_"+str.replace(file,".xlsx",".xlsx"))
                if os.path.exists("..\\"+env.GeneratedFolder+"\\"+now+"_"+str.replace(file,".xlsx",".xls")):
                    os.remove("..\\"+env.GeneratedFolder+"\\"+now+"_"+str.replace(file,".xlsx",".xlsx"))
                try :
                    xlsxFileName = env.GeneratedFolder+"\\xlsx\\"+now+"_"+str.replace(file,".xlsx",".xlsx")
                    xlsProcess.processXlsFiles(env.SourceFolder+"\\"+file,xlsxFileName)
                    convertToXLS(xlsxFileName)
                    logging.applicationLogProcess("Process completed File " + env.SourceFolder+"\\"+file)
                    print("**********************Process completed File " + env.SourceFolder+"\\"+file)
                    file_process.processedFileMoving(file)
                except :
                    print("!!!!!!!!!!!!!!!!Processed File with error : " + file)
                    logging.applicationLogProcess("!!!!!!!!!!!!!!!!Processed File with error : " + file)
                    file_process.corruptedFileMoving(file)
                    
            


def cleanContent(ContentData) :
    ContentData = str.replace(ContentData,"\n","")
    return ContentData

def GetFileExtension(fileName) :
    split_tup = os.path.splitext(fileName)
    return str.lower(split_tup[0]), str.lower(split_tup[1])


def GetFileListing(checkingFolder) :
    f = []
    for (dirpath, dirnames, filenames) in walk(checkingFolder):
        f.extend(filenames)
        break
    return f