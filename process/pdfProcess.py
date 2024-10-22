
import pdfplumber
import pandas as pd
import openpyxl
import numpy
import xlwt
from xlwt import Workbook
from datetime import datetime
from process import logging
from process import file_process
from process import env

def readFullFile(filename) :
    fileContent = []
    with pdfplumber.open(filename) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            for line in text.split('\n'):
                fileContent.append(line)
    return fileContent

def CheckRow(linedata) :
    tmplinedata = linedata.split(" ")
    if tmplinedata[2].isnumeric() != True :
        tmplinedata.insert(2,' ')
    linedata = " ".join(tmplinedata)
    return linedata

def CompanyName(line, omittingContent) :
    cname = False
    if len(omittingContent) >= 3 :
        nm = omittingContent[2].split(" ")
        if len(nm) > 0 :
            if str(line).startswith(nm[0]) == True :
                cname = True
    return cname

def processStepA (fileContent) :
    headerContent = []
    cContentData = []
    footerContent = []
    omittingContent = ['Authorized Signature','Delivery Hours:']
    pType = 1
    cnt = 0
    logging.developerLogProcess(str(fileContent))
    fromname = ""
    for line in fileContent :
        if cnt < 3 :
            omittingContent.append(line)
        if (str(line).startswith("From")) == True :
            fromname = (str(line).strip().replace("From","").replace(":","")).strip()
        cnt = cnt + 1
        if (str(line).find(" AM") > -1 or str(line).find(" PM") > -1) :
            omittingContent.append(line)
        if (line.find("Pno") > -1 and line.find("Description") > -1 and line.find("Total") > -1) :
            pType = 2
        if (line.find("Total Amount") > -1) :
            pType = 3
        if fromname != "" :
            if str(line).strip().startswith(fromname) :
                pType = 4
        if CompanyName(line, omittingContent) == True :
            pType = 1
        
        if pType == 1 :
            headerContent.append(line)
        if pType == 2 :
            arrData = str(line).split(" ")
            # print(arrData)
            # print(len(arrData))
            # if len(arrData) > 1 :
            #     if arrData[0].isnumeric() == True and arrData[1].isnumeric() == True :
            #         # and (arrData[2].isnumeric() == True or len(arrData[2].strip()) == 0) :
            #         # and arrData[len(arrData)-1].isnumeric() == True:
            #         tmpline = CheckRow(line)
            #         cContentData.append(tmpline)
            #     else :
            #         if arrData[0] != "Pno" :
            #             if len(cContentData) >= 1 :
            #                 sarr = str(cContentData[len(cContentData)-1]).split(" ")
            #                 sarr[len(sarr)-5] = sarr[len(sarr)-5] + " " +  line 
            #                 cContentData[len(cContentData)-1] = " ".join(sarr)
            # else :
            #         # print(len(cContentData))
            #         # print(arrData)
            #         if (line not in omittingContent and line not in headerContent and str(line).find("Page") == -1) :                     
            #             if len(cContentData) >= 1 :
            #                 sarr = str(cContentData[len(cContentData)-1]).split(" ")
            #                 sarr[len(sarr)-5] = sarr[len(sarr)-5] + " " +  line 
            #                 cContentData[len(cContentData)-1] = " ".join(sarr)
            
            if len(arrData) > 2 and arrData[0].isnumeric() == True and arrData[1].isnumeric() == True :
                # print("arrData : " + str(arrData))
                tmpline = CheckRow(line)
                cContentData.append(tmpline)
            else :
                # print("arrData a : " + str(arrData))
                if arrData[0] != "Pno" :
                    # print("arrData b : " + str(arrData))
                    # print("omittingContent : "+str(omittingContent))
                    # print("headerContent : "+ str(headerContent))
                    # print("line : " + str(line))
                    # print(str(line).find("Page"))
                    # print(line not in omittingContent)
                    # print(line not in headerContent)
                    if (line not in omittingContent and line not in headerContent and str(line).find("Page") == -1) :
                        # print("arrData c : " + str(arrData))
                        if len(cContentData) >= 1 :
                            # print("arrData d : " + str(arrData))
                            sarr = str(cContentData[len(cContentData)-1]).split(" ")
                            sarr[len(sarr)-5] = sarr[len(sarr)-5] + " " +  line 
                            cContentData[len(cContentData)-1] = " ".join(sarr)     
        if pType == 3 :
            footerContent.append(line)
    # print("-------------------")
    # print(str(footerContent))
    return headerContent, cContentData, footerContent

def CleanData(contentData,keyData) :
    contentData = str(contentData).replace(keyData,"").replace(":","").strip()    
    return contentData

def processHeaderContent(headContent) :
    headerListData = []
    idx = CleanData(CleanData(next(c for c in headContent if str(c).startswith("Supplier") == True),"Supplier"),"PURCHASE ORDER") 
    headerListData.append(idx)
    
    idx = CleanData(next(c for c in headContent if str(c).startswith("Purchaser") == True),"Purchaser") 
    headerListData.append(idx)

    idx = CleanData(next(c for c in headContent if str(c).startswith("Marking") == True),"Marking") 
    headerListData.append(idx)

    headerListData.append(CleanData(headContent[1], "Test") + " " + CleanData(headContent[2],"Test"))

    idx = CleanData(next(c for c in headContent if str(c).startswith("Date") == True),"Date") 
    headerListData.append(idx)

    idx = CleanData(next(c for c in headContent if str(c).startswith("Delivery Date") == True),"Delivery Date") 
    headerListData.append(idx)

    idx = CleanData(next(c for c in headContent if str(c).startswith("Currency") == True),"Currency") 
    headerListData.append(idx)

    idx = CleanData(next(c for c in headContent if str(c).startswith("Terms") == True),"Terms") 
    headerListData.append(idx)

    idx = CleanData(next(c for c in headContent if str(c).startswith("Our Ref") == True),"Our Ref") 
    headerListData.append(idx)
    
    return headerListData

def processCContent(cContentData) :
    cContentListData = ['','','','','','','','']
    cMasterContent = []
    cnt = 1
    for line in cContentData :
        cContentListData = ['','','','','','','','']
        cont = str(line).split(" ")
        # print(line)
        # print(cont)
        if cont[0] != "Pno" :
            # print("a")
            if ((str(cont[0])).strip()).isnumeric() == True and ((str(cont[0])).strip()).isnumeric() == cnt and len(cont) >= 10:
                # print("c")
                # cnt = cnt + 1
                contSize = len(cont)
                cContentListData[0] = cont[0]  
                cContentListData[1] = cont[1]
                cContentListData[2] = cont[2]
                cContentListData[4] = cont[contSize-4]
                cContentListData[5] = cont[contSize-3]
                cContentListData[6] = cont[contSize-2]
                cContentListData[7] = cont[contSize-1]
                cont.pop(contSize-1)
                cont.pop(contSize-2)
                cont.pop(contSize-3)
                cont.pop(contSize-4)
                cont.pop(2)
                cont.pop(1)
                cont.pop(0)
                cContentListData[3] = " ".join(cont)
                cMasterContent.append(cContentListData)
            else :
                # print("d")
                cMasterContent[len(cMasterContent)-1][3] = cMasterContent[len(cMasterContent)-1][3] + " " + " ".join(cont)
        else :
            logging.applicationLogProcess("d")
    return cMasterContent

def processCContentA(cContentData) :
    cContentListData = ['','','','','','','','']
    cMasterContent = []
    for line in cContentData :
        if len(line) > 1 :
            cContentListData = ['','','','','','','','']
            cont = str(line).split(" ")
            contSize = len(cont)
            cContentListData[0] = cont[0]  
            cContentListData[1] = cont[1]
            cContentListData[2] = cont[2]
            cContentListData[4] = cont[contSize-4]
            cContentListData[5] = cont[contSize-3]
            cContentListData[6] = cont[contSize-2]
            cContentListData[7] = cont[contSize-1]
            cont.pop(contSize-1)
            cont.pop(contSize-2)
            cont.pop(contSize-3)
            cont.pop(contSize-4)
            cont.pop(2)
            cont.pop(1)
            cont.pop(0)
            cContentListData[3] = " ".join(cont)
            cMasterContent.append(cContentListData)
    return cMasterContent

def processFooterContent(footerContent,HeaderContent) :
    footerListData = ['','','','','']
    footerListData[1] = ProcessRemarks(footerContent,HeaderContent)
    for line in footerContent :
        tqty = str(line).find("Total Qty")
        if tqty > -1 :
            presub = (str(line))[:tqty]
            subline = (str(line)).replace(presub,"").replace("Total Qty","").replace(":","").strip()
            footerListData[0] = subline
            
        presub = ""
        subline = ""                
        tamt = str(line).find("Total Amount")
        if tamt > -1 :
            presub = (str(line))[:tamt]
            subline = (str(line)).replace(presub,"").replace("Total Amount","").replace(":","").strip()
            footerListData[2] = subline
            
        presub = ""
        subline = ""  
        gst = str(line).find("Add GST 9 %")
        if gst > -1 :
            presub = (str(line))[:gst]
            subline = (str(line)).replace(presub,"").replace("Add GST 9 %","").replace(":","").strip()
            footerListData[3] = subline
            
        presub = ""
        subline = ""  
        adue = str(line).find("Amount Due")
        if adue > -1 :
            presub = (str(line))[:adue]
            subline = (str(line)).replace(presub,"").replace("Amount Due","").replace(":","").strip()
            footerListData[4] = subline
            
    return footerListData
        
def processFooterContentA(footerContent,HeaderContent) :
    footerListData = []
    idx = CleanData(next(c for c in footerContent if str(c).startswith("Total Qty") == True),"Total Qty") 
    footerListData.append(idx)
    footerListData.append(ProcessRemarks(footerContent,HeaderContent))
    for line in footerContent :
        f = str(line).find("Total Amount")
        if f > -1 :
            presub = (str(line))[:f]
            subline = (str(line)).replace(presub,"").replace("Total Amount","").strip()
            footerListData.append(subline)
    idx = CleanData(next(c for c in footerContent if str(c).startswith("Add GST 9 %") == True),"Add GST 9 %") 
    footerListData.append(idx)
    idx = CleanData(next(c for c in footerContent if str(c).startswith("Amount Due") == True),"Amount Due") 
    footerListData.append(idx)
    return footerListData

def ProcessRemarks(footerContent,headContent):
    fromname = CleanData(CleanData(next(c for c in headContent if str(c).startswith("From") == True),"From"),"From") 
    cnt = -1
    remark = ""
    processflag = 0
    for fdata in footerContent :
        cnt = cnt +1
        if fdata == headContent[0] :
            processflag =1
        if str(fdata).startswith(fromname) == True :
            processflag = 1
        if str(fdata).startswith(fromname+":") == True :
            processflag = 0    
        if str(fdata).startswith("Remarks") == True :
            remark = ""
            processflag = 0
        if cnt >= 5 :
            if processflag == 0 :
                remark = remark + " " + fdata
    remark = CleanData(remark,"Remarks")
    remark = CleanData(remark,"\"")            
    return remark
    
def processFinalContent(headerListData, cContentListData, footerListData, filename) :
    print("***********Generating File " + env.SourceFolder+"\\" + filename)
    finalist = []
    tmpFst = []
    tmpFst = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22]
    finalist.append(tmpFst)
    for data in cContentListData :
        tmpFst = []
        for hData in headerListData :
            tmpFst.append(str(hData).replace("\'",""))
        for dData in data :
            tmpFst.append(str(dData).replace("\'",""))
        for fData in footerListData :
            tmpFst.append(str(fData).replace("\'",""))
        finalist.append(tmpFst) 
    fileWrite(finalist,filename)
    filename = str.replace(filename,".xlsx",".xls")
    filename = str.replace(filename,"\\xlsx\\","\\xls\\")
    print("***********Generating File " + env.SourceFolder+"\\" + filename)
    fileWritexls(finalist,filename)

def fileWritexls(finalDataObj,filename) :
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    rowcnt = -1
    colcnt = -1
    for rowdata in finalDataObj:
        rowcnt = rowcnt + 1
        colcnt = -1
        for coldata in rowdata :
            colcnt = colcnt  + 1
            sheet1.write(rowcnt,colcnt,coldata)
    wb.save(filename)
        
def fileWrite(finalDataObj,filename) :
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in finalDataObj:
        sheet.append(row)
    workbook.save(filename)
                                        
def processPDFFiles(file,xlsFile) :
    fileContent = readFullFile(file)
    if len(fileContent) < 10 :
        print("*****************************Processed File with error : " + file)
        logging.applicationLogProcess("*****************************Processed File with error : " + file)
        file_process.corruptedFileMoving(str(file).replace("source_folder","").replace("\\",""))
    else :   
        headerContent, cContentData, footerContent = processStepA(fileContent)
        logging.developerLogProcess("-----------------------------")
        logging.developerLogProcess(str(headerContent))
        logging.developerLogProcess(str(cContentData)) 
        logging.developerLogProcess(str(footerContent))
        logging.developerLogProcess("-----------------------------")
        headerListData = processHeaderContent(headerContent)
        cContentListData = processCContent(cContentData)
        footerListData = processFooterContent(footerContent,headerContent)
        now = str(datetime.now()).replace(" ","_").replace(":","_").replace(".","_").replace("-","_")
        processFinalContent(headerListData, cContentListData, footerListData, xlsFile)
        logging.applicationLogProcess("Process completed File " + env.SourceFolder+"\\"+file)
        print("Process completed File " + env.SourceFolder+"\\"+file)
        file_process.processedFileMoving(str(file).replace("source_folder","").replace("\\",""))
    