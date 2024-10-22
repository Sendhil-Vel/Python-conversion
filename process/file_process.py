import shutil
from process import env
from process import logging

def corruptedFileMoving(fname) :
    sfile = env.SourceFolder+"\\"+fname
    cfile = env.CorruptedFolder+"\\"+fname
    shutil.move(env.SourceFolder+"\\"+fname,env.CorruptedFolder+"\\"+fname)

def processedFileMoving(fname) :
    shutil.move(env.SourceFolder+"\\"+fname,env.CompletedFolder+"\\"+fname)
    logging.applicationLogProcess(fname)
    
def processedFileXLSMoving(fname) :
    shutil.move(env.SourceXLSFolder+"\\"+fname,env.CompletedFolder+"\\"+fname)
    logging.applicationLogProcess(fname)