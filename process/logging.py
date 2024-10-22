from datetime import datetime

def applicationLogProcess(logContent):
    file1 = open("logs/applogprocess.txt", "a")  # append mode
    file1.write(str(datetime.now()) + " : " + logContent + " \n")
    file1.close()
    
def developerLogProcess(logContent):
    file1 = open("logs/logs.txt", "a")  # append mode
    file1.write(str(datetime.now()) + " : " + logContent + " \n")
    file1.close()
    
def pp(cont) :
    print(cont)