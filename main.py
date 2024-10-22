import sched, time
import os
from datetime import datetime
from process import logging
from process import process_verify
from process import env

def main(scheduler) :
    my_scheduler.enter(60, 1, main, (scheduler,))
    if not os.path.exists("source_folder"):
        os.makedirs("source_folder")
    if not os.path.exists("generated_folders"):
        os.makedirs("generated_folders")
    if not os.path.exists("completed_folder"):
        os.makedirs("completed_folder")
    if not os.path.exists("corrupted_folder"):
        os.makedirs("corrupted_folder")
    if not os.path.exists("logs"):
        os.makedirs("logs")
    if not os.path.exists("source_xls"):
        os.makedirs("source_xls")
    if not os.path.exists("generated_folders/xlsx"):
        os.makedirs("generated_folders/xlsx")
    if not os.path.exists("generated_folders/xls"):
        os.makedirs("generated_folders/xls")
    print("Time : " + str(datetime.now()))
    logging.applicationLogProcess("Start Process")
    process_verify.checkForFiles(env.SourceFolder)
    logging.applicationLogProcess("Completed Process")


my_scheduler = sched.scheduler(time.time, time.sleep)
my_scheduler.enter(60, 1, main, (my_scheduler,))
my_scheduler.run()

# main("a")



