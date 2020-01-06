import win32com.client as win32
import argparse
import os
import sys
from datetime import datetime

parser = argparse.ArgumentParser(description='Converts all .xls files in a directory and subdirs to .xlsx.')
parser.add_argument('-i', '--inputdir', help='Parent directory to start in. If at least one directory is not specified\
                        the program prints help.', required=True)

# if no arguments are provided to the program print help
if len(sys.argv) == 1:
    parser.print_help()
    sys.exit(1)
args = parser.parse_args()

# create output and error log file in current working directory
logfile = os.path.join(os.getcwd(), 'xlsxsaver_output_' + datetime.now().strftime("%Y%m%d-%H%M%S") + '.log')
errfile = os.path.join(os.getcwd(), 'xlsxsaver_error_' + datetime.now().strftime("%Y%m%d-%H%M%S") + '.log')

try:
    d = open(logfile, 'w+')
    err = open(errfile, 'w+')
    print ("Searching for .xls files in " + args.inputdir + " and subdirectories.\n")
    d.write ("Searching for .xls files in " + args.inputdir + " and subdirectories.\n")

    for dirName, subdirList, fileList in os.walk(args.inputdir):
        print ("Searching for .xls files in " + dirName + "\n")
        d.write ("Searching for .xls files in " + dirName + "\n")
        for workfile in fileList:
            # make sure we're opening a .xlsx file
            if workfile.endswith('.xls'):
                try:
                    filepath = os.path.join (dirName, workfile)
                    savefilepath = filepath + 'x'
                    # p.save_book_as(file_name=filepath, dest_file_name=savefilepath)
                    excel = win32.gencache.EnsureDispatch('Excel.Application')
                    wb = excel.Workbooks.Open(filepath)
                    wb.SaveAs(savefilepath, FileFormat = 51)
                    excel.Application.Quit()
                    print ('Converted ' + filepath + ' to ' + savefilepath)
                    d.write ('Converted ' + filepath + ' to ' + savefilepath)
                except:
                    print ('Cannot convert ' + filepath + ' to .xlsx')
                    err.write ('Cannot convert ' + filepath + ' to .xlsx')

#  can't open output file due to permissions issues, file locks, etc
except IOError:
    print ('Error opening log and or output file.')