#! /usr/bin/env python

import csv
from openpyxl import Workbook

import tkinter
from tkinter import filedialog
import os

root = tkinter.Tk()
root.withdraw() #use to hide tkinter window

def search_for_file_path ():
    currdir = os.getcwd()
    tempdir = filedialog.askopenfilename(parent=root, initialdir=currdir, title='Please select a directory')
    if len(tempdir) > 0:
        print ("You chose: %s" % tempdir)
    return tempdir



def main():
    filename="BOM.csv"
    wb=Workbook()
    ws=wb.active
    file_path_variable = search_for_file_path()
    print ("\nfile_path_variable = ", file_path_variable)
    with open(file_path_variable) as csv_file:
        csv_reader=csv.reader(csv_file, delimiter=';')
        line_count=0
        for row in csv_reader:
            if line_count==0:
                print (f'Columns are ' + ' ; '.join(row))
                line_count+=1
            else:
                codes=row[2].split('/ ')
                #print(codes)
                count=0
                for i in codes:
                    #print(codes[count])
                    ws.append([codes[count]])
                    count=count+1
                line_count+=1
        name=file_path_variable.split('/')
        count=0
        for i in name:
            final_name=name[count]
            count=count+1
        print(final_name)
        namef=final_name.split('.')
        a='_output.xlsx'
        nameOut=namef[0]+a
                #count=0
                #for i in codes:
                    #print(codes[count])
                    #ws.append([codes[count]])
                   # count=count+1
               # line_count+=1
        print(nameOut)
        wb.save(nameOut)

if __name__ == '__main__':
    main()
