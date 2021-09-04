#! /usr/bin/env python

import csv
from openpyxl import Workbook

import tkinter
from tkinter import filedialog
import os

root = tkinter.Tk()
root.withdraw() #use to hide tkinter window

## function used to search in the file path to obtain the file
def search_for_file_path ():
    currdir = os.getcwd()
    tempdir = filedialog.askopenfilename(parent=root, initialdir=currdir, title='Please select a directory')
    if len(tempdir) > 0:
        print ("You chose: %s" % tempdir)
    return tempdir



def main():
    wb=Workbook()
    ws=wb.active
    #search for the file
    file_path_variable = search_for_file_path()
    print ("\nfile_path_variable = ", file_path_variable)
    # open the file with the csv library
    with open(file_path_variable) as csv_file:
        csv_reader=csv.reader(csv_file, delimiter=';')
        line_count=0
        #read in every row of the file
        for row in csv_reader:
			# if it is the first row print what are the rows
            if line_count==0:
                print (f'Columns are ' + ' ; '.join(row))
                line_count+=1
            # for all other rows, in the third column I separate the elements with the delimitator "/ "
            # I save the separated items in a list (codes)    
            else:
                codes=row[2].split('/ ')
                print(codes)
                count=0
                # for every item in the list I add them to the excell file using the library openpyxl
                for i in codes:
                    print(codes[count])
                    ws.append([codes[count]])
                    count=count+1
                line_count+=1
        # I split the path with delimitator "/"        
        name=file_path_variable.split('/')
        count=0
        # I take the last element of the path, the name of the file with its extension
        for i in name:
            final_name=name[count]
            count=count+1
        print(final_name)
        # I remove the extension by splitting with the "."
        namef=final_name.split('.')
        a='_output.xlsx'
        # I create the final name for the excell file as follows: NameRetrieved_output.xlsx
        nameOut=namef[0]+a  
        print(nameOut)
        # save the excell file
        wb.save(nameOut)

if __name__ == '__main__':
    main()
