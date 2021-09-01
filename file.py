#! /usr/bin/env python

import csv
from openpyxl import Workbook
def main():
    filename="BOM.csv"
    wb=Workbook()
    ws=wb.active
    with open(filename) as csv_file:
        csv_reader=csv.reader(csv_file, delimiter=';')
        line_count=0
        for row in csv_reader:
            if line_count==0:
                print (f'Columns are ' + ' ; '.join(row))
                line_count+=1
            else:
                codes=row[2].split('/ ')
                print(codes)
                count=0
                for i in codes:
                    print(codes[count])
                    ws.append([codes[count]])
                    count=count+1
                line_count+=1
                wb.save("BOM_output.xlsx")

if __name__ == '__main__':
    main()
