#!/usr/bin/python

# Author: Clara Watson
# Date: August, 2022
# Description:  this program takes csv files from one directory and 
# copies them to a new directory and changes them to xlsx files
 

import glob
import openpyxl
import csv
import os

def get():
    """
    this function finds all the csv files in my Local Directory 
    and creates a blank xlsx file in a new directory 
    """
    files = glob.glob("/Users/Clara/Documents/Reports_Local/*")
    # loops through all files in the directory
    for file in files:
        with open(file, 'rt') as f:
            # splits the path string to get file name 
            end = file.split("\\")[1]
            # removes the .csv from the file name 
            end = end.split('.')[0] 
            wb = openpyxl.Workbook() 
            excel_file = "/Users/Clara/Documents/Reports/" +   end + ".xlsx"
            print("excel file " + excel_file)
            wb.save(excel_file)
            csv_to_excel(file, excel_file)


def csv_to_excel(csv_file, excel_file):
    """
    this function takes the csv file and 
    copies the data to the new xlsx file 
    """
    csv_data = []
    with open(csv_file) as file_obj:
        # reads the csv file 
        reader = csv.reader(file_obj)
        reader = csv.reader( (line.replace('\0','') for line in file_obj) )
        # saves data in each row to array
        for row in reader:
            csv_data.append(row)
    # creates a new excel workbook 
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    # populates the workbook from array 
    for row in csv_data:
        sheet.append(row)
    # saves the workbook as xlsx file 
    workbook.save(excel_file)

def delete():
    """
    this file deletes all the csv files from their directory
    """
    files = glob.glob('/Users/Clara/Documents/Reports_Local/*')
    for f in files:
        os.remove(f)
    

if __name__ == "__main__":
    get()
