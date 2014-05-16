# This script allows

import sys
import re
import csv
import os 
import xlwt
import xlrd
import xlsxwriter


# Function: load_menu
#---------------------
#
#
def load_main_menu():
    menu = {}
    menu['1']="Text File" 
    menu['2']="Excel File"
    menu['3'] ="Exit Program"

    while True:
        for key, value in sorted(menu.items()):
            print(key,value)

        selection=input() 
        if selection =='1': 
            text_to_excel()
            break
        elif selection == '2': 
            excel_to_text()
            break
        elif selection == '3':
            break
        else: 
            print("Unknown Option Selected!") 


# Function: load_column_headers
#------------------------------
#
#
def load_column_headers():
    book_header = ["Author (Last, First)","Book Title",
              "Publication City", "Publisher", 
              "Year of Publication", 
              "Medium (Print or Online)"
              ]
    es_header = ["Author (Last, First)", "Article Name", 
                 "Title of Website", "Version Numbers", "Publisher",
                 "Publication Year", "Page Numbers", "Medium of Publication",
                 "Date Accessed", "URL"]
    

    headers = []
    headers.append(book_header)
    headers.append(es_header)
    return headers

# Function: intialize_excelsheet
#--------------------------------
#
#
def initialize_excelsheet(fileName):
    workbook = xlsxwriter.Workbook(fileName)
    sheets = []
    sheets.append(workbook.add_worksheet("Books"))
    sheets.append(workbook.add_worksheet("Electronic Sources"))

    #Initializes formating for header columns
    header_format = workbook.add_format()
    header_format.set_bold()
    header_format.set_bottom()
    header_format.set_align('center')


    headers = load_column_headers()
    for i in range(0,len(headers)):
        sheet = sheets[i]
        header = headers[i]
        for c in range(0,len(header)):
            sheet.write(0,c, header[c], header_format)
            if header[c] == 'Book Title':
                size_factor = 2
            elif header[c] == 'URL':
                size_factor = 8
            else:
                size_factor = 1
            sheet.set_column(c, c, size_factor*(len(header[c])+1))


# Function: text_to_excel
#-------------------------
#
#
def text_to_excel():
    textFileName = input("Enter the name of the input file: (Use .txt) \n")
    excelFileName = input('Name the output file: (Use .xlsx) \n')

    file = open(textFileName, 'r')
    line = file.read()
    words = line.split()
    for word in words:
        print(word) 


#--------------------------------------
#    workbook= xlwt.Workbook()         
#    sheet = workbook.add_sheet("test")
#    sheet.write(0, 0, line)           
#    workbook.save(excelFileName)      
#--------------------------------------

    file.close()


# Function: excel_to_text
#-------------------------
#
#
def excel_to_text():
    excelFileName = input('Name the input file: (Use .xlsx) \n')
    textFileName = input("Name the output file: (Use .txt) \n")

    initialize_excelsheet(excelFileName)
    print("\nEnter the citation information in the Excel Sheet created.")
    print("(Make sure to close & save the document when you are done!)\n")
    while True:
        answer = input("Did you finish? (y/n)")
        if (answer == 'y'):
            break
        elif (answer == 'n'):
            print("\nPlease close & save the document before continuing")
        else:
            print("\nAnswer not recognized.")

    print("Processing File...\n")
    #ADD code HERE!!
    
    textFile = open(textFileName, 'w')
    workbook= xlrd.open_workbook(excelFileName)
    worksheet = workbook.sheet_by_name('Books')

    numRows = worksheet.nrows -1
    curRow = 0
    while curRow < numRows:
        curRow +=1
        data =[]
        for i in range(0,6):
            data.append(worksheet.cell_value(curRow, i))

        textFile.write(data[0] + '. ' + data[1] + '. '
                       + data[2] + ': ' + data[3] + ', ' 
                       + str(int(data[4])) + '.\n\t' + data[5]+ '.\n\n')
    textFile.close()
    print("Process Complete.\nOpen " + textFileName + " to see citations.")


# Function: main
#-----------------
#
#
def main():
    print("Welcome to Biblio! \n")
    print("Please select an input file:")
    load_main_menu()


# Main Program Execution
main()

