#! python3

#Program automates the PDF creation of country charts

import os, openpyxl

country_codes = ["US, FR, AF"] #Need entire list from Claudine

#TODO change to working directory with excel file
os.chdir(r"C:\Users\250250\Documents\Global Risk Assessment\Script Tests")

#TODO open excel file and set variables for wb/sheets
excel_file = input("Please enter the name of the Excel File:")
work_book = openpyxl.load_workbook(os.path.abspath(excel_file)) #Need File Name - Could Prompt User to Enter File Name
threat_sheet = work_book.get_sheet_by_name('Chart-Threat') #Need exact sheet names

#TODO loop through the country codes - approx. 250 and set file to that code
for code in country_codes:
    threat_sheet[G4] = code #Need to confirm the sheet/cell that contains the country code
    work_book.save('DRAFT TEST CERA 2017 - %s - Threat Chart 2017.09.28.xlsm') % (code)
#TODO within the loop file -> save as PDF (create files based on the country code)


#TODO move PDF to specific folder
