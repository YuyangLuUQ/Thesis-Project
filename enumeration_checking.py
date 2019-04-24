import xlrd
import xlwt
import re
import numpy as np
import pandas as pd

max_column = 80 # max number of 

"""
Check if the string is a float number or not
parameter(s):
    string: the chosen string will be checked if it is a float number or not
"""
def is_float(string):
    try:
        float(string)
        return True
    except ValueError:
        return False

"""
get the number of rows of the sheet page
parameter(s):
    sheet: the chosen sheet page of the chosen excel file
"""
def get_row(sheet):
    return sheet.nrows

"""
get the number of columns of the sheet page
parameter(s):
    sheet: the chosen sheet page of the chosen excel file
"""
def get_column(sheet):
    return sheet.ncols

"""
get error result of the column of "Transaction Type"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
"""
def check_transaction_type(row, column, sheet):
    valiable_list = ['N', 'A', 'D'] # the list of available result for "Transaction Type"
    type_column = 0
    no_error = True # if there is any error found, this variable will be set to False
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Transaction Type'\n")
            return
        elif sheet.cell(0, type_column).value == "Transaction Type":
            break
        type_column += 1
    i = 1
    print("Result for 'Transaction Type':")
    while i < row:
        correct = 0
        for letter in valiable_list:
            if sheet.cell(i, type_column).value == letter:
                correct += 1 # the value of correct of every available result should be greater than 0
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Transaction Type'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Transaction Type'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Transaction Type'")
    print('\n')
    
"""
get error result of the column of "Mothers UR number"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
"""
def check_mothers_ur_number(row, column, sheet):
    type_column = 0
    no_error = True
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Mothers UR number'\n")
            return
        elif sheet.cell(0, type_column).value == "Mothers UR number":
            break
        type_column += 1
    i = 1
    print("Result for 'Mothers UR number':")
    while i < row:
        correct = 0
        if sheet.cell(i, type_column).value.isdigit():
            correct += 1
        if len(sheet.cell(i, type_column).value) <= 8 and len(sheet.cell(i, type_column).value) > 0:
            correct += 1
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct < 2 and correct >= 0: # every available result should get a "2" for "correct" after the checking of value type and length
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Mothers UR number'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Mothers UR number'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Mothers UR number'")
    print('\n')

"""
get error result of the column of "Date of confinement"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
"""
def check_Date_of_confinement(row, column, sheet):
    type_column = 0
    no_error = True
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Date of Confinement'\n")
            return
        elif sheet.cell(0, type_column).value == "Date of confinement" or sheet.cell(0, type_column).value == "Date of Confinement":
            break
        type_column += 1
    i = 1
    print("Result for 'Date of confinement':")
    while i < row:
        correct = 0
        sepa_list = sheet.cell(i, type_column).value.split('/')
        
        if sepa_list[2].isdigit():
            correct += 1
        if len(sepa_list[2]) == 4:
            correct += 1
        if int(sepa_list[2]) <= 2019:
            correct += 1

        if sepa_list[1].isdigit():
            correct += 1
        if len(sepa_list[1]) == 2:
            correct += 1
        if int(sepa_list[1]) > 0 and int(sepa_list[1]) <= 12:
            correct += 1
            
        if sepa_list[0].isdigit():
            correct += 1
        if len(sepa_list[0]) == 2:
            correct += 1
        if sepa_list[1] == '01' or sepa_list[1] == '03' or sepa_list[1] == '05' or sepa_list[1] == '07' or sepa_list[1] == '08' or sepa_list[1] == '10' or sepa_list[1] == '12':
            if int(sepa_list[0]) > 0 and int(sepa_list[0]) <= 31:
                correct += 1
        elif sepa_list[1] == '04' or sepa_list[1] == '06' or sepa_list[1] == '09' or sepa_list[1] == '11':
            if int(sepa_list[0]) > 0 and sint(sepa_list[0]) <= 30:
                correct += 1
        elif sepa_list[1] == '02':
            if int(sepa_list[0]) > 0 and int(sepa_list[0]) <= 28:
                correct += 1
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct < 9 and correct >= 0: # every available result should get a "9" for "correct" after the checking of value type, length and size of day, month and year
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Date of confinement'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Date of confinement'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Date of confinement'")
    print('\n')

"""
get error result of the column of "Code Type"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
"""
def check_code_type(row, column, sheet):
    code_list = []
    valiable_list = ['C', 'T', 'M', 'P', 'O', 'L', 'A', 'E'] # the list of available result for "Code Type"
    type_column = 0
    no_error = True
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Code Type'\n")
            return
        elif sheet.cell(0, type_column).value == "Code Type":
            break
        type_column += 1
    i = 1
    print("Result for 'Code Type':")
    while i < row:
        correct = 0
        code_list.append(sheet.cell(i, type_column).value)
        for letter in valiable_list:
            if sheet.cell(i, type_column).value == letter:
                correct += 1
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Code Type'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Code Type'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Code Type'")
    print('\n')
    return code_list

"""
get error result of the column of "Mothers code"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
    code_list: the list contains all the values of column 'Code Type'
"""
def check_mothers_code(row, column, sheet, code_list):
    type_column = 0
    no_error = True
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Mothers code'\n")
            return
        elif sheet.cell(0, type_column).value == "Mothers code":
            break
        type_column += 1
    i = 1
    print("Result for 'Mothers code':")
    while i < row:
        correct = 0
        if code_list[i - 1] == "T" or code_list[i - 1] == "M" or code_list[i - 1] == "P":
            if len(sheet.cell(i, type_column).value) <= 5:
                correct += 1
        elif code_list[i - 1] == "O":
            if len(sheet.cell(i, type_column).value) <= 7:
                correct += 1
        elif code_list[i - 1] == "C":
            ava_list = ['02', '03', '04', '05', '06', '07', '08', '09', '19', '99'] # the list of available result for type "C"
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value == ava_list[j]:
                    correct += 1
                    break
                j += 1
        elif code_list[i - 1] == "L":
            ava_list = ['02', '03', '04', '05', '10', '98', '99'] # the list of available result for type "L"
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value == ava_list[j]:
                    correct += 1
                    break
                j += 1
        elif code_list[i - 1] == "A":
            ava_list = ['06', '07', '08', '03', '04', '99'] # the list of available result for type "A"
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value == ava_list[j]:
                    correct += 1
                    break
                j += 1
        elif code_list[i - 1] == "E":
            ava_list = ['AT', 'MC', 'PC', 'PO'] # the list of available result for type "E"
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value[0:2] == ava_list[j]:
                    correct += 1
                    break
                j += 1
            if len(sheet.cell(i, type_column).value) > 120: # the available length is smaller or equal to 120
                correct -= 1
        else:
            correct = -2
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Mothers code' based on its 'Code Type': " + str(code_list[i - 1]) + ".")
        elif correct == -2:
            no_error = False
            print("The sample in row " + str(i + 1) + " cannot be checked, because it does not have a valid 'Code Type'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Mothers code'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Mothers code'")
    print('\n')

"""
get error result of the column of "Mothers country of birth"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
"""
def check_mothers_country_of_birth(row, column, sheet):
    type_column = 0
    no_error = True
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Mothers country of birth'\n")
            return
        elif sheet.cell(0, type_column).value == "Mothers country of birth":
            break
        type_column += 1
    i = 1
    print("Result for 'Mothers country of birth':")
    while i < row:
        correct = 0
        if len(sheet.cell(i, type_column).value) <= 4 and len(sheet.cell(i, type_column).value) > 0:
            correct += 1
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct > 0 and correct != 1:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Mothers country of birth'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Mothers country of birth'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Mothers UR number'")
    print('\n')

"""
get error result of the column of "Marital status"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
"""
def check_marital_status(row, column, sheet):
    valiable_list = ['1', '2', '3', '4', '5', '9']
    type_column = 0
    no_error = True
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Marital status'\n")
            return
        elif sheet.cell(0, type_column).value == "Marital status":
            break
        type_column += 1
    i = 1
    print("Result for 'Marital status':")
    while i < row:
        correct = 0
        for letter in valiable_list:
            if sheet.cell(i, type_column).value == letter:
                correct += 1
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Marital status'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Marital status'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Marital status'")
    print('\n')


"""
get error result of the column of "Marital status"
parameter(s):
    row: the number of rows of the sheet page
    column: the number of columns of the sheet page
    sheet: the chosen sheet page of the chosen excel file
"""
def check_indigenous_status_mother(row, column, sheet):
    valiable_list = ['1', '2', '3', '4', '9']
    type_column = 0
    no_error = True
    while True:
        if type_column >= max_column:
            print("Cannot find column type 'Indigenous status (Mother)'\n")
            return
        elif sheet.cell(0, type_column).value == "Indigenous status (Mother)":
            break
        type_column += 1
    i = 1
    print("Result for 'MIndigenous status (Mother)':")
    while i < row:
        correct = 0
        for letter in valiable_list:
            if sheet.cell(i, type_column).value == letter:
                correct += 1
        if sheet.cell(i, type_column).value == "":
            correct = -1
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'MIndigenous status (Mother)'.")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " is blank, which is not avaliable for 'Indigenous status (Mother)'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Marital Indigenous status (Mother)'")
    print('\n')

def interact():
    while(True):
        direction = input('Please insert the direction of the file you want to open: ')
        try:
            f = open(r"%s"%(direction))
            break
        except Exception:
            print("The file cannot be opened or not exist\n")
        
    fileName = xlrd.open_workbook(r"%s"%(direction))
    index = input('Please choose sheet page number you want(start from 1): ')
    sheet = fileName.sheet_by_index(int(index) - 1)
    row = get_row(sheet)
    column = get_column(sheet)
    
    check_transaction_type(row, column, sheet)
    check_mothers_ur_number(row, column, sheet)
    check_Date_of_confinement(row, column, sheet)
    check_mothers_country_of_birth(row, column, sheet)
    check_marital_status(row, column, sheet)
    check_indigenous_status_mother(row, column, sheet)
    code_list = check_code_type(row, column, sheet)
    check_mothers_code(row, column, sheet, code_list)

if __name__ == '__main__':
    interact()
