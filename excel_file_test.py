import xlrd
import xlwt
import re

def is_float(string):
    try:
        float(string)
        return True
    except ValueError:
        return False

def get_row(sheet):
    return sheet.nrows

def get_column(sheet):
    return sheet.ncols

def show_basic_info(row, column, sheet):
    print('Idealy:')
    print('This dataset should have %d groups of data (excluded title row) and each has %d attributes and they are: '%(row - 1, column), end="")
    first_row = sheet.row_values(0)
    for i in first_row:
        print("%s, "%(i), end="")
    print('\n')

def check_value_type(row, column, sheet):
    print('Checking for value types....')
    cellNum = 0
    blankNum = 0
    digitNum = 0
    floatNum = 0
    i = 1
    while (i < row):
        j = 1
        while (j < column):
            if sheet.cell(i, j).value == "":
                blankNum += 1
            else:
                if str(sheet.cell(i, j).value).isdigit():
                    digitNum += 1
                if is_float(str(sheet.cell(i, j).value)):
                    floatNum += 1
            cellNum += 1
            j += 1
        i += 1
    if floatNum > digitNum:
        print("The most usual value type is DECIMAL NUMBER, it takes up %f percent of the whole dataset with blanks, %f percent without blanks\n"%(floatNum / cellNum * 100, floatNum / (cellNum - blankNum) * 100))
    else:
        print("The most usual value type is INTEGER NUMBER, it takes up %f percent of the whole dataset with blanks, %f percent without blanks\n"%(digitNum / cellNum * 100, digitNum / (cellNum - blankNum) * 100))

def check_blank(row, column, sheet):
    empty_list = []
    i = 1
    while (i < row):
        j = 1
        while (j < column):
            if sheet.cell(i, j).value == "":
                empty_list.append("(" + str(i + 1) + "," + str(j + 1) + ")")
            j += 1
        i += 1
    if len(empty_list) > 0:
        print("The following cells are currently empty")
        print(empty_list)
        print("{PS: for the columns, the number \"1, 2, 3\" stand for \"A, B, C\"}\n")
    else:
        print("no error found")

def check_null(row, column, sheet):
    null_list = []
    i = 1
    while (i < row):
        j = 1
        while (j < column):
            if sheet.cell(i, j).value == "(Null)":
                null_list.append("(" + str(i + 1) + "," + str(j + 1) + ")")
            j += 1
        i += 1
    if len(null_list) > 0:
        print("The following cells are currently 'null'")
        print(null_list)
        print("{PS: for the columns, the number \"1, 2, 3\" stand for \"A, B, C\"}\n")
    else:
        print("no error found")

def check_date(row, column, sheet, position, babyormother):
    noerror = True
    i = 1
    while (i < row):
        if sheet.cell(i, int(position) - 1).value == "Null":
            print("(" + str(i + 1) + "," + str(int(position)) + ") " + "has some problems, its null")
            i += 1
            noerror = False
            continue
        date_list = re.split(' |/|:|-', sheet.cell(i, int(position) - 1).value)
        if int(date_list[1]) < 1 or int(date_list[1]) > 12:
            print("(" + str(i + 1) + "," + str(int(position)) + ") " + "has some problems, its month is " + str(date_list[1]) + " , which is impossible")
            i += 1
            noerror = False
            continue
        if int(date_list[1]) == 1 or int(date_list[1]) == 3 or int(date_list[1]) == 5 or int(date_list[1]) == 7 or int(date_list[1]) == 8 or int(date_list[1]) == 10 or int(date_list[1]) == 12:
            if int(date_list[0]) < 1 or int(date_list[0]) > 31:
                print("(" + str(i + 1) + "," + str(int(position)) + ") " + "has some problems, its date is " + str(date_list[0]) + " , which is impossible")
                i += 1
                noerror = False
                continue
        if int(date_list[1]) == 4 or int(date_list[1]) == 6 or int(date_list[1]) == 9 or int(date_list[1]) == 11:
            if int(date_list[0]) < 1 or int(date_list[0]) > 30:
                print("(" + str(i + 1) + "," + str(int(position)) + ") " + "has some problems, its date is " + str(date_list[0]) + " , which is impossible")
                i += 1
                noerror = False
                continue
        if int(date_list[1]) == 2:
            if int(date_list[0]) < 1 or int(date_list[0]) > 28:
                print("(" + str(i + 1) + "," + str(int(position)) + ") " + "has some problems, its date is " + str(date_list[0]) + " , which is impossible")
                i += 1
                noerror = False
                continue
        if babyormother == 'mother':
            if int(date_list[2]) < 1970:
                print("(" + str(i + 1) + "," + str(int(position)) + ") " + "has some problems, its year is " + str(date_list[2]) + " , which is impossible")
                i += 1
                noerror = False
                continue
        if babyormother == 'baby':
            if int(date_list[2]) != 2019:
                print("(" + str(i + 1) + "," + str(int(position)) + ") " + "has some problems, its year is " + str(date_list[2]) + " , which is impossible")
                i += 1
                noerror = False
                continue
        i += 1    
    print("{PS: for the columns, the number \"1, 2, 3\" stand for \"A, B, C\"}\n")
    if noerror == True:
        print("No error found")

def check_dup(row, column, sheet, position):
    valueList = []
    i = 1
    while (i < row):
        valueList.append(sheet.cell(i, int(position) - 1).value)
        i += 1
    valueList2 = valueList
    i = 0
    dupList = []
    while (i < len(valueList)):
        found = False
        j = i + 1
        while (j <= len(valueList)):
            if (sheet.cell(i, int(position) - 1).value == sheet.cell(j, int(position) - 1).value):
                dupList.append(sheet.cell(i, int(position) - 1).value)
                found == True
                break
            else:
                j += 1
        if (found == False):
            i += 1
        else:
            for item in valueList:
                if (item == dupList[-1]):
                    valueList.remove(item)
        found = False
    if len(dupList) != 0:
        i = 0
        while (i < len(dupList)):
            posList = []
            j = 0
            while (j < len(valueList2)):
                if (valueList2[j] == dupList[i]):
                    posList.append(j + 2)
                j += 1
            print("the value " + str(dupList[i]) + " is duplicated, it is find in row " + str(posList))
            i += 1
        print("{PS: for the columns, the number \"1, 2, 3\" stand for \"A, B, C\"}\n")
    else:
        print("no error found")

def check_transaction_type(row, column, sheet):
    valiable_list = ['N', 'A', 'D']
    type_column = 0
    no_error = True
    while type_column < column:
        if sheet.cell(0, type_column).value == "Transaction Type":
            break
        type_column += 1
    i = 1
    print("Result for 'Transaction Type':")
    while i < row:
        correct = 0
        for letter in valiable_list:
            if sheet.cell(i, type_column).value == letter:
                correct += 1
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Transaction Type'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Transaction Type'")
    print('\n')

def check_mothers_ur_number(row, column, sheet):
    type_column = 0
    no_error = True
    while type_column < column:
        if sheet.cell(0, type_column).value == "Mothers UR number":
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
        if correct < 2:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Mothers UR number'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Mothers UR number'")
    print('\n')

def check_Date_of_confinement(row, column, sheet):
    type_column = 0
    no_error = True
    while type_column < column:
        if sheet.cell(0, type_column).value == "Date of confinement":
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
            
        if correct < 9:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Date of confinement'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Date of confinement'")
    print('\n')

def check_code_type(row, column, sheet):
    code_list = []
    valiable_list = ['C', 'T', 'M', 'P', 'O', 'L', 'A', 'E']
    type_column = 0
    no_error = True
    while type_column < column:
        if sheet.cell(0, type_column).value == "Code Type":
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
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Code Type'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Code Type'")
    print('\n')
    return code_list

def check_mothers_code(row, column, sheet, code_list):
    type_column = 0
    no_error = True
    while type_column < column:
        if sheet.cell(0, type_column).value == "Mothers code":
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
            ava_list = ['02', '03', '04', '05', '06', '07', '08', '09', '19', '99']
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value == ava_list[j]:
                    correct += 1
                    break
                j += 1
        elif code_list[i - 1] == "L":
            ava_list = ['02', '03', '04', '05', '10', '98', '99']
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value == ava_list[j]:
                    correct += 1
                    break
                j += 1
        elif code_list[i - 1] == "A":
            ava_list = ['06', '07', '08', '03', '04', '99']
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value == ava_list[j]:
                    correct += 1
                    break
                j += 1
        elif code_list[i - 1] == "E":
            ava_list = ['AT', 'MC', 'PC', 'PO']
            j = 0
            while j < len(ava_list):
                if sheet.cell(i, type_column).value[0:2] == ava_list[j]:
                    correct += 1
                    break
                j += 1
            if len(sheet.cell(i, type_column).value) > 120:
                correct -= 1
        else:
            correct = -1
        if correct == 0:
            no_error = False
            print("The sample in row " + str(i + 1) + " does not have a valiable value for 'Mothers code' based on its 'Code Type': " + str(code_list[i - 1]) + ".")
        elif correct == -1:
            no_error = False
            print("The sample in row " + str(i + 1) + " cannot be checked, because it does not have a valid 'Code Type'.")
        i += 1
    if no_error == True:
        print("No error found in column 'Mothers code'")
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
    
    while(True):
        purpose = input('What would you like to know? Please choose from [basic_info / check_value_types / consistency / check_enumeration]: ')
        if(purpose == "basic_info"):
            show_basic_info(row, column, sheet)
        elif(purpose == "check_value_types"):
            check_value_type(row, column, sheet)
        elif(purpose == "consistency"):
            method = input("which type of error would you like to check? Please choose from [uniqueness / blank / null / date]: ")
            if(method == "uniqueness"):
                position = input("please select the column number(start from 1): ")
                check_dup(row, column, sheet, position)
            elif(method == "blank"):
                check_blank(row, column, sheet)
            elif(method == "null"):
                check_null(row, column, sheet)
            elif(method == "date"):
                position = input("please select the column number(start from 1): ")
                babyormother = input("please select baby or mother(baby / mother): ")
                check_date(row, column, sheet, position, babyormother)
        elif(purpose == "check_enumeration"):
            print("""
#################################################################################################
#                                                                                               #
# Attention: this functionality only checks the cell values based on 'Format' and 'description' #
# Any other validations will be checked in other functionalities                                #
#                                                                                               #
#################################################################################################
""")
            check_transaction_type(row, column, sheet)
            check_mothers_ur_number(row, column, sheet)
            check_Date_of_confinement(row, column, sheet)
            code_list = check_code_type(row, column, sheet)
            check_mothers_code(row, column, sheet, code_list)
        else:
            print('Sorry, the selection is not proper!\n')
    
if __name__ == '__main__':
    interact()
