import xlrd
import xlwt

"""
create a excel sheet based on the dictonary in txt file
the txt file is required in following format:
    data = {1:{a: 1, b: 2}, 2: {c: 3, d: 4}}
    variable name can be any, or even just a dictionary by itself without a name is also fine
"""
def interact():
    
    # get the txt content from user's input
    file = input("Please insert the name of the file(it needs to be placed in the same location as this script): ") #get file name from user's input
    f = open(file)
    txt = f.read()
    
    # find out the start of the dictionary
    i = 0
    while i < len(txt): 
        if txt[i] == '{':
            break
        i += 1

    # convert the dictionary from string to dict
    dictionary = eval(txt[i:])
    dict_len = len(dictionary)
    key_list = list(dictionary.get(1).keys())

    # create sheet page
    writebook = xlwt.Workbook()
    sheet = writebook.add_sheet('sheet1')

    # write all the keys into the first row of the sheet
    i = 0
    while i < len(key_list):
        sheet.write(0, i, key_list[i])
        i += 1

    # write the entire dictionary into the sheet
    i = 1
    while i <= dict_len:
        sub_dict = dictionary.get(i)
        j = 0
        while j < len(key_list):
            sheet.write(i, j, sub_dict.get(key_list[j]))
            j += 1
        i += 1

    # save the sheet -> create excel file -> notification
    writebook.save(file[:-4] + '.xls')
    print("the file " + "'" + file[:-4] + ".xls" + "'" + " has been created in the same location")

if __name__ == '__main__':
    interact()
