from openpyxl import load_workbook

FILE_NAME = 'Sinkat Contact Info.xlsx'
CURR_SUB_SHEET = 'Current Subcontractors'
PREV_SUB_SHEET = 'Previous Subcontractors'

def get_name():
    while(True):
        name = input('Enter Name: ')
        name = name.title()
        if(input('Is ' + name + ' ok? 1 for yes, 0 for no: ') == '1'):
            return name
        
def get_sin():
    sin = ''
    while (True):
        sin = ''

        sin = input("Enter SIN (0 TO QUIT): ")   

        sin = sin.replace(' ', '')
        
        print ("SIn is " + sin)
        if (sin == '0'):
            print("No SIN Entry")
            break
        elif (sin.isnumeric() == False):
            print("SIN Contains non-numeric elements")
        elif(len(sin) != 9):
            print("SIN does not contain 9 numbers")
        else:
            sin = (sin[0:3] + ' ' +
                sin[3:6] + ' ' + 
                sin[6:9])
            print(sin)
            return sin
    return sin

def get_address():
    while(True):
        address = input('Enter Address: ')
        address = address.title()
        if(input('Is ' + address + ' ok? 1 for yes, 0 for no: ') == '1'):
            return address

def get_city_province():
    while(True):
        city_province = input('Enter City and Province: ')
        if(input('Is ' + city_province + ' ok? 1 for yes, 0 for no: ') == '1'):
            return city_province

def get_postal_code():
    while(True):
        postal_code = input('Enter postal code: ')
        postal_code = postal_code.capitalize()
        

        if(input('Is ' + postal_code + ' ok? 1 for yes, 0 for no: ') == '1'):
            return postal_code

def get_email():
    while(True):
        city_province = input('Enter City and Province: ')
        if(input('Is ' + city_province + ' ok? 1 for yes, 0 for no: ') == '1'):
            return city_province

def get_phone_number():
    while(True):
        city_province = input('Enter City and Province: ')
        if(input('Is ' + city_province + ' ok? 1 for yes, 0 for no: ') == '1'):
            return city_province
            
def add_new_person(ws):
    new_row = ws.max_row+1
    #Assign
    ws['A' + str(new_row)].value = get_name()
    ws['B' + str(new_row)].value = get_sin()
    ws['C' + str(new_row)].value = get_address()
    ws['D' + str(new_row)].value = get_city_province()
    ws['E' + str(new_row)].value = get_postal_code()
    ws['F' + str(new_row)].value = get_email()
    ws['G' + str(new_row)].value = get_phone_number()


#Workbook
wb = load_workbook(FILE_NAME)
#Load Sheet
curr_sh = wb[CURR_SUB_SHEET]
pre_sh = wb[PREV_SUB_SHEET]

add_new_person(curr_sh)


#Reformat Sheet
wb.save(FILE_NAME)
wb.close()
