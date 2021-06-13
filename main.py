import edit_sheet
from edit_sheet import update_price
from edit_sheet import create_new_xl_file
from edit_sheet import create_new_sheet
from edit_sheet import view_xl_file
from edit_sheet import edit_data
import os

while 1==1:
    print("""
#Excel File Editor
    
1. Create a new excel file
2. Add a new sheet
3. Edit
4. View excel file
5. Update prices in file (xlsheet.xlsx)
6. Help
7. Any other key to exit
    """)
    option = input('Choose your option : ')
    if option == '1':
        create_new_xl_file()
    elif option == '2':
        create_new_sheet()
    elif option == '3':
        edit_data()
    elif option == '4':
        view_xl_file()
    elif option == '5':
        update_price
    elif option == '6':
        os.startfile('readme.txt')
    else:
        print('Thank you!!!')
        break
















