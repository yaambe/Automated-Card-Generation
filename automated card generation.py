import openpyxl as op
import win32com.client
import os

# initial data in the excel sheets

fullname_to_find = 'Ahmed Yameen'
designation_to_find = 'Senior Superintendent'
mobile_to_find = '9916696'
phone_to_find = '7497851'
email_to_find = 'ahmed.yameen'

# The data we are going to capture

fullname = input("Enter Staff Name: ")
designation = input("Enter Designation: ")
mobile = input("Mobile: (960) ")
phone = input("Phone: (960) ")
email = input("Customs E-mail (part before @customs.gov.mv): ")

# editing the template

wb = op.load_workbook(r"excel template.xlsx")
ws = wb['pages']
i = 0
for r in range(1 , ws.max_row + 1):
    for c in range(1 , ws.max_column + 1):
        s = ws.cell(r,c).value
        if s != None and fullname_to_find in s: 
            ws.cell(r,c).value = s.replace(fullname_to_find,fullname)
        if s != None and designation_to_find in s: 
            ws.cell(r,c).value = s.replace(designation_to_find,designation)
        if s != None and mobile_to_find in s: 
            ws.cell(r,c).value = s.replace(mobile_to_find,mobile)
        if s != None and phone_to_find in s: 
            ws.cell(r,c).value = s.replace(phone_to_find,phone)
        if s != None and email_to_find in s: 
            ws.cell(r,c).value = s.replace(email_to_find,email)
            i += 1

# All changes have been made

wb.save('cards.xlsx')

# save a copy

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
wb_path = r'cards.xlsx'
wb = o.Workbooks.Open(wb_path)

# opne the copy to create a pdf

ws_index_list = [1] 
path_to_pdf = r'Sample.pdf'

wb.WorkSheets(ws_index_list).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)


# now have to delete the copy card.xlsx if it exists so the template can be reused again
if os.path.exists('cards.xlsx'):
    os.remove('cards.xlsx')
else:
    print('cards.xlsx does not exist')
