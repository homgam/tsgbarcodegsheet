import requests
import pygsheets
import pymssql


#Authirize Google Sheets
client = pygsheets.authorize(client_secret='xxxxxxx/client_secret.json')
# Open the spreadsheet and the first sheet.
sh = client.open('abc')
wks = sh.sheet1

# Connect TSG Server DB
server = ("xxx.xxx.xxx.xxx")
user = ("xx")
password = ("xxxxx")
conn = pymssql.connect(server, user, password, "xxxxxxx")
cursor = conn.cursor(as_dict=True)


#### START CODING ####

# check last row
lastrow = 0
for row in wks:
    if len(row[0]) >0 :
        lastrow +=1
print (f'lastrow {lastrow}')

# Start loop check 
row = 2
for i in range(1,lastrow):
    bar = wks.cell(f'A{row}').value
    cursor.execute('SELECT * FROM tbm_barcode WHERE code =%s',bar )
    for result in cursor:
        barcodenum = result['code']
        print (barcodenum)
        wks.update_value(f'B{row}',result['product_grade_master_id'])
        wks.update_value(f'C{row}',result['customer_id'])
        wks.update_value(f'D{row}',result['serial_number'])
        print (f'ROW {row} Barcode {barcodenum} write complete')
    row = row + 1
conn.close()
