import requests
import pygsheets
import pymssql


#Authirize Google Sheets
client = pygsheets.authorize(client_secret='xxxxxxxx/client_secret.json') #อย่าลืมแก้
# Open the spreadsheet and the first sheet.
sh = client.open('abc') #อย่าลืมแก้
wks = sh.sheet1

# Connect TSG Server DB
server = ("xxx.xxx.xxx.xxx")
user = ("xxxx")
password = ("xxxxxx")
conn = pymssql.connect(server, user, password, "xxxxxx")
cursor = conn.cursor(as_dict=True)


#### START CODING ####

# check last row
lastrow = 0
for row in wks:
    if len(row[0]) >0 :
        lastrow = lastrow + 1
print (f'lastrow {lastrow}')

# Start loop check 
row = 2
for i in range(1,lastrow):
    barcode = wks.cell(f'A{row}').value
    serialnumber = wks.cell(f'E{row}').value
    if len(serialnumber) > 0 :
        sql = 'UPDATE tbm_barcode SET serial_number = %s WHERE code = %s'
        val = (serialnumber,barcode)
        cursor.execute(sql,val)
        conn.commit()
        wks.update_value(f'F{row}',f'Complete Update serial : {serialnumber}')
        print(f'Update {serialnumber} complete')

        cursor.execute('SELECT * FROM tbm_barcode WHERE code =%s',barcode )
        for result in cursor:
            wks.update_value(f'D{row}',result['serial_number'])
        row = row + 1
    else :
        wks.update_value(f'F{row}',f'No Input from user')
        row = row + 1
conn.close()
wks.update_value(f'A{row}','เมื่อเรียบร้อยแล้ว อย่าลืมลบข้อมูลออกให้หมด จะได้พร้อมใช้ทำงานครั้งถัดไปได้')
