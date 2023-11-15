import PySimpleGUI as sg
import pandas as pd
import mysql.connector
import os
import json
import requests

mysqldb=mysql.connector.connect(host="localhost",user="root",password="",database="aplikasi")
mycursor=mysqldb.cursor()


sg.theme('DarkGreen4')

EXCEL_FILE = 'Pendaftaran.xlsx'

df = pd.read_excel(EXCEL_FILE)

layout=[
[sg.Text('Masukan Data Pasien: ')],
[sg.Text('Nama',size=(15,1)), sg.InputText(key='Nama')],
[sg.Text('No Telp',size=(15,1)), sg.InputText(key='Tlp')],
[sg.Text('Alamat',size=(15,1)), sg.Multiline(key='Alamat')],
[sg.Text('Tgl Lahir',size=(15,1)), sg.InputText(key='Tgl Lahir'),
                                    sg.CalendarButton('Kalender', target='Tgl Lahir', format=('%Y-%m-%d'))],
[sg.Text('Jenis Kelamin',size=(15,1)), sg.Combo(['pria','wanita'],key='Jekel')],
[sg.Text('Pembayaran',size=(15,1)), sg.Checkbox('Tunai',key='Tunai'),
                            sg.Checkbox('Debit',key='Debit'),
                             sg.Checkbox('BPJS',key='BPJS')],
[sg.Submit(), sg.Button('clear'), sg.Button('view data'), sg.Button('open excel'), sg.Exit()]

]

window=sg.Window('Form pendaftaran',layout)

def select():
    results = []
    mycursor.execute("select nama,tlp,alamat,tgl_lahir,jekel,pembayaran from pendaftaran order by id desc")
    for res in mycursor:
        results.append(list(res))

    headings=['Nama','Tlp','Alamat','Tgl Lahir', 'Jekel', 'Pembayaran'] 

    layout2=[
        [sg.Table(values=results,
        headings=headings,
        max_col_width=35,
        auto_size_columns=True,
        display_row_numbers=True,
        justification='right',
        num_rows=20,
        key='-Table-',
        row_height=35)]
    ]   

    window=sg.Window("List Data", layout2)
    event, values = window.read()

def send_gdrive():
    headers={
        "Authorization":"Bearer #####"
    }
    name={
        "name":"Pendaftaran.xlsx",
    }
    files={
        'data':('metadata',json.dumps(name),'application/json;charset=UTF-8'),
        'file':('mimeType',open("./Pendaftaran.xlsx","rb"))
    }
    r=requests.post(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        headers=headers,
        files=files
    )
    print(r.text)

def clear_input():
    for key in values:
        window[key]('')
        return None

while True :
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'EXIT':
        break
    if event == 'Clear':
        clear_input()
    if event == 'view data':
        select()  
    if event == 'open excel':
        os.startfile(EXCEL_FILE)     
    if event == 'Submit':
        nama=values["Nama"]
        tlp=values["Tlp"]
        alamat=values["Alamat"]
        tgl_lahir=values["Tgl Lahir"]
        jekel=values["Jekel"]
        cash=values["Cash"]
        debit=values["Debit"]
        BPJS=values["BPJS"]

        if cash == True:
            pembayaran="Cash"
        if debit == True:
            pembayaran="Debit"
        if BPJS == True:
            pembayaran="BPJS"

        sql="insert into pendaftaran(nama,tlp,alamat,tgl_lahir,jekel,pembayaran) values(%s,%s,%s,%s,%s,%s)"
        val=(nama,tlp,alamat,tgl_lahir,jekel,pembayaran)
        mycursor.execute(sql,val)
        mysqldb.commit()
        
        df =df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Berhasil Di Simpan')
        send_gdrive()
        clear_input()
window.close()       