import pandas as pd
from sqlalchemy import create_engine
import mysql.connector as mys
import PySimpleGUI as sg
import pymysql
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
pymysql.install_as_MySQLdb()

df = pd.read_excel(r'C:\Users\jainr\Downloads\Book1.xlsx')
df.head()
engine = create_engine('mysql://root:1234@localhost/result')
#df.to_sql('result_analysis', con=engine)

mycon=mys.connect(host='localhost',user='root',password='1234',database='result')
crsr=mycon.cursor()
#crsr.execute('select * from result_analysis')
crsr.execute('select name,Science from result_analysis order by Science desc')

Science=crsr.fetchall()
crsr.execute('select name,Maths from result_analysis order by Maths desc')
Maths=crsr.fetchall()
crsr.execute('select name,S_Sc from result_analysis order by S_Sc desc')
S_Sc=crsr.fetchall()
crsr.execute('select name,Eng from result_analysis order by Eng desc')
Eng=crsr.fetchall()
crsr.execute('select name,Hindi from result_analysis order by Hindi desc')
Hindi=crsr.fetchall()



fulltable=crsr.execute('select * from result_analysis')
fulldata=crsr.fetchall()




def make_win1():
    headings=['index','roll no.','name','class','section','Eng marks','Maths marks','Hindi marks','Science marks','S_Sc marks','avg','percentage']
    layout = [[sg.Text('This is the reult analyzer'), sg.Text('      ', k='-OUTPUT-')],
              [sg.Table(fulldata, headings=headings, justification='left', key='-TABLE-')],
              [sg.Button('subject wise result'), sg.Button('Exit')]]
    return sg.Window('Window Title', layout, location=(100,300), finalize=True,resizable=True)


def make_win2():
    font = ('Courier New', 16)
    text = '\n'.join(chr(i)*50 for i in range(65, 91))
    column = [[sg.Text(text, font=font)]]
    headings1=['name','Maths marks']
    headings2=['name','Science marks']
    headings3=['name','S_Sc marks']
    headings4=['name','Eng marks']
    headings5=['name','Hindi marks']
    layout = [[sg.Text('subject wise result')],
              [sg.Text('Maths')],
              [sg.Table(Maths,headings=headings1,justification='left',key='-TABLE-')],
              [sg.Text('Science')],
              [sg.Table(Science,headings=headings2,justification='left',key='-TABLE-')],
              [sg.Text('S_Sc')],
              [sg.Table(S_Sc,headings=headings3,justification='left',key='-TABLE-')],
              [sg.Text('Eng')],
              [sg.Table(Eng,headings=headings4,justification='left',key='-TABLE-')],
              [sg.Text('Hindi')],
              [sg.Table(Hindi,headings=headings5,justification='left',key='-TABLE-')],
              [sg.OK(), sg.Button('Up', key = "up"), sg.Button('Down', key = "down")],
              [sg.Column(column, size=(800, 300), scrollable=True,vertical_scroll_only=True, key = "Column")],
              ]
    return sg.Window('subject wise result', layout ,location=(100,300),resizable=True,finalize=True,)

window1, window2 = make_win1(), None        # start off with 1 window open

while True:             # Event Loop
    window, event, values = sg.read_all_windows()
    if event == sg.WIN_CLOSED or event == 'Exit':
        window.close()
        if window == window2:       # if closing win 2, mark as closed
            window2 = None
        elif window == window1:     # if closing win 1, exit program
            break
    elif event == 'Popup':
        sg.popup('This is a BLOCKING popup','all windows remain inactive while popup active')
    elif event == 'subject wise result' and not window2:
        window2 = make_win2()
        if event == "down":
            window['Column'].Widget.canvas.yview_moveto(1.0)
        elif event == "up":
            window['Column'].Widget.canvas.yview_moveto(0.0)
    elif event == '-IN-':
        window['-OUTPUT-'].update(f'You enetered {values["-IN-"]}')
    elif event == 'Erase':
        window['-OUTPUT-'].update('')
        window['-IN-'].update('')
window.close()

