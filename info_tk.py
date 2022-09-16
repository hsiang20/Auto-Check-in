import tkinter as tk
from tkinter.constants import W
from fake_fetch import fetch
import cv2
from pyzbar import pyzbar
import pygsheets
from playsound import playsound

gc = pygsheets.authorize(service_file='./checkin_2022.json')
sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/19nNC0c9lPmKGCyi0oIKJLwhsyiuXODuVpOfUJ3auOcE/edit#gid=1363448317')
wks = sht.worksheet_by_title("疫苗陰性證明")
# gamename = ["公開男團", "公開女團", "公開男單", "公開女單", "歡樂雙打"]
gamename = ["新生男團", "新生女團", "新生男單", "新生女單"]
IDs = wks.get_col(5)
names = wks.get_col(3)
deps = wks.get_col(4)
# game1 = wks.get_col(12)
# game2 = wks.get_col(13)
# game3 = wks.get_col(14)
# game4 = wks.get_col(15)
# game5 = wks.get_col(16)

game1 = wks.get_col(18)
game2 = wks.get_col(19)
game3 = wks.get_col(20)
game4 = wks.get_col(21)

# games = [game1, game2, game3, game4, game5]
games = [game1, game2, game3, game4]

d = dict()
for i in range(len(IDs)):
    d[IDs[i].upper()] = i
d_name = dict()
for i in range(len(names)):
    d_name[names[i]] = i

camera = cv2.VideoCapture(0)
ret, frame = camera.read()
ret, frame = cv2.threshold(frame, 127, 255, cv2.THRESH_BINARY)

def open_camera():
    global camera
    global ret
    global frame
    global IDcell
    while ret:
        ret, frame = camera.read()
        frame, name, id, dep, contest_list, IDcell = read_barcodes(frame)
        cv2.imshow('Barcode/QR code reader', frame)
           
        if name == False:
            contest='項目：\n\n'
            name = '姓名：無報名資料'
            id = '學號：' + id
            dep = '系級：'
            name_label['text'] = name
            id_label['text'] = id
            dep_label['text'] = dep
            contest_label['text'] = contest
            vaccine_label['text'] = '有無疫苗/陰性證明：'
            complete_label['text'] = ''
            break
           
        if id != '':
            contest='項目：\n\n'
            name = '姓名：' + name
            id = '學號：' + id
            dep = '系級：' + dep
            if contest_list[0]:
                contest += "新生男團\n"
            if contest_list[1]:
                contest += "新生女團\n"
            if contest_list[2]:
                contest += "新生男單\n"
            if contest_list[3]:
                contest += "新生女單\n"
            # if contest_list[4]:
            #     contest += "歡樂雙打\n"
            name_label['text'] = name
            id_label['text'] = id
            dep_label['text'] = dep
            contest_label['text'] = contest
            vaccine_label['text'] = '有無疫苗/陰性證明：'
            complete_label['text'] = ''
            break
        
        if cv2.waitKey(1) & 0xFF == 27:
            break


window = tk.Tk()
# window.title('臺大盃報到')
window.title('新生盃報到')
window.geometry('800x1200')

f, name, id, dep, contest_list, idcell= fetch();
contest = ''

def yes():
    global IDcell
    vaccine_label['text'] = "有無疫苗/陰性證明：有"
    complete_label['text'] = "報到完成！"
    #wks.update_value(IDcell.neighbour((0, 8)).label, 'TRUE')

def no():
    global IDcell
    vaccine_label['text'] = "有無疫苗/陰性證明：無"
    complete_label['text'] = "報到完成！"
    #wks.update_value(IDcell.neighbour((0, 8)).label, 'FALSE')

def enter_id():
    global d
    global IDcell
    name = name_entry_var.get().upper()
    name_entry.delete(0, 'end')
    if name not in names and name not in d:
        contest='項目：\n\n'
        name = '姓名：無報名資料'
        id = '學號：'
        dep = '系級：'
        name_label['text'] = name
        id_label['text'] = id
        dep_label['text'] = dep
        contest_label['text'] = contest
        vaccine_label['text'] = '有無疫苗/陰性證明：'
        complete_label['text'] = ''
    
    elif name in d:
        IDindex = d[name]
        ID = IDs[IDindex]
        name = names[IDindex]
        dep = deps[IDindex]
        game = [False, False, False, False, False]
        for i in range(4):
        # for i in range(5):
            #if IDcell.neighbour((0, 2+i)).value == 'TRUE':
            if games[i][IDindex] == 'TRUE':
                game[i] = True
        IDcell = wks.cell('C' + str(IDindex+1))
        print(IDcell)
        print(f'姓名：{name}')
        print(f'系級：{dep}')
        for i in range(5):
            if game[i]:
                print(gamename[i])
        wks.update_value(IDcell.neighbour((0, 19)).label, 'TRUE')
        contest='項目：\n\n'
        name = '姓名：' + name
        id = '學號：' + ID
        dep = '系級：' + dep
        if game[0]:
            contest += "新生男團\n"
        if game[1]:
            contest += "新生女團\n"
        if game[2]:
            contest += "新生男單\n"
        if game[3]:
            contest += "新生女單\n"
        # if game[4]:
        #     contest += "歡樂雙打\n"
        name_label['text'] = name
        id_label['text'] = id
        dep_label['text'] = dep
        contest_label['text'] = contest
        vaccine_label['text'] = '有無疫苗/陰性證明：'
        complete_label['text'] = ''
        
    else:
        IDindex = d_name[name]
        ID = IDs[IDindex]
        dep = deps[IDindex]
        game = [False, False, False, False, False]
        for i in range(4):
        # for i in range(5):
            if games[i][IDindex] == 'TRUE':
                game[i] = True
        IDcell = wks.cell('C' + str(IDindex+1))
        print(IDcell)
        print(f'姓名：{name}')
        print(f'系級：{dep}')
        for i in range(5):
            if game[i]:
                print(gamename[i])
        wks.update_value(IDcell.neighbour((0, 19)).label, 'TRUE')
        contest='項目：\n\n'
        name = '姓名：' + name
        id = '學號：' + ID
        dep = '系級：' + dep
        if game[0]:
            contest += "新生男團\n"
        if game[1]:
            contest += "新生女團\n"
        if game[2]:
            contest += "新生男單\n"
        if game[3]:
            contest += "新生女單\n"
        # if game[4]:
        #     contest += "歡樂雙打\n"
        name_label['text'] = name
        id_label['text'] = id
        dep_label['text'] = dep
        contest_label['text'] = contest
        vaccine_label['text'] = '有無疫苗/陰性證明：'
        complete_label['text'] = ''

def read_barcodes(frame):
    global IDcell
    global d
    barcodes = pyzbar.decode(frame)
    name, ID, department = '', '', ''
    game = []
    IDcell = None
    for barcode in barcodes:
        playsound('beep.mp3')
        x, y , w, h = barcode.rect
        barcode_info = barcode.data.decode('utf-8')
        cv2.rectangle(frame, (x, y),(x+w, y+h), (0, 255, 0), 2)
        
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, barcode_info, (x + 6, y - 6), font, 2.0, (255, 255, 255), 1)
        ID = barcode_info[0:9].upper()
        print(ID)

        if ID not in d:
            return frame, False, ID, False, False, False


        IDindex = d[ID]

        IDcell = wks.cell('C' + str(IDindex+1))

        name = names[IDindex]
        department = deps[IDindex]
        game = [False, False, False, False, False]
        for i in range(4):
        # for i in range(5):
            if games[i][IDindex] == 'TRUE':
                game[i] = True
        print(IDcell)
        print(f'姓名：{name}')
        print(f'系級：{department}')
        for i in range(4):
        # for i in range(5):
            if game[i]:
                print(gamename[i])
        wks.update_value(IDcell.neighbour((0, 19)).label, 'TRUE')
    return frame, name, ID, department, game, IDcell

name_entry_var = tk.StringVar()
title_label = tk.Label(window, text="111學年度新生盃報到", font=('標楷體', 48), height=2)
name_label = tk.Label(window, text="姓名："+name, font=('標楷體', 36))
id_label = tk.Label(window, text="學號："+id, font=('標楷體', 36))
dep_label = tk.Label(window, text="系級："+dep, font=('標楷體', 36))
contest_label = tk.Label(window, text=contest, font=('標楷體', 36),)
# vaccine_label = tk.Label(window, text='有無疫苗/陰性證明：', font=('標楷體', 36) )
# vaccine_bn_y = tk.Button(window, text='有', command=yes, background='green', highlightcolor='green', font=('標楷體', 36))
# vaccine_bn_n = tk.Button(window, text='無', command=no, background='blue', font=('標楷體', 36))
complete_label = tk.Label(window, text='', font=('標楷體', 36))
complete_bn = tk.Button(window, text='開始掃碼', command=open_camera, background='yellow', font=('標楷體', 36))
name_entry = tk.Entry(window, textvariable=name_entry_var, width=12, font=('標楷體', 36))
name_bn = tk.Button(window, text='輸入', command=enter_id, font=('標楷體, 36'))

title_label.grid(row=0, columnspan=4)
name_label.grid(row=1, column=0, sticky=tk.W, ipady=10, ipadx=80)
id_label.grid(row=2, column=0, sticky=tk.W, ipady=10, ipadx=80)
dep_label.grid(row=3, column=0, sticky=tk.W, ipady=10, ipadx=80)
contest_label.grid(row=1, column=2, rowspan=3, sticky=tk.W)
# vaccine_label.grid(row=4, column=0, ipady=10, sticky=tk.W, ipadx=80)
# vaccine_bn_y.grid(row=5, column=0, columnspan=2)
# vaccine_bn_n.grid(row=5, column=2, columnspan=2, sticky=tk.W)
complete_label.grid(row=6, columnspan=4, ipady=10)
complete_bn.grid(row=7, columnspan=4)
name_entry.grid(row=8, columnspan=2, sticky=tk.E, ipadx=50)
name_bn.grid(row=8, columnspan=2, sticky=tk.E)

window.grid_columnconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)
window.grid_columnconfigure(2, weight=1)
window.grid_columnconfigure(3, weight=1)

window.mainloop()

camera.release()
cv2.destroyAllwindows()

