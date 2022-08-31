import cv2
from pyzbar import pyzbar
import pygsheets
gc = pygsheets.authorize(service_file='./check-in.json')
sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1r7AmqwXxw-MmAT0f-yYiqTnSieinz4tb7N0Onn0C0t8/edit#gid=1154285616')
wks = sht.worksheet_by_title("新生盃報到")
gamename = ["新生男團", "新生女團", "新生男單", "新生女單"]
def read_barcodes(frame):
    barcodes = pyzbar.decode(frame)
    name, ID, department = '', '', ''
    game = []
    IDcell = None
    for barcode in barcodes:
        #print("hi")
        x, y , w, h = barcode.rect
        barcode_info = barcode.data.decode('utf-8')
        cv2.rectangle(frame, (x, y),(x+w, y+h), (0, 255, 0), 2)
        
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, barcode_info, (x + 6, y - 6), font, 2.0, (255, 255, 255), 1)
        ID = barcode_info[0:9]
        print(ID)

        findcell = wks.find(ID.upper())
        if len(findcell) == 0:
            findcell = wks.find(ID.lower())
        if len(findcell) != 1:
            print("found error")
        
        IDcell = wks.cell(findcell[0].label)
        name = IDcell.neighbour("left").value
        department = IDcell.neighbour("right").value
        game = [False, False, False, False]
        for i in range(4):
            if IDcell.neighbour((0, 2+i)).value == 'TRUE':
                game[i] = True
        print(IDcell)
        print(f'姓名：{name}')
        print(f'系級：{department}')
        for i in range(4):
            if game[i]:
                print(gamename[i])
        wks.update_value(IDcell.neighbour((0, 6)).label, 'TRUE')
    return frame, name, ID, department, game, IDcell

def main():
    #1
    camera = cv2.VideoCapture(1)
    ret, frame = camera.read()
    #2
    ret, frame = cv2.threshold(frame, 127, 255, cv2.THRESH_BINARY)
    while ret:
        ret, frame = camera.read()
        frame, name, ID, department, game, IDcell = read_barcodes(frame)
        cv2.imshow('Barcode/QR code reader', frame)
        if cv2.waitKey(1) & 0xFF == 27:
            break
    #3
    camera.release()
    cv2.destroyAllWindows()
#4
if __name__ == '__main__':
    main()
