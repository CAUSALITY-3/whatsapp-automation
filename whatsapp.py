import pyautogui as pg
import webbrowser as web
import time
import xlrd

pageloadtime=8
messagesendtime=1
closetime=1

loc = ("C:/Users/../file_name.xls")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

sheet.cell_value(0, 0)
rows = sheet.nrows-1

first = True
for i in range(rows):
    first=True
    m = sheet.row_values(i + 1)
    n = str(m[0])[-12:-2:]
    phoneNo = "+91" + n
    message = str(m[1])
    time.sleep(1)
    web.open("https://web.whatsapp.com/send?phone="+phoneNo+"&text="+message)
    if first:
        time.sleep(pageloadtime)
        first=False
    width,height = pg.size()

    pg.click(width/1.5,height/1.1)
    time.sleep(messagesendtime)
    pg.press('enter')
    time.sleep(closetime)
    pg.hotkey('ctrl', 'w')