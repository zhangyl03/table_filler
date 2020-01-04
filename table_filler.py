'''
This is a special tool for XMU, Jan. 4, 2020.
If you have any questions, please contact the author
95884234@qq.com
'''
import xlrd
from pynput.keyboard import Key, Controller, Listener

# Read all the data from the exam.xlsx file and save it in a list
def readTable():
    data_file = 'exam.xlsx'
    sheet = 'Sheet1'
    data_list = [] # a data list to contain all the data from .xlsx
    wb = xlrd.open_workbook(data_file) # open exam.xlsx file
    sh = wb.sheet_by_name(sheet) # open Sheet1
    for i in range(1, sh.nrows): # get data from row number 2
        for j in range(4, sh.ncols): # get data from col number 5
            d = sh.cell_value(i,j)
            data_list.append(d)
    return data_list

# Fill in the table on the webpage with the data read from the exam.xlsx file
# Paste all the elements in the list and then press tab
def pasteAndTab(data): # data is a list
    keyboard = Controller()
    for d in data:
        keyboard.type(str(d))
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)

# Nothing to do for the main file now.
def main():
    print('This is the main function.')

def on_press(key):
    if key == Key.space:
        print('Start working! Press Esc to stop.')
        a = readTable()
        pasteAndTab(a)

def on_release(key):
    if key == Key.esc:
        # Stop listener
        return False

if __name__ == "__main__":
    print('Press SPACE to start, and press Esc to stop...')
    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()
    print('Job done!')
