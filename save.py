from time import sleep
import os
import os.path
from glob import glob
import xlrd
import sqlite3
import datetime

def save():
    conn = sqlite3.connect('main.db')
    c = conn.cursor()
    files = glob('excel/*')
    for file in files:
        workbook = xlrd.open_workbook(file)
        print(workbook)
        sheet = workbook.sheet_by_index(0)
        date = sheet.cell_value(rowx=7, colx=6)
        date = datetime.datetime(*xlrd.xldate_as_tuple(date, workbook.datemode))
        for rowidx in range(10,sheet.nrows):
            row = sheet.row(rowidx)
            c.execute('INSERT INTO stocks values (?,?,?,?,?,?,?,?,?,?,?,?)',
                      (row[0].value,
                       row[1].value,
                       row[2].value,
                       row[3].value,
                       row[4].value,
                       row[5].value,
                       row[6].value,
                       row[7].value,
                       row[8].value,
                       row[9].value,
                       row[10].value,
                       date))
        conn.commit()
        workbook.release_resources()
        del workbook
        os.remove(file)
    conn.close()

save()
