"""
Use the instrument risk registers in the curren folder to update the databse for this month.

By default, possible information from a run within the last two weeks is being deleted to allow multiple
runs if some risk registers were delivered too late etc.
"""

import os
import sys
from datetime import datetime, timedelta
from glob import glob

import pyodbc
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table


if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    CUR_PATH = os.path.dirname(os.path.abspath(sys.executable))
else:
    CUR_PATH = os.path.dirname(os.path.abspath(__file__))

DATE_OVERWRITE = None  # datetime(year=2022, month=7, day=27)


def read_excel(fname):
    wb = load_workbook(filename=fname, data_only=True)
    tbl: Table = wb['Risks'].tables['Risk_Reg']
    data = [[ci.value for ci in ri] for ri in wb['Risks'][tbl.ref]]

    return tbl.column_names, data[tbl.headerRowCount:]


def tint(val):
    try:
        return int(val)
    except ValueError:
        return -1


def main():
    print(f'{CUR_PATH}\\full_risk_database.accdb')
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={CUR_PATH}\\full_risk_database.accdb;'
    )
    cnxn = pyodbc.connect(conn_str)
    crsr = cnxn.cursor()

    columns = [c.column_name for c in crsr.columns(table="Full Risk History")]

    # delete entries younger than 2 weeks
    now_date = DATE_OVERWRITE or datetime.now().date()
    ref_data = now_date-timedelta(days=14)
    crsr.execute('DELETE FROM "Full Risk History" WHERE "Date Added">=?', ref_data)
    crsr.execute('DELETE FROM "Latest Risks"')

    for fname in glob(os.path.join(CUR_PATH, 'latest', '*Risks.xlsx')):
        print(fname)
        dcols, data = read_excel(fname)

        insert_data = []

        for di in data:
            row = (di[0].upper(), now_date, tint(di[1]), di[2], di[3],
                   di[4], di[5], di[6], di[7], di[8],
                   tint(di[17]), tint(di[18]), tint(di[19]), tint(di[20]), tint(di[21]), tint(di[22]), tint(di[24])
                   )
            insert_data.append(row)

            res = crsr.execute('SELECT "Risk Rating" FROM "Full Risk History" WHERE Project=? AND "Risk ID"=? '
                               'ORDER BY "Date Added" DESC',
                               (row[0], row[2]))
            risk_history = [fi[0] for fi in res.fetchall()]
            try:
                prev_rating = risk_history[0]
            except IndexError:
                prev_rating = -1
            crsr.execute('INSERT INTO "Latest Risks"'
                         '(Project, "Risk ID", "Risk Title", "Risk and Impact Description",'
                         'Owner, Partner, Status, "Risk Treatment", "Treatment Actions and Notes", '
                         'when, cost, schedule, quality, "max impact", likelyhood, "Risk Rating", '
                         '"Last Rating", "Full History") '
                         'VALUES (?, ?, ?, ?, '
                         '?, ?, ?, ?, ?, '
                         '?, ?, ?, ?, ?, ?, ?,'
                         '?, ?)',
                         (row[0],)+row[2:]+(prev_rating, repr(risk_history)))

        crsr.executemany('INSERT INTO "Full Risk History"'
                         '(Project, "Date Added", "Risk ID", "Risk Title", "Risk and Impact Description",'
                         'Owner, Partner, Status, "Risk Treatment", "Treatment Actions and Notes", '
                         'when, cost, schedule, quality, "max impact", likelyhood, "Risk Rating") '
                         'VALUES (?, ?, ?, ?, ?, '
                         '?, ?, ?, ?, ?, '
                         '?, ?, ?, ?, ?, ?, ?)',
                         insert_data)
    crsr.commit()

    input("Finished, press enter to close program.")


if __name__=='__main__':
    main()
