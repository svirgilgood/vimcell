#!/usr/bin/env python3

import xlrd, xlsxwriter
import os, sys
import subprocess as sp


def xlsx2csv(xlfile, filename):
    dirname = filename.replace('.xlsx', '').replace('.xls', '') + '.vcell'
    print(dirname)
    try:
        os.mkdir(dirname)
    except FileExistsError:
        os.chdir(dirname)
        os.system('vim -p ' + ' '.join(os.listdir('.')))
        sys.exit()
    os.chdir(dirname)
    for sheet in xlfile.sheet_names():
        csv = open(sheet.replace(' ', '_') + '.csv', 'w')
        for rw in range(0, xlfile.sheet_by_name(sheet).nrows):
            values = xlfile.sheet_by_name(sheet).row_values(rw)
            values = [str(x) for x in values]
            new_values = ', '.join(values) + '\n'
            csv.write(new_values)
        csv.close()
    print(os.listdir('.'))
    var = ' '.join(os.listdir('.'))
    os.system('vim -p ' + var)
    #sp.Popen(['vim', '-p'] + os.listdir('.'))

'''
def csv2xlsx(directory):

'''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('workbook', nargs='+')
    ap.add_argument('-x', '--xlsx', help='From xlsx to csv directory',
            action='store_true')
    ap.add_argument('-c', '--csv', help='From csv directory to xlsx workbook',
            action='store_true')
    a = ap.parse_args()
    if a.xlsx:
        xlsx_filename = a.workbook.pop()
        xlsxwb = xlrd.open_workbook(xlsx_filename)
        xlsx2csv(xlsxwb, xlsx_filename)



if __name__ == '__main__':
    import argparse
    main()
