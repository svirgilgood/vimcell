#!/usr/bin/env python3

import xlrd, xlsxwriter
import os, sys
import subprocess as sp
import shlex
import glob 


def openvim(dirname):
    os.chdir(dirname)
    args = shlex.split('vim -p ' + ' '.join(glob.glob('*')))
    #sp.call(args)
    sp.Popen(args).wait()
    sys.exit()


def xlsx2csv(xlfile, filename):
    dirname = filename.replace('.xlsx', '').replace('.xls', '') + '.vcell'
    print(dirname)
    try:
        os.mkdir(dirname)
    except FileExistsError:
        openvim(dirname)
    os.chdir(dirname)
    for sheet in xlfile.sheet_names():
        csv = open(sheet.replace(' ', '_') + '.csv', 'w')
        for rw in range(0, xlfile.sheet_by_name(sheet).nrows):
            values = xlfile.sheet_by_name(sheet).row_values(rw)
            values = [str(x) for x in values]
            new_values = ', '.join(values) + '\n'
            csv.write(new_values)
        csv.close()
    args = shlex.split('vim -p ' + ' '.join(glob.glob('*')))
    #os.system('vim -p ' + var)
    sp.call(args)


def csv2xlsx(directory):
    workbook = xlsxwriter.Workbook('new' + directory.replace('.vcell', '.xlsx'))
    #try:
    owd = os.getcwd()
    os.chdir(directory)
    #except IsAFileError:
    for sheet in glob.glob('*.csv'):
        working_sheet = workbook.add_worksheet(sheet.replace('.csv', ''))
        sheet = open(sheet)
        row = 0
        for line in sheet:
            line = [x.strip() for x in line.split(',')] #line.strip().split(',')
            for x in line:
                working_sheet.write(row, line.index(x), x)
            row += 1
       # row = 0
    os.chdir(owd)
    workbook.close()


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
        try: 
            xlsxwb = xlrd.open_workbook(xlsx_filename)
        except IsADirectoryError:
            openvim(xlsx_filename)
        xlsx2csv(xlsxwb, xlsx_filename)
    if a.csv:
        dirname = a.workbook.pop()
        csv2xlsx(dirname)
    else:
        openvim(a.workbook.pop())


if __name__ == '__main__':
    import argparse
    main()
