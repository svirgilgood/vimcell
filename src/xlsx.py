#!/usr/bin/env python3

import xlrd, xlsxwriter
import sys, os


def xlsx2csv(xlfile):


def csv2xlsx(directory):


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('workbook', nargs='+')
    ap.add_argument('-x', '--xlsx', help='From xlsx to csv directory',
            action='store_true')
    ap.add_argument('-c', '--csv', help='From csv directory to xlsx workbook',
            action='store_true')
    a = ap.parse_args()

    


if '__name__' == '__main__':
    import argparse
    main()
