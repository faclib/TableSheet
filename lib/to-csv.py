#!/usr/bin/python
# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.2'


import sys
import Table
import tempfile
import csv
from argparse import ArgumentParser

reload(sys)
sys.setdefaultencoding('utf-8')
# ------------------------------


def main():
    parser = ArgumentParser(description="преобразует в CSV")

    parser.add_argument("-d", "--delimiter", dest="delimiter", default=",",
      help="delimiter - columns delimiter in csv(default: ',')")
    # parser.add_argument('-a', '--all', action="store_true", help="export all sheets")
    parser.add_argument('infile', type=str, help="входной-файл")
    parser.add_argument('outfile',  metavar='outfile', nargs='?', help="выходной-файл CSV")

    try:
        options = parser.parse_args()
    except:
        return

    outfile = options.outfile or sys.stdout

    convertCsv=Table.ConvertCSV(options.infile)
    convertCsv.convert(outfile, options.delimiter)
    # ---------------------------


if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        print "Error: {0}".format(e)
        sys.exit(1)
