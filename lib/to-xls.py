#!/usr/bin/python
# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.1'

import sys
import Table
from argparse import ArgumentParser

reload(sys)
sys.setdefaultencoding('utf-8')
# ------------------------------


def main():
    parser = ArgumentParser(description="преобразует в XLS")

    parser.add_argument('infile',  type=str, help="входной-файл CSV")
    parser.add_argument('outfile', type=str, help="выходной-файл XLS")
    parser.add_argument("-s", "--sheetname", dest="sheetname", default=None, help="имя сохраняемого листа")

    try:
        args = parser.parse_args()
    except:
        return
    # ---------------------------

    infile = args.infile
    output = args.outfile

    f = Table.utf8_open_file(infile)
    reader = Table.CSVUnicodeReader(f)

    writer = Table.XLSWriter(args.sheetname)
    writer.write_reader(reader)
    writer.frozen()
    writer.save(output)

if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        print "Error: {0}".format(e)
        sys.exit(1)
