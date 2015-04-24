#!/usr/bin/python
# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.2'

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
    parser.add_argument('-H', '--head', action="store_true", help="frozen head")
    parser.add_argument('-c', '--color', dest="color", default=None, help="цвет шапки (F4ECC5)")

    try:
        args = parser.parse_args()
    except:
        return
    # ---------------------------

    infile = args.infile
    output = args.outfile

    infile = Table.unicode_filename(infile)
    f = open(infile, 'rb')

    reader = Table.CSVUnicodeReader(f)
    writer = Table.XLSWriter(args.sheetname)

    if args.head:
        if args.color:
            writer.set_head(args.color)
        else:
            writer.set_head()

    writer.write_reader(reader)

    if args.head:
        writer.frozen()

    writer.save(output)

if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        print "Error: {0}".format(e)
        sys.exit(1)
