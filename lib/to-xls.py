#!/usr/bin/python
# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.0'

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
    args = parser.parse_args()
    # ---------------------------

    infile = args.infile
    output = args.outfile

    f = Table.utf8_file_encode(infile)
    reader = Table.CSVUnicodeReader(f)

    writer = Table.XLSWriter()
    writer.write_reader(reader)
    writer.save(output)

if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        sys.exit("Error: {0}".format(e))
