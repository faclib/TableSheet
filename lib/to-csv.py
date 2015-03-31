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
    parser = ArgumentParser(description="преобразует в CSV")

    parser.add_argument('-w', '--windows', action="store_true", help="формат для Windows")
    parser.add_argument('infile', type=str, help="входной-файл")
    parser.add_argument('outfile',  metavar='outfile', nargs='?', help="выходной-файл CSV")

    try:
        args = parser.parse_args()
    except:
        # sys.exit(2)
        return
    # ---------------------------

    infile = args.infile
    outfile = args.outfile or sys.stdout

    Table.convert_to_csv(infile, outfile, args.windows)

if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        print "Error: {0}".format(e)
        sys.exit(1)
        # sys.exit("Error: {0}".format(e))

        # import errno, traceback
        # traceback.print_exc(file=sys.stdout)
        # sys.exit(errno.EACCES)

