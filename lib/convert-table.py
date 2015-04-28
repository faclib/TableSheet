#!/usr/bin/python
# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.4'

import sys
import Table
from argparse import ArgumentParser
import tempfile

reload(sys)
sys.setdefaultencoding('utf-8')
# ------------------------------


parser = ArgumentParser(description="конвертация таблиц",
    # prog='PROG',
    usage='%(prog)s <command> [options] <infile> <outfile>'
    , add_help=False
    )

parser.add_argument('-h', '--help',  action="store_true", help="help message")

parser.add_argument("command", type=str, choices=['xls', 'csv'], help="команда")
parser.add_argument('infile',  type=str, help="входной-файл")
parser.add_argument('outfile', type=str, help="выходной-файл (CSV, XLS)")

group_csv = parser.add_argument_group('csv', 'конвертация в CSV')
group_csv.add_argument("-d", "--delimiter", dest="delimiter", default=",",
      help="delimiter - columns delimiter in csv (default: ',')")


group_xls = parser.add_argument_group('xls', 'конвертация в XLS')
group_xls.add_argument('-f', '--forse', action="store_true", help="forse convert csv")
group_xls.add_argument('-s', '--sheetname', dest='sheetname', default=None, help="имя сохраняемого листа")
group_xls.add_argument('--head', action="store_true", help="frozen head")
group_xls.add_argument('-c', '--color', dest="color", default=None, help="цвет шапки (F4ECC5)")

# parser.print_help()


def cmd_xls(args):

    if args.forse:
        f = tempfile.TemporaryFile()
        convertCsv=Table.ConvertCSV(args.infile)
        convertCsv.convert(f, ',')
        f.seek(0)
    else:
        infile = Table.unicode_filename(args.infile)
        f = open(infile, 'rb')

    output = args.outfile

    reader = Table.CSVUnicodeReader(f)
    writer = Table.XLSWriter(args.sheetname)
    if args.color:
        writer.set_head(args.color)

    writer.write_reader(reader)
    if args.head:
        writer.frozen()
    writer.save(output)
# ------------------------------------------------


def cmd_csv(args):
    convertCsv=Table.ConvertCSV(args.infile)
    convertCsv.convert(args.outfile, args.delimiter)
# ------------------------------------------------


def main():
    if not len(sys.argv) or '-h' in sys.argv or '--help' in sys.argv:
        parser.print_help()
        return

    try:
        args = parser.parse_args()
    except:
        return

    if args.command == 'csv':
        cmd_csv(args)
    else:
        cmd_xls(args)
# ------------------------------------------------


if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        print "Error: {0}".format(e)
        sys.exit(1)
