#!/usr/bin/python
# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.0'

import chardet
import codecs
import cStringIO
import csv
import os
import re
import subprocess
import sys
import tempfile
import types
import xlrd
import xlwt

# Кодировка по умолчанию
reload(sys)
sys.setdefaultencoding('utf-8')
# ------------------------------


def unicode_name(filename, encoding='utf-8'):
    u"""
    Преобразует имя файла в правельный формат
    """

    filename = filename.decode(encoding)
    return filename
# -----------------------------------------------


def utf8_encode(text):
    u"""
    Конвертирует строку в кодировку utf-8
    """

    enc = chardet.detect(text).get("encoding")
    if enc and enc.lower() != "utf-8":
        try:
            text = text.decode(enc)
            text = text.encode("utf-8")
        except:
            pass
    else:
        pass

    return text
# -----------------------------------------------


def iconv_encode(text, encoding='utf-8'):
    enc = chardet.detect(text).get("encoding")
    if enc and enc.lower() != encoding:
        try:
            text = text.decode(enc)
            text = text.encode(encoding)
        except:
            pass
    else:
        pass

    return text
# -----------------------------------------------


def utf8_file_encode(filename, in_place=False):
    u"""
    Конвертирует файл в кодировку utf-8 и открывает копию 'wb'
    """

    filename = unicode_name(filename)

    with open(filename, 'rb') as F:
        t = F.read()
        text = utf8_encode(t)
        # text = text.replace('\r', '')
        F.close()

        if in_place:
            W = open(filename, 'wb')
        else:
            W = tempfile.TemporaryFile()

        W.write(text)
        W.seek(0)
        return W
# -----------------------------------------------


def get_type_sheet(filename):
    u"""
    Тип таблици (CSV, XLS, XLSX)
    """

    filename = unicode_name(filename)
    p = re.compile('html|plain|csv|xml|office|msword|excel|zip', re.IGNORECASE)

    try:
        proc = subprocess.Popen("/usr/bin/file --mime '%s'" % filename, shell=True, stdout=subprocess.PIPE)
        out = proc.stdout.readlines()
        S = " ".join(out).split(":")[1]
        S = S.lower()

        m = p.search(S).group()
    except Exception, e:
        return ''

    if m == 'html':
        t = 'HTML'
    elif m == 'plain' or m == 'csv':
        t = 'CSV'
    elif m == 'xml' or m == 'msword' or m == 'office':
        t = 'XLS'
    elif m == 'zip' or m == 'excel':
        t = 'XLSX'
    else:
        t = ''

    return t
# -----------------------------------------------


def get_writer(filename=False, windows=False):
    if type(filename) == str or type(filename) == unicode:
        W = open(filename, 'wb')
    elif type(filename) == file:
        W = filename
    elif filename == False:
        W = tempfile.TemporaryFile()
    else:
        raise ValueError("Не верный аргумент файла записи")
    # W = open('cp1251-utf8.csv', 'wb')

    quoting = csv.QUOTE_MINIMAL
    delimiter = ','
    lineterminator = '\n'
    encoding = 'utf-8'

    if windows == True:
        delimiter = ';'
        # lineterminator = '\r\n'
        # encoding = 'cp1251'

    writer = CSVUnicodeWriter(W, delimiter=delimiter, encoding=encoding, quoting=quoting, lineterminator=lineterminator)
    return writer
# -----------------------------------------------


def get_reader(filename, typesheet=None):
    filename = unicode_name(filename)
    if type(typesheet) == types.NoneType:
        typesheet = get_type_sheet(filename)

    if typesheet == 'CSV':
        f = utf8_file_encode(filename)
        reader = CSVUnicodeReader(f)
    elif typesheet == 'XLS':
        reader = XLSReader(filename)
    elif typesheet == 'XLSX':
        temp = tempfile.NamedTemporaryFile()
        proc = subprocess.Popen("/usr/local/bin/xlsx2csv '%s' '%s'" % (filename, temp.name), shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        err = proc.stderr.read()
        if err:
            try:
                reader = XLSReader(filename)
                return reader
            except Exception, e:
                pass

            raise ValueError("Не удалось определить формат файла")
        out = proc.stdout.read()

        f = open(temp.name, 'rb')
        reader = CSVUnicodeReader(f)
    else:
        raise ValueError("Не удалось определить формат файла")

    return reader
# -----------------------------------------------


def detect_dialect(f):
    u"""
    Определить формат разделителей CSV
    """

    try:
        f.seek(0)
        dialect = csv.Sniffer().sniff(f.read(), delimiters=';,')
    except Exception, e:
        f.seek(0)
        dialect=csv.excel
        dialect.lineterminator = '\n'
        row = f.read(1024)
        if len(re.compile(';').findall(row)) > len(re.compile(',').findall(row)):
            dialect.delimiter = ';'
    f.seek(0)
    return dialect
# -----------------------------------------------


def parse_xlsx(filename, output, windows=False):
    u"""
    Конвертирует XLSX в формат CSV
    """

    filename = unicode_name(filename)
    delimiter = ','

    if windows:
        delimiter = ';'

    if output == sys.stdout:
        proc = subprocess.Popen("/usr/local/bin/xlsx2csv -d '%s' '%s'" % (delimiter, filename), shell=True, stdout=sys.stdout, stderr=subprocess.PIPE)
    else:
        output = unicode_name(output)
        proc = subprocess.Popen("/usr/local/bin/xlsx2csv -d '%s' '%s' '%s' " % (delimiter, filename, output), shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    err = proc.stderr.read()
    if err:
        return False
    # out = proc.stdout.read()

    return True
# -----------------------------------------------


def convert_to_csv(filename, output, windows=False):
    if parse_xlsx(filename, output, windows=windows):
        return True

    try:
        reader = XLSReader(filename)
    except:
        try:
            f = utf8_file_encode(filename)
            reader = CSVUnicodeReader(f)
        except:
            raise ValueError("Не удалось преобразовать в CSV")

    writer = get_writer(output, windows=windows)
    writer.write_reader(reader)
    return True
# -----------------------------------------------


class UTF8Recoder:
    u"""
    Итератор, который читает кодированный поток и перекодирует вход для UTF-8
    """

    def __init__(self, f, encoding):
        self.reader = codecs.getreader(encoding)(f)

    def __iter__(self):
        return self

    def next(self):
        return self.reader.next().encode("utf-8")
# -----------------------------------------------


class CSVUnicodeReader:
    """
    CSV Reader, который будет перебирать строки в CSV файл "f",
    который кодируется в данной кодировке.
    """

    def __init__(self, f, encoding="utf-8", **kwds):
        dialect = detect_dialect(f)
        rec = UTF8Recoder(f, encoding)
        self.reader = csv.reader(rec, dialect=dialect, **kwds)

    def next(self):
        row = self.reader.next()
        return [unicode(s, "utf-8") for s in row]

    def __iter__(self):
        return self
# -----------------------------------------------


class CSVUnicodeWriter:
    u"""
    CSV Writer, который напишет строки в CSV файл "f",
    который закодирован в данной кодировке.
    """

    def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
        self.queue = cStringIO.StringIO()
        self.writer = csv.writer(self.queue, dialect=dialect, **kwds)
        self.stream = f
        self.encoder = codecs.getincrementalencoder(encoding)()

    def writerow(self, row):
        self.writer.writerow([s.encode("utf-8") for s in row])
        data = self.queue.getvalue()
        data = data.decode("utf-8")
        data = self.encoder.encode(data)
        self.stream.write(data)
        self.queue.truncate(0)

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)

    def write_reader(self, reader):
        for rowi, row in enumerate(reader):
            self.writerow(row)

    def get_file(self, seek=False):
        F = self.stream
        if type(seek) == int:
            F.seek(seek)

        return F
# -----------------------------------------------


class XLSReader:
    def __init__(self, f, **kwds):
        f = unicode_name(f)
        book = xlrd.open_workbook(f)
        self.sh = book.sheet_by_index(0)
        self.reader = self.get_reader()

    def get_reader(self):
        for rx in xrange(self.sh.nrows):
            row = []
            for cell in self.sh.row(rx):
                data = str(cell.value)
                row.append(data)
            yield row

    def next(self):
        row = self.reader.next()
        return [unicode(s, "utf-8") for s in row]

    def __iter__(self):
        return self
# -----------------------------------------------


class XLSWriter:
    def __init__(self, **kwds):
        self.wb = xlwt.Workbook(encoding='utf-8')
        self.ws = self.wb.add_sheet("Sheet1")

    def write_reader(self, reader):
        for rowi, row in enumerate(reader):
            if rowi == 0:
                self._firsRow(row)
                continue
            for coli, value in enumerate(row):
                value = value.decode('utf-8')
                self.ws.write(rowi, coli, value)
                # print(value)

    def _firsRow(self, row):
        for coli, value in enumerate(row):
            value = value.decode('utf-8')
            self.ws.write(0, coli, value, self._get_style())

    def _get_style(self):
        # Шрифт первой строчки
        fnt = xlwt.Font()
        fnt.name = 'Arial'
        fnt.bold = True

        borders = xlwt.Borders()
        borders.right = 0x1

        style = xlwt.XFStyle()
        style.font = fnt
        style.borders = borders

        return style

    def save(self, filename):
        filename = unicode_name(filename)
        self.wb.save(filename)

    def frozen(self):
        self.ws.panes_frozen = True
        self.ws.horz_split_pos = 1

