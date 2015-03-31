#!/usr/bin/python
# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.1'

import chardet
import codecs
import cStringIO
import csv
import os
import os.path
import re
import subprocess
import sys
import tempfile
import types
import xlrd
import xlwt

# Кодировка по умолчанию
# reload(sys)
# sys.setdefaultencoding('utf-8')
# ------------------------------


def unicode_name(filename, encoding='utf-8'):
    u"""
    Преобразует имя файла в правельный формат
    """

    filename = filename.decode(encoding)
    return filename
# -----------------------------------------------


def utf8_encode(text, encoding='utf-8'):
    u"""
    Конвертирует строку в кодировку utf-8
    """

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


def utf8_open_file(filename):
    u"""
    Открывает файл для чтения в формате utf-8
    """

    filename = unicode_name(filename)
    encoding='utf-8'

    F = open(filename, 'rb')
    text = F.read()
    enc = chardet.detect(text).get("encoding")
    if enc and enc.lower() != encoding:
        try:
            text = text.decode(enc)
            text = text.encode(encoding)
            F.close()
            F = tempfile.TemporaryFile()
            F.write(text)
        except:
            raise RuntimeError("Не удалось преобразовать кодировку")
    F.seek(0)
    return F
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
    u"""
    Создает объект класса CSVUnicodeWriter для записи
    """

    if type(filename) == str or type(filename) == unicode:
        W = open(unicode_name(filename), 'wb')
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
    u"""
    Создает правельный объект класса CSVUnicodeReader для чтения CSV.
    При необходимост преобразует формат.
    """

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

            raise RuntimeError("Не удалось определить формат файла")
        out = proc.stdout.read()

        f = open(temp.name, 'rb')
        reader = CSVUnicodeReader(f)
    else:
        raise RuntimeError("Не удалось определить формат файла")

    return reader
# -----------------------------------------------


def detect_dialect(f):
    u"""
    Определить формат разделителей CSV
    """

    try:
        f.seek(0)
        dialect = csv.Sniffer().sniff(f.read(), delimiters=';,|\t')
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
    filename = unicode_name(filename)

    if not os.path.isfile(filename) or not os.access(filename, os.R_OK):
        raise RuntimeError("Либо файл отсутствует или не читается")

    if parse_xlsx(filename, output, windows=windows):
        return True

    try:
        reader = XLSReader(filename)
    except:
        try:
            f = utf8_open_file(filename)
            reader = CSVUnicodeReader(f)
        except:
            raise RuntimeError("Не удалось преобразовать в CSV")

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


class FitSheetWrapper(object):
    """Try to fit columns to max size of any entry.
    To use, wrap this around a worksheet returned from the
    workbook's add_sheet method, like follows:

        sheet = FitSheetWrapper(book.add_sheet(sheet_name))

    The worksheet interface remains the same: this is a drop-in wrapper
    for auto-sizing columns.
    """
    def __init__(self, sheet):
        self.sheet = sheet
        self.widths = dict()

    def write(self, r, c, label='', *args, **kwargs):
        self.sheet.write(r, c, label, *args, **kwargs)
        # width = str(label).__len__() / 0.02
        width = (str(label).__len__() + 1) * 200

        # 12.29 = 0x0d00 + 6 = 3334
        # 12.57  = 0x0d00 + 79 = 3470

        # 0.28 = 136 * x (x = 0.002)
        if width > self.widths.get(c, 0):
            # print "C: {0} = {1} ({2}) / {3}".format(c, label, str(label).__len__(), width)
            self.widths[c] = width
            self.sheet.col(c).width = width

    def get_sheet(self):
        return self.sheet

    def __getattr__(self, attr):
        return getattr(self.sheet, attr)


class XLSWriter:
    def __init__(self, sheetname=None, **kwds):
        self.head_style = None # Создаем новые стили
        if not sheetname:
            sheetname = "Sheet1"
        sheetname = utf8_encode(sheetname)

        self.book = xlwt.Workbook(encoding='utf-8')
        self.sheet = FitSheetWrapper(self.book.add_sheet(sheetname))

        # добавить новый цвет в палитре и установить RGB
        xlwt.add_palette_colour("custom_colour", 0x21)
        self.book.set_colour_RGB(0x21, 244, 236, 197) # FFF4ECC5 # border FFCCC085

        # Создаем новые стили
        # self.general_style = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
        self.general_style = xlwt.XFStyle()
        self.general_style.num_format_str = '@'

    def write_reader(self, reader):
        for rowi, row in enumerate(reader):
            if rowi == 0:
                self._firs_row(row)
                continue
            for coli, value in enumerate(row):
                value = value.decode('utf-8')
                self.sheet.write(rowi, coli, value, self.general_style)
                # print(value)

    def _firs_row(self, row):
        for coli, value in enumerate(row):
            value = value.decode('utf-8')
            self.sheet.write(0, coli, value, self._get_style())

    def _get_style(self):
        if self.head_style:
            return self.head_style

        # Шрифт первой строчки
        fnt = xlwt.Font()
        fnt.name = 'Arial'
        fnt.bold = True

        borders = xlwt.Borders()
        borders.right = 0x1
        borders.right_colour = 0x13

        style = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
        # style = xlwt.XFStyle()
        style.font = fnt
        style.borders = borders

        self.head_style = style
        return self.head_style

    def save(self, filename):
        try:
            filename = unicode_name(filename)
            self.book.save(filename)
        except:
            raise RuntimeError("Не удалось сохранить файл.")

    def frozen(self):
        sheet = self.sheet.get_sheet()
        sheet.panes_frozen = True
        sheet.horz_split_pos = 1
