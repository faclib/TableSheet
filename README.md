# TableSheet


[![Build Status](https://travis-ci.org/faclib/TableSheet.svg?branch=master)](https://travis-ci.org/faclib/TableSheet)


Конвертор таблиц в стандартный CSV формат



Installation
-------------


Установка

    $ sudo apt-get install python python-pip
    $ sudo pip install --requirement=requirements.txt
    $ composer install 


### requirements.txt

    chardet==2.3.0
    xlrd==0.9.3
    xlsx2csv==0.7.1
    xlutils==1.7.1
    xlwt==0.7.5



## Usage 



### `toCSV()`


```php
Fobia\TableSheet::toCSV($pFilename, $outfile, $options)
```

**Parameters:**

> - `pFilename`   *String*  Имя читаемого файла.
> - `outfile`     *String*    Путь к записываемому файлу.
> - `options`     *Array*    Параметры.
> - - `delimiter`   *String (optional, default: `,`)* разделитель




### `toXLS()`


```php
Fobia\TableSheet::toXLS($csvFile, $outputFile, $options)
```

**Parameters:**

> - `csvFile`      *String*   Имя читаемого файла.
> - `output`       *String*   Путь к записываемому файлу.
> - `options`      *Array*    Параметры.
> - - `sheetname`    *String (optional, default: `Sheet1`)*  Название листа.
> - - `head_color`   *String (optional, default: `null`)*    Установить шапку в цвет (#F4ECC5) [red, yellow, blue]
> - - `forse`        *String (optional, default: `false`)*   Попытаться предварительно преобразовать формат файла



## Usage python

Синтаксис 

```
$ python convert-table.py <command> [options] <infile> <outfile>
```

, где:

> - `command`  *String (`csv`, `xls`)*  команда.
> - `infile`   *String*    входной-файл
> - `outfile`  *String*    выходной-файл (CSV, XLS)


**csv** -  конвертация в CSV

> - `--delimiter <D>` *String (optional, default: `,`)* 
> delimiter columns delimiter in csv (default: ',')


**xls** - конвертация в XLS

> - `--forse`     предварительно преобразовать в csv
> - `--sheetname <S>` *String (optional, default: `Sheet1`)*   имя сохраняемого листа
> - `--head`             заморозить шапку
> - `--color <C>` *String (optional)*   цвет фона шапки




Преобразовать в правельный CSV формат, разделитель ```,```

    $ python convert-table.py csv in.xls out.csv
    $ python convert-table.py csv --delimetr ';' in.xls out.csv


Преобразовать в excel таблицу

    $ python convert-table.py xls in.csv out.xls
    $ python convert-table.py xls --forse  in.oter out.xls
    $ python convert-table.py xls --head --color '#FFCC00' in.csv out.xls