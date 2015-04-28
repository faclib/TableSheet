# TableSheet

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


    Fobia\TableSheet::toCSV('input.xls', 'out.csv');
    Fobia\TableSheet::toXLS('input.csv', 'out.xls');



## Usage python

Синтаксис 

    $ python convert-table.py <command> [options] <infile> <outfile>

, где:

- {xls,csv}             команда
- infile                входной-файл
- outfile               выходной-файл (CSV, XLS)


**csv** -  конвертация в CSV

- ```--delimiter <D>``` - delimiter columns delimiter in csv (default: ',')

**xls** - конвертация в XLS

- ```--forse```          - предварительно преобразовать в csv
- ```--sheetname <S>```  - имя сохраняемого листа
- ```--head```           - заморозить шапку
- ```--color <C>```      - цвет фона шапки




Преобразовать в правельный CSV формат, разделитель ```,```

    $ python convert-table.py csv in.xls out.csv
    $ python convert-table.py csv --delimetr ';' in.xls out.csv


Преобразовать в excel таблицу

    $ python convert-table.py xls in.csv out.xls
    $ python convert-table.py xls --forse  in.oter out.xls
    $ python convert-table.py xls --head --color '#FFCC00' in.csv out.xls