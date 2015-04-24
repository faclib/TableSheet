<?php

/**
 * TableSheet class  - TableSheet.php file
 *
 * @author     Dmitriy Tyurin <fobia3d@gmail.com>
 * @copyright  Copyright (c) 2015 Dmitriy Tyurin
 */

namespace Fobia;

/**
 * TableSheets class
 *
 * @package   Fobia
 */
class TableSheet
{
    const CMD_TO_CSV = 'to-csv.py';
    const CMD_TO_XLS = 'to-xls.py';

    /**
     * Определяет формат таблицы
     *      - Excel2007
     *      - Excel5
     *      - Excel2003XML
     *      - (-) OOCalc
     *      - (-) SYLK
     *      - (-) Gnumeric
     *      - HTML
     *      - CSV
     *
     * @param string $pFilename
     * @return string|FALSE Формат файла
     */
    public static function identifyType($pFilename)
    {
        $result = shell_exec("/usr/bin/file --mime '$pFilename'");
        if (preg_match("/: (.*?)\/(.*?);/", $result, $m)) {
            $type = $m[2];
        } else {
            return false;
        }

        // $type = pathinfo($pFilename, PATHINFO_EXTENSION);

        if (preg_match('/html|plain|csv|xml|office|msword|excel|zip/', $type, $m)) {
            switch ($m[0]) {
                case 'html':      return 'HTML';
                case 'plain':
                case 'csv':       return 'CSV';
                case 'xml':       return 'Excel2003XML';
                case 'msword':
                case 'office':    return 'Excel5';
                case 'zip':
                case 'excel':  return 'Excel2007';
            }
        }
        return false;
    }

    /**
     * Конвертирует в фиксированый CSV формат
     *
     * @param string $pFilename
     * @param string $outfile
     * @return boolean
     */
    public static function toCSV($pFilename, $outfile)
    {
        $cmd = sprintf("%s '%s' '%s'",
            self::getProg(self::CMD_TO_CSV),
            $pFilename, $outfile);
        $res = shell_exec($cmd);
        if (!preg_match('/error:/i', $res)) {
            return true;
        } else {
            trigger_error('Не удалось конфертировать в CSV. ' . $res, E_WARNING);
        }

        return false;
    }

    /**
     * Конвертирует CSV в XML с помощью python утилит
     *
     * @param string $csvFile     входной CSV файл правельного формата
     * @param string $outputFile  сохраняемый файл
     * @param string $sheetname   название листа
     * @param string $head_color  установить шапку в цвете (#F4ECC5)
     * @return boolean
     */
    public static function toXls($csvFile, $outputFile, $sheetname = "Sheet1", $head_color = null)
    {
        $head = '';
        if ($head_color) {
            $head = "--head ";
            if (substr($head_color, 0, 1) == '#') {
                $head .= "--color '{$head_color}'";
            }
        }
        $cmd = sprintf("%s %s --sheetname '%s' '%s' '%s'",
            self::getProg(self::CMD_TO_XLS),
            $head,
            $sheetname,
            $csvFile, $outputFile);
        $res = shell_exec($cmd);

        if (!preg_match('/error:/i', $res)) {
            return true;
        } else {
            trigger_error('Не удалось конфертировать в XLS. ' . $res, E_WARNING);
            return false;
        }
    }


    /**
     * Возвращает команду для выполнения
     *    python prog
     *
     * @param string $prog
     * @return string
     * @throws \RuntimeException
     */
    protected static function getProg($prog)
    {
        $cmd = dirname(dirname(dirname(__FILE__))) . '/lib/' . $prog;
        if (!file_exists($cmd)) {
            throw new \RuntimeException("Не найден файл исполняемого скрипта '{$cmd}'");
        }

        // return sprintf("%s %s", (!self::isWin()) ? "/usr/bin/python" : "python", $cmd);
        return "/usr/bin/python " . $cmd;
    }

    /**
     * Является ли система ядра Windows
     *
     * @return boolean
     */
    protected static function isWin()
    {
        static $win = null;
        if ($win === null) {
            $uname = php_uname( 's' );
            if ( substr( $uname, 0, 7 ) == 'Windows' || preg_match("/win/i", $uname) ) {
                // $os = 'Windows';
                $win = true;
            } else {
                $win = false;
            }
        }
        return $win;
    }
}