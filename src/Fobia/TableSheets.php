<?php

/**
 * TableSheets class  - TableSheets.php file
 *
 * @author     Dmitriy Tyurin <fobia3d@gmail.com>
 * @copyright  Copyright (c) 2015 Dmitriy Tyurin
 */

namespace Fobia;

if (!class_exists('\\PHPExcel')) {
    if(!@include_once('PHPExcel.php')) {
        @include_once('phpexcel/PHPExcel.php');
    }
}

/**
 * TableSheets class
 *
 * @package   Fobia
 */
class TableSheets
{
    const CMD_TO_CSV = 'to-csv.py';

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

    public static function toCSV($pFilename, $outfile, $windows = false)
    {
        $cmd = sprintf("%s %s '%s' '%s'",
            self::getProg(self::CMD_TO_CSV),
            ($windows) ? '--windows' : '',
            $pFilename, $outfile);

        $res = shell_exec($cmd);
        if (!preg_match('/error/i', $res)) {
            return true;
        } else {
            echo $res;
        }

        return false;
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
