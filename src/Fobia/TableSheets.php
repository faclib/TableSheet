<?php

/**
 * TableSheets class  - TableSheets.php file
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
class TableSheets
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

    public static function toCSV($pFilename, $outfile, $windows = false)
    {
        $cmd = sprintf("%s %s '%s' '%s'",
            self::getProg(self::CMD_TO_CSV),
            ($windows) ? '--windows' : '',
            $pFilename, $outfile);

        $res = shell_exec($cmd);
        if (!preg_match('/error:/i', $res)) {
            return true;
        } else {
            // echo $res;
        }

        return false;
    }
    
    /**
     * Конвертирует CSV в XML с помощью python утилит
     * 
     * @param string $csvFile     входной CSV файл правельного формата
     * @param string $outputFile  сохраняемый файл
     * @param string $sheetname   название листа
     * @return boolean
     */
    public static function toXls($csvFile, $outputFile, $sheetname = "Sheet1")
    {
        $cmd = sprintf("%s --sheetname '%s' '%s' '%s'",
            self::getProg(self::CMD_TO_XLS),
            $sheetname,
            $csvFile, $outputFile);
        $res = shell_exec($cmd);
        
        if (!preg_match('/error:/i', $res)) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Конвертирует CSV файл в XLS(Excel) формат
     *
     * @param type $csvFile
     * @param type $outputFile
     */
    public static function toXls2($csvFile, $outputFile = null, $options = array())
    {
        if (!class_exists('\\PHPExcel')) {
            if(!@include_once('PHPExcel.php')) {
                @include_once('phpexcel/PHPExcel.php');
            }
            if (!class_exists('\\PHPExcel', false)) {
                throw new \RuntimeException("Не удалось загрузить класс 'PHPExcel'");
            }
        }

        if ($outputFile === null) {
            $tmp = "/tmp";
            if (isset($_SERVER['tmp'])) {
                if (is_dir($_SERVER['tmp'])) {
                    $tmp = $_SERVER['tmp'];
                }
            }
            $outputFile = tempnam($tmp, "excel_");
        }

        $objReader = new \PHPExcel_Reader_CSV();
        $objReader->setDelimiter(",");
        $objReader->setEnclosure('"');

        $objPHPExcel = $objReader->load($csvFile);
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->setTitle('Sheet1');
        $objPHPExcel->getDefaultStyle()->getFont()
            ->setName('Arial')
            ->setSize(10);

        $objPHPExcel->getActiveSheet()->calculateColumnWidths();

        $rows =  $objPHPExcel->getActiveSheet()->getHighestRow();
        $columns =  $objPHPExcel->getActiveSheet()->getHighestColumn();
        $eCell = $columns.$rows;

        for ($i = \PHPExcel_Cell::columnIndexFromString($columns); $i >= 0; $i--) {
            $objPHPExcel->getActiveSheet()->getColumnDimensionByColumn($i)->setAutoSize(true);
        }

        // $objPHPExcel->getActiveSheet()->getStyle("A1:".$eCell)->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle("A1:".$eCell)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_TOP);
        $objPHPExcel->getActiveSheet()->getStyle("A1:".$eCell)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);


        // HEAD
        $headRow = $objPHPExcel->getActiveSheet()->getStyle('A1:'.$columns.'1');
        $headRow->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $headRow->getFont()->setBold(true);
        // $headRow->getFont()->setColor($pValue);

        $color = array(
            'red' => array(
                'color' => 'FF9C0006',
                'fill' => 'FFFFC7CE',
                'borders' => 'FFD99795',
                ),
            'yellow' => array(
                'color' => 'FF000000',
                'fill' => 'FFF4ECC5',
                'borders' => 'FFCCC085',
                ),
        );
        $headRow->applyFromArray(
                array(
                    'fill'    => array(
                        'type'  => \PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('argb' => 'FFF4ECC5')
                    ),
                    'borders' => array(
                        'inside' => array('style' => \PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => 'FFCCC085')),
                        'bottom' => array('style' => \PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => 'FFCCC085')),
                        'right'  => array('style' => \PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => 'FFCCC085'))
                    )
                )
        );
        $objPHPExcel->getActiveSheet()->freezePane('A2');
        $objPHPExcel->getActiveSheet()->getProtection()->setSheet(true);
        $objPHPExcel->getActiveSheet()
            ->getStyle('A2:'.$eCell)
            ->getProtection()->setLocked(
                \PHPExcel_Style_Protection::PROTECTION_UNPROTECTED
            );
        $objPHPExcel->getActiveSheet()
            ->getStyle('A2')
            ->getProtection()->setLocked(
                \PHPExcel_Style_Protection::PROTECTION_UNPROTECTED
            );
        /**/
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save($outputFile);

        $objPHPExcel->disconnectWorksheets();

        return $outputFile;
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
