<?php

namespace Fobia;

class TableSheet2 extends TableSheet
{

    public static function getProg($prog)
    {
        return parent::getProg($prog);
    }

    public static function isWin()
    {
        return parent::isWin();
    }

}

define("DATA_DIR", dirname(__FILE__) . '/../../data');
define("FIXED_CSV", dirname(__FILE__) . '/../../data/fixed.csv');
define("TMP_CSV", dirname(__FILE__) . '/../../tmp-1.csv');

/**
 * Generated by PHPUnit_SkeletonGenerator on 2015-03-31 at 12:42:36.
 */
class TableSheetTest extends \PHPUnit_Framework_TestCase
{

    protected $head = array('Артикул', 'Наименование');
    protected function getHead()
    {
        $handle = fopen(TMP_CSV, 'r');
        $row = fgetcsv($handle);
        return array_slice($row, 0, 2);
    }

    protected function tearDown()
    {
        @unlink(TMP_CSV);
    }

    public function testIsWin()
    {
        $this->assertEquals(false, TableSheet2::isWin());
    }

    public function testGetProgn()
    {
        $p = sprintf("@^/usr/bin/python .*?%s csv$@", TableSheet::CMD_EXEC_FILE);
        $r = (bool) preg_match($p, TableSheet2::getProg('csv'));
        $this->assertEquals(true, $r);
    }

    /**
     * @dataProvider filesProvider
     */
    public function testAllFiles($file)
    {
        $f = DATA_DIR . '/' . $file;
        $this->assertEquals(true, TableSheet::toCSV($f, TMP_CSV));

//        echo file_get_contents(TMP_CSV);

        $this->assertEquals($this->head, $this->getHead());
    }

    public function filesProvider()
    {
        return array(
            array('файл.csv'),
            array('cp1251.csv'),
            array('excel2003.xls'),
            array('excel2007.xlsx'),
            array('fake.xlsx'),
            array('fixed.csv'),
            // array('long.xlsx'),
            array('table.html'),
        );
    }


    public function testToCsvOneRus()
    {
        $f = DATA_DIR . '/файл.csv';
        $this->assertEquals(true, TableSheet::toCSV($f, TMP_CSV));

        $l1 = file(TMP_CSV);
        $l2 = file(FIXED_CSV);

        $this->assertEquals($l1[0], $l2[0]);
        $this->assertEquals(file_get_contents(TMP_CSV), file_get_contents(FIXED_CSV));
    }   /**/

}
