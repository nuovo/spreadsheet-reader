<?php

namespace SpreadsheetReader\tests;

use SpreadsheetReader\SpreadsheetReader;

include_once WEB_ROOT . 'SpreadsheetReader.php';

class TestXlsImport extends \PHPUnit_Framework_TestCase
{
    /**
     * Simple tests, loads file and compare does library gets the same values
     */
    public function testOne()
    {
        $a = $this->getArrayFromFile('excel2003.xls');
        $b = array(
                array(
                0 => array(
                    0 => 'number',
                    1 => 'FIO'
                ),
                1 => array(
                    0 => 123123,
                    1 => 'Gordon Freeman'
                ),
                2 => array(
                    0 => 123456,
                    1 => 'Eli Vance'
                ),
        ));

		//echo "\n" . json_encode($a) . "\n";
		//echo json_encode($b);

        $this->assertTrue($this->arraysEqual($a, $b));
    }

    public function testTwo()
    {
        $a = $this->getArrayFromFile('file_test_blanks.xls');

        $sheet = 0;
        $b[$sheet][] = array(
            0 => 'Fruit',
            1 => 'Favorite',
            2 => 'Bowl',
            3 => 'Count',
        );

        $b[$sheet][] = array(
            0 => 'Apple',
            1 => 'Y',
            2 => 'Wood',
            3 => 2,
        );

        $b[$sheet][] = array(
            0 => 'Pear',
            1 => '',
            2 => 'Glass',
            3 => 6,
        );

        $b[$sheet][] = array(
            0 => 'Peach',
            1 => 'Y',
            2 => '',
            3 => 2,
        );

        //echo "\n" . json_encode($a) . "\n";
        //echo json_encode($b);

        $this->assertTrue($this->arraysEqual($a, $b));
    }

    public function testThree()
    {
        $a = $this->getArrayFromFile('file_test_blanks3.xls');

        $sheet = 0;
        $b[$sheet][] = array(
            0 => 'Fruit',
            1 => 'Favorite',
            2 => 'Bowl',
            3 => 'Count',
            4 => '',
        );

        $b[$sheet][] = array(
            0 => 'Apple',
            1 => 'Y',
            2 => 'Wood',
            3 => 2,
            4 => '',
        );

        $b[$sheet][] = array(
            0 => 'Pear',
            1 => '',
            2 => 'Glass',
            3 => 6,
            4 => '',
        );

        $b[$sheet][] = array(
            0 => 'Peach',
            1 => 'Y',
            2 => '',
            3 => 2,
            4 => '',
        );

        $b[$sheet][] = array(1 => '', 2 => '', 3 => '', 4 => '', 5 => '',);
        $b[$sheet][] = array(1 => '', 2 => '', 3 => '', 4 => '', 5 => '',);
        $b[$sheet][] = array(1 => '', 2 => '', 3 => '', 4 => '', 5 => '',);
        $b[$sheet][] = array(1 => '', 2 => '', 3 => '', 4 => '', 5 => '',);
        $b[$sheet][] = array(1 => '', 2 => '', 3 => '', 4 => '', 5 => '',);
        $b[$sheet][] = array(1 => '', 2 => '', 3 => '', 4 => '', 5 => '',);

//        echo "\n" . json_encode($a) . "\n";
//        echo json_encode($b);

        $this->assertTrue($this->arraysEqual($a, $b));
    }

    public function testFour()
    {
        $a = $this->getArrayFromFile('formula.xls');

        $sheet = 0;
        $b[$sheet][] = array(
            0 => 1,
            1 => 2,
            2 => 1,
            3 => 2,
            4 => 3,
            5 => 5,
        );

        /*$sheet = 1;
        $b[$sheet] = array(array());

        $sheet = 2;
        $b[$sheet] = array(array());*/

        //echo "\n" . json_encode($a) . "\n";
        //echo json_encode($b);

        $this->assertTrue($this->arraysEqual($a, $b));
    }

    public function testDataTypes()
    {
        $a = $this->getArrayFromFile('datatypes.xls');
        $sheet = 0;
        $b[$sheet][] = array(
            0 => 1,
            1 => '0001',
            2 => 'Some text',
            3 => 2.11,
            4 => true,
            5 => 'January 1, 2014',
            6 => '02:12:00 PM'
        );

//        echo "\n" . json_encode($a) . "\n";
//        echo json_encode($b);

        $this->assertTrue($this->arraysEqual($a, $b));
    }

    /**
     * @param $location
     * @return array
     */
    protected function getArrayFromFile($location)
    {
        $location = DATA_XLS . $location;

        $reader = new SpreadsheetReader($location);
        $sheets = $reader->Sheets();

        $a = array();
        foreach ($sheets as $index => $sheetName) {
            $reader->ChangeSheet($index);

            foreach ($reader as $row) {
                $a[$index][] = $row;
            }
        }

		unset($reader);

        return $a;
    }

    /**
     * Comparing two different arrays
     *
     * @param array $a
     * @param array $b
     * @return bool
     */
    protected function arraysEqual(array $a, array $b)
    {
        return (json_encode($a) === json_encode($b));
    }
}