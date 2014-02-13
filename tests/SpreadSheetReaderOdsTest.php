<?php
namespace SpreadsheetReader\tests;

use SpreadsheetReader\SpreadsheetReader;
use SpreadsheetReader\SpreadsheetReader_ODS;

include_once WEB_ROOT . 'SpreadsheetReader.php';
include_once WEB_ROOT . 'SpreadsheetReader_ODS.php';

class SpreadSheetReaderOdsTest extends \PHPUnit_Framework_TestCase
{
	/**
	 * @var null|SpreadsheetReader
	 */
	protected $reader = null;

	public function setUp()
	{
		$reader = new SpreadsheetReader(DATA_ODS . 'sheets.ods');
		$this->setReader($reader);
	}

	public function testOdsSheets()
	{
		$reader = new SpreadsheetReader(DATA_ODS . 'sheets.ods');

		$sheets = $reader -> Sheets();

		$this->assertTrue(count($sheets) == 3,'Checking sheets count');

		$reader->ChangeSheet(1);
		$this->assertTrue($reader->GetSheetIndex() == 1, 'Checking sheet index');

		unset($reader);
	}

	public function testOdsSheetsIterations()
	{
		$reader = $this->getReader();

		$this->assertTrue($reader->GetSheetIndex() == 0, 'Initial');
		$reader->next();

		$this->assertTrue($reader->GetSheetIndex() == 1, 'After next');
		$reader->rewind();

		$this->assertTrue($reader->GetSheetIndex() == 0, 'After rewind');
	}

	public function testOdsRowsIterations()
	{
		$reader = $this->getReader();
		$reader-> ChangeSheet(1);
		$handle = $reader -> getHandle();
		/** @var array|SpreadsheetReader_ODS $handle */

		$array = $handle -> current();
		$this->assertTrue($array[0] == 2, 'Condition ' . $array[0] . ' == 2 failed');

		$reader -> rewind();

		$this->assertTrue($reader->GetSheetIndex() == 0, 'Is sheet index also set to 0');
		$this->assertTrue($handle->key() == 0, 'Is row index also set to 0');
		$array = $handle -> current();

		$this->assertTrue($array[0] == 1, 'Condition ' . $array[0] . ' == 1 failed');
		//print_r($array);

		$reader -> next();
		$this->assertTrue($reader->GetSheetIndex() == 1, 'Is sheet index also set to 1');
		$this->assertTrue($handle->key() == 0, 'Is row index should be 0 instead of ' . $handle->key());
		$array = $handle -> current();

		$this->assertTrue(isset($array[0]), 'Array value is not set');
		$this->assertTrue($array[0] == 2, 'Condition ' . $array[0] . ' == 2 failed');

		$reader -> ChangeSheet(2);
		$this->assertTrue($reader->GetSheetIndex() == 2, 'Is sheet index also set to 2, is ' . $reader->GetSheetIndex());
		$this->assertTrue($handle->key() == 0, 'Is row index should be 0 instead of ' . $handle->key());

		$array = $handle -> current();
		$this->assertTrue($array[0] == 3, 'Condition ' . $array[0] . ' == 3 failed');
	}

	public function testSeeking()
	{
		$reader = new SpreadsheetReader(DATA_ODS . 'seeking.ods');
		$handle = $reader -> getHandle();
		/** @var array|SpreadsheetReader_ODS $handle */

		$reader -> ChangeSheet(1);

		$this->assertEquals($reader->GetSheetIndex(), 1, 'Expected sheet position is 1, instead of ' . $reader->GetSheetIndex());
		$this->assertEquals($handle-> key(), 0, 'Expected sheet position is 0, instead of ' . $handle -> key());

		$reader -> seek(30);
		$array = $handle -> current();
		$this->assertEquals($array[0], 80, 'Expected value after seeking is 80 instead of ' . $array[0]);

		$sheets = $reader -> Sheets();
		$this->assertEquals(count($sheets), 2, 'seeking.xlsx contains 2 sheets instead of ' . count($sheets));

		$reader -> rewind();

		$reader -> seek(5);
		$array = $handle -> current();
		$this->assertEquals($array[0], 5, 'Expected value after seeking is 5 instead of ' . $array[0]);

		unset($reader);
	}

	public function testSeekingException()
	{
		$reader = new SpreadsheetReader(DATA_ODS . 'seeking.ods');
		$this->setExpectedException('Exception', 'Seek position out of range');

		$reader -> seek(1000);

		unset($reader);
	}

	/**
	 * @param null|\SpreadsheetReader\SpreadsheetReader $reader
	 */
	public function setReader($reader)
	{
		$this->reader = $reader;
	}

	/**
	 * @return null|\SpreadsheetReader\SpreadsheetReader
	 */
	public function getReader()
	{
		return $this->reader;
	}

	public function tearDown()
	{
		unset($this->reader);
	}
}