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
		$handle = $reader->getHandle();
		/** @var bool|SpreadsheetReader_ODS $handle */

		$sheets = $reader -> Sheets();

		$this->assertTrue(count($sheets) == 3,'Expected sheets count value is 3 instead fo ' . count($sheets));

		$reader->ChangeSheet(1);
		$this->assertTrue($handle->getCurrentSheet() == 1, 'Expected sheet index is 1, instead of ' . $handle->getCurrentSheet());

		unset($reader);
	}

	public function testOdsSheetsIterations()
	{
		$reader = $this->getReader();
		$handle = $reader->getHandle();
		/** @var bool|SpreadsheetReader_ODS $handle */

		$this->assertEquals($handle->getCurrentSheet(), 0, 'Expected sheet index is 0, instead of ' . $handle->getCurrentSheet());
		$reader->next();

		$this->assertEquals($handle->getCurrentSheet(), 0, 'Expected sheet index is 0, instead of ' . $handle->getCurrentSheet());
		$reader->rewind();

		$this->assertEquals($handle->getCurrentSheet(), 0, 'Expected sheet index is 0, instead of ' . $handle->getCurrentSheet());
		$reader->ChangeSheet(1);

		$this->assertEquals($handle->getCurrentSheet(), 1, 'Expected sheet index is 1, instead of ' . $handle->getCurrentSheet());
	}

	public function testOdsRowsIterations()
	{
		$reader = $this->getReader();
		$reader-> ChangeSheet(1);
		$handle = $reader -> getHandle();
		/** @var array|SpreadsheetReader_ODS $handle */

		$this->assertEquals($handle->getCurrentSheet(), 1, 'Expected sheet index is 0, instead of ' . $handle->getCurrentSheet());
		$this->assertEquals($handle->key(), 0, 'Expected row index is 0, instead of ' . $handle->key());

		$array = $handle -> current();
		$this->assertEquals($array[0], 2, 'Expected value is 2 instead of ' . $array[0]);

		//Rewinds only records, not sheets
		$reader -> rewind();

		$this->assertEquals($handle->getCurrentSheet(), 1, 'Expected sheet index is 0 instead of ' . $handle->getCurrentSheet());
		$this->assertEquals($handle->key(), 0, 'Expected row index is 0, instead of ' . $handle->key());
		$array = $handle -> current();

		$this->assertEquals($array[0], 2, 'Expected value is 2 instead of ' . $array[0]);
		//print_r($array);

		$reader -> next();
		$this->assertEquals($handle->getCurrentSheet(), 1, 'Expected sheet index is 1 instead of ' . $handle->getCurrentSheet());
		$this->assertEquals($handle->key(), 1, 'Expected row index should be 0 instead of ' . $handle->key());
		$array = $handle -> current();

		//$this->assertTrue(isset($array[0]), 'Array value is not set');
		//$this->assertEquals($array[0], 2, 'Expected value is 2 instead of ' . $array[0]);

		$reader -> ChangeSheet(2);
		$this->assertEquals($handle->getCurrentSheet(), 2, 'Expected sheet index is 2, instead of ' . $handle->getCurrentSheet());
		$this->assertEquals($handle->key(), 0, 'Expected row index should be 0 instead of ' . $handle->key());

		$array = $handle -> current();
		$this->assertEquals($array[0], 3, 'Expected value is 3 instead of ' . $array[0]);
	}

	public function testSeeking()
	{
		$reader = new SpreadsheetReader(DATA_ODS . 'seeking.ods');
		$handle = $reader -> getHandle();
		/** @var array|SpreadsheetReader_ODS $handle */

		$reader -> ChangeSheet(1);

		$this->assertEquals($handle->getCurrentSheet(), 1, 'Expected sheet position is 1, instead of ' . $handle->getCurrentSheet());
		$this->assertEquals($handle-> key(), 0, 'Expected row position is 0, instead of ' . $handle -> key());

		$reader -> seek(30);
		$array = $handle -> current();
		$this->assertEquals($array[0], 80, 'Expected value after seeking is 80 instead of ' . $array[0]);

		$sheets = $reader -> Sheets();
		$this->assertEquals(count($sheets), 2, 'seeking.ods contains 2 sheets instead of ' . count($sheets));

		$reader -> rewind();

		$reader -> seek(5);
		$array = $handle -> current();
		$this->assertEquals($array[0], 55, 'Expected value after seeking is 55 instead of ' . $array[0]);

		$reader -> ChangeSheet(0);

		$this->assertEquals($handle->getCurrentSheet(), 0, 'Expected sheet position is 0, instead of ' . $handle->getCurrentSheet());
		$this->assertEquals($handle-> key(), 0, 'Expected row position is 0, instead of ' . $handle -> key());

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