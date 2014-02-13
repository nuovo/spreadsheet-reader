<?php
namespace SpreadsheetReader\tests;

use SpreadsheetReader\SpreadsheetReader;
use SpreadsheetReader\SpreadsheetReader_CSV;

include_once WEB_ROOT . 'SpreadsheetReader.php';
include_once WEB_ROOT . 'SpreadsheetReader_CSV.php';

class SpreadSheetReaderCsvTest extends \PHPUnit_Framework_TestCase
{
	/**
	 * @var null|SpreadsheetReader
	 */
	protected $reader = null;

	public function setUp()
	{
		$reader = new SpreadsheetReader(DATA_CSV . 'plain.csv');
		$this->setReader($reader);
	}

	public function testCsvSheets()
	{
		$reader = $this->getReader();
		$sheets = $reader -> Sheets();

		$this->assertTrue(count($sheets) == 1, 'Wrong sheets count');

		//@todo implement out of boundary check
		//$reader->ChangeSheet(2);
		//$this->assertTrue($reader->GetSheetIndex() == 1, 'Checking sheet index');
	}

	public function testCsvRowsIterations()
	{
		$reader = $this->getReader();
		$reader-> ChangeSheet(0);
		$handle = $reader -> getHandle();
		/** @var array|SpreadsheetReader_CSV $handle */

		$array = $handle -> current();
		$this->assertTrue($array[0] == 1, 'Condition ' . $array[0] . ' == 1 failed');

		$handle -> next();
		$array = $handle -> current();
		$this->assertTrue($array[0] == 2, 'Condition ' . $array[0] . ' == 2 failed');

		$handle -> rewind();
		$array = $handle -> current();
		$this->assertTrue($array[0] == 1, 'Condition ' . $array[0] . ' == 1 failed');
	}

	public function testSeeking()
	{
		$reader = $this->getReader();
		$handle = $reader -> getHandle();
		/** @var array|SpreadsheetReader_CSV $handle */

		$this->assertEquals($reader->GetSheetIndex(), 0, 'Expected sheet position is 1, instead of ' . $reader->GetSheetIndex());
		$this->assertEquals($handle-> key(), 0, 'Expected sheet position is 0, instead of ' . $handle -> key());

		$reader -> seek(30);
		$array = $handle -> current();
		$this->assertEquals($array[0], 30, 'Expected value after seeking is 30 instead of ' . $array[0]);

		$reader -> rewind();

		$reader -> seek(4);
		$array = $handle -> current();
		$this->assertEquals($array[0], 4, 'Expected value after seeking is 5 instead of ' . $array[0]);
	}

	public function testSeekingException()
	{
		$reader = $this->getReader();
		$this->setExpectedException('Exception', 'Seek position out of range');

		$reader -> seek(1000);
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
}