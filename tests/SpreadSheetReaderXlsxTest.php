<?php
namespace SpreadsheetReader\tests;

use SpreadsheetReader\SpreadsheetReader;
use SpreadsheetReader\SpreadsheetReader_XLSX;

include_once WEB_ROOT . 'SpreadsheetReader.php';
include_once WEB_ROOT . 'SpreadsheetReader_XLSX.php';

class SpreadSheetReaderXlsxTest extends \PHPUnit_Framework_TestCase
{
	/**
	 * @var null|SpreadsheetReader
	 */
	protected $reader = null;

	public function setUp()
	{
		$reader = new SpreadsheetReader(DATA_XLSX . 'sheets.xlsx');
		$this->setReader($reader);
	}

	public function testXlsxSheets()
	{
		$reader = new SpreadsheetReader(DATA_XLSX . 'sheets.xlsx');

		$sheets = $reader -> Sheets();

		$this->assertTrue(count($sheets) == 3,'Checking sheets count');

		$reader->ChangeSheet(1);
		$this->assertTrue($reader->GetSheetIndex() == 1, 'Checking sheet index');
	}

	public function testXlsxSheetsIterations()
	{
		$reader = $this->getReader();

		$this->assertTrue($reader->GetSheetIndex() == 0, 'Initial');
		$reader->next();

		$this->assertTrue($reader->GetSheetIndex() == 1, 'After next');
		$reader->rewind();

		$this->assertTrue($reader->GetSheetIndex() == 0, 'After rewind');
	}

	public function testXlsxRowsIterations()
	{
		$reader = $this->getReader();
		$reader-> ChangeSheet(1);
		$handle = $reader -> getHandle();
		/** @var array|SpreadsheetReader_XLSX $handle */

		$array = $handle -> getCurrentRow();
		$this->assertTrue($array[0] == 2, 'Condition ' . $array[0] . ' == 2 failed');

		$reader -> rewind();

		$this->assertTrue($reader->GetSheetIndex() == 0, 'Is sheet index also set to 0');
		$this->assertTrue($handle->getIndex() == 0, 'Is row index also set to 0');
		$array = $handle -> getCurrentRow();

		$this->assertTrue($array[0] == 1, 'Condition ' . $array[0] . ' == 1 failed');
		//print_r($array);

		$reader -> next();
		$this->assertTrue($reader->GetSheetIndex() == 1, 'Is sheet index also set to 1');
		$this->assertTrue($handle->getIndex() == 0, 'Is row index should be 0 instead of ' . $handle->getIndex());
		$array = $handle -> getCurrentRow();

		$this->assertTrue(isset($array[0]), 'Array value is not set');
		$this->assertTrue($array[0] == 2, 'Condition ' . $array[0] . ' == 2 failed');

		$reader -> ChangeSheet(2);
		$this->assertTrue($reader->GetSheetIndex() == 2, 'Is sheet index also set to 2, is ' . $reader->GetSheetIndex());
		$this->assertTrue($handle->getIndex() == 0, 'Is row index should be 0 instead of ' . $handle->getIndex());

		$array = $handle -> getCurrentRow();
		$this->assertTrue($array[0] == 3, 'Condition ' . $array[0] . ' == 3 failed');
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