<?php
namespace SpreadsheetReader\tests;

use SpreadsheetReader\SpreadsheetReader;
use SpreadsheetReader\SpreadsheetReader_XLSX;

include_once WEB_ROOT . 'SpreadsheetReader.php';
include_once WEB_ROOT . 'SpreadsheetReader_XLSX.php';

class SpreadSheetReaderTest extends \PHPUnit_Framework_TestCase
{
	/**
	 * @var null|SpreadsheetReader
	 */
	protected $reader = null;

	public function setUp()
	{
		//$reader = new SpreadsheetReader(DATA_XLSX . 'sheets.xlsx');
		//$this->setReader($reader);
	}

	public function testNonExistingFileException()
	{

		//try {
		$fileUrl = DATA_XLSX . 'sheetsssss.xlsx';
		$this->setExpectedException('Exception', sprintf('SpreadsheetReader: File (%s) not readable', $fileUrl));

		new SpreadsheetReader($fileUrl);
	}

	/*public function testUKTypeException()
	{
		$this->setExpectedException('Exception', 'Unsupported file type specified');

		$fileUrl = DATA_XLSX . 'sheets.xlsx';
		new SpreadsheetReader($fileUrl, 'xxx.xxx');
	}*/

	public function testParameters()
	{
		$fileUrl = DATA_XLSX . 'sheets.xlsx';
		$reader = new SpreadsheetReader($fileUrl, 'xxx.xxx', 'text/csv');

		$this->assertEquals(
			SpreadsheetReader::TYPE_CSV,
			$reader->getType(),
			'Type should be CSV, but is ' . $reader->getType()
		);

		unset($reader);

		$reader = new SpreadsheetReader($fileUrl);

		$this->assertEquals(
			SpreadsheetReader::TYPE_XLSX,
			$reader->getType(),
			'Type should be XLSX, but is ' . $reader->getType()
		);

		unset($reader);

		$reader = new SpreadsheetReader($fileUrl, 'test.csv');

		$this->assertEquals(
			SpreadsheetReader::TYPE_CSV,
			$reader->getType(),
			'Type should be CSV, but is ' . $reader->getType()
		);

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
}