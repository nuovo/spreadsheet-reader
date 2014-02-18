<?php

namespace SpreadsheetReader;

abstract class AbstractSpreadsheetReader {
	/**
	 * @var int
	 */
	private $CurrentSheet = 0;

	/**
	 * @var array
	 */
	private $CurrentRow = array();

	/**
	 * @var int
	 */
	private $Index = 0;

	/**
	 * @param $index
	 * @return mixed
	 */
	abstract public function ChangeSheet($index);

	/**
	 * @return mixed
	 */
	abstract public function Sheets();

	/**
	 * @param array $CurrentRow
	 */
	public function setCurrentRow($CurrentRow)
	{
		$this->CurrentRow = $CurrentRow;
	}

	/**
	 * @return array
	 */
	public function getCurrentRow()
	{
		return $this->CurrentRow;
	}

	/**
	 * @param int $CurrentSheet
	 */
	public function setCurrentSheet($CurrentSheet)
	{
		$this->CurrentSheet = $CurrentSheet;
	}

	/**
	 * @return int
	 */
	public function getCurrentSheet()
	{
		return $this->CurrentSheet;
	}

	/**
	 * @param int $Index
	 */
	public function setIndex($Index)
	{
		$this->Index = $Index;
	}

	/**
	 * @return int
	 */
	public function getIndex()
	{
		return $this->Index;
	}
}