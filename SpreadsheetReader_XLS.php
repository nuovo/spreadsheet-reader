<?php

namespace SpreadsheetReader;

use SpreadsheetReader\xls\XlsReader as XlsReader;

include_once 'xls/XlsReader.php';

class SpreadsheetReader_XLS extends AbstractSpreadsheetReader implements \Iterator, \Countable {
    /**
     * @var array Options array, pre-populated with the default values.
     */
    private $Options = array(
    );

    /**
     * @var resource File handle
     */
    private $Handle = false;

    private $Index = 0;

    private $Error = false;

    /**
     * @var array Sheet information
     */
    private $Sheets = false;

    /**
     * @var array
     */
    private $SheetIndexes = array();

    /**
     * @var int Current sheet index
     */
    private $CurrentSheet = 0;

    /**
     * @var array Content of the current row
     */
    private $CurrentRow = array();

    /**
     * @var int Column count in the sheet
     */
    private $ColumnCount = 0;
    /**
     * @var int Row count in the sheet
     */
    private $RowCount = 0;

    private $data = null;

    /**
     * @var array Template to use for empty rows. Retrieved rows are merged
     *	with this so that empty cells are added, too
     */
    private $EmptyRow = array();

    /**
     * @param $Filepath
     * @param array $Options
     * @throws \Exception
     */
    public function __construct($Filepath, array $Options = null)
    {
        $this -> Handle = new XlsReader($Filepath, false, 'UTF-8');

        if (function_exists('mb_convert_encoding'))
        {
            $this -> Handle -> setUTFEncoder('mb');
        }

        $sheets = $this -> Handle -> getSheets();
        if (empty($sheets))
        {
            $this -> Error = true;
            return null;
        }

        $this -> ChangeSheet(0);
    }

    public function __destruct()
    {
        unset($this -> Handle);
    }

    /**
     * Retrieves an array with information about sheets in the current file
     *
     * @return array List of sheets (key is sheet index, value is name)
     */
    public function Sheets()
    {
        if ($this -> Sheets === false)
        {
            $this -> Sheets = array();
            $sheets = $this -> Handle -> getSheets();

            $this -> SheetIndexes = array_keys($sheets);

            foreach ($this -> SheetIndexes as $SheetIndex)
            {
                $boundSheet = $this -> Handle -> getBoundsheets();
                $this -> Sheets[] = $boundSheet [$SheetIndex]['name'];
            }
        }

        return $this -> Sheets;
    }

    /**
     * Changes the current sheet in the file to another
     *
     * @param int Sheet index
     *
     * @return bool True if sheet was successfully changed, false otherwise.
     */
    public function ChangeSheet($Index)
    {
        $Index = (int)$Index;
        $Sheets = $this -> Sheets();
		$this->setIndex(0);

        if (isset($Sheets[$Index])) {
			$this->setCurrentSheet($Index);

            $this -> rewind();
            $this -> CurrentSheet = $this -> SheetIndexes[$Index];

            $s = $this-> Handle -> getSheets();

            $this -> ColumnCount = $s[$this -> CurrentSheet]['numCols'];
            $this -> RowCount = $s[$this -> CurrentSheet]['numRows'];

            if ($this -> ColumnCount) {
                $this -> EmptyRow = array_fill(1, $this -> ColumnCount, '');
            } else {
                $this -> EmptyRow = array();
            }

            //Unset data
            unset($this->data);

            return true;
        } else {
            return false;
        }
    }

    public function __get($Name)
    {
        switch ($Name)
        {
            case 'Error':
                return $this -> Error;
                break;
        }
        return null;
    }

    // !Iterator interface methods
    /**
     * Rewind the Iterator to the first element.
     * Similar to the reset() function for arrays in PHP
     */
    public function rewind()
    {
        $this -> Index = 0;

		//Will be sheet offset used for this
		$this -> Handle -> setStartPosition(null);
    }

    /**
     * Return the current element.
     * Similar to the current() function for arrays in PHP
     *
     * @return mixed current element from the collection
     */
    public function current()
    {
        if ($this -> Index == 0)
        {
            $this -> readRecord();
			//$this -> Index--;
        }

        return $this -> CurrentRow;
    }

    /**
     * Move forward to next element.
     * Similar to the next() function for arrays in PHP
     */
    public function next()
    {
        // Internal counter is advanced here instead of the if statement
        //	because apparently it's fully possible that an empty row will not be
        //	present at all
        $this -> Index++;

		return $this->readRecord();
    }

	protected function readRecord()
	{
		if (!isset($this->data[$this->Index])) {
			$this->data = $this -> Handle -> getRowValuesByIndex(
				$this -> CurrentSheet,
				$this -> Index
			);
		}

		if ($this -> Error) {
			return array();
		} else {
			if (!isset($this->data[$this->Index])) {
				$this -> CurrentRow = $this -> EmptyRow;
			} else {
				$this -> CurrentRow = $this->data[$this->Index] + $this -> EmptyRow;
				ksort($this -> CurrentRow);

				$this -> CurrentRow = array_values($this -> CurrentRow);
			}

			return $this -> CurrentRow;
		}
	}

    /**
     * Return the identifying key of the current element.
     * Similar to the key() function for arrays in PHP
     *
     * @return mixed either an integer or a string
     */
    public function key()
    {
        return $this -> Index;
    }

    /**
     * Check if there is a current element after calls to rewind() or next().
     * Used to check if we've iterated to the end of the collection
     *
     * @return boolean FALSE if there's nothing more to iterate over
     */
    public function valid()
    {
        if ($this -> Error)
        {
            return false;
        }

		return ($this -> Index + 1 <= $this -> RowCount);
    }

    // !Countable interface method
    /**
     * Ostensibly should return the count of the contained items but this just returns the number
     * of rows read so far. It's not really correct but at least coherent.
     */
    public function count()
    {
        if ($this -> Error)
        {
            return 0;
        }

        return $this -> RowCount;
    }
}