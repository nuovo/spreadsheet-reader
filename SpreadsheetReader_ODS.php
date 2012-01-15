<?php
/**
 * Class for parsing ODS files
 *
 * @author Martins Pilsetnieks
 */
	class SpreadsheetReader_ODS implements Iterator, Countable
	{
		public function __construct($Filepath)
		{
			if (!is_readable($Filepath))
			{
				throw new Exception('SpreadsheetReader_ODS: File not readable ('.$Filepath.')');
			}
		}

		// !Iterator interface methods
		/** 
		 * Rewind the Iterator to the first element.
		 * Similar to the reset() function for arrays in PHP
		 */ 
		public function rewind()
		{
			if ($this -> Index > 0)
			{
				// If the worksheet was already iterated, XML file is reopened.
				// Otherwise it should be at the beginning anyway
				//$this -> Worksheet -> close();
				//$this -> Worksheet -> open($this -> WorksheetPath);
				//$this -> Valid = true;
			}

			$this -> Index = 0;
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
				$this -> next();
				$this -> Index--;
			}
			return $this -> CurrentRow;
		}

		/** 
		 * Move forward to next element. 
		 * Similar to the next() function for arrays in PHP 
		 */ 
		public function next()
		{
			$this -> Index++;
			$this -> CurrentRow = array();

			return $this -> CurrentRow;
		}

		/** 
		 * Return the identifying key of the current element.
		 * Similar to the key() function for arrays in PHP
		 *
		 * @return mixed either an integer or a string
		 */ 
		public function key()
		{
			return false;
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
			return $this -> Valid;
		}

		// !Countable interface method
		/**
		 * Ostensibly should return the count of the contained items but this just returns the number
		 * of rows read so far. It's not really correct but at least coherent.
		 */
		public function count()
		{
			return $this -> Index + 1;
		}
	}
?>