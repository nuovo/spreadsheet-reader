<?php
/**
 * Class for parsing CSV files
 *
 * @author Martins Pilsetnieks
 */
 	class SpreadsheetReader_CSV implements Iterator, Countable
	{
		/**
		 * @var array Options array, pre-populated with the default values.
		 */
		private $Options = array(
			'Delimiter' => ';',
			'Enclosure' => '"'
		);

		private $Encoding = 'UTF-8';
		private $BOMLength = 0;

		/**
		 * @var resource File handle
		 */
		private $Handle = false;

		private $Filepath = '';

		private $Index = 0;

		private $CurrentRow = array();

		/**
		 * @param string Path to file
		 * @param array Options:
		 *	Enclosure => string CSV enclosure
		 *	Separator => string CSV separator
		 */
		public function __construct($Filepath, array $Options = null)
		{
			$this -> Filepath = $Filepath;

			if (!is_readable($Filepath))
			{
				throw new Exception('SpreadsheetReader_CSV: File not readable ('.$Filepath.')');
			}

			// For safety's sake
			@ini_set('auto_detect_line_endings', true);

			$this -> Options = array_merge($this -> Options, $Options);
			$this -> Handle = fopen($Filepath, 'r');

			// Checking the file for byte-order mark to determine encoding
			$BOM16 = bin2hex(fread($this -> Handle, 2));
			if ($BOM16 == 'fffe')
			{
				$this -> Encoding = 'UTF-16LE';
				//$this -> Encoding = 'UTF-16';
				$this -> BOMLength = 2;
			}
			elseif ($BOM16 == 'feff')
			{
				$this -> Encoding = 'UTF-16BE';
				//$this -> Encoding = 'UTF-16';
				$this -> BOMLength = 2;
			}

			if (!$this -> BOMLength)
			{			
				fseek($this -> Handle, 0);
				$BOM32 = bin2hex(fread($this -> Handle, 4));
				if ($BOM32 == '0000feff')
				{
					//$this -> Encoding = 'UTF-32BE';
					$this -> Encoding = 'UTF-32';
					$this -> BOMLength = 4;
				}
				elseif ($BOM32 == 'fffe0000')
				{
					//$this -> Encoding = 'UTF-32LE';
					$this -> Encoding = 'UTF-32';
					$this -> BOMLength = 4;
				}
			}

			fseek($this -> Handle, 0);
			$BOM8 = bin2hex(fread($this -> Handle, 3));
			if ($BOM8 == 'efbbbf')
			{
				$this -> Encoding = 'UTF-8';
				$this -> BOMLength = 3;
			}

			// Seeking the place right after BOM as the start of the real content
			if ($this -> BOMLength)
			{
				fseek($this -> Handle, $this -> BOMLength);
			}

			// Checking for the delimiter if it should be determined automatically
			if (!$this -> Options['Delimiter'])
			{
				// fgetcsv needs single-byte separators
				$Semicolon = ';';
				$Tab = "\t";
				$Comma = ',';

				// Reading the first row and checking if a specific separator character
				// has more columns than others (it means that most likely that is the delimiter).
				$SemicolonCount = count(fgetcsv($this -> Handle, null, $Semicolon));
				fseek($this -> Handle, $this -> BOMLength);
				$TabCount = count(fgetcsv($this -> Handle, null, $Tab));
				fseek($this -> Handle, $this -> BOMLength);
				$CommaCount = count(fgetcsv($this -> Handle, null, $Comma));
				fseek($this -> Handle, $this -> BOMLength);

				$Delimiter = $Semicolon;
				if ($TabCount > $SemicolonCount || $CommaCount > $SemicolonCount)
				{
					$Delimiter = $CommaCount > $TabCount ? $Comma : $Tab;
				}

				$this -> Options['Delimiter'] = $Delimiter;
			}
		}

		/**
		 * Returns information about sheets in the file.
		 * Because CSV doesn't have any, it's just a single entry.
		 *
		 * @return array Sheet data
		 */
		public function Sheets()
		{
			return array(0 => basename($this -> Filepath));
		}

		/**
		 * Changes sheet to another. Because CSV doesn't have any sheets
		 *	it just rewinds the file so the behaviour is compatible with other
		 *	sheet readers. (If an invalid index is given, it doesn't do anything.)
		 *
		 * @param bool Status
		 */
		public function ChangeSheet($Index)
		{
			if ($Index == 0)
			{
				$this -> rewind();
				return true;
			}
			return false;
		}

		// !Iterator interface methods
		/** 
		 * Rewind the Iterator to the first element.
		 * Similar to the reset() function for arrays in PHP
		 */ 
		public function rewind()
		{
			fseek($this -> Handle, $this -> BOMLength);
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
			// Finding the place the next line starts for UTF-16 encoded files
			// Line breaks could be 0x0D 0x00 0x0A 0x00 and PHP could split lines on the
			//	first or the second linebreak leaving unnecessary \0 characters that mess up
			//	the output.
			if ($this -> Encoding == 'UTF-16LE' || $this -> Encoding == 'UTF-16BE')
			{
				while (!feof($this -> Handle))
				{
					// While bytes are insignificant whitespace, do nothing
					$Char = ord(fgetc($this -> Handle));
					if (!$Char || $Char == 10 || $Char == 13)
					{
						continue;
					}
					else
					{
						// When significant bytes are found, step back to the last place before them
						if ($this -> Encoding == 'UTF-16LE')
						{
							fseek($this -> Handle, ftell($this -> Handle) - 1);
						}
						else
						{
							fseek($this -> Handle, ftell($this -> Handle) - 2);
						}
						break;
					}
				}
			}

			$this -> Index++;
			$this -> CurrentRow = fgetcsv($this -> Handle, null, $this -> Options['Delimiter'], $this -> Options['Enclosure']);

			if ($this -> CurrentRow)
			{
				// Converting multi-byte unicode strings
				// and trimming enclosure symbols off of them because those aren't recognized
				// in the relevan encodings.
				if ($this -> Encoding != 'ASCII' && $this -> Encoding != 'UTF-8')
				{
					$Encoding = $this -> Encoding;
					foreach ($this -> CurrentRow as $Key => $Value)
					{
						$this -> CurrentRow[$Key] = trim(trim(
							mb_convert_encoding($Value, 'UTF-8', $this -> Encoding),
							$this -> Options['Enclosure']
						));
					}

				}
			}

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
			return ($this -> CurrentRow || !feof($this -> Handle));
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