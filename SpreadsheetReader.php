<?php
/**
 * Main class for spreadsheet reading
 *
 * @version 0.4.0
 * @author Martins Pilsetnieks
 */
	class SpreadsheetReader implements Iterator, Countable
	{
		const TYPE_XLSX = 'XLSX';
		const TYPE_XLS = 'XLS';
		const TYPE_CSV = 'CSV';
		const TYPE_ODS = 'ODS';

		private $Options = array(
			'Delimiter' => '',
			'Enclosure' => '"'
		);

		/**
		 * @var int Current row in the file
		 */
		private $Index = 0;

		/**
		 * @var SpreadsheetReader_* Handle for the reader object
		 */
		private $Handle = array();

		/**
		 * @var TYPE_* Type of the contained spreadsheet
		 */
		private $Type = false;

		/**
		 * @param string Path to file
		 * @param string Original filename (in case of an uploaded file), used to determine file type, optional
		 * @param string MIME type from an upload, used to determine file type, optional
		 */
		public function __construct($Filepath, $OriginalFilename = false, $MimeType = false)
		{
			if (!is_readable($Filepath))
			{
				throw new Exception('SpreadsheetReader: File ('.$Filepath.') not readable');
			}

			// 1. Determine type
			if (!$OriginalFilename)
			{
				$OriginalFilename = $Filepath;
			}

			switch ($MimeType)
			{
				case 'text/csv':
				case 'text/comma-separated-values':
				case 'text/plain':
					$this -> Type = self::TYPE_CSV;
					break;
				case 'application/vnd.ms-excel':
				case 'application/msexcel':
				case 'application/x-msexcel':
				case 'application/x-ms-excel':
				case 'application/vnd.ms-excel':
				case 'application/x-excel':
				case 'application/x-dos_ms_excel':
				case 'application/xls':
				case 'application/xlt':
				case 'application/x-xls':
					// Excel does weird stuff
					$Extension = substr($OriginalFilename, -4);
					if (in_array($Extension, array('.csv', '.tsv', '.txt')))
					{
						$this -> Type = self::TYPE_CSV;
					}
					else
					{
						$this -> Type = self::TYPE_XLS;
					}
					break;
				case 'application/vnd.oasis.opendocument.spreadsheet':
				case 'application/vnd.oasis.opendocument.spreadsheet-template':
					$this -> Mode = self::TYPE_ODS;
					break;
				case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
				case 'application/vnd.openxmlformats-officedocument.spreadsheetml.template':
				case 'application/xlsx':
				case 'application/xltx':
					$this -> Mode = self::TYPE_XLSX;
					break;
				case 'application/xml':
					// Excel 2004 xml format uses this
					break;
			}

			if (!$this -> Type)
			{
				if (substr($OriginalFilename, -5) == '.xlsx' || substr($OriginalFilename, -5) == '.xltx')
				{
					$this -> Type = self::TYPE_XLSX;
					$Extension = '.xlsx';
				}
				else
				{
					$Extension = substr($OriginalFilename, -4);
				}

				if ($Extension == '.xls' || $Extension == '.xlt')
				{
					$this -> Type = self::TYPE_XLS;
				}
				elseif ($Extension == '.ods' || $Extension == '.odt')
				{
					$this -> Type = self::TYPE_ODS;
				}
				elseif (!$this -> Type)
				{
					// If type cannot be determined, try parsing as CSV, it might just be a text file
					$this -> Type = self::TYPE_CSV;
				}
			}

			// Pre-checking XLS files, in case they are renamed CSV or XLSX files
			if ($this -> Type == self::TYPE_XLS)
			{
				self::Load(self::TYPE_XLS);
				$this -> Handle = new SpreadsheetReader_XLS($Filepath);
				if ($this -> Handle -> Error)
				{
					$this -> Handle -> __destruct();

					if (is_resource($ZipHandle = zip_open($Filepath)))
					{
						$this -> Type = self::TYPE_XLSX;
						zip_close($ZipHandle);
					}
					else
					{
						$this -> Type = self::TYPE_CSV;
					}
				}
			}

			// 2. Create handle
			switch ($this -> Type)
			{
				case self::TYPE_XLSX:
					self::Load(self::TYPE_XLSX);
					$this -> Handle = new SpreadsheetReader_XLSX($Filepath);
					break;
				case self::TYPE_CSV:
					self::Load(self::TYPE_CSV);
					$this -> Handle = new SpreadsheetReader_CSV($Filepath, $this -> Options);
					break;
				case self::TYPE_XLS:
					// Everything already happens above
					break;
				case self::TYPE_ODS:
					self::Load(self::TYPE_ODS);
					$this -> Handle = new SpreadsheetReader_ODS($Filepath, $this -> Options);
					break;
			}
		}

		/**
		 * Autoloads the required class for the particular spreadsheet type
		 *
		 * @param TYPE_* Spreadsheet type, one of TYPE_* constants of this class
		 */
		private static function Load($Type)
		{
			if (!in_array($Type, array(self::TYPE_XLSX, self::TYPE_XLS, self::TYPE_CSV, self::TYPE_ODS)))
			{
				throw new Exception('SpreadsheetReader: Invalid type ('.$Type.')');
			}

			if (!class_exists('SpreadsheetReader_'.$Type))
			{
				require(dirname(__FILE__).DIRECTORY_SEPARATOR.'SpreadsheetReader_'.$Type.'.php');
			}
		}

		// !Iterator interface methods

		/** 
		 * Rewind the Iterator to the first element.
		 * Similar to the reset() function for arrays in PHP
		 */ 
		public function rewind()
		{
			$this -> Index = 0;
			if ($this -> Handle)
			{
				$this -> Handle -> rewind();
			}
		}

		/** 
		 * Return the current element.
		 * Similar to the current() function for arrays in PHP
		 *
		 * @return mixed current element from the collection
		 */
		public function current()
		{
			if ($this -> Handle)
			{
				return $this -> Handle -> current();
			}
			return null;
		}

		/** 
		 * Move forward to next element. 
		 * Similar to the next() function for arrays in PHP 
		 */ 
		public function next()
		{
			if ($this -> Handle)
			{
				$this -> Index++;

				return $this -> Handle -> next();
			}
			return null;
		}

		/** 
		 * Return the identifying key of the current element.
		 * Similar to the key() function for arrays in PHP
		 *
		 * @return mixed either an integer or a string
		 */ 
		public function key()
		{
			if ($this -> Handle)
			{
				return $this -> Handle -> key();
			}
			return null;
		}

		/** 
		 * Check if there is a current element after calls to rewind() or next().
		 * Used to check if we've iterated to the end of the collection
		 *
		 * @return boolean FALSE if there's nothing more to iterate over
		 */ 
		public function valid()
		{
			if ($this -> Handle)
			{
				return $this -> Handle -> valid();
			}
			return false;
		}

		// !Countable interface method
		public function count()
		{
			if ($this -> Handle)
			{
				return $this -> Handle -> count();
			}
			return 0;
		}
	}
?>