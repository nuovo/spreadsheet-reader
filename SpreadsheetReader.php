<?php
/**
 * Main class for spreadsheet reading
 *
 * @version 0.5.6
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

			// To avoid timezone warnings and exceptions for formatting dates retrieved from files
			$DefaultTZ = @date_default_timezone_get();
			if ($DefaultTZ)
			{
				date_default_timezone_set($DefaultTZ);
			}

			// 1. Determine type
			if (!$OriginalFilename)
			{
				$OriginalFilename = $Filepath;
			}

			$Extension = strtolower(pathinfo($OriginalFilename, PATHINFO_EXTENSION));

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
					if (in_array($Extension, array('csv', 'tsv', 'txt')))
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
					$this -> Type = self::TYPE_ODS;
					break;
				case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
				case 'application/vnd.openxmlformats-officedocument.spreadsheetml.template':
				case 'application/xlsx':
				case 'application/xltx':
					$this -> Type = self::TYPE_XLSX;
					break;
				case 'application/xml':
					// Excel 2004 xml format uses this
					break;
			}

			if (!$this -> Type)
			{
				switch ($Extension)
				{
					case 'xlsx':
					case 'xltx': // XLSX template
					case 'xlsm': // Macro-enabled XLSX
					case 'xltm': // Macro-enabled XLSX template
						$this -> Type = self::TYPE_XLSX;
						break;
					case 'xls':
					case 'xlt':
						$this -> Type = self::TYPE_XLS;
						break;
					case 'ods':
					case 'odt':
						$this -> Type = self::TYPE_ODS;
						break;
					default:
						$this -> Type = self::TYPE_CSV;
						break;
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
		 * Gets information about separate sheets in the given file
		 *
		 * @return array Associative array where key is sheet index and value is sheet name
		 */
		public function Sheets()
		{
			return $this -> Handle -> Sheets();
		}

		/**
		 * Changes the current sheet to another from the file.
		 *	Note that changing the sheet will rewind the file to the beginning, even if
		 *	the current sheet index is provided.
		 *
		 * @param int Sheet index
		 *
		 * @return bool True if sheet could be changed to the specified one,
		 *	false if not (for example, if incorrect index was provided.
		 */
		public function ChangeSheet($Index)
		{
			return $this -> Handle -> ChangeSheet($Index);
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
