<?php
/**
 * Class for parsing XLSX files specifically
 *
 * @author Martins Pilsetnieks
 */
	class SpreadsheetReader_XLSX implements Iterator, Countable
	{
		const CELL_TYPE_BOOL = 'b';
		const CELL_TYPE_NUMBER = 'n';
		const CELL_TYPE_ERROR = 'e';
		const CELL_TYPE_SHARED_STR = 's';
		const CELL_TYPE_STR = 'str';
		const CELL_TYPE_INLINE_STR = 'inlineStr';

		private $Options = array(
			'TempDir' => '',
			'ReturnDateTimeObjects' => false
		);

		private static $RuntimeInfo = array(
			'GMPSupported' => false
		);

		private $Valid = false;

		/**
		 * @var SpreadsheetReader_* Handle for the reader object
		 */
		private $Handle = false;

		// Worksheet file
		/**
		 * @var string Path to the worksheet XML file
		 */
		private $WorksheetPath = false;
		/**
		 * @var XMLReader XML reader object for the worksheet XML file
		 */
		private $Worksheet = false;

		// Shared strings file
		/**
		 * @var string Path to shared strings XML file
		 */
		private $SharedStringsPath = false;
		/**
		 * @var XMLReader XML reader object for the shared strings XML file
		 */
		private $SharedStrings = false;

		// Style data
		/**
		 * @var SimpleXMLElement XML object for the styles XML file
		 */
		private $StylesXML = false;
		/**
		 * @var array Container for cell value style data
		 */
		private $Styles = array();

		private $TempDir = '';

		private $CurrentRow = array();

		// Runtime parsing data
		/**
		 * @var int Current row in the file
		 */
		private $Index = 0;

		private $SharedStringCount = 0;
		private $SharedStringIndex = 0;
		private $LastSharedStringValue = null;

		private $RowOpen = false;
		private $CellOpen = false;
		private $ValueOpen = false;

		private $SSOpen = false;
		private $SSForwarded = false;

		private static $BuiltinFormats = array(
			1 => '0',
			2 => '0.00',
			3 => '#,##0',
			4 => '#,##0.00',

			9 => '0%',
			10 => '0.00%',
			11 => '0.00E+00',
			12 => '# ?/?',
			13 => '# ??/??',
			14 => 'mm-dd-yy',
			15 => 'd-mmm-yy',
			16 => 'd-mmm',
			17 => 'mmm-yy',
			18 => 'h:mm AM/PM',
			19 => 'h:mm:ss AM/PM',
			20 => 'h:mm',
			21 => 'h:mm:ss',
			22 => 'm/d/yy h:mm',

			37 => '#,##0 ;(#,##0)',
			38 => '#,##0 ;[Red](#,##0)',
			39 => '#,##0.00;(#,##0.00)',
			40 => '#,##0.00;[Red](#,##0.00)',

			45 => 'mm:ss',
			46 => '[h]:mm:ss',
			47 => 'mmss.0',
			48 => '##0.0E+0',
			49 => '@',

			// CHT & CHS
			27 => '[$-404]e/m/d',
			30 => 'm/d/yy',
			36 => '[$-404]e/m/d',
			50 => '[$-404]e/m/d',
			57 => '[$-404]e/m/d',

			// THA
			59 => 't0',
			60 => 't0.00',
			61 =>'t#,##0',
			62 => 't#,##0.00',
			67 => 't0%',
			68 => 't0.00%',
			69 => 't# ?/?',
			70 => 't# ??/??'
		);
		private $Formats = array();

		private static $DateReplacements = array(
			'All' => array(
				'\\' => '',
				'am/pm' => 'A',
				'yyyy' => 'Y',
				'yy' => 'y',
				'mmmmm' => 'M',
				'mmmm' => 'F',
				'mmm' => 'M',
				':mm' => ':i',
				'mm' => 'm',
				'm' => 'n',
				'dddd' => 'l',
				'ddd' => 'D',
				'dd' => 'd',
				'd' => 'j',
				'ss' => 's',
				'.s' => ''
			),
			'24H' => array(
				'hh' => 'H',
				'h' => 'G'
			),
			'12H' => array(
				'hh' => 'h',
				'h' => 'G'
			)
		);

		private static $BaseDate = false;
		private static $DecimalSeparator = '.';
		private static $ThousandSeparator = '';
		private static $CurrencyCode = '';

		/**
		 * @var array Cache for already processed format strings
		 */
		private $ParsedFormatCache = array();

		/**
		 * @param string Path to file
		 * @param array Options:
		 *	TempDir => string Temporary directory path
		 *	ReturnDateTimeObjects => bool True => dates and times will be returned as PHP DateTime objects, false => as strings
		 */
		public function __construct($Filepath, array $Options = null)
		{
			if (!is_readable($Filepath))
			{
				throw new Exception('SpreadsheetReader_XLSX: File not readable ('.$Filepath.')');
			}

			$this -> TempDir = isset($Options['TempDir']) && is_writable($Options['TempDir']) ?
				$Options['TempDir'] :
				sys_get_temp_dir();

			$this -> TempDir = rtrim($this -> TempDir, DIRECTORY_SEPARATOR);
			$this -> TempDir = $this -> TempDir.DIRECTORY_SEPARATOR.uniqid().DIRECTORY_SEPARATOR;

			$Zip = new ZipArchive;
			$Status = $Zip -> open($Filepath);

			if ($Status !== true)
			{
				throw new Exception('SpreadsheetReader_XLSX: File not readable ('.$Filepath.') (Error '.$Status.')');
			}

			// Extracting the XMLs from the XLSX zip file
			if ($Zip -> locateName('xl/sharedStrings.xml') !== false)
			{
				$this -> SharedStringsPath = $this -> TempDir.'xl/sharedStrings.xml';
			}

			// 10 tries to check for worksheets should be enough
			$WorksheetIndex = 0;
			for ($i = 0; $i < 10; $i++)
			{
				if ($Zip -> locateName('xl/worksheets/sheet'.$i.'.xml') !== false)
				{
					$WorksheetIndex = $i;
					$this -> WorksheetPath = $this -> TempDir.'xl/worksheets/sheet'.$WorksheetIndex.'.xml';
					break;
				}
			}

			if ($this -> WorksheetPath)
			{
				$Zip -> extractTo($this -> TempDir, 'xl/worksheets/sheet'.$WorksheetIndex.'.xml');
				if ($this -> SharedStringsPath)
				{
					$Zip -> extractTo($this -> TempDir, 'xl/sharedStrings.xml');
				}

				if ($Zip -> locateName('xl/styles.xml') !== false)
				{
					$this -> StylesXML = new SimpleXMLElement($Zip -> getFromName('xl/styles.xml'));
				}
			}

			$Zip -> close();

			if ($this -> WorksheetPath && is_readable($this -> WorksheetPath))
			{
				$this -> Worksheet = new XMLReader;
				$this -> Worksheet -> open($this -> WorksheetPath);
				$this -> Valid = true;
			}
			if ($this -> SharedStringsPath && is_readable($this -> SharedStringsPath))
			{
				$this -> SharedStrings = new XMLReader;
				$this -> SharedStrings -> open($this -> SharedStringsPath);
			}

			// If worksheet is present and is OK, parse the styles already
			if ($this -> Worksheet && $this -> StylesXML)
			{
				if ($this -> StylesXML -> cellXfs && $this -> StylesXML -> cellXfs -> xf)
				{
					foreach ($this -> StylesXML -> cellXfs -> xf as $Index => $XF)
					{
						if ($XF -> attributes() -> applyNumberFormat)
						{
							$FormatId = (int)$XF -> attributes() -> numFmtId;
							// If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
							$this -> Styles[] = $FormatId;
						}
						else
						{
							$this -> Styles[] = false;
						}
					}
				}
				
				if ($this -> StylesXML -> numFmts && $this -> StylesXML -> numFmts -> numFmt)
				{
					foreach ($this -> StylesXML -> numFmts -> numFmt as $Index => $NumFmt)
					{
						$this -> Formats[(int)$NumFmt -> attributes() -> numFmtId] = (string)$NumFmt -> attributes() -> formatCode;
					}
				}

				unset($this -> StylesXML);
			}

			// Setting base date
			if (!self::$BaseDate)
			{
				self::$BaseDate = new DateTime;
				self::$BaseDate -> setTimezone(new DateTimeZone('UTC'));
				self::$BaseDate -> setDate(1900, 1, 0);
				self::$BaseDate -> setTime(0, 0, 0);
			}

			// Decimal and thousand separators
			if (!self::$DecimalSeparator && !self::$ThousandSeparator && !self::$CurrencyCode)
			{
				$Locale = localeconv();
				self::$DecimalSeparator = $Locale['decimal_point'];
				self::$ThousandSeparator = $Locale['thousands_sep'];
				self::$CurrencyCode = $Locale['int_curr_symbol'];
			}

			if (function_exists('gmp_gcd'))
			{
				self::$RuntimeInfo['GMPSupported'] = true;
			}
		}

		/**
		 * Destructor, destroys all that remains (closes and deletes temp files)
		 */
		public function __destruct()
		{
			if ($this -> Worksheet && $this -> Worksheet instanceof XMLReader)
			{
				$this -> Worksheet -> close();
				unset($this -> Worksheet);
			}
			if (file_exists($this -> WorksheetPath))
			{
				@unlink($this -> WorksheetPath);
				unset($this -> WorksheetPath);
			}

			if ($this -> SharedStrings && $this -> SharedStrings instanceof XMLReader)
			{
				$this -> SharedStrings -> close();
				unset($this -> SharedStrings);
			}
			if (file_exists($this -> SharedStringsPath))
			{
				@unlink($this -> SharedStringsPath);
				unset($this -> SharedStringsPath);
			}
		}

		/**
		 * Retrieves a shared string value by its index
		 *
		 * @param int Shared string index
		 *
		 * @return string Value
		 */
		private function GetSharedString($Index)
		{
			// If the desired index is before the current, rewind the XML
			if ($this -> SharedStringIndex > $Index)
			{
				$this -> SSOpen = false;
				$this -> SharedStrings -> close();
				$this -> SharedStrings -> open($this -> SharedStringsPath);
				$this -> SharedStringIndex = 0;
				$this -> LastSharedStringValue = null;
				$this -> SSForwarded = false;
			}

			// Finding the unique string count (if not already read)
			if ($this -> SharedStringIndex == 0 && !$this -> SharedStringCount)
			{
				while ($this -> SharedStrings -> read())
				{
					if ($this -> SharedStrings -> name == 'sst')
					{
						$this -> SharedStringCount = $this -> SharedStrings -> getAttribute('uniqueCount');
						break;
					}
				}
			}

			// If index of the desired string is larger than possible, don't even bother.
			if ($this -> SharedStringCount && ($Index >= $this -> SharedStringCount))
			{
				return '';
			}

			// If an index with the same value as the last already fetched is requested
			// (any further traversing the tree would get us further away from the node)
			if (($Index == $this -> SharedStringIndex) && ($this -> LastSharedStringValue !== null))
			{
				return $this -> LastSharedStringValue;
			}

			// Find the correct <si> node with the desired index
			while ($this -> SharedStringIndex <= $Index)
			{
				// SSForwarded is set further to avoid double reading in case nodes are skipped.
				if ($this -> SSForwarded)
				{
					$this -> SSForwarded = false;
				}
				else
				{
					$ReadStatus = $this -> SharedStrings -> read();
					if (!$ReadStatus)
					{
						break;
					}
				}

				if ($this -> SharedStrings -> name == 'si')
				{
					if ($this -> SharedStrings -> nodeType == XMLReader::END_ELEMENT)
					{
						$this -> SSOpen = false;
						$this -> SharedStringIndex++;
					}
					else
					{
						$this -> SSOpen = true;
	
						if ($this -> SharedStringIndex < $Index)
						{
							$this -> SSOpen = false;
							$this -> SharedStrings -> next('si');
							$this -> SSForwarded = true;
							$this -> SharedStringIndex++;
							continue;
						}
						else
						{
							break;
						}
					}
				}
			}

			$Value = '';

			// Extract the value from the shared string
			if ($this -> SSOpen && ($this -> SharedStringIndex == $Index))
			{
				while ($this -> SharedStrings -> read())
				{
					switch ($this -> SharedStrings -> name)
					{
						case 't':
							if ($this -> SharedStrings -> nodeType == XMLReader::END_ELEMENT)
							{
								continue;
							}
							$Value .= $this -> SharedStrings -> readString();
							break;
						case 'si':
							if ($this -> SharedStrings -> nodeType == XMLReader::END_ELEMENT)
							{
								$this -> SSOpen = false;
								$this -> SSForwarded = true;
								break 2;
							}
							break;
					}
				}
			}

			if ($Value)
			{
				$this -> LastSharedStringValue = $Value;
			}
			return $Value;
		}

		/**
		 * Formats the value according to the index
		 *
		 * @param string Cell value
		 * @param int Format index
		 *
		 * @return string Formatted cell value
		 */
		private function FormatValue($Value, $Index)
		{
			if (!is_numeric($Value))
			{
				return $Value;
			}

			if (!empty($this -> Styles[$Index]))
			{
				$Index = $this -> Styles[$Index];
			}
			else
			{
				return $Value;
			}

			$Format = array();

			if (isset($this -> ParsedFormatCache[$Index]))
			{
				$Format = $this -> ParsedFormatCache[$Index];
			}

			if (!$Format)
			{
				$Format = array(
					'Code' => false,
					'Type' => false,
					'Scale' => 1,
					'Thousands' => false,
					'Currency' => false
				);

				if (isset(self::$BuiltinFormats[$Index]))
				{
					$Format['Code'] = self::$BuiltinFormats[$Index];
				}
				elseif (isset($this -> Formats[$Index]))
				{
					$Format['Code'] = $this -> Formats[$Index];
				}

				// Format code found, now parsing the format
				if ($Format['Code'])
				{
					$Sections = explode(';', $Format['Code']);
					$Format['Code'] = $Sections[0];
	
					switch (count($Sections))
					{
						case 2:
							if ($Value < 0)
							{
								$Format['Code'] = $Sections[1];
							}
							break;
						case 3:
						case 4:
							if ($Value < 0)
							{
								$Format['Code'] = $Sections[1];
							}
							elseif ($Value == 0)
							{
								$Format['Code'] = $Sections[2];
							}
							break;
					}
				}

				// Stripping colors
				$Format['Code'] = trim(preg_replace('{^\[[[:alpha:]]+\]}i', '', $Format['Code']));

				// Percentages
				if (substr($Format['Code'], -1) == '%')
				{
					$Format['Type'] = 'Percentage';
				}
				elseif (preg_match('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]}i', $Format['Code']))
				{
					$Format['Type'] = 'DateTime';

					$Format['Code'] = trim(preg_replace('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])}i', '', $Format['Code']));
					$Format['Code'] = strtolower($Format['Code']);

					$Format['Code'] = strtr($Format['Code'], self::$DateReplacements['All']);
					if (strpos($Format['Code'], 'A') === false)
					{
						$Format['Code'] = strtr($Format['Code'], self::$DateReplacements['24H']);
					}
					else
					{
						$Format['Code'] = strtr($Format['Code'], self::$DateReplacements['12H']);
					}
				}
				elseif ($Format['Code'] == '[$EUR ]#,##0.00_-')
				{
					$Format['Type'] = 'Euro';
				}
				else
				{
					// Removing skipped characters
					$Format['Code'] = preg_replace('{_.}', '', $Format['Code']);
					// Removing unnecessary escaping
					$Format['Code'] = preg_replace("{\\\\}", '', $Format['Code']);
					// Removing string quotes
					$Format['Code'] = str_replace(array('"', '*'), '', $Format['Code']);
					// Removing thousands separator
					if (strpos($Format['Code'], '0,0') !== false || strpos($Format['Code'], '#,#') !== false)
					{
						$Format['Thousands'] = true;
					}
					$Format['Code'] = str_replace(array('0,0', '#,#'), array('00', '##'), $Format['Code']);

					// Scaling (Commas indicate the power)
					$Scale = 1;
					$Matches = array();
					if (preg_match('{(0|#)(,+)}', $Format['Code'], $Matches))
					{
						$Scale = pow(1000, strlen($Matches[2]));
						// Removing the commas
						$Format['Code'] = preg_replace(array('{0,+}', '{#,+}'), array('0', '#'), $Format['Code']);
					}

					$Format['Scale'] = $Scale;

					if (preg_match('{#?.*\?\/\?}', $Format['Code']))
					{
						$Format['Type'] = 'Fraction';
					}
					else
					{
						$Format['Code'] = str_replace('#', '', $Format['Code']);

						$Matches = array();
						if (preg_match('{(0+)(\.?)(0*)}', preg_replace('{\[[^\]]+\]}', '', $Format['Code']), $Matches))
						{
							$Integer = $Matches[1];
							$DecimalPoint = $Matches[2];
							$Decimals = $Matches[3];

							$Format['MinWidth'] = strlen($Integer) + strlen($DecimalPoint) + strlen($Decimals);
							$Format['Decimals'] = $Decimals;
							$Format['Precision'] = strlen($Format['Decimals']);
							$Format['Pattern'] = '%0'.$Format['MinWidth'].'.'.$Format['Precision'].'f';
						}
					}

					$Matches = array();
					if (preg_match('{\[\$(.*)\]}u', $Format['Code'], $Matches))
					{
						$CurrFormat = $Matches[0];
						$CurrCode = $Matches[1];
						$CurrCode = explode('-', $CurrCode);
						if ($CurrCode)
						{
							$CurrCode = $CurrCode[0];
						}

						if (!$CurrCode)
						{
							$CurrCode = self::$CurrencyCode;
						}

						$Format['Currency'] = $CurrCode;
					}
					$Format['Code'] = trim($Format['Code']);
				}

				$this -> ParsedFormatCache[$Index] = $Format;
			}

			// Applying format to value
			if ($Format)
			{
				// Percentages
				if ($Format['Type'] == 'Percentage')
				{
					if ($Format['Code'] === '0%')
					{
						$Value = round(100 * $Value, 0).'%';
					}
					else
					{
						$Value = sprintf('%.2f%%', round(100 * $Value, 2));
					}
				}
				// Dates and times
				elseif ($Format['Type'] == 'DateTime')
				{
					$Days = (int)$Value;
					// Correcting for Feb 29, 1900
					if ($Days > 60)
					{
						$Days--;
					}

					// At this point time is a fraction of a day
					$Time = ($Value - (int)$Value);
					$Seconds = 0;
					if ($Time)
					{
						// Here time is converted to seconds
						// Some loss of precision will occur
						$Seconds = (int)($Time * 86400);
					}

					$Value = clone self::$BaseDate;
					$Value -> add(new DateInterval('P'.$Days.'D'.($Seconds ? 'T'.$Seconds.'S' : '')));

					if (!$this -> Options['ReturnDateTimeObjects'])
					{
						$Value = $Value -> format($Format['Code']);
					}
					else
					{
						// A DateTime object is returned
					}
				}
				elseif ($Format['Type'] == 'Euro')
				{
					$Value = 'EUR '.sprintf('%1.2f', $Value);
				}
				else
				{
					// Fractional numbers
					if ($Format['Type'] == 'Fraction' && ($Value != (int)$Value))
					{
						$Integer = floor(abs($Value));
						$Decimal = fmod(abs($Value), 1);
						// Removing the integer part and decimal point
						$Decimal *= pow(10, strlen($Decimal) - 2);
						$DecimalDivisor = pow(10, strlen($Decimal));

						if (self::$RuntimeInfo['GMPSupported'])
						{
							$GCD = gmp_strval(gmp_gcd($Decimal, $DecimalDivisor));
						}
						else
						{
							$GCD = self::GCD($Decimal, $DecimalDivisor);
						}

						$AdjDecimal = $DecimalPart/$GCD;
						$AdjDecimalDivisor = $DecimalDivisor/$GCD;

						if (
							strpos($Format['Code'], '0') !== false || 
							strpos($Format['Code'], '#') !== false ||
							substr($Format['Code'], 0, 3) == '? ?'
						)
						{
							// The integer part is shown separately apart from the fraction
							$Value = ($Value < 0 ? '-' : '').
								$Integer ? $Integer.' ' : ''.
								$AdjDecimal.'/'.
								$AdjDecimalDivisor;
						}
						else
						{
							// The fraction includes the integer part
							$AdjDecimal += $Integer * $AdjDecimalDivisor;
							$Value = ($Value < 0 ? '-' : '').
								$AdjDecimal.'/'.
								$AdjDecimalDivisor;
						}
					}
					else
					{
						// Scaling
						$Value = $Value / $Format['Scale'];

						if ($Format['MinWidth'] && $Format['Decimals'])
						{
							if ($Format['Thousands'])
							{
								$Value = number_format($Value, $Format['Precision'],
									self::$DecimalSeparator, self::$ThousandSeparator);
							}
							else
							{
								$Value = sprintf($Format['Pattern'], $Value);
							}

							$Value = preg_replace('{(0+)(\.?)(0*)}', $Value, $Format['Code']);
						}
					}

					// Currency/Accounting
					if ($Format['Currency'])
					{
						$Value = preg_replace('', $Format['Currency'], $Value);
					}
				}
				
			}

			return $Value;
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
				$this -> Worksheet -> close();
				$this -> Worksheet -> open($this -> WorksheetPath);
				$this -> Valid = true;

				$this -> RowOpen = false;
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

			if (!$this -> RowOpen)
			{
				while ($this -> Valid = $this -> Worksheet -> read())
				{
					if ($this -> Worksheet -> name == 'row')
					{
						// Getting the row spanning area (stored as e.g., 1:12)
						// so that the last cells will be present, even if empty
						$RowSpans = $this -> Worksheet -> getAttribute('spans');
						if ($RowSpans)
						{
							$RowSpans = explode(':', $RowSpans);
							$CurrentRowColumnCount = $RowSpans[1];
						}
						else
						{
							$CurrentRowColumnCount = 0;
						}

						$this -> RowOpen = true;
						break;
					}
				}
			}

			// Reading the necessary row, if found
			if ($this -> RowOpen)
			{
				if ($CurrentRowColumnCount > 0)
				{
					$this -> CurrentRow = array_fill(0, $CurrentRowColumnCount, '');
				}
				else
				{
					$this -> CurrentRow = array();
				}

				// These two are needed to control for empty cells
				$MaxIndex = 0;
				$CellCount = 0;

				while ($this -> Valid = $this -> Worksheet -> read())
				{
					switch ($this -> Worksheet -> name)
					{
						// End of row
						case 'row':
							if ($this -> Worksheet -> nodeType == XMLReader::END_ELEMENT)
							{
								$this -> RowOpen = false;
								break 2;
							}
						// Cell
						case 'c':
							// If it is a closing tag, skip it
							if ($this -> Worksheet -> nodeType == XMLReader::END_ELEMENT)
							{
								continue;
							}

							$this -> CellOpen = !$this -> CellOpen;

							// Determine cell type and get value
							if ($this -> Worksheet -> getAttribute('t') == self::CELL_TYPE_SHARED_STR)
							{
								$SharedStringIndex = $this -> Worksheet -> readString();
								$Value = $this -> GetSharedString($SharedStringIndex);
							}
							else
							{
								$Value = $this -> Worksheet -> readString();
							}

							// Format value if necessary
							if ($Value !== '')
							{
								$StyleId = (int)$this -> Worksheet -> getAttribute('s');
								if ($StyleId && isset($this -> Styles[$StyleId]))
								{
									$Value = $this -> FormatValue($Value, $StyleId);
								}
							}

							// Get the index of the cell
							$Index = $this -> Worksheet -> getAttribute('r');
							$Letter = preg_replace('{[^[:alpha:]]}S', '', $Index);
							$Index = self::IndexFromColumnLetter($Letter);

							$this -> CurrentRow[$Index] = $Value;

							$CellCount++;
							if ($Index > $MaxIndex)
							{
								$MaxIndex = $Index;
							}

							break;
					}
				}

				// Adding empty cells, if necessary
				// Only empty cells inbetween and on the left side are added
				if ($MaxIndex + 1 > $CellCount)
				{
					$this -> CurrentRow = $this -> CurrentRow + array_fill(0, $MaxIndex + 1, '');
					ksort($this -> CurrentRow);
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

		/**
		 * Takes the column letter and converts it to a numerical index (0-based)
		 *
		 * @param string Letter(s) to convert
		 *
		 * @return mixed Numeric index (0-based) or boolean false if it cannot be calculated
		 */
		public static function IndexFromColumnLetter($Letter)
		{
			$Powers = array();

			$Letter = strtoupper($Letter);

			$Result = 0;
			for ($i = strlen($Letter) - 1, $j = 0; $i >= 0; $i--, $j++)
			{
				$Ord = ord($Letter[$i]) - 64;
				if ($Ord > 26)
				{
					// Something is very, very wrong
					return false;
				}
				$Result += $Ord * pow(26, $j);
			}
			return $Result - 1;
		}

		/**
		 * Helper function for greatest common divisor calculation in case GMP extension is
		 *	not enabled
		 *
		 * @param int Number #1
		 * @param int Number #2
		 *
		 * @param int Greatest common divisor
		 */
		public static function GCD($A, $B)
		{
			$A = abs($A);
			$B = abs($B);
			if ($A + $B == 0)
			{
				return 0;
			}
			else
			{
				$C = 1;

				while ($A > 0)
				{
					$C = $A;
					$A = $B % $A;
					$B = $C;
				}

				return $C;
			}
		}
	}
?>