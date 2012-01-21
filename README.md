Nuovo/Nouveau spreadsheet-reader is a PHP spreadsheet reader with the difference that my only goal for it was efficient data extraction that could handle large (as in really large) files. So far I cannot definitely say that it is CPU, time or IO-efficient but at least it won't run out of memory (except for XLS files).

So far XLSX, ODS and text/CSV file parsing should be memory-efficient. XLS file parsing is done with php-excel-reader from http://code.google.com/p/php-excel-reader/ which, sadly, has memory issues with bigger spreadsheets.

### Requirements:
*  PHP 5.3.0 or newer
*  PHP must have Zip file support (see http://php.net/manual/en/zip.installation.php)

### Usage:

Very simple:

	<?php
		// If you need to parse XLS files, include php-excel-reader
		require('php-excel-reader/excel_reader2.php');
	
		require('SpreadsheetReader.php');
	
		$Reader = new SpreadsheetReader('example.xlsx');
		foreach ($Reader as $Row)
		{
			print_r($Row);
		}
	?>

### Testing

From the command line:

	php test.php path-to-spreadsheet.xls

### TODOs:
*  ODS date formats;
*  XLSX XML parsing suffers from an occasional Shliemel the painter moment (sharedStrings.xml)

http://www.nuovo.lv