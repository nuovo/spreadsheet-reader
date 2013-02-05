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

### Notes about library performance
*  CSV and text files are read strictly sequentially so performance should be O(n);
*  When parsing XLS files, all of the file content is read into memory so large XLS files can lead to "out of memory" errors;
*  XLSX files use so called "shared strings" internally to optimize for cases where the same string is repeated multiple times.
	Internally XLSX is an XML text that is parsed sequentially to extract data from it, however, in some cases these shared strings are a problem -
	sometimes Excel may put all, or nearly all of the strings from the spreadsheet in the shared string file (which is a separate XML text), and not necessarily in the same
	order. Worst case scenario is when it is in reverse order - for each string we need to parse the shared string XML from the beginning, if we want to avoid keeping the data in memory.
	To that end, the XLSX parser has a cache for shared strings that is used if the total shared string count is not too high. In case you get out of memory errors, you can
	try adjusting the *SHARED_STRING_CACHE_LIMIT* constant in SpreadsheetReader_XLSX to a lower one.

### TODOs:
*  ODS date formats;

### Licensing
All of the code in this library is licensed under the MIT license as included in the LICENSE file, however, for now the library relies on php-excel-reader library for XLS file parsing which is licensed under the PHP license.

http://www.nuovo.lv