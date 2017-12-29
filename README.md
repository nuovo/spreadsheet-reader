# Установка - Install Composer

```

composer require jackmartin\readerexcel@dev-v0.5.11-fix

```

# Описание - Description

**spreadsheet-reader** is a PHP spreadsheet reader that differs from others in that the main goal for it was efficient
data extraction that could handle large (as in really large) files. So far it may not definitely be CPU, time
or I/O-efficient but at least it won't run out of memory (except maybe for XLS files).

# Требуется - Requirements:
*  PHP 5.3.0 or newer
*  PHP must have Zip file support (see http://php.net/manual/en/zip.installation.php)
*  Composer

# Использовать - Usage:


However, now also multiple sheet reading is supported for file formats where it is possible. (In case of CSV, it is handled as if
it only has one sheet.)

You can retrieve information about sheets contained in the file by calling the `Sheets()` method which returns an array with
sheet indexes as keys and sheet names as values. Then you can change the sheet that's currently being read by passing that index
to the `ChangeSheet($Index)` method.

### Пример - Example: ###

```
include_once __DIR__ . '/vendor/autoload.php';

use ReaderExcel\SpreadsheetReader;

$file = __DIR__ . '/test.xlsx';

$reader = new SpreadsheetReader($file);

$Sheets = $reader->Sheets();

foreach ($Sheets as $Index => $Name) {

    $reader->ChangeSheet($Index);

    foreach ($reader as $Key => $Row) {

        print_r($Row);

    }

}

```

If a sheet is changed to the same that is currently open, the position in the file still reverts to the beginning, so as to conform
to the same behavior as when changed to a different sheet.

# Notes about library performance
*  CSV and text files are read strictly sequentially so performance should be O(n);
*  When parsing XLS files, all of the file content is read into memory so large XLS files can lead to "out of memory" errors;
*  XLSX files use so called "shared strings" internally to optimize for cases where the same string is repeated multiple times.
	Internally XLSX is an XML text that is parsed sequentially to extract data from it, however, in some cases these shared strings are a problem -
	sometimes Excel may put all, or nearly all of the strings from the spreadsheet in the shared string file (which is a separate XML text), and not necessarily in the same
	order. Worst case scenario is when it is in reverse order - for each string we need to parse the shared string XML from the beginning, if we want to avoid keeping the data in memory.
	To that end, the XLSX parser has a cache for shared strings that is used if the total shared string count is not too high. In case you get out of memory errors, you can
	try adjusting the *SHARED_STRING_CACHE_LIMIT* constant in SpreadsheetReader_XLSX to a lower one.

# TODOs:
*  ODS date formats;

# Licensing
All of the code in this library is licensed under the MIT license as included in the LICENSE file, however, for now the library
relies on php-excel-reader library for XLS file parsing which is licensed under the PHP license.