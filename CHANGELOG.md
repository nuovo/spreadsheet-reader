### v.0.5.11  2015-04-30

- Added a special case for cells formatted as text in XLSX. Previously leading zeros would get truncated if a text cell contained only numbers.

### v.0.5.10  2015-04-18

- Implemented SeekableIterator. Thanks to [paales](https://github.com/paales) for suggestion ([Issue #54](https://github.com/nuovo/spreadsheet-reader/issues/54) and [Pull request #55](https://github.com/nuovo/spreadsheet-reader/pull/55)).
- Fixed a bug in CSV and ODS reading where reading position 0 multiple times in a row would result in internal pointer being advanced and reading the next line. (E.g. reading row #0 three times would result in rows #0, #1, and #2.). This could have happened on multiple calls to `current()` while in #0 position, or calls to `seek(0)` and `current()`.

### v.0.5.9  2015-04-18

- [Pull request #85](https://github.com/nuovo/spreadsheet-reader/pull/85): Fixed an index check. (Thanks to [pa-m](https://github.com/pa-m)).

### v.0.5.8  2015-01-31

- [Issue #50](https://github.com/nuovo/spreadsheet-reader/issues/50): Fixed an XLSX rewind issue. (Thanks to [osuwariboy](https://github.com/osuwariboy))
- [Issue #52](https://github.com/nuovo/spreadsheet-reader/issues/52), [#53](https://github.com/nuovo/spreadsheet-reader/issues/53): Apache POI compatibility for XLSX. (Thanks to [dimapashkov](https://github.com/dimapashkov))
- [Issue #61](https://github.com/nuovo/spreadsheet-reader/issues/61): Autoload fix in the main class. (Thanks to [i-bash](https://github.com/i-bash))
- [Issue #60](https://github.com/nuovo/spreadsheet-reader/issues/60), [#69](https://github.com/nuovo/spreadsheet-reader/issues/69), [#72](https://github.com/nuovo/spreadsheet-reader/issues/72): Fixed an issue where XLSX ChangeSheet may not work. (Thanks to [jtresponse](https://github.com/jtresponse), [osuwariboy](https://github.com/osuwariboy))
- [Issue #70](https://github.com/nuovo/spreadsheet-reader/issues/70): Added a check for constructor parameter correctness.


### v.0.5.7  2013-10-29

- Attempt to replicate Excel's "General" format in XLSX files that is applied to otherwise unformatted cells.
Currently only decimal number values are converted to PHP's floats.

### v.0.5.6  2013-09-04

- Fix for formulas being returned along with values in XLSX files. (Thanks to [marktag](https://github.com/marktag))

### v.0.5.5  2013-08-23

- Fix for macro sheets appearing when parsing XLS files. (Thanks to [osuwariboy](https://github.com/osuwariboy))

### v.0.5.4  2013-08-22

- Fix for a PHP warning that occurs with completely empty sheets in XLS files.
- XLSM (macro-enabled XLSX) files are recognized and read, too.
- composer.json file is added to the repository (thanks to [matej116](https://github.com/matej116))

### v.0.5.3  2013-08-12

- Fix for repeated columns in ODS files not reading correctly (thanks to [etfb](https://github.com/etfb))
- Fix for filename extension reading (Thanks to [osuwariboy](https://github.com/osuwariboy))

### v.0.5.2  2013-06-28

- A fix for the case when row count wasn't read correctly from the sheet in a XLS file.

### v.0.5.1  2013-06-27

- Fixed file type choice when using mime-types (previously there were problems with  
XLSX and ODS mime-types) (Thanks to [incratec](https://github.com/incratec))

- Fixed an error in XLSX iterator where `current()` would advance the iterator forward  
with each call. (Thanks to [osuwariboy](https://github.com/osuwariboy))

### v.0.5.0  2013-06-17

- Multiple sheet reading is now supported:
	- The `Sheets()` method lets you retrieve a list of all sheets present in the file.
	- `ChangeSheet($Index)` method changes the sheet in the reader to the one specified.

- Previously temporary files that were extracted, were deleted after the SpreadsheetReader  
was destroyed but the empty directories remained. Now those are cleaned up as well.  

### v.0.4.3  2013-06-14

- Bugfix for shared string caching in XLSX files. When the shared string count was larger  
than the caching limit, instead of them being read from file, empty strings were returned.  

### v.0.4.2  2013-06-02

- XLS file reading relies on the external Spreadsheet_Excel_Reader class which, by default,  
reads additional information about cells like fonts, styles, etc. Now that is disabled  
to save some memory since the style data is unnecessary anyway.  
(Thanks to [ChALkeR](https://github.com/ChALkeR) for the tip.)

Martins Pilsetnieks  <pilsetnieks@gmail.com>