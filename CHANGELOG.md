### v.0.5.4  2013-08-22

- Fix for a PHP warning that occurs with completely empty sheets in XLS files.
- XLSM (macro-enabled XLSX) files are recognized and read, too.

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