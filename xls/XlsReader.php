<?php

namespace SpreadsheetReader\xls;

include_once 'OLEReader.php';

define('SPREADSHEET_EXCEL_READER_BIFF8',			 0x600);
define('SPREADSHEET_EXCEL_READER_BIFF7',			 0x500);
define('SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS',   0x5);
define('SPREADSHEET_EXCEL_READER_WORKSHEET',		 0x10);
define('SPREADSHEET_EXCEL_READER_TYPE_BOF',		  0x809);
define('SPREADSHEET_EXCEL_READER_TYPE_EOF',		  0x0a);
define('SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET',   0x85);
define('SPREADSHEET_EXCEL_READER_TYPE_DIMENSION',	0x200);
define('SPREADSHEET_EXCEL_READER_TYPE_ROW',		  0x208);
define('SPREADSHEET_EXCEL_READER_TYPE_DBCELL',	   0xd7);
define('SPREADSHEET_EXCEL_READER_TYPE_FILEPASS',	 0x2f);
define('SPREADSHEET_EXCEL_READER_TYPE_NOTE',		 0x1c);
define('SPREADSHEET_EXCEL_READER_TYPE_TXO',		  0x1b6);
define('SPREADSHEET_EXCEL_READER_TYPE_RK',		   0x7e);
define('SPREADSHEET_EXCEL_READER_TYPE_RK2',		  0x27e);
define('SPREADSHEET_EXCEL_READER_TYPE_MULRK',		0xbd);
define('SPREADSHEET_EXCEL_READER_TYPE_MULBLANK',	 0xbe);
define('SPREADSHEET_EXCEL_READER_TYPE_INDEX',		0x20b);
define('SPREADSHEET_EXCEL_READER_TYPE_SST',		  0xfc);
define('SPREADSHEET_EXCEL_READER_TYPE_EXTSST',	   0xff);
define('SPREADSHEET_EXCEL_READER_TYPE_CONTINUE',	 0x3c);
define('SPREADSHEET_EXCEL_READER_TYPE_LABEL',		0x204);
define('SPREADSHEET_EXCEL_READER_TYPE_LABELSST',	 0xfd);
define('SPREADSHEET_EXCEL_READER_TYPE_NUMBER',	   0x203);
define('SPREADSHEET_EXCEL_READER_TYPE_NAME',		 0x18);
define('SPREADSHEET_EXCEL_READER_TYPE_ARRAY',		0x221);
define('SPREADSHEET_EXCEL_READER_TYPE_STRING',	   0x207);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMULA',	  0x406);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMULA2',	 0x6);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMAT',	   0x41e);
define('SPREADSHEET_EXCEL_READER_TYPE_XF',		   0xe0);
define('SPREADSHEET_EXCEL_READER_TYPE_BOOLERR',	  0x205);
define('SPREADSHEET_EXCEL_READER_TYPE_FONT',	  0x0031);
define('SPREADSHEET_EXCEL_READER_TYPE_PALETTE',	  0x0092);
define('SPREADSHEET_EXCEL_READER_TYPE_UNKNOWN',	  0xffff);
define('SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR', 0x22);
define('SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS',  0xE5);
define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS' ,	25569);
define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904', 24107);
define('SPREADSHEET_EXCEL_READER_MSINADAY',		  86400);
define('SPREADSHEET_EXCEL_READER_TYPE_HYPER',	     0x01b8);
define('SPREADSHEET_EXCEL_READER_TYPE_COLINFO',	     0x7d);
define('SPREADSHEET_EXCEL_READER_TYPE_DEFCOLWIDTH',  0x55);
define('SPREADSHEET_EXCEL_READER_TYPE_STANDARDWIDTH', 0x99);
define('SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT',	"%s");

class XlsReader {

    /**
     * Default number formats used by Excel
     */
    protected $numberFormats = array(
        0x1 => "0",
        0x2 => "0.00",
        0x3 => "#,##0",
        0x4 => "#,##0.00",
        0x5 => "\$#,##0;(\$#,##0)",
        0x6 => "\$#,##0;[Red](\$#,##0)",
        0x7 => "\$#,##0.00;(\$#,##0.00)",
        0x8 => "\$#,##0.00;[Red](\$#,##0.00)",
        0x9 => "0%",
        0xa => "0.00%",
        0xb => "0.00E+00",
        0x25 => "#,##0;(#,##0)",
        0x26 => "#,##0;[Red](#,##0)",
        0x27 => "#,##0.00;(#,##0.00)",
        0x28 => "#,##0.00;[Red](#,##0.00)",
        0x29 => "#,##0;(#,##0)",  // Not exactly
        0x2a => "\$#,##0;(\$#,##0)",  // Not exactly
        0x2b => "#,##0.00;(#,##0.00)",  // Not exactly
        0x2c => "\$#,##0.00;(\$#,##0.00)",  // Not exactly
        0x30 => "##0.0E+0"
    );

    // END PUBLIC API
    protected $boundsheets = array();
    protected $formatRecords = array();
    protected $fontRecords = array();
    protected $xfRecords = array();
    protected $colInfo = array();

    protected $rowInfo = array();
    protected $sst = array();

    protected $sheets = array();
    protected $data;
    protected $store_extended_info;

    /**
     * @var OLEReader
     */
    protected $ole;

    protected $_defaultEncoding = "UTF-8";
    protected $_defaultFormat = SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT;
    protected $_columnsFormat = array();
    protected $_rowoffset = 1;
    protected $_coloffset = 1;

    // MK: Added to make data retrieval easier
    protected $colnames = array();
    protected $colindexes = array();
    protected $standardColWidth = 0;
    protected $defaultColWidth = 0;
    protected $startPosition = null;

    /**
     * List of default date formats used by Excel
     */
    protected $dateFormats = array (
        0xe => "m/d/Y",
        0xf => "M-d-Y",
        0x10 => "d-M",
        0x11 => "M-Y",
        0x12 => "h:i a",
        0x13 => "h:i:s a",
        0x14 => "H:i",
        0x15 => "H:i:s",
        0x16 => "d/m/Y H:i",
        0x2d => "i:s",
        0x2e => "H:i:s",
        0x2f => "i:s.S"
    );

    function __construct($fileName, $storeExtendedInfo = true, $outputEncoding = '') {


        $this->setUTFEncoder('iconv');

        if ($outputEncoding != '') {
            $this->setOutputEncoding($outputEncoding);
        }

        for ($i=1; $i<245; $i++) {
            $name = strtolower(( (($i-1)/26>=1)?chr(($i-1)/26+64):'') . chr(($i-1)%26+65));
            $this->colnames[$name] = $i;
            $this->colindexes[$i] = $name;
        }

        $this->storeExtendedInfo = $storeExtendedInfo;

        if ($fileName) {
            $this->ole = new OLEReader($fileName);
            $this->read();
        } else {
            throw new Exception('Source file missing');
        }
    }

    function myHex($d) {
        if ($d < 16) return "0" . dechex($d);
        return dechex($d);
    }

    function dumpHexData($data, $pos, $length) {
        $info = "";
        for ($i = 0; $i <= $length; $i++) {
            $info .= ($i==0?"":" ") . $this->myHex(ord($data[$pos + $i])) . (ord($data[$pos + $i])>31? "[" . $data[$pos + $i] . "]":'');
        }
        return $info;
    }

    function getCol($col) {
        if (is_string($col)) {
            $col = strtolower($col);
            if (array_key_exists($col,$this->colnames)) {
                $col = $this->colnames[$col];
            }
        }
        return $col;
    }

    // PUBLIC API FUNCTIONS
    // --------------------

    function val($row,$col,$sheet=0) {
        $col = $this->getCol($col);
        if (array_key_exists($row,$this->sheets[$sheet]['cells']) && array_key_exists($col,$this->sheets[$sheet]['cells'][$row])) {
            return $this->sheets[$sheet]['cells'][$row][$col];
        }
        return "";
    }

    function value($row,$col,$sheet=0) {
        return $this->val($row,$col,$sheet);
    }

    function info($row,$col,$type='',$sheet=0) {
        $col = $this->getCol($col);
        if (array_key_exists('cellsInfo',$this->sheets[$sheet])
            && array_key_exists($row,$this->sheets[$sheet]['cellsInfo'])
            && array_key_exists($col,$this->sheets[$sheet]['cellsInfo'][$row])
            && array_key_exists($type,$this->sheets[$sheet]['cellsInfo'][$row][$col])) {
            return $this->sheets[$sheet]['cellsInfo'][$row][$col][$type];
        }
        return "";
    }

    function type($row,$col,$sheet=0) {
        return $this->info($row,$col,'type',$sheet);
    }

    function raw($row,$col,$sheet=0) {
        return $this->info($row,$col,'raw',$sheet);
    }

    // CELL (XF) PROPERTIES
    // ====================
    function xfRecord($row,$col,$sheet=0) {
        $xfIndex = $this->info($row,$col,'xfIndex',$sheet);
        if ($xfIndex!="") {
            return $this->xfRecords[$xfIndex];
        }
        return null;
    }
    function xfProperty($row,$col,$sheet,$prop) {
        $xfRecord = $this->xfRecord($row,$col,$sheet);
        if ($xfRecord!=null) {
            return $xfRecord[$prop];
        }
        return "";
    }

    function read16bitstring($data, $start) {
        $len = 0;
        while (ord($data[$start + $len]) + ord($data[$start + $len + 1]) > 0) $len++;
        return substr($data, $start, $len);
    }

    // ADDED by Matt Kruse for better formatting
    function _format_value($format,$num,$f) {
        // 49==TEXT format
        // http://code.google.com/p/php-excel-reader/issues/detail?id=7
        if ( (!$f && $format=="%s") || ($f==49) || ($format=="GENERAL") ) {
            return array('string'=>$num, 'formatColor'=>null);
        }

        // Custom pattern can be POSITIVE;NEGATIVE;ZERO
        // The "text" option as 4th parameter is not handled
        $parts = explode(";",$format);
        $pattern = $parts[0];
        // Negative pattern
        if (count($parts)>2 && $num==0) {
            $pattern = $parts[2];
        }
        // Zero pattern
        if (count($parts)>1 && $num<0) {
            $pattern = $parts[1];
            $num = abs($num);
        }

        $color = "";
        $matches = array();
        $color_regex = "/^\[(BLACK|BLUE|CYAN|GREEN|MAGENTA|RED|WHITE|YELLOW)\]/i";
        if (preg_match($color_regex,$pattern,$matches)) {
            $color = strtolower($matches[1]);
            $pattern = preg_replace($color_regex,"",$pattern);
        }

        // In Excel formats, "_" is used to add spacing, which we can't do in HTML
        $pattern = preg_replace("/_./","",$pattern);

        // Some non-number characters are escaped with \, which we don't need
        $pattern = preg_replace("/\\\/","",$pattern);

        // Some non-number strings are quoted, so we'll get rid of the quotes
        $pattern = preg_replace("/\"/","",$pattern);

        // TEMPORARY - Convert # to 0
        $pattern = preg_replace("/\#/","0",$pattern);

        // Find out if we need comma formatting
        $has_commas = preg_match("/,/",$pattern);
        if ($has_commas) {
            $pattern = preg_replace("/,/","",$pattern);
        }

        // Handle Percentages
        if (preg_match("/\d(\%)([^\%]|$)/",$pattern,$matches)) {
            $num = $num * 100;
            $pattern = preg_replace("/(\d)(\%)([^\%]|$)/","$1%$3",$pattern);
        }

        // Handle the number itself
        $number_regex = "/(\d+)(\.?)(\d*)/";
        if (preg_match($number_regex,$pattern,$matches)) {
            $left = $matches[1];
            $dec = $matches[2];
            $right = $matches[3];
            if ($has_commas) {
                $formatted = number_format($num,strlen($right));
            }
            else {
                $sprintf_pattern = "%1.".strlen($right)."f";
                $formatted = sprintf($sprintf_pattern, $num);
            }
            $pattern = preg_replace($number_regex, $formatted, $pattern);
        }

        return array(
            'string'=>$pattern,
            'formatColor'=>$color
        );
    }

    /**
     * Set the encoding method
     */
    function setOutputEncoding($encoding) {
        $this->_defaultEncoding = $encoding;
    }

    /**
     *  $encoder = 'iconv' or 'mb'
     *  set iconv if you would like use 'iconv' for encode UTF-16LE to your encoding
     *  set mb if you would like use 'mb_convert_encoding' for encode UTF-16LE to your encoding
     */
    function setUTFEncoder($encoder = 'iconv') {
        $this->_encoderFunction = '';
        if ($encoder == 'iconv') {
            $this->_encoderFunction = function_exists('iconv') ? 'iconv' : '';
        } elseif ($encoder == 'mb') {
            $this->_encoderFunction = function_exists('mb_convert_encoding') ? 'mb_convert_encoding' : '';
        }
    }

    function setRowColOffset($iOffset) {
        $this->_rowoffset = $iOffset;
        $this->_coloffset = $iOffset;
    }

    /**
     * Set the default number format
     */
    function setDefaultFormat($sFormat) {
        $this->_defaultFormat = $sFormat;
    }

    /**
     * Force a column to use a certain format
     */
    function setColumnFormat($column, $sFormat) {
        $this->_columnsFormat[$column] = $sFormat;
    }

    /**
     * Read the spreadsheet file using OLE, then parse
     */
    function read() {
        // check error code
        if($this->ole->getError() == 1) {
            throw new Exception('The filename ' . $this->ole->getFileName() . ' is not readable');
        }

        $this->data = $this->ole->getWorkBook();

        $this->createDocumentIndex();

        //Destroying object to save some memory
        unset($this->ole);
    }

    protected function createDocumentIndex()
    {
        $pos = 0;
        $data = $this->data;

        $code = v($data,$pos);
        $length = v($data,$pos+2);
        $version = v($data,$pos+4);
        $substreamType = v($data,$pos+6);

        $this->version = $version;

        if (($version != SPREADSHEET_EXCEL_READER_BIFF8) &&
            ($version != SPREADSHEET_EXCEL_READER_BIFF7)) {
            return false;
        }

        if ($substreamType != SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS){
            return false;
        }

        $pos += $length + 4;
        $code = v($data,$pos);
        $length = v($data,$pos+2);

        while ($code != SPREADSHEET_EXCEL_READER_TYPE_EOF) {

            switch ($code) {
                case SPREADSHEET_EXCEL_READER_TYPE_SST:
                    $spos = $pos + 4;
                    $limitpos = $spos + $length;
                    $uniqueStrings = $this->_GetInt4d($data, $spos+4);
                    $spos += 8;

                    for ($i = 0; $i < $uniqueStrings; $i++) {
                        // Read in the number of characters
                        if ($spos == $limitpos) {
                            $opcode = v($data,$spos);
                            $conlength = v($data,$spos+2);
                            if ($opcode != 0x3c) {
                                return -1;
                            }
                            $spos += 4;
                            $limitpos = $spos + $conlength;
                        }
                        $numChars = ord($data[$spos]) | (ord($data[$spos+1]) << 8);
                        $spos += 2;
                        $optionFlags = ord($data[$spos]);
                        $spos++;
                        $asciiEncoding = (($optionFlags & 0x01) == 0) ;
                        $extendedString = ( ($optionFlags & 0x04) != 0);

                        // See if string contains formatting information
                        $richString = ( ($optionFlags & 0x08) != 0);

                        if ($richString) {
                            // Read in the crun
                            $formattingRuns = v($data,$spos);
                            $spos += 2;
                        }

                        if ($extendedString) {
                            // Read in cchExtRst
                            $extendedRunLength = $this->_GetInt4d($data, $spos);
                            $spos += 4;
                        }

                        $len = ($asciiEncoding)? $numChars : $numChars*2;
                        if ($spos + $len < $limitpos) {
                            $retstr = substr($data, $spos, $len);

                            $spos += $len;
                        } else {
                            // found countinue
                            $retstr = substr($data, $spos, $limitpos - $spos);
                            $bytesRead = $limitpos - $spos;
                            $charsLeft = $numChars - (($asciiEncoding) ? $bytesRead : ($bytesRead / 2));
                            $spos = $limitpos;

                            while ($charsLeft > 0){
                                $opcode = v($data,$spos);
                                $conlength = v($data,$spos+2);
                                if ($opcode != 0x3c) {
                                    return -1;
                                }
                                $spos += 4;
                                $limitpos = $spos + $conlength;
                                $option = ord($data[$spos]);
                                $spos += 1;
                                if ($asciiEncoding && ($option == 0)) {
                                    $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len;
                                    $asciiEncoding = true;
                                }
                                elseif (!$asciiEncoding && ($option != 0)) {
                                    $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len/2;
                                    $asciiEncoding = false;
                                }
                                elseif (!$asciiEncoding && ($option == 0)) {
                                    // Bummer - the string starts off as Unicode, but after the
                                    // continuation it is in straightforward ASCII encoding
                                    $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                    for ($j = 0; $j < $len; $j++) {
                                        $retstr .= $data[$spos + $j].chr(0);
                                    }
                                    $charsLeft -= $len;
                                    $asciiEncoding = false;
                                }
                                else{
                                    $newstr = '';
                                    for ($j = 0; $j < strlen($retstr); $j++) {
                                        $newstr = $retstr[$j].chr(0);
                                    }
                                    $retstr = $newstr;
                                    $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len/2;
                                    $asciiEncoding = false;
                                }
                                $spos += $len;
                            }
                        }
                        $retstr = ($asciiEncoding) ? $retstr : $this->_encodeUTF16($retstr);

                        if ($richString){
                            $spos += 4 * $formattingRuns;
                        }

                        // For extended strings, skip over the extended string data
                        if ($extendedString) {
                            $spos += $extendedRunLength;
                        }

                        //$fileReader->appendLine($retstr);

                        $this->sst[] = $retstr;
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FILEPASS:
                    return false;
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_NAME:
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_FORMAT:
                    $indexCode = v($data,$pos+4);
                    if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $numchars = v($data,$pos+6);
                        if (ord($data[$pos+8]) == 0){
                            $formatString = substr($data, $pos+9, $numchars);
                        } else {
                            $formatString = substr($data, $pos+9, $numchars*2);
                        }
                    } else {
                        $numchars = ord($data[$pos+6]);
                        $formatString = substr($data, $pos+7, $numchars*2);
                    }
                    $this->formatRecords[$indexCode] = $formatString;
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_XF:
                    $fontIndexCode = (ord($data[$pos+4]) | ord($data[$pos+5]) << 8) - 1;
                    $fontIndexCode = max(0,$fontIndexCode);
                    $indexCode = ord($data[$pos+6]) | ord($data[$pos+7]) << 8;
                    $alignbit = ord($data[$pos+10]) & 3;
                    $bgi = (ord($data[$pos+22]) | ord($data[$pos+23]) << 8) & 0x3FFF;
                    $bgcolor = ($bgi & 0x7F);
//						$bgcolor = ($bgi & 0x3f80) >> 7;
                    $align = "";
                    if ($alignbit==3) { $align="right"; }
                    if ($alignbit==2) { $align="center"; }

                    $fillPattern = (ord($data[$pos+21]) & 0xFC) >> 2;
                    if ($fillPattern == 0) {
                        $bgcolor = "";
                    }

                    $xf = array();
                    $xf['formatIndex'] = $indexCode;

                    if (array_key_exists($indexCode, $this->dateFormats)) {
                        $xf['type'] = 'date';
                        $xf['format'] = $this->dateFormats[$indexCode];
                        if ($align=='') { $xf['align'] = 'right'; }
                    }elseif (array_key_exists($indexCode, $this->numberFormats)) {
                        $xf['type'] = 'number';
                        $xf['format'] = $this->numberFormats[$indexCode];
                        if ($align=='') { $xf['align'] = 'right'; }
                    }else{
                        $isdate = FALSE;
                        $formatstr = '';
                        if ($indexCode > 0){
                            if (isset($this->formatRecords[$indexCode]))
                                $formatstr = $this->formatRecords[$indexCode];
                            if ($formatstr!="") {
                                $tmp = preg_replace("/\;.*/","",$formatstr);
                                $tmp = preg_replace("/^\[[^\]]*\]/","",$tmp);
                                if (preg_match("/[^hmsday\/\-:\s\\\,AMP]/i", $tmp) == 0) { // found day and time format
                                    $isdate = TRUE;
                                    $formatstr = $tmp;
                                    $formatstr = str_replace(array('AM/PM','mmmm','mmm'), array('a','F','M'), $formatstr);
                                    // m/mm are used for both minutes and months - oh SNAP!
                                    // This mess tries to fix for that.
                                    // 'm' == minutes only if following h/hh or preceding s/ss
                                    $formatstr = preg_replace("/(h:?)mm?/","$1i", $formatstr);
                                    $formatstr = preg_replace("/mm?(:?s)/","i$1", $formatstr);
                                    // A single 'm' = n in PHP
                                    $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                    $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                    // else it's months
                                    $formatstr = str_replace('mm', 'm', $formatstr);
                                    // Convert single 'd' to 'j'
                                    $formatstr = preg_replace("/(^|[^d])d([^d]|$)/", '$1j$2', $formatstr);
                                    $formatstr = str_replace(array('dddd','ddd','dd','yyyy','yy','hh','h'), array('l','D','d','Y','y','H','g'), $formatstr);
                                    $formatstr = preg_replace("/ss?/", 's', $formatstr);
                                }
                            }
                        }
                        if ($isdate){
                            $xf['type'] = 'date';
                            $xf['format'] = $formatstr;
                            if ($align=='') { $xf['align'] = 'right'; }
                        }else{
                            // If the format string has a 0 or # in it, we'll assume it's a number
                            if (preg_match("/[0#]/", $formatstr)) {
                                $xf['type'] = 'number';
                                if ($align=='') { $xf['align']='right'; }
                            }
                            else {
                                $xf['type'] = 'other';
                            }
                            $xf['format'] = $formatstr;
                            $xf['code'] = $indexCode;
                        }
                    }
                    $this->xfRecords[] = $xf;
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR:
                    $this->nineteenFour = (ord($data[$pos+4]) == 1);
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET:

                    $rec_offset = $this->_GetInt4d($data, $pos+4);
                    $rec_typeFlag = ord($data[$pos+8]);
                    $rec_visibilityFlag = ord($data[$pos+9]);
                    $rec_length = ord($data[$pos+10]);

                    if ($version == SPREADSHEET_EXCEL_READER_BIFF8){
                        $chartype =  ord($data[$pos+11]);
                        if ($chartype == 0){
                            $rec_name	= substr($data, $pos+12, $rec_length);
                        } else {
                            $rec_name	= $this->_encodeUTF16(substr($data, $pos+12, $rec_length*2));
                        }
                    }elseif ($version == SPREADSHEET_EXCEL_READER_BIFF7){
                        $rec_name	= substr($data, $pos+11, $rec_length);
                    }
                    $this->boundsheets[] = array('name'=>$rec_name,'offset'=>$rec_offset);
                    break;

            }

            $pos += $length + 4;
            $code = ord($data[$pos]) | ord($data[$pos+1])<<8;
            $length = ord($data[$pos+2]) | ord($data[$pos+3])<<8;
        }

        foreach ($this->boundsheets as $index => $boundsheet) {
            $this->createSheetIndex($index, $boundsheet['offset']);
        }
    }

    protected function createSheetIndex($sn, $spos)
    {
        $cont = true;
        $data = $this->data;
        // read BOF
        $code = ord($data[$spos]) | ord($data[$spos+1])<<8;
        $length = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
        $version = ord($data[$spos + 4]) | ord($data[$spos + 5])<<8;

        if (!$this->isSheetVersionAppropriate($sn, $data)) {
            return -1;
        }

        if (!$this->isSubStreamTypeAppropriate($sn, $data)){
            return -2;
        }

        $spos += $length + 4;
        $rowsCounter = 0;
        while($cont) {


            $lowcode = ord($data[$spos]);

            if ($lowcode == SPREADSHEET_EXCEL_READER_TYPE_EOF) {
                break;
            }

            $code = $lowcode | ord($data[$spos+1])<<8;
            $length = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
            $spos += 4;

            $this->sheets[$sn]['maxrow'] = $this->_rowoffset - 1;
            $this->sheets[$sn]['maxcol'] = $this->_coloffset - 1;

            unset($this->rectype);
            $previousLength = 0;

            switch ($code) {
                case SPREADSHEET_EXCEL_READER_TYPE_DIMENSION:
                    if (!isset($this->numRows)) {
                        if (($length == 10) ||  ($version == SPREADSHEET_EXCEL_READER_BIFF7)){
                            $this->sheets[$sn]['numRows'] = ord($data[$spos+2]) | ord($data[$spos+3]) << 8;
                            $this->sheets[$sn]['numCols'] = ord($data[$spos+6]) | ord($data[$spos+7]) << 8;
                        } else {
                            $this->sheets[$sn]['numRows'] = ord($data[$spos+4]) | ord($data[$spos+5]) << 8;
                            $this->sheets[$sn]['numCols'] = ord($data[$spos+10]) | ord($data[$spos+11]) << 8;
                        }
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_ROW:
                    $rowsCounter++;
                    /*$row = ord($data[$spos]) | ord($data[$spos+1])<<8;

                    $this->rowInfo[$sn][$row] = Array(
                        'spos' => $spos + $length,
                    );*/

                    break;

/*                case SPREADSHEET_EXCEL_READER_TYPE_MULRK:
                    $colFirst   = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $colLast	= ord($data[$spos + $length - 2]) | ord($data[$spos + $length - 1])<<8;
                    $columns	= $colLast - $colFirst + 1;
                    $tmppos = $spos+4;
                    for ($i = 0; $i < $columns; $i++) {
                        $numValue = $this->_GetIEEE754($this->_GetInt4d($data, $tmppos + 2));
                        $info = $this->_getCellDetails($tmppos-4,$numValue,$colFirst + $i + 1);
                        $tmppos += 6;

                        $rowData[$colFirst + $i] = $info['string'];
                    }

                    break;*/
                case SPREADSHEET_EXCEL_READER_TYPE_RK:
                case SPREADSHEET_EXCEL_READER_TYPE_RK2:
                case SPREADSHEET_EXCEL_READER_TYPE_LABELSST:
                case SPREADSHEET_EXCEL_READER_TYPE_NUMBER:
                case SPREADSHEET_EXCEL_READER_TYPE_LABEL:
                $index  = $this->_GetInt4d($data, $spos + 6);

                    if (!$this->getStartPosition()) {
                        $column	 = ord($data[$spos + 2]) | ord($data[$spos+3])<<8;

                        if ($column == 0) {
                            $this->setStartPosition($spos - $previousLength - 4);
                        }
                    }
                break;
                case SPREADSHEET_EXCEL_READER_TYPE_EOF:
                    $cont = false;
                    break;

                default:
                    break;
            }

            //Need for position calculation
            $previousLength = $length;
            $spos += $length;
        }

        $this->sheets[$sn]['numRows'] = $rowsCounter;

        if (!isset($this->sheets[$sn]['numCols'])) {
            $this->sheets[$sn]['numCols'] = $this->sheets[$sn]['maxcol'];
        }
    }

    public function whichBlock($code, $ending = "\n")
    {
        $retval = '';

        switch ($code) {
            case SPREADSHEET_EXCEL_READER_TYPE_DIMENSION:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_DIMENSION' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_RK:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_RK' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_RK2:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_RK2' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_LABELSST:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_LABELSST' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_MULRK:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_MULRK' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_NUMBER:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_NUMBER' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_FORMULA:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_FORMULA' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_FORMULA2:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_FORMULA2' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_BOOLERR:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_BOOLERR' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_STRING:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_STRING' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_ROW:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_ROW' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_DBCELL:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_DBCELL' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_MULBLANK:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_MULBLANK' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_LABEL:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_LABEL' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_EOF:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_EOF' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_HYPER:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_HYPER' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_DEFCOLWIDTH:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_DEFCOLWIDTH' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_STANDARDWIDTH:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_STANDARDWIDTH' . $ending;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_COLINFO:
                $retval = $code . 'SPREADSHEET_EXCEL_READER_TYPE_COLINFO' . $ending;

                break;

            default:
                $retval = $code . 'UK' . $ending;
                break;
        }

        return $retval;
    }

    public function isDataBlock($code)
    {
        $dataTypes = array(
            SPREADSHEET_EXCEL_READER_TYPE_RK,
            SPREADSHEET_EXCEL_READER_TYPE_RK2,
            SPREADSHEET_EXCEL_READER_TYPE_LABELSST,
            SPREADSHEET_EXCEL_READER_TYPE_MULRK,
            SPREADSHEET_EXCEL_READER_TYPE_NUMBER,
            SPREADSHEET_EXCEL_READER_TYPE_LABEL,
        );

        return in_array($code, $dataTypes);
    }

    public function isSheetVersionAppropriate($sheetIndex, $data)
    {
        $spos = $this->boundsheets[$sheetIndex]['offset'];
        $version = ord($data[$spos + 4]) | ord($data[$spos + 5])<<8;

        if (($version != SPREADSHEET_EXCEL_READER_BIFF8) && ($version != SPREADSHEET_EXCEL_READER_BIFF7)) {
            return false;
        }

        return true;
    }

    public function isSubStreamTypeAppropriate($sheetIndex, $data)
    {
        $spos = $this->boundsheets[$sheetIndex]['offset'];
        $substreamType = ord($data[$spos + 6]) | ord($data[$spos + 7])<<8;

        if ($substreamType != SPREADSHEET_EXCEL_READER_WORKSHEET) {
            return false;
        }

        return true;
    }

    /**
     * Parse a worksheet
     */
    public function getRowValuesByIndex($sp, $rowIndex) {

        $cont = true;
        $data = $this->data;
        $spos = $this->getStartPosition();

        if ($spos === null || $sp > 0) {
            $spos = $this->boundsheets[$sp]['offset'];
        }

        $version = ord($data[$spos + 4]) | ord($data[$spos + 5])<<8;
        if (!$this->isSheetVersionAppropriate($sp, $data)) {
            return -1;
        }

        if (!$this->isSubStreamTypeAppropriate($sp, $data)){
            return -2;
        }

        $rowData = array();
        $iterations = 0;
        while($cont) {
            $iterations++;
            $lowcode = ord($data[$spos]);

            if ($lowcode == SPREADSHEET_EXCEL_READER_TYPE_EOF) break;

            $code = $lowcode | ord($data[$spos+1])<<8;
            $length = ord($data[$spos+2]) | ord($data[$spos+3])<<8;

            $spos += 4;
            $this->sheets[$sp]['maxrow'] = $this->_rowoffset - 1;
            $this->sheets[$sp]['maxcol'] = $this->_coloffset - 1;
            unset($this->rectype);

            $row = ord($data[$spos]) | ord($data[$spos+1])<<8;

            switch ($code) {
                case SPREADSHEET_EXCEL_READER_TYPE_RK:
                case SPREADSHEET_EXCEL_READER_TYPE_RK2:
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $rknum = $this->_GetInt4d($data, $spos + 6);
                    $numValue = $this->_GetIEEE754($rknum);
                    $info = $this->_getCellDetails($spos,$numValue,$column);

                    $rowData[$row][$column+1] = $info['string'];

                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_LABELSST:
                    $column	 = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $index  = $this->_GetInt4d($data, $spos + 6);
                    $sstValue = $this->sst[$index];

                    $rowData[$row][$column+1] = $sstValue;
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_MULRK:
                    $colFirst   = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $colLast	= ord($data[$spos + $length - 2]) | ord($data[$spos + $length - 1])<<8;
                    $columns	= $colLast - $colFirst + 1;
                    $tmppos = $spos+4;
                    for ($i = 0; $i < $columns; $i++) {
                        $numValue = $this->_GetIEEE754($this->_GetInt4d($data, $tmppos + 2));
                        $info = $this->_getCellDetails($tmppos-4,$numValue,$colFirst + $i + 1);
                        $tmppos += 6;

                        $rowData[$row][($colFirst+1) + $i] = $info['string'];
                    }
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_NUMBER:
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $tmp = unpack("ddouble", substr($data, $spos + 6, 8)); // It machine machine dependent
                    if ($this->isDate($spos)) {
                        $numValue = $tmp['double'];
                    }
                    else {
                        $numValue = $this->createNumber($spos);
                    }
                    $info = $this->_getCellDetails($spos,$numValue,$column);

                    $rowData[($row)][($column+1)] = $info['string'];
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FORMULA:
                case SPREADSHEET_EXCEL_READER_TYPE_FORMULA2:
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    if ((ord($data[$spos+6])==0) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //String formula. Result follows in a STRING record
                        // This row/col are stored to be referenced in that record
                        // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                        $previousRow = $row;
                        $previousCol = $column;
                    } elseif ((ord($data[$spos+6])==1) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //Boolean formula. Result is in +2; 0=false,1=true
                        // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                        if (ord($this->data[$spos+8])==1) {
                            $rowData[($row)][($column+1)] = "TRUE";
                        } else {
                            $rowData[($row)][($column+1)] = "FALSE";
                        }
                    } elseif ((ord($data[$spos+6])==2) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //Error formula. Error code is in +2;
                    } elseif ((ord($data[$spos+6])==3) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //Formula result is a null string.
                        $rowData[($row)][($column+1)] = "";
                    } else {
                        // result is a number, so first 14 bytes are just like a _NUMBER record
                        $tmp = unpack("ddouble", substr($data, $spos + 6, 8)); // It machine machine dependent
                              if ($this->isDate($spos)) {
                                $numValue = $tmp['double'];
                              }
                              else {
                                $numValue = $this->createNumber($spos);
                              }
                        $info = $this->_getCellDetails($spos,$numValue,$column);
                        $rowData[($row)][($column+1)] = $info['string'];
                    }
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_BOOLERR:
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $string = ord($data[$spos+6]);
                    $rowData[($row)][($column+1)] = $string;
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_STRING:
                    // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                    if ($version == SPREADSHEET_EXCEL_READER_BIFF8){
                        // Unicode 16 string, like an SST record
                        $xpos = $spos;
                        $numChars =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
                        $xpos += 2;
                        $optionFlags =ord($data[$xpos]);
                        $xpos++;
                        $asciiEncoding = (($optionFlags &0x01) == 0) ;
                        $extendedString = (($optionFlags & 0x04) != 0);
                        // See if string contains formatting information
                        $richString = (($optionFlags & 0x08) != 0);
                        if ($richString) {
                            // Read in the crun
                            $formattingRuns =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
                            $xpos += 2;
                        }
                        if ($extendedString) {
                            // Read in cchExtRst
                            $extendedRunLength =$this->_GetInt4d($this->data, $xpos);
                            $xpos += 4;
                        }
                        $len = ($asciiEncoding)?$numChars : $numChars*2;
                        $retstr =substr($data, $xpos, $len);
                        $xpos += $len;
                        $retstr = ($asciiEncoding)? $retstr : $this->_encodeUTF16($retstr);
                    }
                    elseif ($version == SPREADSHEET_EXCEL_READER_BIFF7){
                        // Simple byte string
                        $xpos = $spos;
                        $numChars =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
                        $xpos += 2;
                        $retstr =substr($data, $xpos, $numChars);
                    }

                    $rowData[$previousRow][$previousCol+1] = $retstr;
                    //$this->addcell($sp, $previousRow, $previousCol, $retstr);
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_MULBLANK:
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $cols = ($length / 2) - 3;
                    for ($c = 0; $c < $cols; $c++) {
                        $rowData[$row][($column+1) + $c] = "";
                    }
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_LABEL:
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $rowData[$row][$column+1] = substr($data, $spos + 8, ord($data[$spos + 6]) | ord($data[$spos + 7]) << 8);
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_ROW:
                    if ($row == ($rowIndex + 32)) {
                        $this->setStartPosition($spos + $length);
                        $cont = false;
                    }
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_EOF:
                    $cont = false;
                    break;
                default:
                    break;
            }

            $spos += $length;
        }

        //echo 'Iterations:' . $iterations . "\n";

        return $rowData;
    }

    function isDate($spos) {
        $xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;

        return ($this->xfRecords[$xfindex]['type'] == 'date');
    }

    // Get the details for a particular cell
    function _getCellDetails($spos,$numValue,$column) {
        $xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
        $xfrecord = $this->xfRecords[$xfindex];
        $type = $xfrecord['type'];

        $format = $xfrecord['format'];
        $formatIndex = $xfrecord['formatIndex'];
//        $fontIndex = $xfrecord['fontIndex'];
        $formatColor = "";
        $rectype = '';
        $string = '';
        $raw = '';

        if (isset($this->_columnsFormat[$column + 1])){
            $format = $this->_columnsFormat[$column + 1];
        }

        if ($type == 'date') {
            // See http://groups.google.com/group/php-excel-reader-discuss/browse_frm/thread/9c3f9790d12d8e10/f2045c2369ac79de
            $rectype = 'date';
            // Convert numeric value into a date
            $utcDays = floor($numValue - ($this->nineteenFour ? SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904 : SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS));
            $utcValue = ($utcDays) * SPREADSHEET_EXCEL_READER_MSINADAY;
            $dateinfo = gmgetdate($utcValue);

            $raw = $numValue;
            $fractionalDay = $numValue - floor($numValue) + .0000001; // The .0000001 is to fix for php/excel fractional diffs

            $totalseconds = floor(SPREADSHEET_EXCEL_READER_MSINADAY * $fractionalDay);
            $secs = $totalseconds % 60;
            $totalseconds -= $secs;
            $hours = floor($totalseconds / (60 * 60));
            $mins = floor($totalseconds / 60) % 60;
            $string = date ($format, mktime($hours, $mins, $secs, $dateinfo["mon"], $dateinfo["mday"], $dateinfo["year"]));
        } else if ($type == 'number') {
            $rectype = 'number';
            $formatted = $this->_format_value($format, $numValue, $formatIndex);
            $string = $formatted['string'];
//            $formatColor = $formatted['formatColor'];
            $raw = $numValue;
        } else{
            if ($format=="") {
                $format = $this->_defaultFormat;
            }
            $rectype = 'unknown';
            $formatted = $this->_format_value($format, $numValue, $formatIndex);
            $string = $formatted['string'];
//            $formatColor = $formatted['formatColor'];
            $raw = $numValue;
        }

        return array(
            'string'=>$string,
            'raw'=>$raw,
            'rectype'=>$rectype,
            'format'=>$format,
//            'formatIndex'=>$formatIndex,
//            'fontIndex'=>$fontIndex,
            'formatColor'=>$formatColor,
            'xfIndex'=>$xfindex
        );

    }


    function createNumber($spos) {
        $rknumhigh = $this->_GetInt4d($this->data, $spos + 10);
        $rknumlow = $this->_GetInt4d($this->data, $spos + 6);
        $sign = ($rknumhigh & 0x80000000) >> 31;
        $exp =  ($rknumhigh & 0x7ff00000) >> 20;
        $mantissa = (0x100000 | ($rknumhigh & 0x000fffff));
        $mantissalow1 = ($rknumlow & 0x80000000) >> 31;
        $mantissalow2 = ($rknumlow & 0x7fffffff);
        $value = $mantissa / pow( 2 , (20- ($exp - 1023)));
        if ($mantissalow1 != 0) $value += 1 / pow (2 , (21 - ($exp - 1023)));
        $value += $mantissalow2 / pow (2 , (52 - ($exp - 1023)));
        if ($sign) {$value = -1 * $value;}
        return  $value;
    }

    function _GetIEEE754($rknum) {
        if (($rknum & 0x02) != 0) {
            $value = $rknum >> 2;
        } else {
            //mmp
            // I got my info on IEEE754 encoding from
            // http://research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
            // The RK format calls for using only the most significant 30 bits of the
            // 64 bit floating point value. The other 34 bits are assumed to be 0
            // So, we use the upper 30 bits of $rknum as follows...
            $sign = ($rknum & 0x80000000) >> 31;
            $exp = ($rknum & 0x7ff00000) >> 20;
            $mantissa = (0x100000 | ($rknum & 0x000ffffc));
            $value = $mantissa / pow( 2 , (20- ($exp - 1023)));
            if ($sign) {
                $value = -1 * $value;
            }
            //end of changes by mmp
        }
        if (($rknum & 0x01) != 0) {
            $value /= 100;
        }
        return $value;
    }

    function _encodeUTF16($string) {
        $result = $string;
        if ($this->_defaultEncoding){
            switch ($this->_encoderFunction){
                case 'iconv' :	 $result = iconv('UTF-16LE', $this->_defaultEncoding, $string);
                    break;
                case 'mb_convert_encoding' : $result = mb_convert_encoding($string, $this->_defaultEncoding, 'UTF-16LE' );
                    break;
            }
        }
        return $result;
    }

    function _GetInt4d($data, $pos) {
        $value = ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | (ord($data[$pos+3]) << 24);
        if ($value >= 4294967294) {
            $value =- 2;
        }
        return $value;
    }

    /**
     * @param array $sheets
     */
    public function setSheets($sheets)
    {
        $this->sheets = $sheets;
    }

    /**
     * @return array
     */
    public function getSheets()
    {
        return $this->sheets;
    }

    /**
     * @param array $boundsheets
     */
    public function setBoundsheets($boundsheets)
    {
        $this->boundsheets = $boundsheets;
    }

    /**
     * @return array
     */
    public function getBoundsheets()
    {
        return $this->boundsheets;
    }

    /**
     * @param null $startPosition
     */
    public function setStartPosition($startPosition)
    {
        $this->startPosition = $startPosition;
    }

    /**
     * @return null
     */
    public function getStartPosition()
    {
        return $this->startPosition;
    }
}