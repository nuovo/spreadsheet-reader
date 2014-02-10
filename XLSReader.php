<?php
/**
* Excel XLS file reader class for reading XLS files efficiently (without parsing all of the file and
*	only getting the requested bit at a time. No formatting, graphs, charts, images, etc, only data.
* Because all of the existing ones suck.
*
* Requirements: iconv or mbstring
*
* @author Martins Pilsetnieks
*/

class Reader implements \Iterator
{
    const ERROR_NO_ERROR = 0;
    const ERROR_NOT_READABLE = 1;
    const ERROR_FILE_EMPTY = 2;
    const ERROR_INVALID_FILE = 4;
    const ERROR_UNKNOWN_VERSION = 8;
    const ERROR_NOT_A_WORKBOOK = 16;

    const OLE_NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
    const OLE_ROOT_START_BLOCK_POS = 0x30;
    const OLE_SMALL_BLOCK_DEPOT_BLOCK_POS = 0x3c;
    const OLE_SMALL_BLOCK_SIZE = 0x40;
    const OLE_EXTENSION_BLOCK_POS = 0x44;
    const OLE_NUM_EXTENSION_BLOCK_POS = 0x48;
    const OLE_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x4c;
    const OLE_PROPERTY_STORAGE_BLOCK_SIZE = 0x80;
    const OLE_BIG_BLOCK_SIZE = 0x200;

    const OLE_SIZE_OF_NAME_POS = 0x40;
    const OLE_TYPE_POS = 0x42;
    const OLE_START_BLOCK_POS = 0x74;
    const OLE_SIZE_POS = 0x78;
    const OLE_SMALL_BLOCK_THRESHOLD = 0x1000;

    const VERSION_BIFF7 = 0x500;
    const VERSION_BIFF8 = 0x600;
    const WORKBOOKGLOBALS = 0x5;

    const TYPE_EOF = 0x0a;
    const TYPE_BOF = 0x809;
    const TYPE_BOUNDSHEET = 0x85;
    const TYPE_DIMENSION = 0x200;

    private static $WorkbookMemory = 5242880;

    private $Error = 0;

    /**
     * @var resource Internal file handle
     */
    private $Handle = false;

    private $Workbook = false;

    private $OLE_RootStartBlock = 0;

    private $OLE_BigBlockDepotBlocks = array();
    private $OLE_NumBigBlockDepotBlocks = 0;

    private $OLE_ExtensionBlock = 0;
    private $OLE_NumExtensionBlocks = 0;

    private $OLE_SmallBlockChain = array();
    private $OLE_BigBlockChain = array();

    private $OLE_Props = array();
    private $OLE_WorkbookPropIndex = -1;
    private $OLE_RootEntryPropIndex = -1;

    /**
     * @param string Path to file
     */
    public function __construct($File)
    {
        // Error: Unreadable file, exit
        if (!is_readable($File))
        {
            $this -> Error = self::ERROR_NOT_READABLE;
            throw new Exception('Cannot read file');
        }

        // Error: Empty file, exit
        if (!filesize($File))
        {
            $this -> Error = self::ERROR_FILE_EMPTY;
            return null;
        }

        $this -> Handle = fopen($File, 'rb');

        // Error: Not an OLE file
        // 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 = OLE identifier
        if (fread($this -> Handle, 8) != pack('CCCCCCCC', 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1))
        {
            $this -> Error = self::ERROR_INVALID_FILE;
            fclose($this -> Handle);

            return null;
        }

        $this -> Workbook = fopen('php://temp/maxmemory='.self::$WorkbookMemory, 'wrb');

        $this -> OLE_NumBigBlockDepotBlocks = $this -> int4d(self::OLE_NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);
        $this -> OLE_ExtensionBlock = $this -> int4d(self::OLE_EXTENSION_BLOCK_POS);
        $this -> OLE_NumExtensionBlocks = $this -> int4d(self::OLE_NUM_EXTENSION_BLOCK_POS);
        $this -> OLE_RootStartBlock = $this -> int4d(self::OLE_ROOT_START_BLOCK_POS);
        $this -> OLE_SBDStartBlock = $this -> int4d(self::OLE_SMALL_BLOCK_DEPOT_BLOCK_POS);

        $BBDBlocks = $this -> OLE_NumBigBlockDepotBlocks;
        if ($this -> OLE_NumExtensionBlocks)
        {
            $BBDBlocks = (self::OLE_BIG_BLOCK_SIZE - self::OLE_BIG_BLOCK_DEPOT_BLOCKS_POS) / 4;
        }

        for ($i = 0, $Position = self::OLE_BIG_BLOCK_DEPOT_BLOCKS_POS; $i < $BBDBlocks; $i++, $Position += 4)
        {
            $this -> OLE_BigBlockDepotBlocks[$i] = $this -> int4d($Position);
        }

        $BlocksToRead = min($this -> OLE_NumBigBlockDepotBlocks - $BBDBlocks, self::OLE_BIG_BLOCK_SIZE / 4 - 1);
        for ($i = 0; $i < $this -> OLE_NumExtensionBlocks; $i++)
        {
            for (
                $j = $BBDBlocks, $Position = ($this -> OLE_ExtensionBlock + 1) * self::OLE_BIG_BLOCK_SIZE;
                $j < $BBDBlocks + $BlocksToRead;
                $j++, $Position += 4
            )
            {
                $this -> OLE_BigBlockDepotBlocks[$j] = $this -> int4d($Position);
            }

            $BBDBlocks += $BlocksToRead;
            if ($BBDBlocks < $this -> OLE_NumBigBlockDepotBlocks)
            {
                $this -> OLE_ExtensionBlock = $this -> int4d($Position);
            }
        }

        // Reading big block chain
        $Index = 0;
        for ($i = 0; $i < $this -> OLE_NumBigBlockDepotBlocks; $i++)
        {
            $Position = ($this -> OLE_BigBlockDepotBlocks[$i] + 1) * self::OLE_BIG_BLOCK_SIZE;
            for ($j = 0; $j < self::OLE_BIG_BLOCK_SIZE / 4; $j++, $Position += 4)
            {
                $this -> OLE_BigBlockChain[$Index++] = $this -> int4d($Position);
            }
        }

        $Block = $this -> OLE_RootStartBlock;
        $RSBData = '';
        do
        {
            fseek($this -> Handle, ($Block + 1) * self::OLE_BIG_BLOCK_SIZE);
            $RSBData .= fread($this -> Handle, self::OLE_BIG_BLOCK_SIZE);
        }
        while (($Block = $this -> OLE_BigBlockChain[$Block]) != -2);

        // Reading small block chain
        $SBDBlock = $this -> OLE_SBDStartBlock;
        $Index = 0;
        do
        {
            for (
                $i = 0, $Position = ($SBDBlock + 1) * self::OLE_BIG_BLOCK_SIZE;
                $i < self::OLE_BIG_BLOCK_SIZE / 4;
                $i++, $Position += 4
            )
            {
                $this -> OLE_SmallBlockChain[$Index++] = $this -> int4d($Position);
            }
        }
        while (($SBDBlock = $this -> OLE_BigBlockChain[$SBDBlock]) != -2);

        // Reading properties
        $Offset = 0;
        do
        {
            $Data = substr($RSBData, $Offset, self::OLE_PROPERTY_STORAGE_BLOCK_SIZE);
            $NameSize = ord($Data[self::OLE_SIZE_OF_NAME_POS]) | (ord($Data[self::OLE_SIZE_OF_NAME_POS + 1]) << 8);
            $Type = ord($Data[self::OLE_TYPE_POS]);
            $StartBlock = $this -> int4d_2(substr($Data, self::OLE_START_BLOCK_POS, 4));
            $Size = $this -> int4d_2(substr($Data, self::OLE_SIZE_POS, 4));

            $Name = substr($Data, 0, $NameSize);
            $Name = str_replace("\x00", '', $Name);

            $this -> OLE_Props[] = array(
                'Name' => $Name,
                'Type' => $Type,
                'StartBlock' => $StartBlock,
                'Size' => $Size
            );

            $Name = strtolower($Name);
            if ($Name == 'workbook' || $Name == 'book')
            {
                $this -> OLE_WorkbookPropIndex = count($this -> OLE_Props) - 1;
            }
            elseif ($Name == 'root entry')
            {
                $this -> OLE_RootEntryPropIndex = count($this -> OLE_Props) - 1;
            }
        }
        while (($Offset += self::OLE_PROPERTY_STORAGE_BLOCK_SIZE) < strlen($RSBData));

        // Getting workbook
        if ($this -> OLE_Props[$this -> OLE_WorkbookPropIndex]['Size'] < self::OLE_SMALL_BLOCK_THRESHOLD)
        {
            $RootData = $this -> OLE_ReadData($this -> OLE_Props[$this -> OLE_RootEntryPropIndex]['StartBlock']);

            $Block = $this -> OLE_Props[$this -> OLE_WorkbookPropIndex]['StartBlock'];
            do
            {
                fwrite($this -> Workbook,
                    substr($RootData, $Block * self::OLE_SMALL_BLOCK_SIZE, self::OLE_SMALL_BLOCK_SIZE),
                    self::OLE_SMALL_BLOCK_SIZE);
            }
            while (($Block = $this -> OLE_SmallBlockChain[$Block]) != -2);
        }
        else
        {
            $NumBlocks = $this -> OLE_Props[$this -> OLE_WorkbookPropIndex]['Size'] / self::OLE_BIG_BLOCK_SIZE;
            if ($this -> OLE_Props[$this -> OLE_WorkbookPropIndex]['Size'] % self::OLE_BIG_BLOCK_SIZE)
            {
                $NumBlocks++;
            }

            if ($NumBlocks)
            {
                $Block = $this -> OLE_Props[$this -> OLE_WorkbookPropIndex]['StartBlock'];
                do
                {
                    fseek($this -> Handle, ($Block + 1) * self::OLE_BIG_BLOCK_SIZE);
                    fwrite($this -> Workbook,
                        fread($this -> Handle, self::OLE_BIG_BLOCK_SIZE), self::OLE_BIG_BLOCK_SIZE);
                }
                while (($Block = $this -> OLE_BigBlockChain[$Block]) != -2);
            }
        }

        // $this -> Workbook now contains what Spreadsheet_Excel_Reader contains in $this -> data

        fseek($this -> Workbook, 0);

        $Code = $this -> int2d(null, false);
        $Length = $this -> int2d(null, false);
        $Version = $this -> int2d(null, false);
        $SubstreamType = $this -> int2d(null, false);

        if ($Version != self::VERSION_BIFF7 && $Version != self::VERSION_BIFF8)
        {
            $this -> Error = self::ERROR_UNKNOWN_VERSION;
            return null;
        }

        if ($SubstreamType != self::WORKBOOKGLOBALS)
        {
            $this -> Error = self::ERROR_NOT_A_WORKBOOK;
            return null;
        }

        $Position = $Length + 4;

        $Code = $this -> int2d($Position);
        $Length = $this -> int2d($Position + 2);

        while ($Code != self::TYPE_EOF)
        {
            switch ($Code)
            {
                case SPREADSHEET_EXCEL_READER_TYPE_XF:
                    $IndexCode = $this -> int2d($Position + 6);
                    $XF = array();

                    if (array_key_exists($IndexCode, self::$DateFormats))
                    {
                        $XF['Type'] = 'Date';
                        $XF['Format'] = self::$DateFormats[$IndexCode];
                    }
                    elseif (array_key_exists($IndexCode, self::$NumberFormats))
                    {
                        $XF['Type'] = 'Number';
                        $XF['Format'] = self::$NumberFormats[$IndexCode];
                    }
                    else
                    {
                        if ($IndexCode > 0)
                        {

                        }
                    }
/*
                    else
                    {
                        $isdate = FALSE;
                        $formatstr = '';
                        if ($indexCode > 0){
                            if (isset($this->formatRecords[$indexCode]))
                                $formatstr = $this->formatRecords[$indexCode];
                            if ($formatstr!="") {*/
                                //$tmp = preg_replace("/\;.*/","",$formatstr);
/*									$tmp = preg_replace("/^\[[^\]]*\]/","",$tmp);
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
                    $this->xfRecords[] = $xf;*/
                break;
            }
        }
    }

    public function Sheets()
    {
        if ($this -> Sheets === false)
        {
            $this -> Sheets = array();
            foreach ($this -> WorkbookXML -> sheets -> sheet as $Index => $Sheet)
            {

                $Attributes = $Sheet -> attributes();
                foreach ($Attributes as $Name => $Value)
                {
                    if ($Name == 'sheetId') {
                        $SheetID = (int)$Value;
                        break;
                    }
                }

                $this -> Sheets[$SheetID] = (string)$Sheet['name'];
            }
            ksort($this -> Sheets);
        }

        return array_values($this -> Sheets);
    }

    public function __get($Name)
    {
        if ($Name == 'Error')
        {
            return $this -> Error;
        }
        return null;
    }

    public function __destruct()
    {
        if ($this -> Handle)
        {
            fclose($this -> Handle);
        }
        if ($this -> Workbook)
        {
            fclose($this -> Workbook);
        }
    }

    // !Iterator interface methods
    public function rewind()
    {
    }

    public function key()
    {
    }

    public function valid()
    {
    }

    public function current()
    {
    }

    public function next()
    {
    }

    private function GetBytes($Position, $Bytes, $KeepPosition)
    {
        if ($KeepPosition)
        {
            $OldPosition = ftell($this -> Handle);
        }

        if (!is_null($Position))
        {
            fseek($this -> Handle, $Position);
        }

        $Val = fread($this -> Handle, $Bytes);

        if ($KeepPosition)
        {
            fseek($this -> Handle, $OldPosition);
            unset($OldPosition);
        }

        return $Val;
    }

    private function int4d($Position = null, $KeepPosition = false)
    {
        return $this -> int4d_2($this -> GetBytes($Position, 4, $KeepPosition));
    }

    private function int4d_2($Val)
    {
        $Val = ord($Val[0]) | (ord($Val[1]) << 8) | (ord($Val[2]) << 16) | (ord($Val[3]) << 24);
        return $Val >= 4294967294 ? -2 : $Val;
    }

    private function int2d($Position = null, $KeepPosition = false)
    {
        return $this -> int2d_2($this -> GetBytes($Position, 4, $KeepPosition));
    }

    private function int2d_2($Val)
    {
        return ord($Val[0]) | (ord($Val[1]) << 8);
    }

    private function OLE_ReadData($Block)
    {
        if ($Block == -2)
        {
            return '';
        }

        $Result = '';
        do
        {
            fseek($this -> Handle, ($Block + 1) * self::OLE_BIG_BLOCK_SIZE);
            $Result .= fread($this -> Handle, self::OLE_BIG_BLOCK_SIZE);
        }
        while (($Block = $this -> OLE_BigBlockChain[$Block]) != -2);

        return $Result;
    }
}