<?php

namespace SpreadsheetReader\xls;

define('NUM_BIG_BLOCK_DEPOT_BLOCKS_POS', 0x2c);
define('SMALL_BLOCK_DEPOT_BLOCK_POS', 0x3c);
define('ROOT_START_BLOCK_POS', 0x30);
define('BIG_BLOCK_SIZE', 0x200);
define('SMALL_BLOCK_SIZE', 0x40);
define('EXTENSION_BLOCK_POS', 0x44);
define('NUM_EXTENSION_BLOCK_POS', 0x48);
define('PROPERTY_STORAGE_BLOCK_SIZE', 0x80);
define('BIG_BLOCK_DEPOT_BLOCKS_POS', 0x4c);
define('SMALL_BLOCK_THRESHOLD', 0x1000);
// property storage offsets
define('SIZE_OF_NAME_POS', 0x40);
define('TYPE_POS', 0x42);
define('START_BLOCK_POS', 0x74);
define('SIZE_POS', 0x78);
define('IDENTIFIER_OLE', pack("CCCCCCCC",0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1));

function getInt4d($data, $pos) {
    $value = ord($data[$pos]) | (ord($data[$pos+1])	<< 8) | (ord($data[$pos+2]) << 16) | (ord($data[$pos+3]) << 24);

    if ($value >= 4294967294) {
        $value = -2;
    }

    return $value;
}

// http://uk.php.net/manual/en/function.getdate.php
function gmgetdate($ts = null){
    $k = array('seconds', 'minutes', 'hours','mday','wday','mon','year','yday','weekday','month',0);

    return (
        array_comb($k, explode(":", gmdate('s:i:G:j:w:n:Y:z:l:F:U',is_null($ts)?time():$ts)))
    );
}

// Added for PHP4 compatibility
function array_comb($array1, $array2) {
    $out = array();
    foreach ($array1 as $key => $value) {
        $out[$value] = $array2[$key];
    }

    return $out;
}

function v($data,$pos) {
    return ord($data[$pos]) | ord($data[$pos+1])<<8;
}

function humanFileSize($size)
{
    if ($size >= 1073741824) {
        $fileSize = round($size / 1024 / 1024 / 1024,1) . 'GB';
    } elseif ($size >= 1048576) {
        $fileSize = round($size / 1024 / 1024,1) . 'MB';
    } elseif($size >= 1024) {
        $fileSize = round($size / 1024,1) . 'KB';
    } else {
        $fileSize = $size . ' bytes';
    }
    return $fileSize;
}

/**
 * Class OLEReader
 */
class OLEReader {
    /**
     * @var string
     */
    protected $data = '';

    /**
     * @var int
     */
    protected $numBigBlockDepotBlocks = 0;

    /**
     * @var int
     */
    protected $sbdStartBlock = 0;

    /**
     * @var int
     */
    protected $rootStartBlock = 0;

    /**
     * @var int
     */
    protected $extensionBlock = 0;

    /**
     * @var int
     */
    protected $numExtensionBlocks = 0;

    /**
     * @var null|int
     */
    protected $error = null;

    /**
     * @var array
     */
    protected $bigBlockChain = array();

    /**
     * @var array
     */
    protected $smallBlockChain = array();

    /**
     * @var string
     */
    protected $entry = '';

    /**
     * @var array
     */
    protected $properties = array();

    /**
     * @var null|string
     */
    protected $fileName = null;

    /**
     * @var int
     */
    protected $wrkBookCnt = 0;

    /**
     * @var int
     */
    protected $rootEntryCnt = 0;

    /**
     * @param $fileName
     * @throws Exception
     */
    function __construct($fileName)
    {
        if (file_exists($fileName) && is_readable($fileName)) {
            $this->fileName = $fileName;
            $this->read($fileName);
        } else {
            throw new Exception('File does not exist or it is not accessible');
        }
    }

    /**
     * @param $sFileName
     * @return bool
     */
    protected function read($sFileName){
        $this->setData(file_get_contents($sFileName));

        //File empty
        if (!$this->data) {
            $this->error = 1;
            return false;
        }

        //File isn't OLE document
        if (substr($this->data, 0, 8) != IDENTIFIER_OLE) {
            $this->error = 1;
            return false;
        }

        $this->setNumBigBlockDepotBlocks(getInt4d($this->getData(), NUM_BIG_BLOCK_DEPOT_BLOCKS_POS));
        $this->setSbdStartBlock(getInt4d($this->getData(), SMALL_BLOCK_DEPOT_BLOCK_POS));
        $this->setRootStartBlock(getInt4d($this->getData(), ROOT_START_BLOCK_POS));
        $this->setExtensionBlock(getInt4d($this->getData(), EXTENSION_BLOCK_POS));
        $this->setNumExtensionBlocks(getInt4d($this->getData(), NUM_EXTENSION_BLOCK_POS));

        $bigBlockDepotBlocks = array();
        $pos = BIG_BLOCK_DEPOT_BLOCKS_POS;
        $bbdBlocks = $this->numBigBlockDepotBlocks;

        if ($this->getNumExtensionBlocks() != 0) {
            $bbdBlocks = (BIG_BLOCK_SIZE - BIG_BLOCK_DEPOT_BLOCKS_POS)/4;
        }

        for ($i = 0; $i < $bbdBlocks; $i++) {
            $bigBlockDepotBlocks[$i] = getInt4d($this->getData(), $pos);
            $pos += 4;
        }

        for ($j = 0; $j < $this->getNumExtensionBlocks(); $j++) {
            $pos = ($this->getExtensionBlock() + 1) * BIG_BLOCK_SIZE;
            $blocksToRead = min($this->getNumBigBlockDepotBlocks() - $bbdBlocks, BIG_BLOCK_SIZE / 4 - 1);

            for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; $i++) {
                $bigBlockDepotBlocks[$i] = getInt4d($this->getData(), $pos);
                $pos += 4;
            }

            $bbdBlocks += $blocksToRead;
            if ($bbdBlocks < $this->getNumBigBlockDepotBlocks()) {
                $this->setExtensionBlock(getInt4d($this->getData(), $pos));
            }
        }

        // readBigBlockDepot
        $pos = 0;
        $index = 0;
        $bigBlockChain = array();

        for ($i = 0; $i < $this->getNumBigBlockDepotBlocks(); $i++) {
            $pos = ($bigBlockDepotBlocks[$i] + 1) * BIG_BLOCK_SIZE;
            //echo "pos = $pos";
            for ($j = 0 ; $j < BIG_BLOCK_SIZE / 4; $j++) {
                $bigBlockChain[$index] = getInt4d($this->getData(), $pos);
                $pos += 4 ;
                $index++;
            }
        }

        $this->setBigBlockChain($bigBlockChain);

        // readSmallBlockDepot();
        $pos = 0;
        $index = 0;
        $sbdBlock = $this->getSbdStartBlock();
        $smallBlockChain = array();

        while ($sbdBlock != -2) {
            $pos = ($sbdBlock + 1) * BIG_BLOCK_SIZE;
            for ($j = 0; $j < BIG_BLOCK_SIZE / 4; $j++) {
                $smallBlockChain[$index] = getInt4d($this->getData(), $pos);
                $pos += 4;
                $index++;
            }
            $sbdBlock = $bigBlockChain[$sbdBlock];
        }

        $this->setSmallBlockChain($smallBlockChain);

        $block = $this->getRootStartBlock();
        $this->entry = $this->readData($block);
        $this->readPropertySets();
    }

    /**
     * @param $bl
     * @return string
     */

    function readData($bl) {
        $block = $bl;
        $pos = 0;
        $data = '';
        while ($block != -2)  {
            $pos = ($block + 1) * BIG_BLOCK_SIZE;
            $data = $data . substr($this->getData(), $pos, BIG_BLOCK_SIZE);
            $block = $this->bigBlockChain[$block];
        }

        return $data;
    }

    /**
     *
     */
    protected function readPropertySets(){

        $offset = 0;
        while ($offset < strlen($this->getEntry())) {
            $d = substr($this->entry, $offset, PROPERTY_STORAGE_BLOCK_SIZE);
            $nameSize = ord($d[SIZE_OF_NAME_POS]) | (ord($d[SIZE_OF_NAME_POS+1]) << 8);
            $type = ord($d[TYPE_POS]);

            $startBlock = getInt4d($d, START_BLOCK_POS);
            $size = getInt4d($d, SIZE_POS);

            $name = '';
            for ($i = 0; $i < $nameSize ; $i++) {
                $name .= $d[$i];
            }

            $name = str_replace("\x00", "", $name);

            $this->properties[] = array (
                'name' => $name,
                'type' => $type,
                'startBlock' => $startBlock,
                'size' => $size);

            if ((strtolower($name) == "workbook") || (strtolower($name) == "book")) {
                $this->setWrkBookCnt(count($this->properties) - 1);
            }

            if ($name == "Root Entry") {
                $this->setRootEntryCnt(count($this->properties) - 1);
            }

            $offset += PROPERTY_STORAGE_BLOCK_SIZE;
        }

    }

    /**
     * @return string
     */
    function getWorkBook(){
        if ($this->properties[$this->getWrkBookCnt()]['size'] < SMALL_BLOCK_THRESHOLD){
            $rootData = $this->readData($this->properties[$this->getRootEntryCnt()]['startBlock']);

            $streamData = '';
            $block = $this->properties[$this->getWrkBookCnt()]['startBlock'];

            $pos = 0;
            while ($block != -2) {
                $pos = $block * SMALL_BLOCK_SIZE;
                $streamData .= substr($rootData, $pos, SMALL_BLOCK_SIZE);
                $block = $this->smallBlockChain[$block];
            }

            return $streamData;
        }else {
            $numBlocks = $this->properties[$this->getWrkBookCnt()]['size'] / BIG_BLOCK_SIZE;
            if ($this->properties[$this->getWrkBookCnt()]['size'] % BIG_BLOCK_SIZE != 0) {
                $numBlocks++;
            }

            if ($numBlocks == 0) {
                return '';
            }

            $streamData = '';
            $block = $this->properties[$this->getWrkBookCnt()]['startBlock'];

            $pos = 0;
            while ($block != -2) {
                $pos = ($block + 1) * BIG_BLOCK_SIZE;
                $streamData .= substr($this->data, $pos, BIG_BLOCK_SIZE);
                $block = $this->bigBlockChain[$block];
            }

            return $streamData;
        }
    }

    /**
     * @param int $wrkBookCnt
     */
    public function setWrkBookCnt($wrkBookCnt)
    {
        $this->wrkBookCnt = $wrkBookCnt;
    }

    /**
     * @return int
     */
    public function getWrkBookCnt()
    {
        return $this->wrkBookCnt;
    }

    /**
     * @param int $rootEntry
     */
    public function setRootEntryCnt($rootEntry)
    {
        $this->rootEntryCnt = $rootEntry;
    }

    /**
     * @return int
     */
    public function getRootEntryCnt()
    {
        return $this->rootEntryCnt;
    }

    /**
     * @param array $bigBlockChain
     */
    public function setBigBlockChain($bigBlockChain)
    {
        $this->bigBlockChain = $bigBlockChain;
    }

    /**
     * @return array
     */
    public function getBigBlockChain()
    {
        return $this->bigBlockChain;
    }

    /**
     * @param string $data
     */
    public function setData($data)
    {
        $this->data = $data;
    }

    /**
     * @return string
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * @param string $entry
     */
    public function setEntry($entry)
    {
        $this->entry = $entry;
    }

    /**
     * @return string
     */
    public function getEntry()
    {
        return $this->entry;
    }

    /**
     * @param int|null $error
     */
    public function setError($error)
    {
        $this->error = $error;
    }

    /**
     * @return int|null
     */
    public function getError()
    {
        return $this->error;
    }

    /**
     * @param int $extensionBlock
     */
    public function setExtensionBlock($extensionBlock)
    {
        $this->extensionBlock = $extensionBlock;
    }

    /**
     * @return int
     */
    public function getExtensionBlock()
    {
        return $this->extensionBlock;
    }

    /**
     * @param null|string $fileName
     */
    public function setFileName($fileName)
    {
        $this->fileName = $fileName;
    }

    /**
     * @return null|string
     */
    public function getFileName()
    {
        return $this->fileName;
    }

    /**
     * @param int $numBigBlockDepotBlocks
     */
    public function setNumBigBlockDepotBlocks($numBigBlockDepotBlocks)
    {
        $this->numBigBlockDepotBlocks = $numBigBlockDepotBlocks;
    }

    /**
     * @return int
     */
    public function getNumBigBlockDepotBlocks()
    {
        return $this->numBigBlockDepotBlocks;
    }

    /**
     * @param int $numExtensionBlocks
     */
    public function setNumExtensionBlocks($numExtensionBlocks)
    {
        $this->numExtensionBlocks = $numExtensionBlocks;
    }

    /**
     * @return int
     */
    public function getNumExtensionBlocks()
    {
        return $this->numExtensionBlocks;
    }

    /**
     * @param array $properties
     */
    public function setProperties($properties)
    {
        $this->properties = $properties;
    }

    /**
     * @return array
     */
    public function getProperties()
    {
        return $this->properties;
    }

    /**
     * @param int $rootStartBlock
     */
    public function setRootStartBlock($rootStartBlock)
    {
        $this->rootStartBlock = $rootStartBlock;
    }

    /**
     * @return int
     */
    public function getRootStartBlock()
    {
        return $this->rootStartBlock;
    }

    /**
     * @param int $sbdStartBlock
     */
    public function setSbdStartBlock($sbdStartBlock)
    {
        $this->sbdStartBlock = $sbdStartBlock;
    }

    /**
     * @return int
     */
    public function getSbdStartBlock()
    {
        return $this->sbdStartBlock;
    }

    /**
     * @param array $smallBlockChain
     */
    public function setSmallBlockChain($smallBlockChain)
    {
        $this->smallBlockChain = $smallBlockChain;
    }

    /**
     * @return array
     */
    public function getSmallBlockChain()
    {
        return $this->smallBlockChain;
    }
}