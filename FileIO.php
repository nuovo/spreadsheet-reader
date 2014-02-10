<?php


/**
 * Class FileIO
 */
class FileIO {
    /**
     * @var null
     */
    protected $handle = null;

    /**
     * @var null
     */
    protected $fileName = null;

    /**
     * @var null
     */
    protected $mode = null;

    /**
     * @var null
     */
    protected $tmpLocation = null;

    /**
     * @param string $mode
     * @param null $fileName
     */
    function __construct($mode = 'a+b', $fileName = null)
    {
        $this->setMode($mode);
        $this->setTmpLocation(sys_get_temp_dir());
        $this->setFileName($fileName);
        $this->setMode($mode);
    }

    /**
     * @param $line
     * @throws Exception
     */
    public function appendLine($line)
    {
        if (!$this->getHandle()) {
            throw new Exception('Invalid file handle');
        }

        if (!is_writable($this->getFullFilePath())) {
            throw new Exception('File is not writable');
        }

        fwrite($this->getHandle(), $line . PHP_EOL);
    }

    /**
     * @param $position
     * @return bool|string
     */
    public function getLine($position)
    {
        $file = new SplFileObject($this->getFullFilePath());
        $file->seek($position);

        echo $file->current() . " - " . $position . '<br>';
        return $file->current();

//        $currentPosition = 0;
//
//        //Making sure that handler position is at the beginning of the file
//        if (ftell($this->getHandle()) > 0) {
//            rewind($this->getHandle());
//        }
//
//        while($content = stream_get_line($this->getHandle(), 4096, "\n")) {
//            echo "$position : $currentPosition == $position <br>";
//            if ($currentPosition == $position) {
//                return $content;
//            }
//
//            $currentPosition++;
//        }
//
//        return false;
    }

    /**
     * @param null $fileName
     */
    public function setFileName($fileName)
    {
        if (!$fileName) {
            $fileName = 'sp_' . md5(time());
        }

        $this->fileName = $fileName;
    }

    /**
     * @return null
     */
    public function getFileName()
    {
        return $this->fileName;
    }

    /**
     * @param null $handle
     */
    public function setHandle($handle)
    {
        $this->handle = $handle;
    }

    /**
     * @return null
     * @throws Exception
     */
    public function getHandle()
    {
        if ($this->handle === null) {
            $fullPath = $this->getFullFilePath();

            $handle = fopen($fullPath, $this->getMode());

            if (!$handle) {
                throw new Exception('Invalid handle');
            }

            $this->setHandle($handle);
        }

        return $this->handle;
    }

    /**
     * @param null $mode
     */
    public function setMode($mode)
    {
        $this->mode = $mode;
    }

    /**
     * @return null
     */
    public function getMode()
    {
        return $this->mode;
    }

    /**
     * @param $tmpLocation
     * @throws Exception
     */
    public function setTmpLocation($tmpLocation)
    {
        if (!$tmpLocation) {
            throw new Exception('Temp directory missing');
        }

        //In some of the systems function sys_get_temp_dir() does not adds trailing slash
        //to be sure
        if (substr($tmpLocation, -1) !== DIRECTORY_SEPARATOR) {
            $tmpLocation .= DIRECTORY_SEPARATOR;
        }

        $this->tmpLocation = $tmpLocation;
    }

    /**
     * @return null
     */
    public function getTmpLocation()
    {
        return $this->tmpLocation;
    }

    /**
     * @return string
     */
    public function getFullFilePath()
    {
        return $this->getTmpLocation() . $this->getFileName();
    }
}