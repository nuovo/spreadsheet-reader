<?php

//namespace SpreadsheetReader;

header('Content-Type: text/html');
ini_set('display_errors',1);
ini_set('display_startup_errors',1);
error_reporting(-1);

class XlsParser_filter extends \php_user_filter
{
    function filter ($in, $out, &$consumed, $closing)
    {
        while ($bucket = stream_bucket_make_writeable($in)) {
            $bucket->data = strtoupper($bucket->data);
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }

        return PSFS_PASS_ON;
    }
}

stream_filter_register("xlsparser", "XlsParser_filter")
or die("Failed to register filter");

$filename = __DIR__ . DIRECTORY_SEPARATOR . 'test.txt';
echo $filename;

if (!is_writable($filename)) {
    throw new Exception('File is not writable');
}

$writeHandle = fopen($filename, 'w+');
var_dump($writeHandle);

$readHandle = fopen('test/airports.xls', 'rb');
var_dump($readHandle);

fwrite($writeHandle, 'aaaaa');

stream_filter_append($writeHandle, "xlsparser");

$contents = '';
while (!feof($readHandle)) {
    $contents .= fread($readHandle, 81920);

    fwrite($writeHandle, $contents);
}

fclose($writeHandle);
fclose($readHandle);

