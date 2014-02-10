<?php

require('php-excel-reader/excel_reader2.php');
require('SpreadsheetReader.php');

$spreadsheet = new SpreadsheetReader('test/test3.xlsx');

foreach ($spreadsheet as $Key => $Row)
{
    if ($Row)
    {
        print_r($Row);
    } else
    {
        var_dump($Row);
    }
}