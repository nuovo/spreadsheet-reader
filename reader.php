<?php

header('Content-Type: text/plain');

set_time_limit(90);

use ReaderExcel\SpreadsheetReader;

include_once __DIR__ . '/vendor/autoload.php';

$read = false;

if (isset($argv[1])) {

    $Filepath = __DIR__ . '/' . $argv[1];

    $read = true;

} else {

    echo 'Please enter name file excel';

}

if ($read) {

    date_default_timezone_set('UTC');

    $StartMem = memory_get_usage();
    echo '---------------------------------' . PHP_EOL;
    echo 'Starting memory: ' . $StartMem . PHP_EOL;
    echo '---------------------------------' . PHP_EOL;

    try
    {
        $Spreadsheet = new SpreadsheetReader($Filepath);
        $BaseMem     = memory_get_usage();

        echo $BaseMem . "\n";

        $Sheets = $Spreadsheet->Sheets();

        echo '---------------------------------' . PHP_EOL;
        echo 'Spreadsheets:' . PHP_EOL;
        print_r($Sheets);
        echo '---------------------------------' . PHP_EOL;
        echo '---------------------------------' . PHP_EOL;

        foreach ($Sheets as $Index => $Name) {
            echo '---------------------------------' . PHP_EOL;
            echo '*** Sheet ' . $Name . ' ***' . PHP_EOL;
            echo '---------------------------------' . PHP_EOL;

            $Time = microtime(true);

            $Spreadsheet->ChangeSheet($Index);

            foreach ($Spreadsheet as $Key => $Row) {
                echo $Key . ': ';
                if ($Row) {
                    print_r($Row);
                } else {
                    var_dump($Row);
                }
                $CurrentMem = memory_get_usage();

                echo 'Memory: ' . ($CurrentMem - $BaseMem) . ' current, ' . $CurrentMem . ' base' . PHP_EOL;
                echo '---------------------------------' . PHP_EOL;

                if ($Key && ($Key % 500 == 0)) {
                    echo '---------------------------------' . PHP_EOL;
                    echo 'Time: ' . (microtime(true) - $Time);
                    echo '---------------------------------' . PHP_EOL;
                }
            }

            echo PHP_EOL . '---------------------------------' . PHP_EOL;
            echo 'Time: ' . (microtime(true) - $Time);
            echo PHP_EOL;

            echo '---------------------------------' . PHP_EOL;
            echo '*** End of sheet ' . $Name . ' ***' . PHP_EOL;
            echo '---------------------------------' . PHP_EOL;
        }

    } catch (Exception $E) {
        echo $E->getMessage();
    }

}
