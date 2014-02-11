<?php

namespace SpreadsheetReader;

header('Content-Type: text/plain');
ini_set('display_errors',1);
ini_set('display_startup_errors',1);
error_reporting(-1);

// Excel reader from http://code.google.com/p/php-excel-reader/
//require('php-excel-reader/excel_reader2.php');
require('SpreadsheetReader.php');

if (!function_exists('humanFileSize')) {
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
}

/**
 * XLS parsing uses php-excel-reader from http://code.google.com/p/php-excel-reader/
 */

	if (isset($argv[1]))
	{
		$Filepath = $argv[1];
	}
	elseif (isset($_GET['File']))
	{
		$Filepath = $_GET['File'];
	}
	else
	{
		if (php_sapi_name() == 'cli')
		{
			echo 'Please specify filename as the first argument'.PHP_EOL;
		}
		else
		{
			echo 'Please specify filename as a HTTP GET parameter "File", e.g., "/test.php?File=test.xlsx"';
		}
		exit;
	}



	date_default_timezone_set('UTC');

	$StartMem = memory_get_usage();
	echo '---------------------------------'.PHP_EOL;
	echo 'Starting memory: ' . humanFileSize($StartMem) . PHP_EOL;
	echo '---------------------------------'.PHP_EOL;

	try
	{
		$Spreadsheet = new SpreadsheetReader($Filepath);
		$BaseMem = memory_get_usage();

		$Sheets = $Spreadsheet -> Sheets();

		echo '---------------------------------'.PHP_EOL;
		echo 'Spreadsheets:'.PHP_EOL;
		print_r($Sheets);
		echo '---------------------------------'.PHP_EOL;
		echo '---------------------------------'.PHP_EOL;

		foreach ($Sheets as $Index => $Name)
		{
			echo '---------------------------------'.PHP_EOL;
			echo '*** Sheet '.$Name.' ***'.PHP_EOL;
			echo '---------------------------------'.PHP_EOL;

			$Time = microtime(true);

			$Spreadsheet -> ChangeSheet($Index);
            $counter = 0;
			foreach ($Spreadsheet as $Key => $Row)
			{
				echo $Key.': ';
				if ($Row)
				{
					print_r($Row);
				}
				else
				{
					var_dump($Row);
				}
				$CurrentMem = memory_get_usage();
		
				echo 'Memory: '. humanFileSize($CurrentMem - $BaseMem).' current, '. humanFileSize($CurrentMem).' base'.PHP_EOL;
				echo '---------------------------------'.PHP_EOL;
		
				//if ($Key && ($Key % 500 == 0))
				//{
					echo '---------------------------------'.PHP_EOL;
					echo 'Time: '.(microtime(true) - $Time) .PHP_EOL;
					echo '---------------------------------'.PHP_EOL;
				//}

//                if ($counter == 1) {
//                    die();
//                }

                $counter++;
			}
		
			echo PHP_EOL.'---------------------------------'.PHP_EOL;
			echo 'Time: '.(microtime(true) - $Time);
			echo PHP_EOL;

			echo '---------------------------------'.PHP_EOL;
			echo '*** End of sheet '.$Name.' ***'.PHP_EOL;
			echo '---------------------------------'.PHP_EOL;
		}
		
	}
	catch (Exception $E)
	{
		echo $E -> getMessage();
	}



