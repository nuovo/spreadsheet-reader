<?php
/**
 * XLS parsing uses php-excel-reader from http://code.google.com/p/php-excel-reader/
 */
	if (isset($argv[1]))
	{
		$Filepath = $argv[1];
	}
	else
	{
		echo 'Please specify filename as the first argument'.PHP_EOL;
		exit;
	}

	// Excel reader from http://code.google.com/p/php-excel-reader/
	require('php-excel-reader/excel_reader2.php');
	require('SpreadsheetReader.php');

	date_default_timezone_set('UTC');

	$Spreadsheet = new SpreadsheetReader($Filepath);
	$BaseMem = memory_get_usage();

	$Time = microtime(true);

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

		echo ($CurrentMem - $BaseMem).' current, '.$CurrentMem.' base'.PHP_EOL;
		echo '---------------------------------'.PHP_EOL;

		if ($Key && ($Key % 500 == 0))
		{
			echo '---------------------------------'.PHP_EOL;
			echo 'Time: '.(microtime(true) - $Time);
			echo '---------------------------------'.PHP_EOL;
		}
	}

	echo PHP_EOL.'---------------------------------'.PHP_EOL;
	echo 'Time: '.(microtime(true) - $Time);
	echo PHP_EOL;
?>
