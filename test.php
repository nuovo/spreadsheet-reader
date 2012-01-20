<?php
/**
 * XLS parsing uses php-excel-reader from http://code.google.com/p/php-excel-reader/
 * The latest version, 2.21, didn't work for me but 2.2 was just fine
 */
	if (isset($argv[1]))
	{
		$Filepath = $argv[1];
	}
	else
	{
		echo 'Please specify filename as the first argument'."\n";
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
		var_dump($Row);
		$CurrentMem = memory_get_usage();

		echo ($CurrentMem - $BaseMem).' current, '.$CurrentMem." base\n";
		echo '---------------------------------'."\n";

		if ($Key && ($Key % 500 == 0))
		{
			echo '---------------------------------'."\n";
			echo 'Time: '.(microtime(true) - $Time);
			echo '---------------------------------'."\n";
		}
	}

	echo "\n".'---------------------------------'."\n";
	echo 'Time: '.(microtime(true) - $Time);
?>
