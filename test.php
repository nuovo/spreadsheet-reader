<?php
/**
 * XLS parsing uses php-excel-reader from http://code.google.com/p/php-excel-reader/
 */
	header('Content-Type: text/plain');

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

	// Excel reader from http://code.google.com/p/php-excel-reader/
	require('php-excel-reader/excel_reader2.php');
	require('SpreadsheetReader.php');

	date_default_timezone_set('UTC');

	$StartMem = memory_get_usage();
	echo '---------------------------------'.PHP_EOL;
	echo 'Starting memory: '.$StartMem.PHP_EOL;
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
		
				echo 'Memory: '.($CurrentMem - $BaseMem).' current, '.$CurrentMem.' base'.PHP_EOL;
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

			echo '---------------------------------'.PHP_EOL;
			echo '*** End of sheet '.$Name.' ***'.PHP_EOL;
			echo '---------------------------------'.PHP_EOL;
		}
		
	}
	catch (Exception $E)
	{
		echo $E -> getMessage();
	}
?>
