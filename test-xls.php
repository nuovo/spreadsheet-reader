<?php

header('Content-Type: text/html');
ini_set('display_errors',1);
ini_set('display_startup_errors',1);
error_reporting(-1);

require 'xls/XlsReader.php';


$reader = new XlsReader('test/airports.xls', false, 'UTF-8');

