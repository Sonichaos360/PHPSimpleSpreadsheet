<?php

require('../../DefinitiveExcel.class.php');

$xls = new DefinitiveExcel();

$xls
//Sheet Name
->setName('test') 
//Author Name
->setAuthor('Luciano Vergara') 
//Set Columns range
->setRange(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']);

//Start Sheet
$xls->startExcel();

//Instant Pause Excel
$xls->pauseSheet("test");