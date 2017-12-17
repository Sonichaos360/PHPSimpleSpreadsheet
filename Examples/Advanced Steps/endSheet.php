<?php

require('../../DefinitiveExcel.class.php');

$xls = new DefinitiveExcel();

//Set The original sheet name
$xls->continueSheet("test");

//Close sheet data
$xls->endExcel();

/*
* Save file (ZIP TO XLSX) if there are so many regs ON your sheet and PHP can't zip the files 
* you can use any other program on your terminal, zip the files and rename the resultant file as NAME.xlsx
* That's all
*/
$xls->doXmlx('test.xlsx');
