<?php

require('../../src/PHPSimpleSpreadsheet.php');

$xls = new Sonichaos360\PHPSimpleSpreadsheet\PHPSimpleSpreadsheet();

$xls
//Sheet Name
->setName('test')
//Author Name
->setAuthor('Luciano Vergara')
//Set Columns range
->setRange(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']);

//Start Sheet
$xls->startSheet();

//SetData
$count = 1;

//Set header row, style bold
$xls->insertRow(['A DATA', 'B DATA', 'C DATA', 'D DATA', 'E DATA', 'F DATA', 'G DATA', 'H DATA', 'I DATA'], "bold");

//Increment row count
$count++;

//Add data using insertRow and pass range ordered array values
while ($count <= 10) {
    //Set row data
    $xls->insertRow(['A DATA', 'B DATA', 'C DATA', 'D DATA', 'E DATA', 'F DATA', 'G DATA', 'H DATA', 'I DATA']);

    //Show row number on console
    $count++;
}

//Close sheet data
$xls->endSheet();

/*
* Save file (ZIP TO XLSX) if there are so many regs ON your sheet and PHP can't zip the files
* you can use any other program on your terminal, zip the files and rename the resultant file as NAME.xlsx
* That's all
*/
if ($xls->doXmlx('test.xlsx')) {
    echo "File generated successfully. <a href=\"test.xlsx\">OPEN FILE<a>";
} else {
    throw new Exception('There was a problem generating the Spreadsheet.');
}
