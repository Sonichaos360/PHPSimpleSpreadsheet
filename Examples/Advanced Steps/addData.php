<?php

require('../../DefinitiveExcel.class.php');

$xls = new DefinitiveExcel();

//Set The original sheet name
$xls->continueSheet("test");

//SetData
$count = 1;

//Add data using insertRow and pass range ordered array values
while($count <= 100)
{
    //Set row data
    $xls->insertRow(['A DATA', 'B DATA', 'C DATA', 'D DATA', 'E DATA', 'F DATA', 'G DATA', 'H DATA', 'I DATA']);

    //Show row number on console
    echo $count."\n";
    $count++;
}

//Pause Sheet
$xls->pauseSheet();