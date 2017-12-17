<?php

require('DefinitiveExcel.class.php');

function cadenaAleatoria($numero)
{
    $caracter= "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    srand((double)microtime()*1000000);
    $rand = null;

    for($i=0; $i<$numero; $i++) 
    {
        $rand .= $caracter[rand()%strlen($caracter)];
    }

    return $rand;
}

$xls = new DefinitiveExcel();

$xls->setName('test')
->setAuthor('Luciano Vergara')
->setRange(array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB'));

//Start
$xls->startExcel();

//SetData
$count = 1;
$xls->setRowCount(1);

while($count <= 50000)
{
    $xls->insertRow(array(cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8), cadenaAleatoria(8),cadenaAleatoria(8),cadenaAleatoria(8)));
    echo $count."\n";
    $count++;
    $xls->setRowCount($count);
}

//End
$xls->endExcel();

//Save file
$xls->doXmlx('test.xlsx');

