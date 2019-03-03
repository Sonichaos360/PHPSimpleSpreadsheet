#PowerfullPHPSpreadsheetGeneratorTESTS

## PHPSimpleSpreadsheet
Very simple and powerful tool for generate XLSX spreadsheets in PHP with low memory usage.

##Basic Spreadsheet

```php
require('../../src/PHPSimpleSpreadsheet.php');

$xls = new diblet\PHPSimpleSpreadsheet\PHPSimpleSpreadsheet();

$xls
//Sheet Name
->setName('test') 
//Author Name
->setAuthor('Luciano Joan Vergara') 
//Set Columns range
->setRange(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']);

//Start Sheet
$xls->startExcel();

//SetData
$count = 1;

//Add data using insertRow and pass range ordered array values
while($count <= 10)
{
    //Set row data
    $xls->insertRow(['A DATA', 'B DATA', 'C DATA', 'D DATA', 'E DATA', 'F DATA', 'G DATA', 'H DATA', 'I DATA']);

    $count++;
}

//Close sheet data
$xls->endExcel();

/*
* Save file (ZIP TO XLSX) if there are so many regs ON your sheet and PHP can't zip the files 
* you can use any other program on your terminal, zip the files and rename the resultant file as NAME.xlsx
* That's all
*/
if($xls->doXmlx('test.xlsx'))
{
    echo "File generated successfully. <a href=\"test.xlsx\">OPEN FILE<a>";
}
else
{
    throw new Exception('There was a problem generating the Spreadsheet.');
}
```

##Advanced Paginated Spreadsheet (Lots of rows)
```php
/**
 * Require Classes
 */
require('../../src/PHPSimpleSpreadsheet.php');

/**
 * Create object
 */
$xls = new diblet\PHPSimpleSpreadsheet\PHPSimpleSpreadsheet();

/**
 * Define number of elements per page
 */
$elements = 100;

/**
 * Defile total items to export
 */
$total_elements = 500;

/**
 * Get the pointer paremeter, 
 * just a counter to know the current page
 * If it is NULL then we know this is the first page
 */
$pointer = ( !isset($_GET["pointer"]) ? $elements : $_GET["pointer"]);

/**
 * If this is the first page
 * then we should start the spreadsheet
 * Else we just continue the current spreadsheet
 */
 if($pointer == $elements)
 {
    /**
     * Start class
     */
    $xls
    ->setName('test') 
    ->setAuthor('Luciano Vergara') 
    ->setRange(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'])
    ->startExcel();
 }
 else
 {
    /**
    * You shound use the sheet name as parameter to continue, 
    * in this case "test" is the name we defined before
    */
    $xls->continueSheet("test");
 }

//SetData
$count = 1;

//Add data using insertRow and pass range ordered array values
while($count <= $elements)
{
    /**
     * Here you can do your querys or something like that to
     * obtain the data using LIMIT clauses or similar indicating 
     * the $pointer variable as LIMIT and set data to the ROW
     */

    //Set row data
    $xls->insertRow(['A '.$pointer.' DATA', 'B '.$pointer.' DATA', 'C '.$pointer.' DATA', 'D '.$pointer.' DATA', 'E '.$pointer.' DATA', 'F '.$pointer.' DATA', 'G '.$pointer.' DATA', 'H '.$pointer.' DATA', 'I '.$pointer.' DATA']);

    //Show row number on console
    $count++;
}

/**
 * If we can not reach the pag limit then pause sheet
 * Else we should just end the sheet
 */
if($pointer < $total_elements)
{
    $xls->pauseSheet();

    ?>
    <strong><?php echo $pointer; ?> OF <?php echo $total_elements; ?> ELEMENTS PROCESSED...</strong>
    <script>
        /**
        * In JS just reload page and increase counter
        */
        setTimeout(function(){
            window.location.href = 'paginated.php?pointer=<?php echo ($pointer+$elements); ?>';
        }, 3000);
    </script>
    <?php
}
else
{
    /**
    * End Sheet
    */
    $xls->endExcel();

    /*
    * Save file (ZIP TO XLSX) if there are so many regs ON your sheet and PHP can't zip the files 
    * you can use any other program on your terminal, zip the files and rename the resultant file as NAME.xlsx
    * That's all
    */
    if($xls->doXmlx('test.xlsx'))
    {
        echo "File generated successfully. <a href=\"test.xlsx\">OPEN FILE<a>";
    }
    else
    {
        throw new Exception('There was a problem generating the Spreadsheet.');
    }
}
```


## Licence
The Software is provided "as is", without warranty of any kind. Please refer to [the license file](LICENSE).
