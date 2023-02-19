<?php
/**
 * This file is part of Sonichaos360/PHPSimpleSpreadsheet
 *
 * Sonichaos360/PHPSimpleSpreadsheet is open source software: you can
 * distribute it and/or modify it under the terms of the MIT License
 * (the "License"). You may not use this file except in
 * compliance with the License.
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
 * implied. See the License for the specific language governing
 * permissions and limitations under the License.
 *
 * @copyright Copyright (c) Luciano Vergara <contacto@lucianovergara.com>
 * @license https://opensource.org/licenses/MIT MIT License
 */

 require('../../src/PHPSimpleSpreadsheet.php');

 use Sonichaos360\PHPSimpleSpreadsheet\PHPSimpleSpreadsheet;
 
 // create a new instance of the class
 $sheet = new PHPSimpleSpreadsheet();
 
 // set sheet name and author
 $sheet->setName('Product Prices');
 $sheet->setAuthor('Your Name');
 
 // set column range and column width
 $sheet->setRange(['A', 'B']);
 $sheet->setColumnsWidth([25, 15]);
 
 // start the sheet creation
 $sheet->startSheet();
 
 // insert headers
 $headers = ['Product', 'Price'];
 $sheet->insertRow($headers, 'bold', []);
 
 // insert data
 $data = [
     ['Product A', 10, 'number'],
     ['Product B', 15, 'number'],
     ['Product C', 20, 'number'],
     ['Product D', 25, 'number'],
 ];
 
 foreach ($data as $row) {
     $row[1] = (float)$row[1];
     $sheet->insertRow($row, '', [$row[2]]);
 }
 
 // insert total formula
 $formula = '=SUM(B2:B'.(count($data) + 1).')';
 $totalRow = ['Total', $formula];
 $sheet->insertRow($totalRow, 'bold', ['formula']);
 
 // end the sheet creation
 $sheet->endSheet();
 
 // create the XLSX file
 if ($sheet->doXmlx('product_prices.xlsx')) {
     echo 'File created successfully!  <a href="product_prices.xlsx">OPEN FILE<a>';
 } else {
     echo 'Error creating file.';
 }
 