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

namespace Sonichaos360\PHPSimpleSpreadsheet;

class PHPSimpleSpreadsheet
{
    public $name;
    public $author;
    public $range;
    public $columnsWidth;
    public $rowCount;
    public $tempDir;
    public $defStyles;
    public $classPath;

    public function __construct()
    {
        $this->tempDir = sys_get_temp_dir().DIRECTORY_SEPARATOR;
        $this->classPath = dirname(__FILE__);
        $this->columnsWidth = "";
        $this->defstyles = array(
            "normal" => 0,
            "bold" => 1,
            "italic" => 2
        );
    }

    /**
     * Set sheet name prop
     *
     * @return this Current object
     */
    public function setName($name = '')
    {
        $this->name = $name;
        return $this;
    }

    /**
     * Set sheet author prop
     *
     * @return this Current object
     */
    public function setAuthor($author = '')
    {
        $this->author = $author;
        return $this;
    }

    /**
     * Set sheet columns range
     *
     * @return this Current object
     */
    public function setRange($range = array())
    {
        $this->range = $range;
        return $this;
    }

    /**
     * Collect created temp files and zips into unique XLSX file
     *
     * @return bool ZIP process result
     */
    public function doXmlx($destination)
    {
        $source = $this->tempDir.$this->name;

        if (!extension_loaded('zip') || !file_exists($source)) {
            throw new \Exception('Can not create ZIP file. Please enable ZIP PHP Extension.');
        }
    
        $zip = new \ZipArchive();

        if (!$zip->open($destination, \ZIPARCHIVE::CREATE)) {
            return false;
        }

        if (strtolower(PHP_OS) == 'winnt') {
            $source = str_replace('\\', '/', realpath($source));
        }
    
        if (is_dir($source) === true) {
            $files = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($source), \RecursiveIteratorIterator::SELF_FIRST);

            foreach ($files as $file) {
                if (strtolower(PHP_OS) == 'winnt') {
                    $file = str_replace('\\', '/', $file);
                }
                
                if (in_array(substr($file, strrpos($file, '/')+1), array('.', '..'))) {
                    continue;
                }
                   
                $file = realpath($file);
    
                if (is_dir($file) === true) {
                    if (strtolower(PHP_OS) == 'winnt') {
                        $zip->addEmptyDir(explode("\\".$this->name."\\", str_replace($source . '/', '', $file . '/'))[1]);
                    } else {
                        $zip->addEmptyDir(explode("/".$this->name."/", $file)[1]);
                    }
                } elseif (is_file($file) === true) {
                    if (strtolower(PHP_OS) == 'winnt') {
                        $zip->addFromString(explode("\\".$this->name."\\", str_replace($source . '/', '', $file))[1], file_get_contents($file));
                    } else {
                        $zip->addFromString(explode("/".$this->name."/", $file)[1], file_get_contents($file));
                    }
                }
            }
        } elseif (is_file($source) === true) {
            $zip->addFromString(basename($source), file_get_contents($source));
        }
        
        $this->cleanTemp();

        return $zip->close();
    }

    /**
     * Clear all past temp data it is necessary in case there is an old sheet using the same file name
     *
     * @return bool Delete process result
     */
    public function cleanTemp()
    {
        //Clean
        if (file_exists($this->tempDir.$this->name.'/_rels/.rels')) {
            unlink($this->tempDir.$this->name.'/_rels/.rels');
        }

        if (file_exists($this->tempDir.$this->name.'.json')) {
            unlink($this->tempDir.$this->name.'.json');
        }

        $this->removeDirectory($this->tempDir.$this->name);
    }

    /**
     * Delete main XML temp directory and files
     *
     * @return bool Delete process result
     */
    public function removeDirectory($path)
    {
        $files = glob($path . '/*');
        
        foreach ($files as $file) {
            is_dir($file) ? $this->removeDirectory($file) : unlink($file);
        }

        if (is_dir($path)) {
            return rmdir($path);
        }

        return false;
    }
    
    /**
     * Create base xml files in order to start sheet creation
     *
     * @return void
     */
    public function startSheet()
    {
        //Delete old generated file
        if (file_exists($this->tempDir.$this->name.".xlsx")) {
            unlink($this->tempDir.$this->name.".xlsx");
        }

        //Set First Row
        $this->rowCount = 1;

        //Clean
        $this->cleanTemp();

        //Create Dirs
        mkdir($this->tempDir.$this->name.'');
        mkdir($this->tempDir.$this->name.'/_rels');
        mkdir($this->tempDir.$this->name.'/docProps');
        mkdir($this->tempDir.$this->name.'/xl');
        mkdir($this->tempDir.$this->name.'/xl/_rels');
        mkdir($this->tempDir.$this->name.'/xl/worksheets');

        //Create Temp Files
        file_put_contents($this->tempDir.$this->name.'/_rels/.rels', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."rels.xml"), FILE_APPEND | LOCK_EX);
        file_put_contents($this->tempDir.$this->name.'/docProps/app.xml', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."app.xml"), FILE_APPEND | LOCK_EX);
        file_put_contents($this->tempDir.$this->name.'/docProps/core.xml', str_replace("[[DATE]]", date("Y-m-d", time()).'T'.date("H:i:s", time()).'.00Z', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."core.xml")), FILE_APPEND | LOCK_EX);
        file_put_contents($this->tempDir.$this->name.'/[Content_Types].xml', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."types.xml"), FILE_APPEND | LOCK_EX);
        file_put_contents($this->tempDir.$this->name.'/xl/styles.xml', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."styles.xml"), FILE_APPEND | LOCK_EX);
        file_put_contents($this->tempDir.$this->name.'/xl/_rels/workbook.xml.rels', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."workbook_rels.xml"), FILE_APPEND | LOCK_EX);
        file_put_contents($this->tempDir.$this->name.'/xl/workbook.xml', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."workbook.xml"), FILE_APPEND | LOCK_EX);
        file_put_contents($this->tempDir.$this->name.'/xl/worksheets/sheet1.xml', str_replace("[[COLUMNS_WIDTH]]", $this->columnsWidth, file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."sheet_start.xml")), FILE_APPEND | LOCK_EX);
    }

    /**
     * Close XML tags in order to finish sheet creation
     *
     * @return void
     */
    public function endSheet()
    {
        file_put_contents($this->tempDir.$this->name.'/xl/worksheets/sheet1.xml', file_get_contents($this->classPath.DIRECTORY_SEPARATOR."xml".DIRECTORY_SEPARATOR."sheet_end.xml"), FILE_APPEND | LOCK_EX);
    }

    /**
     * Insert new row in workbook and increase row count
     *
     * @return void
     */
    public function insertRow($row, $styles = "")
    {
        $finalRow = '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.$this->rowCount.'">';

        $i = 0;
        foreach ($this->range as $item) {
            $finalRow .= '<c r="'.$item.$this->rowCount.'" s="'.$this->defstyles[ ($styles != "") ? $styles : "normal" ].'" t="inlineStr"><is><t>'.$this->clean($row[$i]).'</t></is></c>';
            $i++;
        }

        $finalRow .= '</row>';

        file_put_contents($this->tempDir.$this->name.'/xl/worksheets/sheet1.xml', $finalRow, FILE_APPEND | LOCK_EX);

        $this->rowCount++;
    }

    /**
     * Creates a temp file including data related to the current sheet creation, this is necessary in case of big multiple parts sheet creation
     *
     * @return bool Process result
     */
    public function pauseSheet()
    {
        return file_put_contents($this->tempDir.$this->name.'.json', json_encode(['range' => $this->range, 'rowcount' => $this->rowCount]), LOCK_EX);
    }

    /**
     * Load current temp paused file and continue sheet creation
     *
     * @return bool Process result
     */
    public function continueSheet($sheetname)
    {
        if (file_exists($this->tempDir.$sheetname.'.json')) {
            $data = json_decode(file_get_contents($this->tempDir.$sheetname.'.json'), true);
            $this->range = $data["range"];
            $this->rowCount = $data["rowcount"];
            $this->name = $sheetname;
            return true;
        } else {
            throw new \Exception('Sheet config file missing..');
        }
    }

    /**
     * Parse special chars encoding in row content
     *
     * @return string Parsed string
     */
    public function clean($str)
    {
        return str_replace(
            array("&","<",">","'",'"'),
            array("&amp;","&lt;", "&gt;","&apos;",'&quot;'),
            $str
        );
    }

    /**
     * Set sheet columns width
     *
     * @return void
     */
    public function setColumnsWidth($cols)
    {
        $i = 1;
        $this->columnsWidth .= '<cols>';
        foreach($cols AS $item) {
            $this->columnsWidth .= '<col min="'.$i.'" max="'.$i.'" width="'.$item.'" customWidth="1"/>';
            $i++;
        }
        $this->columnsWidth .= '</cols>';
    }
}
