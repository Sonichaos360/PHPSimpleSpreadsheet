<?php

Class DefinitiveExcel
{
    public $name;
    public $author;
    public $totalCount;
    public $range;
    public $rowCount;
    
    public function setName($name = '')
    {
        $this->name = $name;
        return $this;
    }

    public function setAuthor($author = '')
    {
        $this->author = $author;
        return $this;
    }

    public function setTotalCount($totalCount = '')
    {
        $this->totalCount = $totalCount;
        return $this;
    }

    public function setRange($range = array())
    {
        $this->range = $range;
        return $this;
    }

    public function setRowCount($rowCount = array())
    {
        $this->rowCount = $rowCount;
        return $this;
    }

    function doXmlx($destination)
    {
        $source = $this->name;
        if (!extension_loaded('zip') || !file_exists($source)) {
            return false;
        }
    
        $zip = new ZipArchive();
        if (!$zip->open($destination, ZIPARCHIVE::CREATE)) {
            return false;
        }
    
        $source = str_replace('\\', '/', realpath($source));
    
        if (is_dir($source) === true)
        {
            $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($source), RecursiveIteratorIterator::SELF_FIRST);

            foreach ($files as $file)
            {
                $file = str_replace('\\', '/', $file);
                // $file = explode($this->name."", $file)[1];
                
                // Ignore "." and ".." folders
                if( in_array(substr($file, strrpos($file, '/')+1), array('.', '..')) )
                    continue;
    
                $file = realpath($file);
    
                if (is_dir($file) === true)
                {
                    $zip->addEmptyDir(explode("\\".$this->name."\\", str_replace($source . '/', '', $file . '/'))[1]);
                }
                else if (is_file($file) === true)
                {
                    $zip->addFromString(explode("\\".$this->name."\\", str_replace($source . '/', '', $file))[1], file_get_contents($file));
                }
            }
        }
        else if (is_file($source) === true)
        {
            $zip->addFromString(basename($source), file_get_contents($source));
        }
        
        $this->cleanTemp();
        return $zip->close();
    }

    public function cleanTemp()
    {
        //Clean
        if(file_exists($this->name.'/_rels/.rels'))
        {
            unlink($this->name.'/_rels/.rels');
        }
        $this->removeDirectory($this->name);
    }

    public function removeDirectory($path) 
    {
        $files = glob($path . '/*');
        
        foreach ($files as $file) 
        {
            is_dir($file) ? $this->removeDirectory($file) : unlink($file);
        }

        if(is_dir($path))
        {
            rmdir($path);
        }
        return;
    }

    public function startExcel()
    {
        //Clean
        $this->cleanTemp();
        //Create Dirs
        mkdir($this->name.'');
        mkdir($this->name.'/_rels');
        mkdir($this->name.'/docProps');
        mkdir($this->name.'/xl');
        mkdir($this->name.'/xl/_rels');
        mkdir($this->name.'/xl/worksheets');

        //Create Temp Files
        file_put_contents($this->name.'/_rels/.rels', 
        '<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
        ', FILE_APPEND | LOCK_EX);

        file_put_contents($this->name.'/docProps/app.xml', 
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime></Properties>
        ', FILE_APPEND | LOCK_EX);

        file_put_contents($this->name.'/docProps/core.xml', 
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">'.date("Y-m-d", time()).'T'.date("H:i:s", time()).'.00Z</dcterms:created><dc:title>Doc Title</dc:title><dc:creator>Doc Author</dc:creator><cp:revision>0</cp:revision></cp:coreProperties>
        ', FILE_APPEND | LOCK_EX);

        file_put_contents($this->name.'/[Content_Types].xml', 
        '<?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
        </Types>
        ', FILE_APPEND | LOCK_EX);

        file_put_contents($this->name.'/xl/styles.xml', 
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="GENERAL" /></numFmts><fonts count="4"><font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font><font><name val="Arial"/><family val="0"/><sz val="10"/></font><font><name val="Arial"/><family val="0"/><sz val="10"/></font><font><name val="Arial"/><family val="0"/><sz val="10"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="20"><xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164"><alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/><protection hidden="false" locked="true"/></xf><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/></cellStyleXfs><cellXfs count="1"><xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0">	<alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/>	<protection locked="true" hidden="false"/></xf></cellXfs><cellStyles count="6"><cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/><cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/><cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/><cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/><cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/><cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/></cellStyles></styleSheet>
        ', FILE_APPEND | LOCK_EX);

        file_put_contents($this->name.'/xl/_rels/workbook.xml.rels', 
        '<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
        ', FILE_APPEND | LOCK_EX);

        file_put_contents($this->name.'/xl/workbook.xml', 
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/><bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" state="visible" r:id="rId2"/></sheets><calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>
        ', FILE_APPEND | LOCK_EX);

        file_put_contents($this->name.'/xl/worksheets/sheet1.xml', 
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetPr filterMode="false">
                <pageSetUpPr fitToPage="false"/>
            </sheetPr>
            <dimension ref="A1:AJ10"/>
            <sheetViews>
                <sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="true" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">
                    <selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>
                </sheetView>
            </sheetViews>
            <cols>
                <col collapsed="false" hidden="false" max="1024" min="1" style="0" width="11.5"/>
            </cols>
            <sheetData>', FILE_APPEND | LOCK_EX);

    }

    public function endExcel()
    {
        file_put_contents($this->name.'/xl/worksheets/sheet1.xml', 
        '</sheetData>
        <printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>
        <pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>
        <pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>
        <headerFooter differentFirst="false" differentOddEven="false">
            <oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>
            <oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>
        </headerFooter>
        </worksheet>', FILE_APPEND | LOCK_EX);
    }

    public function insertRow($row)
    {
        $finalRow = '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.$this->rowCount.'">';

        $i = 0;
        foreach($this->range AS $item)
        {
            $finalRow .= '<c r="'.$item.$this->rowCount.'" s="0" t="inlineStr"><is><t>'.$row[$i].'</t></is></c>';
            $i++;
        }        

        $finalRow .= '</row>';

        file_put_contents($this->name.'/xl/worksheets/sheet1.xml', $finalRow, FILE_APPEND | LOCK_EX);
    }
}