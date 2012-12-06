<?php

/**
 * Class for read XLSX Data to array
 */

Class XLSXReader {
    
    protected $filename;
    protected $extractdir;
    protected $xlsxObject;
    protected $sharedStrings;
    protected $maxFileSize;
    protected $sheetRelationsSheet;
    
    public function __construct($file_path, $extractionPath) {
        $this->xlsxObject = new stdClass();
        $this->filename = $file_path;
        $this->extractdir = $extractionPath;
        
        try{
            $this->mapXlsxFile();
            //$this->echoInFriendlyFormat($this->xlsxObject);
        }catch (Exception $e){
            throw new Exception($e->getMessage());
        }//end try
        
    }//end __construct()
    
    public function setMaxFileSize($fileSize){
        return true;
    }//end setMaxFileSize()
    
    public function getSheetNumber () {
        if (isset($this->xlsxObject->sheetCount) === true){
            return $this->xlsxObject->sheetCount;
        }else{ 
            return 0;
        }
    }//end getSheetNumber()
    
    
    public function read(){
        //$this->echoInFriendlyFormat($this->xlsxObject);
        $xmlDoc = new DOMDocument();
        
        if ($this->sharedStringsCheck() === true) {

          try{
                $this->loadSharedStrings();
                
                for($c = 0; $c < $this->xlsxObject->sheetCount; $c++) {
                    $sheet = 'Sheet'.$c;
                    
                    if($this->xlsxObject->$sheet->hasData === true) { //van adat
                        $emptyTable = $this->createBaseArrayForSheet($this->xlsxObject->$sheet);
                        
                         if (file_exists($this->getExtractDir().'/xl/'.$this->xlsxObject->$sheet->file) === true) {
                             if($xmlDoc->load($this->getExtractDir().'/xl/'.$this->xlsxObject->$sheet->file)){
                                 $SheetObject = simplexml_import_dom($xmlDoc);
                                 //$this->echoInFriendlyFormat($SheetObject);
                                 $rowCount = 0;
                                 foreach ($SheetObject->sheetData->row as $row) {
                                     $rowAttributes = $row->attributes();
                                     $rowNumber = $rowAttributes->r;
                                        
                                        foreach($row->c as $column){
                                            $columnAttributes = $column->attributes();
                                            
                                            $cellId = $columnAttributes['r'];
                                            $columnLetter = substr($cellId,0,1);
                                            if($columnAttributes['t'] == "n") {
                                                $value = $column->v;
                                            }else if($columnAttributes['t'] == "s") {
                                                $value = $this->sharedStrings->data[(int)$column->v];
                                            }//endif
                                            $a = "A";
                                            $emptyTable[(int)$rowNumber][$columnLetter] = $value;
                                        }//end foreach
                                     $rowCount++;
                                 }//end foreach
                                 
                                 $numericIndexTable = array();
                                 $sorcounter = $oszlopcounter = 0;
                                 foreach($emptyTable as $egysor) {
                                     $oszlopcounter = 0;
                                     foreach($egysor as $egycella){
                                         $numericIndexTable[$sorcounter][$oszlopcounter] = $egycella;
                                         $oszlopcounter++;
                                     }
                                     $sorcounter++;
                                 }
                                 //$this->echoInFriendlyFormat($numericIndexTable);
                                 $this->xlsxObject->$sheet->data->rowCount = $rowCount;
                                 $this->xlsxObject->$sheet->data->realTable = $emptyTable;
                                 $this->xlsxObject->$sheet->data->numericIndexTable = $numericIndexTable;
                             }else {
                                 throw new Exception('Unable to load '.$this->xlsxObject->$sheet->file);
                             }
                         }else{
                             throw new Exception($this->xlsxObject->$sheet->file.' not found');
                         }//end if
                        
                    }else {
                        $this->xlsxObject->$sheet->data->rowCount = 0;
                        $this->xlsxObject->$sheet->data->realTable = array();
                        $this->xlsxObject->$sheet->data->numericIndexTable = array();
                    }//end if
                    
                }//end for

          }catch (Exception $e) {
                throw new Exception($e->getMessage());
          }//end try
        
        }//end if          
        
        //$this->echoInFriendlyFormat($this->xlsxObject);
        
        return $this->xlsxObject;
        
    }//end read
    
    /**
     * Create an empty array for sheet
     * @todo csak akkor működik, ha 
     * @param type $sheet
     * @return type 
     */
    protected function createBaseArrayForSheet($sheet){
        $dimension = $sheet->dimension;    
        //$this->echoInFriendlyFormat($dimension);
        $dimensionDarabok = explode(':', $dimension);
        $emptyTable = array();

        if (count($dimensionDarabok) == 2 ) {
            $startColumn = $dimensionDarabok[0][0];
            $endCloumn = $dimensionDarabok[1][0];
            $startRow = substr($dimensionDarabok[0], 1);
            $endRow = substr($dimensionDarabok[1], 1);
            
            for ($sor = (int)$startRow; $sor <= (int)$endRow; $sor++) {
                for($oszlop = $startColumn; $oszlop <= $endCloumn; $oszlop++){
                    $emptyTable[$sor][$oszlop] = null;
                }
            }
            //$this->echoInFriendlyFormat($emptyTable);
        
            return $emptyTable;
            
        }else {
            return array();
        }//end if
        
    }//end createBaseArrayForSheet()
    
    
    /**
     * feltérképezi, hogy hány lap van, milyen filenévvel, hol stb
     */
    protected function mapXlsxFile(){
        
       // $xlsxObject = new stdClass();
        
        try {
            $this->ExtractZipArchive();
            $xmlDoc = new DOMDocument();
            $this->getSheetRelationsSheet();
            
            if (is_file($this->getExtractDir().'/xl/workbook.xml') === true) {
                $xmlDoc->load($this->getExtractDir().'/xl/workbook.xml');
                $workBook = simplexml_import_dom($xmlDoc);
                //echo $workBook->workbookView->activeTab;
                $workBookAttributes = $workBook->bookViews->workbookView->attributes();
                $this->xlsxObject->activeTab = $workBookAttributes['activeTab'];
                //echo $xlsxObject->activeTab;
                $sheetCount = 0;
                
                foreach($workBook->sheets->sheet as $sheet) {
                    $sheetClass = 'Sheet'.$sheetCount;
                    $this->xlsxObject->$sheetClass = new stdClass();
                    $sheetAttributes = $sheet->attributes();
                    $this->xlsxObject->$sheetClass->name = $sheetAttributes['name'];
                    $this->xlsxObject->$sheetClass->id = $sheetAttributes["sheetId"];
                    $sheetCount++;
                    
                }//end foreach()
                //$this->echoInFriendlyFormat($this->xlsxObject);
                $this->xlsxObject->sheetCount = $sheetCount;
                
                //beolvassuk a sheet-ek fileneveit és a típusokat
                for($c = 0 ; $c < $this->xlsxObject->sheetCount ; $c++){
                            $sheet = 'Sheet'.$c;
                            $sheetRelAttributes = $this->sheetRelationsSheet[(int)($this->xlsxObject->$sheet->id -1 )]->attributes();
                            $this->xlsxObject->$sheet->type = $sheetRelAttributes['Type'];
                            $this->xlsxObject->$sheet->file = $sheetRelAttributes['Target'];
                            $this->xlsxObject->$sheet->hasData = $this->hasSheetData($this->xlsxObject->$sheet);
                            $this->xlsxObject->$sheet->dimension = $this->getSheetDimension($this->xlsxObject->$sheet);
                 }//end for
                
            }else {
                throw new Exception('xl/workbook.xml not found in the extraction directory');
            }//end if
            
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    return true;
    
    }//end mapmapXlsxFile();
    
    
    protected function getSheetRelationsSheet() {
        $xmlDoc = new DOMDocument();
         //beolvassuk a sheet-ek fileneveit és a típusokat
        if (is_file($this->getExtractDir().'/xl/_rels/workbook.xml.rels') === true) {

            if ($xmlDoc->load($this->getExtractDir().'/xl/_rels/workbook.xml.rels') === true) {
                $workBookRels = simplexml_import_dom($xmlDoc);
                
                foreach($workBookRels as $workBookRel) {
                    $attributes = $workBookRel->attributes();
                        if(strpos($attributes['Target'], 'sheet')){
                            $this->sheetRelationsSheet[] = $workBookRel;
                        }
                   
                }//end foreach
                    
            }else {
                throw new Exception('unable to load /xl/_rels/workbook.xml.rels');
            }//end if

        }else{
            throw new Exception('file /xl/_rels/workbook.xml.rels not found in extraction directory');
        }//end if
        
    }//sheetRelationsSheet();
    
    
    /**
     * Ellenőrzi, hogy van-e sharedStrings
     * @return boolean 
     */
    public function sharedStringsCheck() {
        
        if (file_exists($this->getExtractDir().'/xl/sharedStrings.xml') === true) {
            return true;
        }else {
            return false;
        }//end if
        
    }//end sharedStringsCheck
    
    /**
     * Eldönti egy sheet objektumról, hogy van-e benne adat
     * @param type $sheet 
     */
    protected function hasSheetData($sheet){
        //$this->echoInFriendlyFormat($sheet);
        $xmlDoc = new DOMDocument();
        if (file_exists($this->getExtractDir().'/xl/'.$sheet->file) === true) {
            
            if ($xmlDoc->load($this->getExtractDir().'/xl/'.$sheet->file) === true) {
                $sheetObj = simplexml_import_dom($xmlDoc);
                $sheetAttributes = $sheetObj->dimension->attributes();

                if (strpos($sheetAttributes['ref'],':') === false) {
                    return false;
                }else {
                    return true;
                }//end if
                                
            }else{
                throw new Exception('Unable to load '.$sheet->file);
            }//end if
        
        }else {
            throw new Exception($sheet->file.' not found');
        }//end if
        
    }//hasSheetData
    
    
    /**
     * get The dimension of sheet
     * @param type $sheet 
     */
    protected function getSheetDimension ($sheet) {
        $xmlDoc = new DOMDocument();
        
        if (file_exists($this->getExtractDir().'/xl/'.$sheet->file) === true) {
            
            if ($xmlDoc->load($this->getExtractDir().'/xl/'.$sheet->file) === true) {
                $sheetObj = simplexml_import_dom($xmlDoc);
                $sheetAttributes = $sheetObj->dimension->attributes();
                
                return $sheetAttributes['ref'];
                                               
            }else{
                throw new Exception('Unable to load '.$sheet->file);
            }//end if
        
        }else {
            throw new Exception($sheet->file.' not found');
        }//end if
    
        
    }//end getSheetDimension
    
    
    protected function loadSharedStrings(){
        
        $xmlDoc = new DOMDocument();
        if ($xmlDoc->load($this->getExtractDir().'/xl/sharedStrings.xml') === true) {
            $sharedStringsTmp = simplexml_import_dom($xmlDoc);
            $sharedStringsAttribute = $sharedStringsTmp->attributes();
            $this->sharedStrings->count = $sharedStringsAttribute["count"];
            $this->sharedStrings->uniqueCount = $sharedStringsAttribute["uniqueCount"];
                        
            foreach($sharedStringsTmp->si as $val){
                 $this->sharedStrings->data[] = $val->t;
            }
            
        }else{
            throw new Exception('Unable to load sharedStrings.xml');
        }//end if

        return true;
    }//end loadSharedStrings();
    
    
    
    /**
     * echoes $mit in friendly format to the screen
     * @param mixed $mit 
     */
    protected function echoInFriendlyFormat($mit) {
        echo '<PRE>',var_dump($mit),'</PRE>';
    }
    
    protected function setExtractDir($extractDirPath) {
        $this->extractdir = $extractDirPath;
    }
    
    public function getExtractDir() {
        return $this->extractdir;
    }
    
    public function getFileName(){
        return $this->filename;
    }
    
    protected function ExtractZipArchive() {
        
        $zip = new ZipArchive;
        
        $extractDir = $this->getExtractDir();
        
        if(empty($extractDir) === true) {
            throw new Exception('Extraction directory path not defined, user the setExtractDir() method');
        }
                
        $errorCode = $zip->open($this->getFileName());
        
        if ($errorCode !== true){
            $error = '';
            switch($errorCode){
               case ZIPARCHIVE::ER_EXISTS: $error = 'File already exists.';break;
               case ZIPARCHIVE::ER_INCONS: $error = 'Zip archive inconsistent.';break;
               case ZIPARCHIVE::ER_INVAL:  $error = 'Invalid argument.';break;
               case ZIPARCHIVE::ER_MEMORY: $error = 'Malloc failure.';break;
               case ZIPARCHIVE::ER_NOENT:  $error = 'No such file.';break;
               case ZIPARCHIVE::ER_NOZIP:  $error = 'Not a zip archive.';break;
               case ZIPARCHIVE::ER_OPEN:   $error = "Can't open file.";break;
               case ZIPARCHIVE::ER_READ:   $error = "Read error.";break;
               case ZIPARCHIVE::ER_SEEK:   $error = "Seek error.";break;
            } 
            throw new Exception($error);
        }//end if;
        
        if ($zip->extractTo($this->getExtractDir()) === false) {
            throw new Exception ('Unable to extract to the specified directory');
        }//end if
        
        if ($zip->close() === false) {
            throw new Exception('Unable to close Zip archive');
        }//end if
            
    }//end ExtractZipArchive()
    
    
}//end class
?>
