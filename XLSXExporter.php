<?php

require_once JPATH_COMPONENT_ADMINISTRATOR.DS.'Classes/pclzip.lib.php';


Class XLSXExporter {
    
    protected $db;
    protected $destinationFolder;
    protected $filename;
    protected $tmpDirname;
    protected $productsTable;
    protected $extractionDirectory;
    protected $columnLetters = array();
    protected $header;
    protected $sharedStrings;
    protected $sheetObject;
    protected $hasHeader;
    protected $defaultFilesDirectory;
    protected $total=NULL;
    
    public function __construct($filename, $destinationFolder) {
        
        if (file_exists(JPATH_SITE.DS.'administrator/components/com_virtuemart/virtuemart.cfg.php') === true) {
                require_once JPATH_SITE.DS.'administrator/components/com_virtuemart/virtuemart.cfg.php';
        }else {
            throw new Exception('You have to install virtueMart component!');
        }//end if
        
        $counter = 0;
        for ($c = 'A' ; $c <= 'Z'; $c++) {
            $this->columnLetters[$counter] = $c;
            $counter++;
        }//end for
        
        $this->filename = $filename;
        $this->destinationFolder = $destinationFolder;
        $this->extractionDirectory = time();
        $this->vm_prefix = VM_TABLEPREFIX;
        $this->db =& JFactory::getDBO();
        $this->header = array(
            'Category',
            'Product Name',
            'Product SKU (Required) ',
            'Product short description',
            'product description',
            'Product price',
            'Product currency',
            'Product in stock',
            'Product thumbnail image path ',
            'Product full image',
            'Product special',
            'Product discount',
            'Discount is percent (1 = "yes", 0 = "no")',
            'product tax id (default 2)',
            'Product weight',
            'Product weight unit of measure',
            'Product length',
            'Product width',
            'Product height',
            'Product dimensions unit of measure',
            'Product Publish'
            ); 
        
        
        $this->hasHeader = false;
        
    }//end __constuct
    
    public function export($start, $step) {
        
        try{
            $this->extractDefaultXlsx();
            $this->getProductTableArray();
            $this->productsTable = $this->getProductTableArray($start, $step);
                        
            $this->loadSharedStrings();
            $this->loadSheetobject();
            $this->clearDefaultData();
            
            $this->loadHeaders();           
            
            $this->loadData();
            
            $this->saveToXml($this->sheetObject,$this->getExtractDirectoryPath().'/xl/worksheets/sheet1.xml');
            $this->saveToXml($this->sharedStrings,$this->getExtractDirectoryPath().'/xl/sharedStrings.xml');
            
            $fileClass = $this->saveToFile();
            
            return $fileClass;
            //$this->echoInFriendlyFormat($this->columnLetters);
            //$this->echoInFriendlyFormat($this->sharedStrings);
            
        }catch(Exception $e){
            throw new Exception($e->getMessage());
        }//end try
        
    }//end export()
    
    /**
     * Loads the sheet data
     */
    protected function loadData(){
        
        $rowOffset = 0;
        if($this->hasHeader === true){
            $rowOffset = 2;
        }//end if
        
        for ($c = $rowOffset ; $c < ($this->productsTable->data->rowCount + $rowOffset); $c++) {
            
            //adding row
            $row = $this->sheetObject->sheetData->addChild('row');
            $row->addAttribute('collapsed','false');
            $row->addAttribute('customFormat','false');
            $row->addAttribute('customHeight','false');
            $row->addAttribute('hidden','false');
            $row->addAttribute('ht','12.8');
            $row->addAttribute('outlineLevel','0');
            $row->addAttribute('r',$c);
            
            $columnCounter = 0;
            foreach($this->productsTable->data->rows[($c - $rowOffset)] as $fieldName => $value){
               $cell = $row->addChild('c');
                             
               $cell->addAttribute('r',$this->columnLetters[$columnCounter].$c);
               $cell->addAttribute('s','0');
                   
               if (is_numeric($value) == true) {
                   $cellType = 'n';
                   $cellData = $value;
               }else{
                   $cellType = 's';
                    $cellData = $this->getSharedStringId($value);
               }    
                   $cell->addAttribute('t',$cellType);
                   $cellValue = $cell->addChild('v',$cellData);
               
                   $columnCounter++;
            }
        }//end for
        
        
        $endCoord =  $this->columnLetters[($columnCounter-1)].($c-1);
        $this->sheetObject->dimension->attributes()->ref = 'A1:'.$endCoord;
        $this->echoInFriendlyFormat($this->sheetObject->dimension);
        
        return true;
    }//end loadData()
    
    /**
     * save data to xlsxX file and returns the filename
     */
    protected function saveToFile () {
        $filename = '';
        // $errorCode = $zip->open(JPATH_COMPONENT_ADMINISTRATOR.DS."defaultXlsx/default.xlsx");
                      
        $zip = new ZipArchive();
        $errorCode = $zip->open($this->getExtractDirectoryPath().'/'.$this->filename.'.xlsx',ZIPARCHIVE::CREATE);
        
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
        
        $extractDir = $this->getExtractDirectoryPath();

        $zip->addEmptyDir('docProps');
        $zip->addEmptyDir('_rels');
        $zip->addEmptyDir('xl');
        $zip->addEmptyDir('xl/_rels');
        $zip->addEmptyDir('xl/worksheets');

        //root files
        $zip->addFile($extractDir.'/[Content_Types].xml', '[Content_Types].xml');   

        //docProps files
        $zip->addFile($extractDir.'/docProps/app.xml', 'docProps/app.xml');   
        $zip->addFile($extractDir.'/docProps/core.xml', 'docProps/core.xml');   

        //_rels files
        $zip->addFile($extractDir.'/_rels/.rels', '_rels/.rels'); 

        //xl files
        $zip->addFile($extractDir.'/xl/styles.xml', 'xl/styles.xml'); 
        $zip->addFile($extractDir.'/xl/workbook.xml', 'xl/workbook.xml'); 
        $zip->addFile($extractDir.'/xl/sharedStrings.xml', 'xl/sharedStrings.xml');

        //xl_rel files
        $zip->addFile($extractDir.'/xl/_rels/workbook.xml.rels', 'xl/_rels/workbook.xml.rels'); 

        //worksheet files 
        $zip->addFile($extractDir.'/xl/worksheets/sheet1.xml', 'xl/worksheets/sheet1.xml'); 

        $zip->close();

        $fileClass = new stdClass();
        $fileClass->fileName = $this->filename.'.xlsx';
        $fileClass->filePath = $this->getExtractDirectoryPath().'/'.$this->filename.'.xlsx';
        
        return $fileClass;
        
    }//saveToFile
    
    
    /**
     * Load the headers to the xlsx object
     * @return type 
     */
    protected function loadHeaders() {
        
        for ($c = 0; $c < count($this->header) ; $c++ ) {
                
                $child = $this->sheetObject->sheetData->row->addChild('c');
                $childValue = $child->addChild('v',$this->getSharedStringId($this->header[$c]) );
                $child->addAttribute('r',$this->columnLetters[$c].'1');
                $child->addAttribute('s','0');
                $child->addAttribute('t','s');
                
        }// end for
        
        $this->hasHeader = true;
        
        return true;
    }//loadHeaders()
    
    
    /**
     * Clears the initial data from sharedStrings array and sheet array
     */
    protected function clearDefaultData(){
        unset($this->sharedStrings->si[0]);
        $this->sharedStrings->attributes()->count = 0;
        $this->sharedStrings->attributes()->uniqueCount = 0;
        unset($this->sheetObject->sheetData->row->c[0]);
    }//clearDefaultData()
    
    protected function saveToXml($xmlObject, $filePath) {
        if ($xmlObject->asXML($filePath) === false){
            throw new Exception('Unable to save '.$filePath);
        }
    }
    
    
    protected function loadSheetobject() {
        $this->sheetObject = $this->readXml($this->getExtractDirectoryPath().'/xl/worksheets/sheet1.xml');
    }//end loadSheetobject()
    
    protected function loadSharedStrings() {
        $this->sharedStrings = $this->readXml($this->getExtractDirectoryPath().'/xl/sharedStrings.xml');
        
    }//end loadSharedStrings()    
    
    /**
     * Method to get SharedStringId from sharedString array. Insert new strings and manage the counters
     */
    protected function getSharedStringId($searchString) {
        $counter = 0;
        
        $sharedStringsAttributes = $this->sharedStrings->attributes();
        $totalCounter  = (int)$sharedStringsAttributes['count'];
        $uniqueCounter = (int)$sharedStringsAttributes['uniqueCount'];
        
        foreach ($this->sharedStrings as $oneSharedString) {
            
            
            if ( $oneSharedString->t == $searchString) {
                $totalCounter++;
                $this->sharedStrings->attributes()->count = $totalCounter;
                return $counter;
            }// end if
            
            $counter++;
        }//end foreach
        
        $totalCounter++;
        $this->sharedStrings->attributes()->count = $totalCounter;
        
        $uniqueCounter++;
        $this->sharedStrings->attributes()->uniqueCount = $uniqueCounter;
        
        $child = $this->sharedStrings->addChild('si');
        $child->addChild('t',$searchString);
        
        return $counter;
    }//end getSharedStringId()
    
    
    public function extractDefaultXlsx(){
        //JPATH_COMPONENT_ADMINISTRATOR.DS."Classes/XLSXReader.php
        
        $zip = new ZipArchive;
        
        $extractDir = $this->getExtractDirectoryPath();
        
        $errorCode = $zip->open(JPATH_COMPONENT_ADMINISTRATOR.DS."defaultXlsx/default.xlsx");
        
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
        
        if ($zip->extractTo($extractDir) === false) {
            throw new Exception ('Unable to extract to the specified directory');
        }//end if
        
        if ($zip->close() === false) {
            throw new Exception('Unable to close Zip archive');
        }//end if
        
    }//end extractDefaultXlsx();
    
    
    protected function setExtractionDirectory($dirName){
       $this->extractionDirectory = $dirName; 
    }//end setExtractionDirectory
    
    
    protected function getExtractDirectoryPath(){
       
       $path = $this->destinationFolder.'/'.$this->extractionDirectory;
        
       return $path;
    }//createDirectoryStructure()
    
    /**
     * Gets the product table into an object for further usage
     */
    
    protected function getProductTableArray($start, $step){
        $headerRow = array(
            'category_id',
            'product_name',
            'product_sku',
            'product_s_desc',
            'product_desc',
            'product_price',
            'product_currency',
            'product_in_stock',
            'product_thumb_image',
            'product_full_image',
            'product_special',
            'product_discount_id',
            'is_percent',
            'product_tax_id',
            'product_weight',
            'product_weight_uom',
            'product_length',
            'product_width',
            'product_height',
            'product_lwh_uom',
            'product_publish'
            );
        
        $productList = $this->getProductList($start, $step);
        
        $ProductsTable = new stdClass();
        
        foreach($productList as $Product){
            $categoryString  = $this->getCategories($Product->product_id);
            $productPrice    = $this->getProductPrice($Product->product_id);
            $productDiscount = $this->getDiscount($Product->product_discount_id);
            
            $row = new stdClass();
            
            $row->category_id           = $categoryString;
            $row->product_name          = $Product->product_name;
            $row->product_sku           = $Product->product_sku;
            $row->product_s_desc        = $Product->product_s_desc;
            $row->product_desc          = $Product->product_desc;
            $row->product_price         = $productPrice->price;
            $row->product_currency      = $productPrice->currency;
            $row->product_in_stock      = $Product->product_in_stock;
            $row->product_thumb_image   = $Product->product_thumb_image;
            $row->product_full_image    = $Product->product_full_image;
            $row->product_special       = $Product->product_special;
            $row->product_discount_id   = $productDiscount->amount;
            $row->is_percent            = $productDiscount->is_percent;
            $row->product_tax_id        = $Product->product_tax_id;
            $row->product_weight        = $Product->product_weight;
            $row->product_weight_uom    = $Product->product_weight_uom;
            $row->product_length        = $Product->product_length;
            $row->product_width         = $Product->product_width;
            $row->product_height        = $Product->product_height;
            $row->product_lwh_uom       = $Product->product_lwh_uom;
            $row->product_publish       = $Product->product_publish;
            
            
            $ProductsTable->data->rows[] = $row;
        }//end foreach
        
        $ProductsTable->data->rowCount = count($productList);
        
        return $ProductsTable;
    }//end getProductTableArray()
    
    protected function getProductList($start = 0, $step = null) {
        
        if (isset($step) === true && $step > 0) {
            $limit = ' limit '.$start.', '.$step;
        }
//        Ha nincs értéke a totalnak akkor kérdezzük le
        if(empty ($this->total) == true){
            $productdb = 'select count(*) as darab from #__'.$this->vm_prefix.'_product';
            $this->db->setQuery($productdb);
            $this->total= $this->db->loadResult();
        }
        if($this->total > $start){
            $productSelect = 'select * from #__'.$this->vm_prefix.'_product'.$limit;
            
            $this->db->setQuery($productSelect);
            $productObject = $this->db->loadObjectlist();
        }
        
        return $productObject;
    }//end getProductList()
    
    public function getTotal() {
        
    }
    
    /**
     * Gets the product categories in string format ( mainCategory1/subcategory1#mainCategory2/subcategory2 )
     * @param integer $productId
     * @return string 
     */
    protected function getCategories($productId){
        
        $kategoriakSelect = 'select * from #__'.$this->vm_prefix.'_product_category_xref t where t.product_id = '.$productId;
        $this->db->setQuery($kategoriakSelect);
        $productKategoriak = $this->db->loadObjectlist();
            
        $katGyujto = array();
        
        foreach($productKategoriak as $productKategoria) {
                $kategoriaNevek = array();
                $vizsgaltId = $productKategoria->category_id;
                
                do {
                    $kategoriaSelect = 'SELECT t.*, t1.category_name FROM #__'.$this->vm_prefix.'_category_xref t, #__'.$this->vm_prefix.'_category t1
                                        where t.category_child_id = t1.category_id
                                        and t.category_child_id = '.$vizsgaltId;
                    $this->db->setQuery($kategoriaSelect);
                    $kategoriaXref = $this->db->loadObjectlist();
                    $kategoriaNevek[] = $kategoriaXref[0]->category_name;
                    $parent_id = $kategoriaXref[0]->category_parent_id;
                    $vizsgaltId = $parent_id;
                }while($parent_id != 0);
                
                $katNevekGoodOrder = array();
                
                for($c = count($kategoriaNevek) - 1 ; $c >= 0; $c--){
                    $katNevekGoodOrder[] = $kategoriaNevek[$c];
                }
                
                $katGyujto[] = implode('/',$katNevekGoodOrder);
                
            }//foreach()
            
            $katGyujtoString = implode('#',$katGyujto);
                        
            return $katGyujtoString;
    }//end getCategories()
    
    /**
     * ges the product first available price
     * @param int $productId 
     */
    protected function getProductPrice($productId){
        
        $productPriceSelect = 'select t.product_currency, t.product_price, t1.* from 
                                    #__'.$this->vm_prefix.'_product_price t,
                                    #__'.$this->vm_prefix.'_product t1, 
                                    #__'.$this->vm_prefix.'_vendor t2
                                    where 
                                    t.product_id = t1.product_id
                                    and t1.vendor_id = t2.vendor_id
                                    and t.product_currency = t2.vendor_currency
                                    and '.time().' between t.product_price_vdate and t.product_price_edate
                                    and t.product_id = '.$productId.'

                                    union

                                    select t.product_currency, t.product_price,  t1.* from 
                                        #__'.$this->vm_prefix.'_product_price t,
                                        #__'.$this->vm_prefix.'_product t1, 
                                        #__'.$this->vm_prefix.'_vendor t2
                                        where 
                                        t.product_id = t1.product_id
                                        and t1.vendor_id = t2.vendor_id
                                        and t.product_currency = t2.vendor_currency
                                        and t.product_price_vdate = 0
                                        and t.product_price_edate = 0
                                        and t.product_id = '.$productId.'

                                    union 

                                    select t.product_currency, t.product_price, t1.* from 
                                        #__'.$this->vm_prefix.'_product_price t,
                                        #__'.$this->vm_prefix.'_product t1 
                                        where 
                                        t.product_id = t1.product_id
                                        and t.product_id = '.$productId;    
            
            $this->db->setQuery($productPriceSelect);
            
            $productPrices = $this->db->loadObjectlist();
            $productPrice = $productPrices[0];
            
            $retPrice = new stdClass();
            $retPrice->price    = $productPrice->product_price;
            $retPrice->currency = $productPrice->product_currency;
            
            return $retPrice;
    }//end getProductPrice()
    
    /**
     * Gets the product discount in string format
     * @param integer $productId
     * @return string 
     */
    protected function getDiscount($productDiscountId) {
        
        $productDiscount = new stdClass();
        
        $amount = 0;
        
        if ($productDiscountId != 0) {
           $discountSelect = 'select * from #__'.$this->vm_prefix.'_product_discount t where t.discount_id = '.$productDiscountId; 
           $this->db->setQuery($discountSelect);
           $discount = $this->db->loadObjectList();
           $productDiscount->amount = $discount[0]->amount;
           $productDiscount->is_percent = $discount[0]->is_percent;
        }else{
           $productDiscount->amount = 0;
           $productDiscount->is_percent = 0;
        }
       return $productDiscount;  
    }//getDiscount()
    
    
    
    
    /**
    * echoes $mit in friendly format to the screen
    * @param mixed $mit 
    */
    protected function echoInFriendlyFormat($mit) {
        echo '<PRE>',var_dump($mit),'</PRE>';
    }//end echoInFriendlyFormat()
    
    /**
     * Method to read xslsx Sheet xml into an simplexml object
     * @param int $sheetId 
     */
    protected function readXml($XmlPath) {
         $xmlDoc = new DOMDocument();
         if (file_exists($XmlPath) === true){    
             if ($xmlDoc->load($XmlPath) === true){
                 $retObject = simplexml_import_dom($xmlDoc);
             }else{
                 throw new Exception('Unable to load '.$xmlPath);
             }//end if
         }else {
             throw new Exception('Unable to open '.$XmlPath);
         }//end if
    
         return $retObject;
         
    }//end readSheet
    
    
    
    
}//end class

?>
