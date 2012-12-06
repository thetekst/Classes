<?php
    require_once('translitIt.php');
	
    class XLSXImporter {
       
        protected $xlsxObject; 
        protected $vm_prefix;
        protected $db;
        protected $statistics;
		
        public function __construct($xlsxObject) {
            
            if (file_exists(JPATH_SITE.DS.'administrator/components/com_virtuemart/virtuemart.cfg.php') === true) {
                require_once JPATH_SITE.DS.'administrator/components/com_virtuemart/virtuemart.cfg.php';
            }else {
                throw new Exception('You have to install virtueMart component!');
            }
        
            $this->vm_prefix = VM_TABLEPREFIX;
            $this->db =& JFactory::getDBO();
            $this->xlsxObject = $xlsxObject;
            $this->statistics->inserted = 0;
            $this->statistics->updated = 0;
            
        }//end __construct
        
        /**
         * Method to import XLSX spreadsheet data to database
         * @todo átalakítani olyanra, hogy ne csak az első sheet-ből dolgozzon mert most csak ezt tudja
         */
        
        public function import(){

            //$this->echoInFriendlyFormat($this->xlsxObject);
            if ($this->xlsxObject->Sheet0->hasData === true) {
               $sheet = $this->xlsxObject->Sheet0;
               //the first line contains the field names in this version. We start the reading from the second
               
               for ($c = 1; $c <  $sheet->data->rowCount; $c++) {
                   $oneProduct = $this->loadOneRecord($sheet, $c);
			   
                   //$this->echoInFriendlyFormat($oneProduct->product_name);
                   if (empty($oneProduct->product_sku) === false ) {

                       if ($this->isThereProductLikeThis($oneProduct->product_sku) === true) { //van ilyen termék update
							
                           $productId = $this->getProductIdFromSku($oneProduct->product_sku);

						   
                           /*$this->manageCategory($oneProduct->category_id, $productId);*/
                           /*$this->managePriceAndCurrency($productId, $oneProduct->product_price, $oneProduct->product_currency);*/
                           /*$discountId = $this->manageDiscount($oneProduct->product_discount_id, $oneProduct->is_percent);*/
                           $this->loadProductIntoDb($oneProduct, $productId);
						   $this->loadProductIntoDbRuRU($oneProduct, $productId);
						   //$this->getProductName($productId, $oneProduct->product_name, $oneProduct->product_s_desc, $oneProduct->product_desc);
						   


                       }else {//nincs ilyen termék, insert
							
                           /*$discountId = $this->manageDiscount($oneProduct->product_discount_id, $oneProduct->is_percent);*/
						   
                           $productId = $this->loadProductIntoDb($oneProduct, null);

						   $this->loadProductIntoDbRuRU($oneProduct, $productId);
						   //$this->getProductName($productId, $oneProduct->product_name, $oneProduct->product_s_desc, $oneProduct->product_desc);

                           /*$this->manageManufacturer($productId);*/
                           /*$this->manageCategory($oneProduct->category_id, $productId);*/
                           /*$this->managePriceAndCurrency($productId, $oneProduct->product_price, $oneProduct->product_currency);*/

						   

                       }//end if
                       
					   
						
                   }//end if 
               }//end for 
                
               
            }else{
                throw new Exception('The first Sheet has no data. Nothing to do...'); 
            }
            
            return $this->statistics;
        }//end import
        
        /**
         * load one product into an stdClass
         */
        protected function loadOneRecord($sheet, $c) {
                  
            $item = new stdClass();
            
            /*$item->category_id          = $sheet->data->numericIndexTable[$c][0];*/
            $item->product_sku          = $sheet->data->numericIndexTable[$c][1];
			$item->product_weight       = $sheet->data->numericIndexTable[$c][2];
			$item->published      		= $sheet->data->numericIndexTable[$c][3];
			$item->product_name         = $sheet->data->numericIndexTable[$c][4];
            $item->product_s_desc       = $sheet->data->numericIndexTable[$c][5];
            $item->product_desc         = $sheet->data->numericIndexTable[$c][6];
			$item->slug         		= $sheet->data->numericIndexTable[$c][7];
			
            /*$item->product_price        = $sheet->data->numericIndexTable[$c][5];
            $item->product_currency     = $sheet->data->numericIndexTable[$c][6];
            $item->product_in_stock     = $sheet->data->numericIndexTable[$c][7];
            $item->product_thumb_image  = $sheet->data->numericIndexTable[$c][8];
            $item->product_full_image   = $sheet->data->numericIndexTable[$c][9];
            $item->product_special      = $sheet->data->numericIndexTable[$c][10];
            $item->product_discount_id  = $sheet->data->numericIndexTable[$c][11];
            $item->is_percent           = $sheet->data->numericIndexTable[$c][12];
            $item->product_tax_id       = $sheet->data->numericIndexTable[$c][13];
            
            $item->product_weight_uom   = $sheet->data->numericIndexTable[$c][15];
            $item->product_length       = $sheet->data->numericIndexTable[$c][16];
            $item->product_width        = $sheet->data->numericIndexTable[$c][17];
            $item->product_height       = $sheet->data->numericIndexTable[$c][18];
            $item->product_lwh_uom      = $sheet->data->numericIndexTable[$c][19];*/
            
            
            return $item;
        }//end loadOneRecord()
        
        
        protected function isThereProductLikeThis($product_sku) {
            $isThereProductQuery = 'select count(*) from #__'.$this->vm_prefix.'_products t where t.product_sku = "'.$product_sku.'"';
            $this->db->setQuery($isThereProductQuery);
            $productCount = $this->db->loadResult();
            
            if ($productCount == 0) {
                return false;
            }else {
                return true;
            }
        }//end isThereProductLikeThis()
        
        protected function getProductIdFromSku($product_sku) {
            $productIdQuery = 'select * from #__'.$this->vm_prefix.'_products t where t.product_sku = "'.$product_sku.'"';

            $this->db->setQuery($productIdQuery);

            return $this->db->loadResult();
        }//end getProductIdFromSku()
		
		protected function getProductName($productId, $product_name, $product_s_desc, $product_desc) {
		
			if (empty($product_name) === false && empty($product_s_desc) === false && empty($product_desc) === false) {
			
				/*$productIdQuery = 'select p.virtuemart_product_id from #__'.$this->vm_prefix.'_products_ru_ru AS p where p.virtuemart_product_id = "'.$productId.'"';
            $this->db->setQuery($productIdQuery);*/
			
			/*$productsRuRuNameId = $this->db->loadResult();
			$productsRuRuProductSDescId = $this->db->loadResult();
			$productsRuRuProductDescId = $this->db->loadResult();*/
			
			$productsRuRuClass = new stdClass();
			
			/*if (empty($productsRuRuNameId) === false) {    
				$productsRuRuNameClass->product_name = $productsRuRuNameId;
				$productsRuRuNameClass->product_s_desc = productsRuRuProductADescId();
				$productsRuRuNameClass->product_desc = productsRuRuProductDescId();
			}else {
				$productsRuRuNameClass->product_name = "false";
				$productsRuRuNameClass->product_s_desc = "";
				$productsRuRuNameClass->product_desc = "";
			}//end if*/
			
			$productsRuRuClass->virtuemart_product_id       = (int)$productId;
			$productsRuRuClass->product_name    			= $product_name;
			$productsRuRuClass->product_s_desc    			= $product_s_desc;
			$productsRuRuClass->product_desc    			= $product_desc;
			
            if (empty($productsRuRuClass->virtuemart_product_id) === false) {  
                    if (!$this->db->updateObject("#__".$this->vm_prefix."_products_ru_ru",$productsRuRuClass,'virtuemart_product_id') ) {
                                throw new Exception( $this->db->stderr() );
                    }//end if
					
                }else {
                    if (!$this->db->insertObject("#__".$this->vm_prefix."_products_ru_ru",$productsRuRuClass) ) {
                                throw new Exception( $this->db->stderr() );
                    }//end if
					
                }//end if
                
            }//end if
			return true;
        }//end getProductName()
		      
        
        /**
        * echoes $mit in friendly format to the screen
        * @param mixed $mit 
        */
		
		/*-------------
		
        protected function echoInFriendlyFormat($mit) {
            echo '<PRE>',var_dump($mit),'</PRE>';
        }//end echoInFriendlyFormat()
        ---------------*/
        /**
         * Method to manage (insert or update) categories, category_xref, and product_category_xref
         * @param string $categoryString 
         */
		 
		 
		 /*-------------
		 
        protected function manageCategory ($categoryString, $productId = null) {
            
            $db2 =& JFactory::getDBO();
            
            if (empty($categoryString) === false) {
                    
                    $command2 = 'delete from #__'.$this->vm_prefix.'_product_category_xref where product_id = '.$productId;
                    $this->db->setQuery($command2);
                    $this->db->query();
                    
                    $osszKategoraiaDarabok = explode('#',$categoryString);

                    for ($d = 0; $d < count($osszKategoraiaDarabok); $d++) {

                        $kategoriaDarabok = explode('/', $osszKategoraiaDarabok[$d]);
                        $parent_id = 0;

                        for($c = 0; $c < count($kategoriaDarabok); $c++) {

                            $kategoriaVanQuery = 'select t.category_id from #__'.$this->vm_prefix.'_category t where LOWER(t.category_name) = "'.  strtolower($kategoriaDarabok[$c]).'"';

                            $this->db->setQuery($kategoriaVanQuery);
                            $kategoriaId = $this->db->loadResult();

                            if (empty($kategoriaId) === false) {

                                $kategoriaXrefQuery = 'select count(*) from #__'.$this->vm_prefix.'_category_xref t where t.category_parent_id = '.$parent_id.' and t.category_child_id = '.$kategoriaId;

                                $this->db->setQuery($kategoriaXrefQuery);
                                $vaneKategoriaXref = $this->db->loadResult();
                                //echo ' <br>'.$vaneKategoriaXref;

                                if($vaneKategoriaXref == 0) {//bár van ilyen kategoria, de nem olyan kontextusban ahogy az excelben
                                    $kategoriaClass = new stdClass();
                                    $kategoriaClass->category_name = $kategoriaDarabok[$c];
                                    $kategoriaClass->virtuemart_vendor_id = 1;
                                    $this->db->insertObject("#__".$this->vm_prefix."_category",$kategoriaClass,'category_id');
                                    $new_kategory_id = $this->db->insertId();
                                    $kategoriaXrefClass = new stdClass();
                                    $kategoriaXrefClass->category_parent_id = $parent_id;
                                    $kategoriaXrefClass->category_child_id = $new_kategory_id;

                                    if (!$this->db->insertObject("#__".$this->vm_prefix."_category_xref",$kategoriaXrefClass) ) {
                                        throw new Exception( $this->db->stderr() );
                                    }
                                    $parent_id = $new_kategory_id;
                                }else {
                                    $parent_id = $kategoriaId;
                                }

                            }else {
                                $kategoriaClass = new stdClass();
                                $kategoriaClass->category_name = $kategoriaDarabok[$c];
                                $kategoriaClass->virtuemart_vendor_id = 1;
                                $kategoriaClass->category_publish = 'Y';
                                $this->db->insertObject("#__".$this->vm_prefix."_category",$kategoriaClass,'category_id');
                                $new_kategory_id = $this->db->insertId();

                                $kategoriaXrefClass = new stdClass();
                                $kategoriaXrefClass->category_parent_id = $parent_id;
                                $kategoriaXrefClass->category_child_id = $new_kategory_id;
                                if (!$this->db->insertObject("#__".$this->vm_prefix."_category_xref",$kategoriaXrefClass) ) {
                                    throw new Exception( $this->db->stderr() );
                                }
                                $parent_id = $new_kategory_id;
                            }//end if

                        }//end for
                        
                        if ( (empty($new_kategory_id) === true ) && (empty($kategoriaId) === false) ) {
                            
                            $categoryInsertid = $kategoriaId;
                            $termekVanEQuery = 'select count(*) from #__'.$this->vm_prefix.'_product_category_xref t where t.product_id ='.$productId.' and t.category_id = '.$categoryInsertid;
                            $db2->setQuery($termekVanEQuery);
                            $van = $db2->loadResult();

                            if ($van == 0) {
                                
                                $productCategoryXrefClass = new stdClass();
                                $productCategoryXrefClass->category_id = $categoryInsertid;
                                $productCategoryXrefClass->product_id = (int)$productId;

                                if (!$this->db->insertObject("#__".$this->vm_prefix."_product_category_xref",$productCategoryXrefClass)) {
                                    throw new Exception($this->db->stdErr()); 
                                }//end if

                            }
                        } else if (empty($new_kategory_id) === false && empty($kategoriaId) === true) {
                            $categoryInsertid = $new_kategory_id;
                            $productCategoryXrefClass = new stdClass();
                            $productCategoryXrefClass->category_id = $categoryInsertid;
                            $productCategoryXrefClass->product_id = (int)$productId;

                            if (!$this->db->insertObject("#__".$this->vm_prefix."_product_category_xref",$productCategoryXrefClass)) {
                                throw new Exception($this->db->stdErr()); 
                            }//end if
                        } //end if
                        
                        if(isset($new_kategory_id) === true) {
                            unset($new_kategory_id);
                        }
                    }  //end for      
                    //kategoriavisszafejtés és insert kész
                }//end if
            return true;
        }//manageCategory()
        
        protected function managePriceAndCurrency($productId, $productPrice, $productCurrency) {
            
            if (empty($productPrice) === false && empty($productCurrency) === false) {
                
                $productPriceQuery = 'select t.virtuemart_product_price_id from #__'.$this->vm_prefix.'_product_prices t where t.virtuemart_product_id = '.$productId.' and t.virtuemart_shoppergroup_id = 5';
                $this->db->setQuery($productPriceQuery);

                $productPriceId = $this->db->loadResult();

                $productPriceClass = new stdClass();

                if (empty($productPriceId) === false) {    
                    $productPriceClass->product_price_id = $productPriceId;
                }else {
                    $productPriceClass->shopper_group_id = 5;
                }//end if

                $productPriceClass->product_id       = (int)$productId;
                $productPriceClass->product_price    = (float)str_replace(',','.',$productPrice);
                $productPriceClass->product_currency = (string)$productCurrency;


                if (empty($productPriceId) === false) {  
                    if (!$this->db->updateObject("#__".$this->vm_prefix."_product_prices",$productPriceClass,'virtuemart_product_price_id') ) {
                                throw new Exception( $this->db->stderr() );
                    }//end if
                }else {
                    if (!$this->db->insertObject("#__".$this->vm_prefix."_product_prices",$productPriceClass) ) {
                                throw new Exception( $this->db->stderr() );
                    }//end if
                    
                }//end if
                
            }//end if
            
            return true;
        }//end managePriceAndCurrency()
        
		----------*/
		
        /**
         * 
         * @param string $productDiscount 
         */
		 
		 /*--------------
        protected function manageDiscount($productDiscount, $isPercent) {
                        
            if ( empty($productDiscount) === false && $productDiscount != 0) {
                //discount cuccok
                
                $productDiscountQuery = 'select count(*) from #__'.$this->vm_prefix.'_products_ru_ru t where t.amount = '.$productDiscount.' and t.is_percent = '.$isPercent;
                $this->db->setQuery($productDiscountQuery);
                $vanDiscount = $this->db->loadResult();

                if($vanDiscount > 0){ //van ilyen discount lekérdezzük az Id-jét
                    $productDiscountIdQuery = 'select t.discount_id from #__'.$this->vm_prefix.'_product_discount t where t.amount = '.$productDiscount.' and t.is_percent = '.$isPercent;
                    $this->db->setQuery($productDiscountIdQuery);
                    $productDiscountId = $this->db->loadResult();

                }else {//nincs még ilyen discount, létrehozzuk

                    $discountClass = new stdClass();
                    $discountClass->amount = str_replace(',','.',$productDiscount);
                    $discountClass->is_percent = (int)$isPercent;
                    $discountClass->start_date = time();
                    $discountClass->end_date = 0;
                    //var_dump($discountClass);
                    if (!$this->db->insertObject("#__".$this->vm_prefix."_product_discount",$discountClass) ) {
                            throw new Exception( $db->stderr() );
                    }//end if

                    $productDiscountId = $this->db->insertId();

                }//end if
                //discount cuccok vége

                return $productDiscountId;
            }else {
                return null;
            }
        }//end manageDiscount()
        
        protected function manageManufacturer($productId) {
            
            $productManufacturerClass = new stdClass();
            $productManufacturerClass->product_id = $productId;
            $productManufacturerClass->manufacturer_id = 1;

            if (!$this->db->insertObject("#__".$this->vm_prefix."_product_mf_xref",$productManufacturerClass) ) {
                        throw new Exception( $db->stderr() );
            }//end if
            
        }//end manageManufacturer()
        
        ---------------*/
        
        protected function loadProductIntoDb($product, $productId = null){
            
            $productClass = new stdClass();
	
            if(empty($productId) == false) {
                $productClass->virtuemart_product_id = $productId;
            }
			

			
			if (empty($product->published) === false) {
                $productClass->published = (bool)$product->published;
            }else {
                $productClass->published = 1;
            }
            
            $productClass->product_sku = (string)$product->product_sku;
            
                       
            $productClass->virtuemart_vendor_id = 1;
			
			if (empty($product->product_weight) === false) {
                $productClass->product_weight      = (float)$product->product_weight;
            }	
			 
            /*
            if (empty($product->product_in_stock) === false) {
                $productClass->product_in_stock = (int)$product->product_in_stock;
            }
            
            if (empty($product->product_thumb_image) === false) {
                $productClass->product_thumb_image = (string)$product->product_thumb_image;
            }
            
            if (empty($product->product_full_image) === false) {
                $productClass->product_full_image  = (string)$product->product_full_image;
            }
            
            if (empty($product->product_special) === false) {
                $productClass->product_special     = (string)$product->product_special;
            }
            
            if (empty($product->product_discount_id) === false) {
                $productClass->product_discount_id = $discountId;
            }else {
                $productClass->product_discount_id = 0;
            }
           
            if (empty($product->product_tax_id) === false) {
                $productClass->product_tax_id      = (int)$product->product_tax_id;
            }
            
            
            
            if (empty($product->product_weight) === false) {
                $productClass->product_weight_uom      = (string)$product->product_weight_uom;
            }
            
            if (empty($product->product_length) === false) {
                $productClass->product_length      = (float)$product->product_length;
            }
            
            if (empty($product->product_width) === false) {
                $productClass->product_width      = (float)$product->product_width;
            }
            
            if (empty($product->product_height) === false) {
                $productClass->product_height      = (float)$product->product_height;
            }
            
            if (empty($product->product_lwh_uom) === false) {
                $productClass->product_lwh_uom      = (string)$product->product_lwh_uom;
            }
            */

//////////////////////////////////////////////////////////////////////////////////////////////////////
						file_put_contents('log.txt',var_export($productClass,true), FILE_APPEND);
						file_put_contents('log.txt',"\n", FILE_APPEND);
//////////////////////////////////////////////////////////////////////////////////////////////////////
           
            
            if (empty ($productId) === false) {
                if ($this->db->updateObject("#__".$this->vm_prefix."_products",$productClass, 'virtuemart_product_id') ) {


                }//end if
				else {
					throw new Exception( $this->db->stderr() );
				}
                $this->statistics->updated++;
            }else {
                if ($this->db->insertObject("#__".$this->vm_prefix."_products",$productClass) ) {

                    
                }//end if
				else {
					throw new Exception( $this->db->stderr() );
				}
                $this->statistics->inserted++;        
                $productId = $this->db->insertId();
            }//end if
            

            return $productId;
        }//end loadProductIntoDb
        
		protected function loadProductIntoDbRuRU($product, $productId = null){

			$productRuRuClass = new stdClass();

            if(empty($productId) == false) {
                $productRuRuClass->virtuemart_product_id = $productId;
            }

			if (empty($product->product_name) === false) {
                $productRuRuClass->product_name        = (string)$product->product_name;
            }
            
            if (empty($product->product_s_desc) === false) {
                $productRuRuClass->product_s_desc      = (string)$product->product_s_desc;
            }
            
            if (empty($product->product_desc) === false) {
                $productRuRuClass->product_desc        = (string)$product->product_desc;
            }
			
			if (empty($product->slug) === false) {
				$pN = (string)$product->slug;
				$pN = translitIt($pN);
				//$pN = mb_convert_case($pN, MB_CASE_TITLE, 'UTF-8');
                $productRuRuClass->slug        = $pN;
            }
			else
			{
				$pN = (string)$product->product_name;
				$pN = translitIt($pN);
				//$pN = mb_convert_case($pN, MB_CASE_TITLE, 'UTF-8');
				$productRuRuClass->slug		= $pN;
			}
				
			
//////////////////////////////////////////////////////////////////////////////////////////////////////
						file_put_contents('log.txt',var_export($productRuRuClass,true), FILE_APPEND);
						file_put_contents('log.txt',"\n\n", FILE_APPEND);
//////////////////////////////////////////////////////////////////////////////////////////////////////

            if (empty ($productId) === true) {
                if (!$this->db->updateObject("#__".$this->vm_prefix."_products_ru_ru",$productRuRuClass, 'virtuemart_product_id')) {
					throw new Exception( $this->db->stderr() );
				}
            }else {
                if (!$this->db->insertObject("#__".$this->vm_prefix."_products_ru_ru",$productRuRuClass)) {
					throw new Exception( $this->db->stderr() );
                     
                }//end if
            }//end if

        }//end loadProductIntoDbRuRU
		
        
    }//end class
    
    
    
    /*				
//---------------------------
				$dbc = new mysqli('localhost', 'root', '', 'vmart_db');

				if (mysqli_connect_errno()) {
					printf("Connect failed: %s\n", mysqli_connect_error());
					exit();
				}

				$my_query = "INSERT INTO vmart_virtuemart_products_ru_ru (virtuemart_product_id, product_name, product_s_desc,product_desc,slug)".
																							"VALUES ('$productRuRuClass->virtuemart_product_id',".
																							"'$productRuRuClass->product_name',".
																							"'$productRuRuClass->product_s_desc',".
																							"'$productRuRuClass->product_desc',".
																							"'$productRuRuClass->slug')";
				$result = $dbc->query($my_query);

				$dbc->close();

//---------------------------
*/
?>
