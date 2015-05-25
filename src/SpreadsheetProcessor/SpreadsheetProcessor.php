<?php
namespace cymapgt\core\application\spreadsheet\SpreadsheetProcessor;

use cymapgt\Exception\SpreadsheetProcessorException;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelWrapperType;


/**
 * @category   CYMAPGTReporting
 * @package    cymapgt.core.application.spreadsheet
 * @copyright  Copyright (c) 2014 CYMAP
 * @license   http://www.opensource.org/licenses/bsd-license.php
 * @version    2.0.0, 2014-05-05
 */

/**
 * PHPExcel_CachedObjectStorageFactory
 * - This class encapsulates a number of key PHPExcel methods in order to make
 *   creation and manipulation of files less verbose
 * - It also provides an organized structure of managing temporary files on disk. 
 * - Extends PHPExcel capability by adding a SWF filetype to the stack of file formats
 *
 * @category	CYMAPGTReporting
 * @package		CYMAPGT_PHPExcelWrapper
 * @copyright  Copyright (c) 2012 CYMAP-GT
 */ 
class SpreadsheetProcessor extends PHPExcelWrapper
{
    /**
    * function createSheet()  =   This function allows one to add a new worksheet on the PHPExcel Object

    * Cyril Ogana - 2013-02-06
    * @param  int/null  $index    Numeric index to where you want to add the worksheet
    * @return bool      True on success
    * @access public
    */
    public function newSheet($index = null) {
        //fetch the excel object
        $excelObj = $this->phpXl;

        try {
            $excelObj->createSheet($index);
            return true;			 
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    } 

    /**
    * function addSheet()  =   Append a worksheet object at a position in the worksheet
    *
    * Cyril Ogana - 2013-02-06
    * @param  object    $wSheetObj  The worksheet object to be appended to the worksheet
    * @param  bool      $external   Whether the worksheet is internal or external
    * @param  int/null  $index      Numeric index to where you want to add the worksheet
    * @return bool      True on success
    * @access public
    */
    public function addSheet(\PHPExcel_Worksheet $wSheetObj, $external = false, $index = null) {
        //fetch the excel object
        $excelObj = $this->phpXl;

        try {
            if ($external) {
                $excelObj->addExternalSheet($wSheetObj, $index);
            } else {
                $excelObj->addSheet($wSheetObj, $index);
            }
            
            return true;			 
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }

    /**
    * function getActiveWorksheet()  =  Get the current active worksheet
    *
    * Cyril Ogana - 2013-02-06
    * @return object                 Return PHPExcel worksheet object
    * @access public
    */
    public function getActiveSheet() {
        //fetch the excel object
        $excelObj = $this->phpXl;
        return $excelObj->getActiveSheet();
    }

    /**
    * function setActiveSheet() - Set the active sheet
    *
    * Cyril Ogana - 2013-02-06
    * @param  int   index    Get the index as integer or name
    * @parm   bool  byName   If true, run getSheetByName else run get by index
    * @return bool           True on success, false on failure
    * @access public
    */
    public function setActiveSheet($index = 0, $byName = false) {
        //fetch the excel object
        $excelObj = $this->phpXl;

        try {
            if ($byName) {
                    return $excelObj->setActiveSheetIndexByName($index);
                } else {
                    return $excelObj->setActiveSheetIndex($index);
                }
        } catch (SpreadsheetProcessorException $exception) {
            return $exception->getMessage();
        }
    }
	 
	 
    /**
    * function getAllSheets()  =  Get all worksheets
    *
    * Cyril Ogana - 2013-02-06
    * @return array        Return array collection of the worksheets
    * @access public
    */
    public function getAllSheets() {
        $excelObj = $this->phpXl;
        return $excelObj->getAllSheets();
    }
	 
    /**
    * function getSheet()  =  Get a worksheet by name or its index
    *
    * Cyril Ogana - 2013-02-06
    * @param  int   index    Get the index as integer
    * @parm   bool  byName   If true, run getSheetByName else run get by index
    * @return object/faalse  Return the phpexcel object or false if exception
    * @access public
    */
    public function getSheet($index = 0, $byName = false) {
        //fetch the excel object
        $excelObj = $this->phpXl;

        try {
            if ($byName) {
                return $excelObj->getSheetByName($index);			 
            } else {
                return $excelObj->getSheet($index);			     
            }
        } catch (\PHPExcel_Exception $exception) {
            $exception->getMessage();
        }
    }
	 
    /**
    * function removeSheet() - This method removes worksheets from the file
    *
    * Cyril Ogana - 2014-05-04 - Fix the removing sheet by index number
    * @param  int   index    Get the index as integer or name
    * @parm   bool  byName   If true, run getSheetByName else run get by index
    * @return bool           True on success, false on failure
    * @access public
    */
    public function removeSheet($index = 0, $byName = false) {
        //fetch the excel object
        $excelObj = $this->phpXl;

        try {
            if ($byName) { 
                $excelObj->setActiveSheetIndexByName($index);
            } else {
                $excelObj->setActiveSheetIndex($index);
            }
            $activeSheetIndex = $excelObj->getActiveSheetIndex();
            $excelObj->removeSheetByIndex($activeSheetIndex);            
            return true;
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
    * function getSheetCount()
    *
    * Cyril Ogana - 2013-02-06
    * @return int  Number of worksheets in the workbook object
    * @access public
    */
    public function getSheetCount() {
        //fetch the excel object
        $excelObj = $this->phpXl;
        return $excelObj->getSheetCount();
    }
	 
	 
    /**
    * function setProperties()  = This function sets one or more properties for a spreadsheet. If the
    *                             filetype is not Excel5 or Excel2007, the function returns FALSE, else
    *                             returns TRUE if the file property was successfully set 
    *                             which is primarily used to write to PDF, HTML and SWF
    * Cyril Ogana - 2014-05-04    - Add array type hint
    * @param  array  $properties  This is an associative array, keys being the properties to be set
    * @return bool                True on success, false on failure
    * @access public
    */
    public function setProperties(array $properties = null) {
        $excelObj = $this->phpXl;
        //validate that if the property is custom, its index 
        if (isset($properties['custom']) && !is_array($properties['custom'])) {
            return false;		     	
        }

        //validate that the $fileObj is a subclass or child of PHPExcel
        if(!($excelObj instanceof \PHPExcel)) {
            return false;
        }

        $fileProperties = array(
            'category'       => 'setCategory',
            'company'        => 'setCompany',
            'created'        => 'setCreated',
            'creator'        => 'setCreator',
            'description'    => 'setDescription',
            'keywords'       => 'setKeywords',
            'lastmodifiedby' => 'setLastModifiedBy',
            'manager'        => 'setManager',
            'modified'       => 'setModified',
            'subject'        => 'setSubject',
            'title'          => 'setTitle',
            'custom'         => 'setCustomProperty'
        );

        //iterate the properties array and excecute the property setter method
        foreach ($properties as $property=>$value) {
           $func = $fileProperties[$property];		
           
           if (array_key_exists($property, $fileProperties) && $property=='custom') {
                $propertyName  = $value[0];
                $propertyValue = $value[1];
                $excelObj->getProperties()->{$func}($propertyName, $propertyValue);
           }else{
                $excelObj->getProperties()->{$func}($value);
           }
        }
        
        return true;
    }
	
    /**
	 * function setHyperlink()  = This function handles adding a hyperlink to a cell. If the
	 *                            filetype is not Excel5 or Excel2007, the function returns FALSE, else
	 * Cyril Ogana - 2012-07-03
         * @param  string $link      This is the link to be created
	 * @param  string $tooltip   This is the tooltip to the link
	 * @param  string $coordType This coordinate type takes "TEXT" or "NUMERIC" excel cell coords
	 * @param  string $cellCoord This if $coordType is "TEXT", $cellCoord is used to locate cell
	 * @param  int    $cellR     This if $coordType is "NUMERIC", $cellR is used to locate row
	 * @param  int    $cellC     This if $coordType is "NUMERIC", $cellC is used to locate column
	 * @return bool              True on success, false on failure
	 * @access public
	 */
	public function setHyperlink($link, $tooltip, $coordType, $cellCoord = '', $cellR = 1, $cellC = 0) {
            //coordType mustbe TEXT or NUMERIC	
            if (!($coordType == 'TEXT') && !($coordType == 'NUMERIC')) {
                return false;
            }

            //Get the excel object
            $excelObj = $this->phpXl;

            if ($coordType == 'NUMERIC') {
                //validate that the coords are integers
                if (
                    !is_int($cellR) 
                    && !is_int($cellC)
                    && !$cellR >= 1
                    && !$cellC >= 0
                ) {
                    return false;
                }
                
                $cellCoord = $this->GetExcelAlphanumericColumnRow($cellR, $cellC);
            }

            $excelObj->getActiveSheet()->getCell($cellCoord)->getHyperlink()->setUrl($link);
            $excelObj->getActiveSheet()->getCell($cellCoord)->getHyperlink()->setTooltip($tooltip);
            return true;
	}
	
    /**
    * function getHyperlink()  = This function fetches a hyperlink from a cell if already set
    * Cyril Ogana - 2012-07-03
    * @param  string $coordType This coordinate type takes "TEXT" or "NUMERIC" excel cell coords
    * @param  string $cellCoord This if $coordType is "TEXT", $cellCoord is used to locate cell
    * @param  int    $cellR     This if $coordType is "NUMERIC", $cellR is used to locate row
    * @param  int    $cellC     This if $coordType is "NUMERIC", $cellR is used to locate column
    * @return mixed             Return array of hyperlink + tooltip or false if fail
    * @access public
    */
    public function getHyperlink($coordType, $cellCoord = '', $cellR = 1, $cellC = 0) {
        //coordType mustbe TEXT or NUMERIC	
        if (!($coordType == 'TEXT') && !($coordType == 'NUMERIC')) {
            return false;
        }

        //Get the excel object
        $excelObj = $this->phpXl;

        if ($coordType == 'NUMERIC') {
            //validate that the coords are integers
            if (!is_int($cellR) 
                && !is_int($cellC)
                && !$cellR >= 1
                && !$cellC >= 0
            ) {
                return false;
            }
            
            $cellCoord = $this->GetExcelAlphanumericColumnRow($cellR, $cellC);
        }

        $link    = $excelObj->getActiveSheet()->getCell($cellCoord)->getHyperlink()->getUrl();
        $tooltip = $excelObj->getActiveSheet()->getCell($cellCoord)->getHyperlink()->getTooltip();
        $hyperlink = array('link' => $link, 'tooltip' => $tooltip);
        return $hyperlink;
    }

    /**
    * function setFirstPageNumber()  = This sets the first page number of default worksheet to an integer value
    * Cyril Ogana - 2014-05-02 - Add Exception handling
    * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise
    * @param  bool   $reset      If $reset is true, we need to reset page number to default 
    * @param  int    $pageNumber This is the integer value to which to set the page number
    * @return bool               True on success, false on failure
    * @access public
    */	
    public function setFirstPageNumber($wSheetKey = '' , $keyIsIndex = false, $reset = true, $pageNumber = 1) {
        try {       
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }

            if (!is_bool($reset)) {
                return false;
            }

            if (!is_int($pageNumber)) {
                return false;
            }

            if (!$reset) {
                $wSheet->getPageSetup()->setFirstPageNumber($pageNumber);
            } else {
                $wSheet->getPageSetup()->resetFirstPageNumber();
            }           
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }

        return true;
    }
	 
    /**
    * function getFirstPageNumber()  = Get the first page number of a wrk sheet
    * Cyril Ogana - 2012-07-03
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return integer             Return the first page number as int
    * @access public
    */
    public function getFirstPageNumber($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }
            return $wSheet->getPageSetup()->getFirstPageNumber();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }

    /**
    * function setFitToHeight()  = This fits a worksheet to height of the pages specified
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise
    * @param  int    $pValue     This is the number of pages across which to spread the height
    * @param  bool   $pUpdate    This is TRUE to fit the worksheet to page once height set
    * @return bool               True on success, false on failure
    * @access public
    */		 
    public function setFitToHeight($wSheetKey = '', $keyIsIndex = false, $pValue = 1, $pUpdate = true) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }

            $wSheet->getPageSetup()->setFitToHeight($pValue, $pUpdate);
            return true;
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	 
    /**
    * function getFitToHeight()  = Get whether fit to height setting for the wrk sheet
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return integer             return the value of fitToHeight
    * @access public
    */
    public function getFitToHeight($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }

            return $wSheet->getPageSetup()->getFitToHeight();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	 
    /**
    * function setFitToPage()  = This sets the fit to page number (to which #of pages print will adjust to)
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise
    * @param  int    $pValue     This is the number of pages across which to spread the height
    * @return bool               True on success, false on failure
    * @access public
    */		 
    public function setFitToPage($wSheetKey = '', $keyIsIndex = false, $pValue = 1) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }

            $wSheet->getPageSetup()->setFitToPage($pValue);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
    
    /**
    * function getFitToPage()  = Get the fit to page setting of a worksheet
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return integer             Return the fitToPage value
    * @access public
    */
    public function getFitToPage($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }

            return $wSheet->getPageSetup()->getFitToPage();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }


    /**
    * function setFitToWidth()  = This fits a worksheet to width of a page
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise
    * @param  int    $pValue     This is the number of pages across which to spread the width
    * @param  bool   $pUpdate    This is TRUE to fit the worksheet to page once width set
    * @return bool               True on success, false on failure
    * @access public
    */		 
    public function setFitToWidth($wSheetKey = '', $keyIsIndex = false, $pValue = 1, $pUpdate = true) {
        try {
            if (!$wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex)) {
                return false;
            }

            $wSheet->getPageSetup()->setFitToWidth($pValue, $pUpdate);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }

    }
	 
    /**
    * function getFitToWidth()  = Get whether fit to width setting for the wrk sheet
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return integer           return the value of fitToWidth
    * @access public
    */
    public function getFitToWidth($wSheetKey = '', $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            return $wSheet->getPageSetup()->getFitToWidth();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	 
    /**
    * function setHorizontalCentered()  = This ensures that pages are horizontal centered for printing
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise
    * @param  bool   $value      This is the number of pages across which to spread the height
    * @return bool               True on success, false on failure
    * @access public
    */			 
    public function setHorizontalCentered($wSheetKey = '', $keyIsIndex = false, $value = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            $wSheet->getPageSetup()->setHorizontalCentered($value);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }	 	
    }//TODO: Check whey setHorizontalCentered not passing unit tests

    /**
    * function getHorizontalCentered()  = Get horizontal centering configuration for a worksheet obj
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return bool                true if is on, false if off
    * @access public
    */
    public function getHorizontalCentered($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            return $wSheet->getPageSetup()->getHorizontalCentered();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
    * function setOrientation()  = This sets the page orientation for the worksheet for print
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  string $orientation Orientation is default, landscape or portrait
    * @return bool                True on success, false on failure
    * @access public
    */	 
    public function setOrientation($wSheetKey = '', $keyIsIndex = false, $orientation = 'default') {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }

            if (
                !$orientation == 'default'
                && !$orientation == 'landscape'
                && !$orientation == 'portrait'
            ) {
                return false;
            }//TODO: More validation on this wrappers setter methods

            $wSheet->getPageSetup()->setOrientation($orientation);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }	 	
    }

    /**
    * function getOrientation()  = Get the print orientation of the worksheet
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise	 
    * @return string              returns DEFAULT, LANDSCAPE or PORTRAIT
    * @access public
    */
    public function getOrientation($wSheetKey = '', $keyIsIndex = false) {
        if (!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex)) {
            return false;
        }

        return $wSheet->getPageSetup()->getOrientation();
    }

    /**
    * function setPaperSize()  = This sets the page orientation for the worksheet for print
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  string $pValue      This is the papersize value (OpenXML format ) taken from
    *                             Office Open XML Part 4 - Markup Language Reference, page 1988
    * @return bool                True on success, false on failure
    * @access public
    */	 
    public function setPaperSize($wSheetKey = '', $keyIsIndex = false, $pValue = 'LETTER') {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey, $keyIsIndex))) {
                return false;
            }
            
            $pConst = '\PHPExcel_Worksheet_PageSetup::PAPERSIZE_'.strtoupper(str_replace(' ','_',$pValue));

            if (!is_string($pValue) || !isset($pConst)) {
                return false;
            }

            $wSheet->getPageSetup()->setPapersize($pValue);
            return true;           
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	 
    /**
    * function getPaperSize()  = Get the paper size of the worksheet pages
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise 
    * @return int                 returns an integer representing one of the PAPERSIZE_*** constants
    * @access public
    */
    public function getPaperSize($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            return $wSheet->getPageSetup()->getPaperSize();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }

    /**
    * function setPrintArea()  = This sets a certain region as the print area using worksheet range coordinates
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  bool   $isAdd       
    * @param  string $cellCoord   The cell coordinates as excel range e.g A1:B10
    * @param  int    $cellR1      The start row coordinate as integer
    * @param  int    $cellC1      The start col coordinate as integer
    * @param  int    $cellR2      The end row coordinate as integer
    * @param  int    $cellC2      The end col coordinate as integer
    * @param  int    $index       Index argument
    * @param  string $method      O = Overwrite, I = Insert
    * @return bool                true/false
    * @access public
    */	 
    public function setPrintArea(
        $coordType,
        $wSheetKey = '',
        $keyIsIndex = false,
        $isAdd = false,
        $cellCoord = '',
        $cellR1 = 1,
        $cellC1 = 0, 
        $cellR2 = 1,
        $cellC2 = 0,
        $index = 0,
        $method = 'O'
    ) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_int($index)) {
                return false;
            }

            if(!is_bool($isAdd)){
                return false;
            }

            $paClassMethod = array();        //The page setup printarea methods to use (add, or set)

            if ($isAdd) {
                $paClassMethod[] = "addPrintAreaByColumnAndRow";
                $paClassMethod[] = "addPrintArea";
            } else {
                $paClassMethod[] = "setPrintAreaByColumnAndRow";
                $paClassMethod[] = "setPrintArea";
            }

            if (!($method == 'O') && !($method == 'I')) {
                return false;
            }

            //coordType mustbe TEXT or NUMERIC	
            if (!($coordType == 'TEXT') && !($coordType == 'NUMERIC')) {
                return false;
            }

            if ($coordType == 'NUMERIC') {
                //validate that the coords are integers
                if (
                    !is_int($cellR1)
                    && !is_int($cellR2)
                    && !is_int($cellC1)
                    && !is_int($cellC2)
                    && !$cellR1 >= 1
                    && !$cellR2 >= 1
                    && !$cellC1 >= 0
                    && !$cellC2 >= 0
                ) {
                    return false;
                }   

                $wSheet->getPageSetup()->{$paClassMethod[0]}($cellC1, $cellR1, $cellC2, $cellR2, $index, $method);
                /*TODO: We have taken advantage of php's variable variables but also one method takes one less
                  argument. PHP doesn't complain but this is bad programming practice. search 4 better solution*/
            } else {
                $wSheet->getPageSetup()->{$paClassMethod[1]}($cellCoord, $index, $method);
            }
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }			     
    }

    /**
    * function getPrintArea()  = Get the print area settings for the worksheet
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  int    $index       The index of print area is several exist
    * @param  bool   $oCheck      If true,only check if it exists and return bool, else return int
    * @return mixed               depending on oCheck, returns bool or returns a string containing the print ranges
    * @access public
    */
    public function getPrintArea($wSheetKey = '' , $keyIsIndex = false, $index = 0, $oCheck = false) {
        try {
            if(!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))){
                return false;
            }

            if (!is_int($index) && !($index >= 0)) {
                return false;
            }

            if (!is_bool($oCheck)) {
                return false;
            }		 

            if (!$oCheck) {
                return $wSheet->getPageSetup()->getPrintArea($index);		 	
            } else {
                return $wSheet->getPageSetup()->isPrintAreaSet($index);
            }           
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
    * function clearPrintArea()  = Clear a print area
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  int    $index       The index of print area is several exist. if 0, all will be cleared
    * @return obj                 returns pagesetup obj
    * @access public
    */
    public function clearPrintArea($wSheetKey = '' , $keyIsIndex = false, $index = 0) {
        try {
            if  (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_int($index) && !($index >= 0)) {
                return false;
            }

            return $wSheet->getPageSetup()->clearPrintArea($index);            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }		 	
    }
	
    /**
    * function setRepeatCols()  = Set columns to be repeated on each page of the report
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  string $repeatType  RepeatType = "ARRAY" or "STARTEND"
    * @param  array  $cellArray   Array whose first and second indices $a[0] and $a[1] give start & end
    *                             columns to be repeated
    * @param  string $cellC1      The start col coordinate as integer or string
    * @param  string $cellC2      The end col coordinate as integer or string
    * @return bool                True on success, false on failure
    * @access public
    */	 	 
    public function setRepeatCols($wSheetKey = '', $keyIsIndex = false, $repeatType = '', $cellArray = null, $cellC1 = '', $cellC2 = '') {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!($repeatType == 'ARRAY') && !($repeatType == 'STARTEND')) {
                return false;
            }

            if ($repeatType == 'ARRAY') {
                if (
                    !(is_array($cellArray))
                    || !(isset($cellArray[0]))
                    || !(is_int($cellArray[0])) 
                    || !($cellArray[0] >= 0)
                    || !(isset($cellArray[1]))
                    || !(is_int($cellArray[1]))
                    || !($cellArray[1] >= 1)
                ) {
                    return false;
                }
                $wSheet->getPageSetup()->setColumnsToRepeatAtLeft($cellArray);
            } else {
                if (!is_string($cellC1) && !is_string($cellC2)) { //TODO: Ensure these are only A-Z
                    return false;
                }

                $wSheet->getPageSetup()->setColumnsToRepeatAtLeftByStartAndEnd($cellC1, $cellC2);
            }
            return true;            
        } catch (Exception $exception) {
            return $exception->getMessage();
        }
    }
	 
    /**
    * function getRepeatCols()  = Get the columns to repeat settings (left & left by start and end)
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  bool   $oCheck      If true,only check if it exists and return bool, else return array 
    * @return array/bool          If oCheck is true returns bool else returns an array of the cols marked as repeatable
    * @access public
    */
    public function getRepeatCols($wSheetKey = '' , $keyIsIndex = false, $oCheck = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_bool($oCheck)) {
                return false;
            }

            if ($oCheck == false) {
                return $wSheet->getPageSetup()->getColumnsToRepeatAtLeft();		 	
            } else {
                return $wSheet->getPageSetup()->isColumnsToRepeatAtLeftSet();
            }            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
    * function setRepeatRows()  = Set rows to be repeated on each page of the report
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  string $repeatType  RepeatType = "ARRAY" or "STARTEND"
    * @param  array  $cellArray   Array whose first and second indices $a[0] and $a[1] give start & end
    *                             rows to be repeated
    * @param  string $cellR1      The start row coordinate as integer
    * @param  string $cellR2      The end row coordinate as integer
    * @return bool                True on success, false on failure
    * @access public
    */ 
    public function setRepeatRows($wSheetKey = '', $keyIsIndex = false, $repeatType = '', $cellArray = null, $cellR1 = 1, $cellR2 =  1) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!($repeatType == 'ARRAY') && !($repeatType == 'STARTEND')) {
                return false;
            }

            if ($repeatType == 'ARRAY') {
                if (
                    !is_array($cellArray)
                    || !isset($cellArray[0])
                    || !is_int($cellArray[0])
                    || !($cellArray[0] >= 0)
                    || !isset($cellArray[1])
                    || !is_int($cellArray[1])
                    || !($cellArray[1] >= 1)
                ) {
                    return false;
                }

                $wSheet->getPageSetup()->setRowsToRepeatAtTop($cellArray);
            } else {
                if (!is_int($cellR1)
                    && !is_int($cellR2)
                    && !$cellR1 >= 1
                    && !$cellR2 >= 1
                ) { //TODO: Ensure these are only A-Z
                    return false;
                }

                $wSheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd($cellR1, $cellR2);
            }

            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        } 	
    } 

    /**
    * function getRepeatRows()  = Get the rows to repeat settings (top & top by start and end)
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  bool   $oCheck      If true,only check if it exists and return bool, else return array 
    * @return array/bool          If oCheck is true, returns bool else returns an array of the rows marked as repeatable
    * @access public
    */
    public function getRepeatRows($wSheetKey = '' , $keyIsIndex = false, $oCheck = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_bool($oCheck)) {
                return false;
            }

            if ($oCheck == false) {
                return $wSheet->getPageSetup()->getRowsToRepeatAtTop();		 	
            } else {
                return $wSheet->getPageSetup()->isRowsToRepeatAtTopSet();
            }           
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
		 
    /**
    * function setScale()  = Print scaling. Valid values range from 10 to 400 This setting is overridden
    *                        when fitToWidth and/or fitToHeight are in use
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  int    $pValue      The scale value as integer
    * @param  bool   $pUpdate     Update flag, true or false
    * @return bool                True on success, false on failure
    * @access public
    */ 
    public function setScale($wSheetKey = '', $keyIsIndex = false, $pValue = 100, $pUpdate = false) {
        try {
            if (!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex)) {
                return false;
            }

            if (!is_int($pValue) && !($pValue >= 10) && !($pValue <= 100)) {
                return false;
            }

            if (!is_bool($pUpdate)) {
                return false;
            }

            $wSheet->getPageSetup()->setScale($pValue, $pUpdate);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }

    /**
    * function getScale()  = Get the scale set for setting worksheet content on a page
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return int                 Returns an integer value between 10 and 100
    * @access public
    */
    public function getScale($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }
            return $wSheet->getPageSetup()->getScale();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
    
    /**
     * function setVerticalCentered()  = This function sets the vertical centering of data on page
     * Cyril Ogana - 2012-07-19
     * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
     * @param  bool   $value       True to activate the centering , false to deactivate it
     * @return bool                True on success, false on failure
     * @access public
     */ 	 
    public function setVerticalCentered($wSheetKey = '', $keyIsIndex = false, $value = false) {
        try {
            if (!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex)) {
                return false;
            }

            if (!is_bool($value)) {
                return false;
            }

            $wSheet->getPageSetup()->setVerticalCentered($value);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }


     /**
    * function getVerticalCentered()  = Check if vertical centering of page content is set to true
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return bool                Returns a boolean true or false
    * @access public
    */
    public function getVerticalCentered($wSheetKey = '' , $keyIsIndex = false) {
        if (!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex)) {
            return false;
        }

        return $wSheet->getPageSetup()->getVerticalCentered();
    }

    /**
    * function getWorksheet()  =  Returns a PHPExcel_Worksheet object as per params passed
    * Cyril Ogana - 2012-07-18
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @return object or false     worksheet Obj on success, false on failure
    * @access public
    */	 
    public function getWorksheet($wSheetKey = '', $keyIsIndex = false) {
        if ($this->type == PHPExcelWrapperType::CSV) {        //if it is csv, can't get worksheet
            return false;
        }

        if (!isset($this->phpXl)) {                             //all the same, PhpXl might not be set
            return false;
        }

        if (!$wSheetKey) {
            return $this->phpXl->getActiveSheet();
        } else {
            if ($keyIsIndex && is_int($wSheetKey)) {
                return $this->phpXl->getSheet($wSheetKey);
            } elseif (!$keyIsIndex && is_string($wSheetKey)) {
                return $this->phpXl->getSheetByName($wSheetKey);
            } else {
                return false;
            }
        }
    }
	 
    /**
    * function setMargin()  =  Sets setting for one or several page margins
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  mixed  $margin      If 0, set all. If string, set specified. If array, set indicated
    * @param  double $pValue      This is the value of the margin. Default = 0.75 like in excel
    * @return obj
    * @access public
    */	 
    public function setMargin($wSheetKey = '' , $keyIsIndex = false, $margin = 'top', $pValue = 0.75) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            $marginArray = array (
                'top'    => 'setTop',
                'bottom' => 'setBottom',
                'right'  => 'setRight',
                'left'   => 'setLeft',
                'header' => 'setHeader',
                'footer' => 'setFooter'
            );

            if (is_string($margin)) {     //if margin is string, set only the particular part e.g. 'top'
                if (!is_float($pValue) && !is_int($pValue)) {
                    return false;
                }

                if (!isset($marginArray[$margin])) {
                    return false;
                }

                 $wSheet->getPageMargins()->{$marginArray[$margin]}($pValue);
                 return true;
            }

            if ($margin === 0) {          //if margin is 0, set all margins
                if (!is_array($pValue)) {
                    return false;
                }

                if (count(array_diff_key($marginArray, $pValue))) {
                    return false;
                }

                foreach ($pValue as $key => $value) {
                    if (!is_float($value) && !is_int($value)) {
                        return false;
                    }
                }

                foreach ($marginArray as $key => $value) {
                    $wSheet->getPageMargins()->{$value}($pValue[$key]);
                }

                return true;
            }

            if (is_array($margin)) {     //if marign is an array, set those in associative array e.g. $margin['top']
                if (count($margin) > 6) {
                    return false;
                }

                if (count(array_diff_key($margin, $pValue))) {
                    return false;
                }

                foreach ($pValue as $key => $value) {
                    if (!is_float($value) && !is_int($value)) {
                        return false;
                    }
                }

                foreach ($marginArray as $key => $value) {
                    if (array_key_exists($key, $margin)) {
                        $wSheet->getPageMargins()->{$value}($pValue[$key]);
                    }
                }

                return true;
            }

            return false;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	 
    /**
    * function getMargin()  =  Gets setting for one or several page margins
    * Cyril Ogana - 2012-07-19
    * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
    * @param  mixed  $margin      If 0, set all. If string, set specified. If array, set indicated
    * @return mixed               False if error, array or string if margin found
    * @access public
    */	 
    public function getMargin($wSheetKey = '' , $keyIsIndex = false, $margin = 'top') {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }	 

            $marginArray = array(
                'top'    => 'getTop',
                'bottom' => 'getBottom',
                'right'  => 'getRight',
                'left'   => 'getLeft',
                'header' => 'getHeader',
                'footer' => 'getFooter'
            );

            $resultArr = array();

            if (is_string($margin)) {
                if (!isset($marginArray[$margin])) {
                    return false;
                }

                return $wSheet->getPageMargins()->{$marginArray[$margin]}();
            }

            if ($margin === 0) {
                foreach ($marginArray as $key => $value) {
                    $resultArr[$key] = $wSheet->getPageMargins()->{$value}();
                }

                return $resultArr;
            }

            if (is_array($margin)) {
                if (count($margin) > 6) {
                    return false;
                }

                foreach ($marginArray as $key => $value) {
                    if (array_key_exists($key, $margin)) {
                        $resultArr[$key] = $wSheet->getPageMargins()->{$value}();
                    }
                }

                return $resultArr;
            }

            return false;            
        } catch (\PHPExcel_Exception $exception) {
            return $message->getMessage();
        }
    }	 

    /**
     * function setHFAlignWithMargins()
     * Set align for header/footer with margins
     *
     * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @param  bool    $pValue
     * @return bool
     * @return obj
     */
    public function setHFAlignWithMargins($wSheetKey = '' , $keyIsIndex = false, $pValue = true) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_bool($pValue)) {
                return false;
            }

            $wSheet->getHeaderFooter()->setAlignWithMargins($pValue);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
     * function getHFAlignWithMargins()
     * Get align for header/footer with margins
     *
     * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @return bool
     */
    public function getHFAlignWithMargins($wSheetKey = '' , $keyIsIndex = false) {
        try {
             if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            return $wSheet->getHeaderFooter()->getAlignWithMargins();           
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
    * function setHFDifferentFirst()
    * Toggle the flag to set different first page header on worksheet page
    * 
    * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise	 * @param  bool    $pValue
    * @return obj
    */
    public function setHFDifferentFirst($wSheetKey = '' , $keyIsIndex = false, $pValue = true) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_bool($pValue)) {
                return false;
            }

            $wSheet->getHeaderFooter()->setDifferentFirst($pValue);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
     * function getHFDifferentFirst()
     * Get flag for whether different first page section is set for header/footer is set
     *
     * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @return bool
     */
    public function getHFDifferentFirst($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            return $wSheet->getHeaderFooter()->getDifferentFirst();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }	
	

    /**
    * function setHFDifferentOddEven()
    * Toggle the flag to set different odd and even page numbers on worksheet pages
    * 
    * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
    * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise	 
    * @param  bool    $pValue
    * @return bool
    */
    public function setHFDifferentOddEven($wSheetKey = '' , $keyIsIndex = false, $pValue = true) {
        try {
            if ((!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_bool($pValue)) {
                return false;
            }
            $wSheet->getHeaderFooter()->setDifferentOddEven($pValue);
            return true;
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }

    /**
     * function getHFDifferentOddEven()
     * Get flag for whether different first odd & even page section is set for header/footer is set
     *
     * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @return bool
     */
    public function getHFDifferentOddEven($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex)) {
                return false;
            }

            return $wSheet->getHeaderFooter()->getDifferentOddEven();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }	

    /**
     * function setHFSections()  =  Sets headers and footes on the worksheet page
     * Cyril Ogana - 2012-07-19
     * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
     * @param  mixed  $section     If 0, set all. If string, set specified. If array, set indicated
     * @param  mixed  $pValue      This is the text to place in the section. Default = "" like in excel
     * @return bool
     * @access public
     */	 
    public function setHFSections($wSheetKey = '' , $keyIsIndex = false, $section = 'oddheader', $pValue = '') {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            $sectionArray = array(
                'evenfooter'  => 'setEvenFooter',
                'evenheader'  => 'setEvenHeader',
                'firstfooter' => 'setFirstFooter',
                'firstheader' => 'setFirstHeader',
                'oddfooter'   => 'setOddFooter',
                'oddheader'   => 'setOddHeader'
            );

            if (is_string($section)) {
                if (!isset($sectionArray[$section])) {
                    return false;
                }

                $wSheet->getHeaderFooter()->{$sectionArray[$section]}($pValue);
                return true;
            }

            if ($section === 0) {
                if (!is_array($pValue)) {
                    return false;
                }

                if (count(array_diff_key($sectionArray, $pValue))) {
                    return false;
                }

                foreach ($sectionArray as $key => $value) {
                    $wSheet->getHeaderFooter()->{$value}($pValue[$key]);
                }

                return true;
            }

            if (is_array($section)) {
                if (count($section) > 6) {
                    return false;
                }

                if (count(array_diff_key($section, $pValue))) {
                    return false;
                }

                foreach ($sectionArray as $key => $value) {
                    if (array_key_exists($key, $section)) {
                        $wSheet->getHeaderFooter()->{$value}($pValue[$key]);
                    }
                }

                return true;
            }
            return false;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	 
    /**
     * function getHFSections()  =  Gets headers and footer sections on the worksheet page
     * Cyril Ogana - 2012-07-19
     * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
     * @param  mixed  $section     If 0, set all. If string, set specified. If array, set indicated
     * @return mixed               string or bool
     * @access public
     */	 
    public function getHFSections($wSheetKey = '' , $keyIsIndex = false, $section = 'oddheader') {
        try {
            if (!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex)) {
                return false;
            }

            $sectionArray = array (
                'evenfooter'  => 'getEvenFooter',
                'evenheader'  => 'getEvenHeader',
                'firstfooter' => 'getFirstFooter',
                'firstheader' => 'getFirstHeader',
                'oddfooter'   => 'getOddFooter',
                'oddheader'   => 'getOddHeader'
            );

            $resultArr = array();

            if (is_string($section)) {
                if (!isset($sectionArray[$section])) {
                    return false;
                }

                return $wSheet->getHeaderFooter()->{$sectionArray[$section]}();
            }

            if ($section === 0) {
                foreach($sectionArray as $key=>$value){
                    $resultArr[$key] = $wSheet->getHeaderFooter()->{$value}();
                }

                return $resultArr;
            }

            if (is_array($section)) {
                if (count($section) > 6) {
                    return false;
                }

                foreach ($sectionArray as $key=>$value) {
                    if (array_key_exists($key, $section)) {
                        $resultArr[$key] = $wSheet->getHeaderFooter()->{$value}();
                    }
                }

                return $resultArr;
            }
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            $exception->getMessage();
        }
    }
	 	 
    /**
    * function setHFScaleWithDocument()
    * Toggle the flag to set scaling header and footer section with the document
    * on or off
    * @param  mixed   $wSheetKey   - numeric index of worksheet or tab name
    * @param  bool    $keyIsIndex  - true if $wSheetKey is to be index, false if wsheet name
    * @param  bool    $pValue      - boolean
    * @return bool
    */
    public function setHFScaleWithDocument($wSheetKey = '' , $keyIsIndex = false, $pValue = true) {
        try {
            if ((!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if (!is_bool($pValue)) {
                return false;
            }

            $wSheet->getHeaderFooter()->setScaleWithDocument($pValue);
            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }	

    /**
     * function getHFScaleWithDocument()
     * Get flag for whether scale with document is set for header/footer
     *
     * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @return bool
     */
    public function getHFScaleWithDocument($wSheetKey = '' , $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            return $wSheet->getHeaderFooter()->getScaleWithDocument();            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
     * function addDrawingObject()
     * Take in a name and path for a drawing object (image), add it and attach it to a worksheet
     *
     * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @param  string $objPath    This is the filepath and filename to the drawing object being added
     * @param  string $objName    The name of the drawing object
     * @return mixed			  return array of object name and PHPExcel_Worksheet_Drawing object or false
     */
    public function addDrawingObject($wSheetKey = '' , $keyIsIndex = false, $objPath = '', $objName = 'imageObj') {
        try {
            $objDrawingArr = array();       /*if drawing object created successfully, this array() associatively
                                            contain the 'name' i.e object name and the actual drawing object*/

            if ((!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            //get the drawing collection and check if imageObj has been set
            $drawingArray = $wSheet->getDrawingCollection();

            if ($objName == 'imageObj') {          //if using default, generate next name e.g. imageObj, imageObj1, imageObj2
                $objName = $this->getNextDefaultObjName($drawingArray, 'imageObj', 'getName');

                if ($objName == false) {
                    return false;
                }
            }

            $objDrawing = new \PHPExcel_Worksheet_Drawing();
            $objDrawing->setName($objName);      //Create a worksheet drawing object

            $objDrawing->setPath($objPath);	 //TODO: USE data service to register correct location. Also see parameter default on function header

            //populate $objDrawingArr and return results
            $objDrawingArr['object'] = $objDrawing->setWorksheet($wSheet);
            $objDrawingArr['name']   = $objName;
            return $objDrawingArr;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }

    /**
     * function getDrawingObjects()
     * Return the entire array of drawing objects, or just one object
     *
     * @param  string $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @param  string $objName    The name of the drawing object. If blank, return all objects
     * @return mixed			  Return array of drawing objects, or just one drawing object, or false if name of
     *                            the object you specified was not found
     */
    public function getDrawingObjects($wSheetKey = '' , $keyIsIndex = false, $objName = '') {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            $drawingName     = '';                      //initialize this to empty string
            $drawingObjArray = array();                 /*associative array where objName is the index
                                                          and drawing object is the element
                                                                                                     */ 
            //get the drawing collection and check if imageObj has been set
            $drawingArray = $wSheet->getDrawingCollection();

            foreach ($drawingArray as $drawing) {
                $drawingName = $drawing->getName();     //get the drawing name

                if ($objName == '') {
                    $drawingObjArray[$drawingName] = $drawing;
                } else {
                    if ($drawingName == $objName) {
                        return $drawing;
                    }				
                }
            }

            if ($objName == '') {
                return $drawingObjArray;
            } else {
                return false;                           //if user gave wrong objName, return null
            }            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
     * function getNextDefaultObjName()
     * Generate default object name by appending numeric suffix if objects with default 
     * name alread exist in the object collection
     *
     * @param  array  $objArray             This is the array of objects we want to get default name
     * @param  string $objName              This is the name of the object
     * @param  string $objNameMethodCall    The method that must exist in an object instance of the objects
     *                                      contained in $objArray
     * @return mixed                        Return an available objName or false if error
     */	
    protected function getNextDefaultObjName($objArray, $objName, $objNameMethodCall) {
        $currentName = $objName;           //hold the current name as we iterate
        $compareName = $objName;
        $a           = 0;                  //counter for default objects added

        try {
            foreach ($objArray as $obj) {    //iterate $objArray and call the $objNameMethodCall
                $currentName = $obj->{$objNameMethodCall}();

                if($currentName == $compareName) {
                    $a++;
                    $currentName = $objName . $a;    //append suffix and continue checking if is unique
                    $compareName = $currentName;     
                }
            }
            return $currentName;          //Return the next available default name that's possible
        } catch (SpreadsheetProcessorException $objError){    //TODO: Beef up exception handling (can we log it?)
            return false;
        }
    }
	
    /**
     * function setStyle() 
     * Set variety of styles and layout on a cell or cell range
     * 
     * @param  string   $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string   $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @param  string   $col   	    The column of the cell to set the style
     * @param  string   $row   	    The row of the cell to set the style
     * @param  array	$style      We will only be supporting applyFromArray for now to
     *                              encourage batch procesing :)
     * @param  bool     $isCoordinate   True to indicate $col.$row is a cellrange
     * @return bool						True/false
     */
    public function setStyle($wSheetKey, $keyIsIndex, $col, $row, $style, $isCoordinate = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            $cellCoord = $this->GetCellCoord($col, $row, $isCoordinate);        
            $wSheet->getStyle($cellCoord)->applyFromArray($style);

            return true;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }

    /**
     * function getStyle()
     * Get the style object of a cell or cell range
     * 
     * @param  string   $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string   $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @param  string	$col   			The column of the cell to set the style
     * @param  string	$row   			The row of the cell to set the style
     * @param  bool	$isCoordinate   True to indicate $col.$row is a cellrange
     * @return mixed	Object or boolean false
     */
    public function getStyle($wSheetKey, $keyIsIndex, $col, $row, $isCoordinate = false) {
        try {
            if ((!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            $cellCoord = $this->GetCellCoord($col, $row, $isCoordinate);
            $styleObj = $wSheet->getStyle($cellCoord);

            return $styleObj;           
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
     * function setMergeCells()
     * Set variety of styles and layout on a cell or cell range
     * 
     * @param  string $coordType   This is the coordinate type, either "TEXT" or "NUMERIC"
     * @param  string $wSheetKey   This is the key of the worksheet, either associative or indexed
     * @param  string $keyIsIndex  This is true if numeric key is expected, false if otherwise
     * @param  string $cellCoord   The cell coordinates as excel range e.g A1:B10
     * @param  int    $cellR1      The start row coordinate as integer
     * @param  int    $cellC1      The start col coordinate as integer
     * @param  int    $cellR2      The end row coordinate as integer
     * @param  int    $cellC2      The end col coordinate as integer
     * @return bool                true/false
     * @access public
     */	 
    public function setMergeCells (
        $coordType,
        $wSheetKey = '',
        $keyIsIndex = false,
        $cellCoord = '',
        $cellR1 = 1,
        $cellC1 = 0, 
        $cellR2 = 1,
        $cellC2 = 0
    ) {
        try {
            if ((!$wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }

            if ($coordType == 'TEXT') {
                $wSheet->mergeCells($cellCoord);                   
            } elseif ($coordType == 'NUMERIC') {
                $wSheet->mergeCellsByColumnAndRow($cellC1, $cellR1, $cellC2, $cellR2);
            } elseif ($coordType == 'ARRAY') {
                $wSheet->setMergeCells($cellCoord);
            } else {
                return false;
            }            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
        
        return true;
    }
									 
    /**
     * function getMergeCells()
     * Get all merged cell ranges as an array
     * 
     * @param  string   $wSheetKey  This is the key of the worksheet, either associative or indexed
     * @param  string   $keyIsIndex This is true if numeric key is expected, false if otherwise* 
     * @return mixed	Array of merge coordinates or boolean false
     */
    public function getMergeCells($wSheetKey = '', $keyIsIndex = false) {
        try {
            if (!($wSheet = $this->getWorksheet($wSheetKey,$keyIsIndex))) {
                return false;
            }        

            $mergeArr = $wSheet->getMergeCells();

            return $mergeArr;            
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
	
    /**
     *  function getRangeCoordsFromArray()
     *  Get the coordinates for multiple ranges input as array. 
     *  TODO: - introduce range checker so that conflicting ranges in differnet array
     *          elements are either ignored/flagged or trigger an exception
     *  TODO: - add option of returning ranges instead of string coords
     *  TODO: - if the array passed in (coordsArray) is an associative array, add
     *          param option for caller to specify whether a named range should be
     *          returned
     * 
     *
     * @param  array coordsArray    - Array of numeric coordinates of ranges to adjust
     * @param  bool  adjustColIndex - If true, column values are reduced by value of 1
     *                                because column A in excel is 0 and not 1
     * @return mixed                - return string or false on error
     */
    public function getRangeCoordsFromArray($coordsArray, $adjustColIndex = false) {
        try {
            //return fals if it is not an array
            if (!is_array($coordsArray)) {
                return false; 	
            }

            //declare the coordinates string that we will return
            $coordString = '';

             //loop through each coords
            foreach ($coordsArray as $coord) {
                //$coord arrays should be associative arrays with indices "col" and "row"
                if (
                    !isset($coord['col'])
                    || !isset($coord['row'])
                    || $coord['col'] < 1 
                    || $coord['row'] < 1
                ) {
                    return false;
                }

                //adjust for excel col beginning from 0. But if col was already 0 its error
                if ($adjustColIndex) {
                    if ($coord['col'] == 0) {
                        return false;
                    } else {
                        --$coord['col'];
                    }
                }

                $coordStringChild = $this->GetExcelAlphanumericColumnRow($coord['col'], $coord['row']);				

                //$append the range to coordsstring
                if ($coordString == '') {
                    $coordString = $coordStringChild;		
                } else {
                    $coordString = $coordString . ':' . $coordStringChild;
                }
            }

            return $coordString;      
        } catch (\PHPExcel_Exception $exception) {
            return $exception->getMessage();
        }
    }
}
