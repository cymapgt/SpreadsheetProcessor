<?php
namespace cymapgt\core\application\spreadsheet\SpreadsheetProcessor;

use cymapgt\Exception\SpreadsheetProcessorFactoryException;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelTmpFilePath;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelWrapperType;

/**
* SpreadsheetProcessorFactory
* This class is a singleton factory. It creates file handlers to the excel
* workbook object and to csv, and also wraps important static methods from
* the PHPExcel library e.g caching
*
* @category	Spreadsheet
* @package	cymapgt.core.application.spreadsheet
* @license   http://www.opensource.org/licenses/bsd-license.php
* @copyright    Copyright (c) 2014 - cymap
*/
class SpreadsheetProcessorFactory
{
    private static $cacheArray         = array();   //store cache methods
    //private static $cacheMethod      = "Memcache";  
    //private static $cacheConfigArray = null;
    private static $rwFileTypes        = array();   //store file types
    private static $valueBinder        = 'default'; //setting for value binder to use  by generated file handles
    private static $valueBinders       = array();   //store available value binders here

    //Hold an instance of the class
    private static $instance;
    
    //Sandobx mode
    private static $sandboxMode;
    
    //Base directory
    private static $baseDir;

    //A private constructor
    private function __construct() {
    }

    /**
     * Returns the singleton instance, or creates it if not existing
     * 
     * @return object
     * 
     * @access public
     */
    public static function getInstance() {
        if (!isset(self::$instance)) {
            $class = __CLASS__;
            self::$instance = new $class;
        }
        return self::$instance;
    }

    /**
     * Returns the singleton instance, or creates it if not existing
     * 
     * @access public
     */
    public function __clone() {
        throw new SpreadsheetProcessorFactoryException('Class '.__CLASS__.' is singleton');
    }

    /**
     * Initialize the factory with common PHPExcel settings
     * 
     * @param bool $sandboxMode - Set sandbox mode 0 is local, 1 is remote
     * 
     * @access public
     */    
    public static function initialize($sandboxMode = 0) {
        //set Sandbox mode
        self::$sandboxMode = $sandboxMode;
        
        //set Base Dir
        self::$baseDir = dirname(__FILE__) . '/';
        
        //initialize CachedObjectStorageFactory and array to store cache methods
        $cacheObj = '\PHPExcel_CachedObjectStorageFactory';

        /**
        *1) CACHING
        */	
        //assign the values to array dictionary
        self::$cacheArray['Memory']           = $cacheObj::cache_in_memory;
        self::$cacheArray['MemoryGZip']       = $cacheObj::cache_in_memory_gzip;
        self::$cacheArray['MemorySerialized'] = $cacheObj::cache_in_memory_serialized;
        self::$cacheArray['Igbinary']         = $cacheObj::cache_igbinary;
        self::$cacheArray['DiscISAM']         = $cacheObj::cache_to_discISAM;
        self::$cacheArray['APC']              = $cacheObj::cache_to_apc;
        self::$cacheArray['Memcache']         = $cacheObj::cache_to_memcache;
        self::$cacheArray['PHPTemp']          = $cacheObj::cache_to_phpTemp;
        self::$cacheArray['Wincache']         = $cacheObj::cache_to_wincache;
        self::$cacheArray['SQLite']           = $cacheObj::cache_to_sqlite;
        self::$cacheArray['SQLite3']          = $cacheObj::cache_to_sqlite3;

        /**
        *2) READING & WRITING FILES
        *   - We already have classes with lovely functionality for handling PHPExcel's
        *     complex properties
        *   - The lovely PHPExcelWrapper class by Zeriph - http://www.codeplex.com/site/users/view/zeriph
        *     will handle manipulation of Excel 2007, Excel 5 and CSV files. It is located in
        *     PHPExcelWrapper.php file, which we have included here
        *   - We extend this class with CYMAPGT_PHPExcelWrapper to enable handling of PDF, SWF,
        *     and HTML output. This class is included in CYMAPGT_PHPExcelWrapper.php, included here
        *     Again, as a factory, we return instances of objects, or make use of their static methods
        */
        static $rwFileTypes = array();

        //rwFileTypes array contains list of possible file types
        self::$rwFileTypes['Excel5'] = array('WrapperType' => PHPExcelWrapperType::Excel5,
            'FilePath' => PHPExcelTmpFilePath::TMP_EXCEL,
            'FileExt'  => 'xls'
        );
        self::$rwFileTypes['Excel2007'] = array('WrapperType' =>PHPExcelWrapperType::Excel2007,
            'FilePath'    => PHPExcelTmpFilePath::TMP_EXCEL,
            'FileExt'     => 'xlsx'		                                        
        );
        self::$rwFileTypes['CSV']       = array('WrapperType' => PHPExcelWrapperType::CSV,
            'FilePath'    => PHPExcelTmpFilePath::TMP_TEXT,
            'FileExt'     => 'csv'
        );
        self::$rwFileTypes['HTML']      = array('WrapperType' => PHPExcelWrapperType::HTML,
            'FilePath'    => PHPExcelTmpFilePath::TMP_HTML,
            'FileExt'     => 'html'
        );
        self::$rwFileTypes['PDF']       = array('WrapperType' => PHPExcelWrapperType::HTML,
            'FilePath'    => PHPExcelTmpFilePath::TMP_PDF,
            'FileExt'     => 'pdf'
        );
        self::$rwFileTypes['SWF']       = array("WrapperType" => PHPExcelWrapperType::HTML,
            'FilePath'    => PHPExcelTmpFilePath::TMP_SWF,
            'FileExt'     => 'swf'
        );

        //initialize the value binders										  
        self::$valueBinders['default']  = '\PHPExcel_Cell_DefaultValueBinder';
        self::$valueBinders['advanced'] = '\PHPExcel_Cell_AdvancedValueBinder';
    }
    
    /**
     * function setBaseDir() - Set the directory from which files are accessible
     *                         from PHP. In sandbox mode, must use the /files
     *                         sandbox folder set for this class
     * 
     * @param string $dir - Directory where to store
     * 
     * @access publc
     * @static
     */
    public static function setBaseDir($dir) {
        //should be writable
        if(!file_exists($dir)
            || !is_writeable($dir)
        ) {
            throw new SpreadsheetProcessorFactoryException('Base directory must be writeable');
        }
        
        self::$baseDir = $dir;
    }

    /**
     * function setCacheMethod() = This function sets the default cache method for use
     * by the CYMAPGT_PHPExcelWrapper class
     * TODO: provide config.inc or ini file for advanced configuration of cache methods
     * Cyril Ogana - 2012-07-03
     * 
     * @param  string $cacheMethod       The cache method, which should be one of the indices
     *                                   of $cacheArray
     * @param  string $cacheConfigArray  The cache method settings. default is null
     * @return void
     * @access public
     */
    public function setCacheMethod($cacheMethod = 'Memory', $cacheConfigArray = null) {
        //Get the const, which PHPExcel understands
        $cacheMethodConst = self::$cacheArray[$cacheMethod];
        return \PHPExcel_Settings::setCacheStorageMethod($cacheMethodConst);	
    }

    /**
     * function getCacheMethod() = This function gets the current cache method
     * Cyril Ogana - 2012-07-03
     * 
     * @param  int    $type       the type, using the ..CACHEGET..constants defined in this file
     * @return mixed              string or false
     * @access public
     */	
    public function getCacheMethod($type = \SPREADSHEETPROCESSOR_CACHEGETCUR) {
        switch ($type) {
            case \SPREADSHEETPROCESSOR_CACHEGETALL:
                return \PHPExcel_Settings::getAllCacheStorageMethods();
            case \SPREADSHEETPROCESSOR_CACHEGETCUR:
                return \PHPExcel_Settings::getCacheStorageMethod();
            case \SPREADSHEETPROCESSOR_CACHEGETAVL:
                return \PHPExcel_Settings::getCacheStorageMethods();
            default:
                return false;
        }	    
    }

    /**
     * function resetCacheMethod() = This function resets the cache method to default of Memory
     * Cyril Ogana - 2012-07-03
     *
     * @return void
     * @access public
     */	
    public function resetCacheMethod() {
        self::setCacheMethod();
    }

    /**
    * function createFile() = This function returns a handler to one of the stream wrapper
    *                             objects either PHPExcelWrapper or its child CYMAPGT_PHPExcelWrapper
    *                             which is primarily used to write to PDF, HTML and SWF
    * Cyril Ogana - 2012-07-03
    * @param  string $fileName    This is the file name of the file to create
    * @param  bool   $spreadsheet This is flag to indicate whether it is a spreadsheet type or other
    *                             type. Spreadsheet is xlsx, xls or csv (csv hypothetically)
    * @param  string $type        This is the actual file type as per $rwFileTypes array configuration
    * @parma  bool   $ovrFlag     If true, any file with the given name will be overwritten
    * @return object PHPExcelWrapper or CYMAPGT_PHPExcelWrapper
    * @access public
    */	
    public function createFile($fileName = 'tmp' , $spreadsheet = true, $type = 'CSV', $ovrFlag = false) {
        //we only will allow creating Excel 2007 files or CSV for now .. CO 20130702
        if($spreadsheet && self::$rwFileTypes[$type]['WrapperType'] <= 2) {
            $fileDir  = self::$baseDir.self::$rwFileTypes['Excel2007']['FilePath'];
            //$fileDir  = realpath($fileDir);
            $fileExt  = self::$rwFileTypes['Excel2007']['FileExt'];
            $fileType = self::$rwFileTypes['Excel2007']['WrapperType']; 
        } else {
            $fileDir  = self::$baseDir.self::$rwFileTypes['CSV']['FilePath'];
            //$fileDir  = realpath($fileDir);
            $fileExt  = self::$rwFileTypes['CSV']['FileExt'];
            $fileType = self::$rwFileTypes['CSV']['WrapperType'];
        }
        //die($fileDir);        
        return new SpreadsheetProcessor($fileName, $fileDir, $fileExt, $fileType, $ovrFlag);
    }

    /**
    * function openFile() = This function returns a handler to one of the stream wrapper
    *                             objects either PHPExcelWrapper or its child CYMAPGT_PHPExcelWrapper
    *                             which is primarily used to write to PDF, HTML and SWF
    * Cyril Ogana - 2012-07-03
    * @param  string $fileName    This is the file name of the file to create
    * @param  string $fileDir     This is the filedir. Or default of one of the object types
    * @return object PHPExcelWrapper or CYMAPGT_PHPExcelWrapper
    * @access public
    */
    public function openFile($fileName, $fileDir = 'Excel2007') {
        if ($fileDir == 'Excel2007') {
            $fileTypeIndex = 'Excel2007';
        } elseif($fileDir == 'Excel5') {
            $fileTypeIndex = 'Excel5';
        } elseif($fileDir == 'CSV') {
            $fileTypeIndex = 'CSV';
        } else {/*Do nothing*/}

        //$fileDir       = self::$baseDir.self::$rwFileTypes[($fileTypeIndex)]["FilePath"];
	$fileDir       = self::$baseDir.self::$rwFileTypes[($fileTypeIndex)]['FilePath'];
        //$fileDir       = realpath($fileDir);
        $fileExt       = self::$rwFileTypes[($fileTypeIndex)]['FileExt'];
        $fileType      = self::$rwFileTypes[($fileTypeIndex)]['WrapperType'];
        $ovrFlag       = true;
        $fileNameQual  = $fileName;
        $fileNameFqual = $fileDir . '/' . $fileName;   //make the fileName fully qualified
        $fileNameFqext = $fileNameFqual. '.' . $fileExt;

        //Open a reader
        /*$inputFileType = \PHPExcel_IOFactory::identify($fileNameFqext);
        $objReader = \PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($fileNameFqext);
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel,$fileTypeIndex);*/
        $phpExcelObj = new SpreadsheetProcessor($fileNameQual, $fileDir, $fileExt, $fileType,$ovrFlag);
        /*$phpExcelObj->setXlObj($objPHPExcel,0);
		$copiedFileName = $phpExcelObj->fileName;
		$phpExcelObj->fileName = $fileNameFqext;
		//unlink($copiedFileName);
        $phpExcelObj->setXlObj($objReader,1);
        $phpExcelObj->setXlObj($objWriter,2);*/
        $phpExcelObj->Save();

        return $phpExcelObj;   
    }

   /**
     *3) FORMATTING AND LAYOUT
     *   - This section will involve handling calls to formatting methods such as setting of the
     *     cell value binder
     */

     /**
     * function addValueBinder() = The value binder static array can be updated with new value binder
     *                             class if it was not picked up during initialization, or is required to
     *                             be loaded at runtime
     * Cyril Ogana - 2012-07-17
    * @param  string $binderName  This is the name of the value binder to register
     * @return void
     * @access public
     */	
    public function addValueBinder($binderName, $binderClassName) {
        //TODO: VALIDATE IS VALID BINDER NAME AND CLASSNAME (THOROUGH))...Cyril
        if(!is_string($binderName)){
            return false;
        }

        if(!is_string($binderClassName)){
            return false;
        }

        if (!array_key_exists($binderName, self::$valueBinders)                 //if the array key exists, exit
            || (array_search($binderClassName,self::$valueBinders)===false) //if binder class exists, exit
        ) {
            return false;
        }
        
        self::$valueBinders['$binderName'] = $binderClassName;
        return true;
    }

    /**
    * function setValueBinder() = Set the 'default' value binder to be used by objects generated
    *                             whose file handler was instantiated by this factory class
    * Cyril Ogana - 2012-07-17
    * @param  string $binderName  This is the name of the value binder to activate
    * @return void
    * @access public
    */	
    public function setValueBinder($binderName = 'default') {
        if(!is_string($binderName) || !array_key_exists($binderName, self::$valueBinders)) {
            return false;
        }

        \PHPExcel_Cell::setValueBinder( new self::$valueBinders[$binderName]());
        self::$valueBinder = $binderName;      //append binder name to static property
        return true;
    }

    /**
    * function getValueBinder() = Get the current value binder set in static valueBinder parameter
    * Cyril Ogana - 2012-07-17
    * @return object  Return a PHPExcel_ValueBinder object
    * @access public
    */	
    public function getValueBinder(){
        self::setValueBinder(self::$valueBinder);
        return self::$valueBinder;	
    }
}
