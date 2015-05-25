<?php
namespace cymapgt\core\application\spreadsheet\SpreadsheetProcessor;

use cymapgt\Exception\SpreadsheetProcessorException;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelTmpFilePath;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelWrapperType;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelBorderStyle;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelBorderType;
use cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper\PHPExcelTextDirection;

/**
*
* PHPExcelWrapper class is a wrapper for the PHPExcel library
*
* The PHPExcelWrapper class aims to make reading/writing Excel files easier by
* creating an overall stream object. The PHPExcelWrapper class can be either
* a CSV or an Excel5/2007 type. There are no 'reading' or 'get' methods within
* this class becuase of the idea that it is much easier to read in data from
* a CSV than an Excel file. Thus we implement the AutoConvert static function
* which you can use to convert (if needed) an Excel file to a CSV file and
* read in the data via a file object and explode(',', $line)
* NOTE: if the Type is CSV, then a lot of functions in this class do not
*       do anything (e.g. autoFit() simply returns out)
*
* @license   http://www.opensource.org/licenses/bsd-license.php
*/
class PHPExcelWrapper
{
    //The underlying PHPExcel object
    protected $phpXl;
    
    //The underlying PHPExcelWriter interface
    protected $phpXlWriter;
	
    //The underlying PHPExcelReader interface
    protected $phpXlReader;
	
    //The underlying file handle (for CSV only)
    protected $handle;
	
    //(string) The current file name
    public $fileName;
	
    //(bool) Value indicating if the current stream is open
    public $isOpen;
	
    //(int) The current row of the Excel file the underlying stream is on
    public $currentRow;
	
    //(string) The type this object is (CSV/Excel5/Excel2007)
    public $type;

    /*
    * The overloaded constructor for the PHPExcelWrapper class
    *
    * @param 	string	$fileName	The file name to use (can be relative or absolute)
    * @param 	int		$type		(OPTIONAL) The type of wrapper to load (0 for Excel5 (default), 1 for Excel2007, 2 for CSV)
    */
    public function __construct (
        $fileName,
        $fileDir,                
        $fileExt = 'xlsx',
        $type    = PHPExcelWrapperType::Excel2007,
        $ovrFlag = false
    ) {
        $this->isOpen = false;
        $fileName     = self::getNewFileName($fileDir, $fileName, $fileExt,$ovrFlag);
        $this->open($fileName, $type);
    }

    /**
     * function setXlObj()        = This function sets the protected PhpXl property. We only use this
     *                              when we are injecting a file from objReader for editing 
     * Cyril Ogana - 2013-01-29
     * @param object $obj          This is the phpexcel object being opened for reading      
     */
     public function setXlObj($obj, $objType) { 	
        switch ($objType) {
            case 0:
                if (!($obj instanceof \PHPExcel)) {
                    throw new SpreadsheetProcessorException('
                        Object of wrong type passed to setXlObj. Type expected is Excel
                    ');
                }            
                $this->PhpXl = $obj;				
                break;
            case 1:
                if (!($obj instanceof \PHPExcel_Reader_Abstract)) {
                    throw new SpreadsheetProcessorException('
                        Object of wrong type passed to setXlObj. Type expected is Excel Reader
                    ');
                }
                $this->PhpXlReader = $obj;
                break;
            case 2:
                if (!($obj instanceof \PHPExcel_Writer_Abstract)) {
                    throw new SpreadsheetProcessorException('
                        Object of wrong type passed to setXlObj. Type expected is Excel Writer
                    ');
                }
                $this->PhpXlWriter = $obj;
                break;
        }
     }
         
    /**
    * Automatically convert a file type to another
    *
    * Automatically convert a file type to another. (From CSV to XLS/XLSX and back)
    * The purpose of this function is primarly to convert an XLS/XLSX file to a CSV
    * file for ease of reading data (just open a file handle, read line by line
    * and do an explode(',', $line)
    *
    * @param 	string  $fileToConvert                  The file name to convert. It doesn't have to have an extnesion as 
    *				
    * @param    string  $baseDir                        The base directory from the factory
    * 							PHPExcel can auto open in the proper format
    * @param	string	$newFileName                    (OPTIONAL) The name of the new file to save.
    *                                                   This is file name ONLY (no folder path or extension)
    *                                                   Default is tmp.
    * @param 	PHPExcelWrapperType	$typeTo		(OPTIONAL) The type to convert to, either Excel5, Excel2007 or CSV. 
    *													Default is Excel5.
    * @param 	bool					$deleteOldFile	(OPTIONAL) True to automatically delete the $fileToConvert file.
    *													Default is false.
    * @return	string					The new file name of the converted file
    */
    public static function autoConvert (
        $fileToConvert,
        $baseDir,
        $newFileName   = 'tmp',
        $typeTo        = PHPExcelWrapperType::Excel5,
        $deleteOldFile = false
    ) {
	// NOTE: Any saving/reading to the Excel2007 format needs php_zip.so or php_zip.dll to operate
	$writerType  = '';
        $newFileName = '';
        $ext         = '';
        $tmpDir      = '';
		
        switch ($typeTo) {
            case PHPExcelWrapperType::Excel5:
                $writerType = 'Excel5'; $ext = 'xls';
                $tmpDir     = PHPExcelTmpFilePath::TMP_EXCEL;
                break;
            case PHPExcelWrapperType::Excel2007:
                $writerType = 'Excel2007'; $ext = 'xlsx';
                $tmpDir     = PHPExcelTmpFilePath::TMP_EXCEL;
                break;
            case PHPExcelWrapperType::CSV:
                $writerType = 'CSV'; $ext = 'csv';
                $tmpDir     = PHPExcelTmpFilePath::TMP_TEXT;
                break;
            case PHPExcelWrapperType::PDF:
                $writerType = 'PDF'; $ext = 'pdf';
                $tmpDir     = PHPExcelTmpFilePath::TMP_PDF;
                break;
            case PHPExcelWrapperType::HTML:
                $writerType = 'HTML'; $ext = 'html';
                $tmpDir     = PHPExcelTmpFilePath::TMP_HTML;
                break;
        }
        
        $newFileName = self::getNewFileName(($baseDir.$tmpDir), $newFileName, $ext);

        $auto = PHPExcel_IOFactory::load($fileToConvert);
        $writer = PHPExcel_IOFactory::createWriter($auto, $writerType);
        $writer->save($newFileName);
        if ($deleteOldFile) {
            unlink($fileToConvert);
        }
        return $newFileName;
    }
	
    /**
     * Gets a new file name for a relevently named temp file
     *
     * This function will get a new file name based on the parameters passed in
     * If a file exists in the directory it will increment a counter and append
     * it between the file name and extension.
     * 
     * @param      string	$dir   			The directory to look at for a new file name
     * @param      string	$oldFileName	The old file name
     * @param      string	$ext 			The extension of the file
     * @param      bool     $ovrFlag        If true, we do not create file indexes, but overwrite
     * 
     * @returns	   string value of the new file name
     */
    public static function getNewFileName($dir, $oldFileName, $ext, $ovrFlag) {
        $idX = 0;
        if (substr($dir, (strlen($dir) - 1), 1) != '/') { 
            $dir .= '/';
        }

        if (substr($ext, 0, 1) != ".") { 
            $ext = ".".$ext;
        }

        if (!file_exists(($dir.$oldFileName.$ext)) || ($ovrFlag)) {
            return ($dir.$oldFileName.$ext);
        }

        do {
            $fullName = $dir.$oldFileName.'.'.($idX++).$ext;
        } while (file_exists($fullName));

        $idX--;

        return $dir.$oldFileName.'.'.$idX.$ext;
    }

    /**
     * Flushes out and saves any data and closes all underlying streams
     */
    public function close() {
        $this->flush();
        
        if ($this->type == PHPExcelWrapperType::CSV) {
            fclose($this->handle);
        } else {
            $this->phpXl->disconnectWorksheets();
            $this->phpXl->garbageCollect();
            $this->currentRow = 1;
        }
        
        $this->isOpen = false;
        unset($this->phpXl);
        unset($this->phpXlWriter);
        unset($this->handle);
    }
	
    /**
     * Flushes any data to the file (saves the file)
     */
    public function flush() {
        // CSV type doesn't need flush since it was open withw w+
        if ($this->type != PHPExcelWrapperType::CSV) {
            $this->phpXlWriter->setPHPExcel($this->phpXl);
            $this->phpXlWriter->setPreCalculateFormulas(false);
            $this->phpXlWriter->save($this->fileName);
        }
        return true;
    }
	
    /**
     * Get the underlying stream object
     *
     * @returns 	Either the file object if Type is CSV or the underlying PHPExcel object
     */
    public function getBaseStream() {
        if ($this->type == PHPExcelWrapperType::CSV) { 
            return $this->handle;
        }
        
        return $this->phpXl;
    }
	
    /**
     * Gets the column name from a number (e.g. 2='B', 27='AA', etc.)
     * 
     * @param	 int   $col   The column number to convert
     *
     * @returns	A string representation of the column number
     */
    public function getExcelAlphaColumn($col) {
        //disallow negative col number...Cyril Ogana 2014.04.30
        if (
            $col < 0
            || !(is_int($col))
        ) {
            throw new SpreadsheetProcessorException('Input for calculating Excel alpha column should be a positive integer');
        }
        
        $div = $col;
        $mod = 0;
        $name = '';

        while ($div > 0) {
            $mod = ($div - 1) % 26;
            $name = 
            $name = chr(65 + $mod).$name;
            $div = (int)(($div - $mod) / 26);
        }

        return $name;
    }

    /**
     * Gets the column number from column name (e.g. 'B'=2, 'AA'=27, etc.)
     * 
     * @param    string	$col    The column name to convert
     *
     * @returns	An integer value representation of the column name
     */
    public function getExcelColumnFromAlpha($col) {
        if (is_numeric($col)) {
            return $col;
        }

        $col = strtoupper($col);
        $len = strlen($col);
        $tot = 0;

        for ($i = 0; $i < $len; $i++) {
            $num = ord(substr($col, $i, 1)) - 64;
            $pow = pow(26, $i) - 1;
            $tot += ($num + $pow);
        }

        return $tot;
    }
	
    /**
     * Gets the Excel column name from a numeric column and row (e.g. 2 and 1 = 'B1', 27 and 2 = 'AA2', etc.)
     * 
     * @param		int		$col   The column to convert
     * @param		int		$row   The row
     *
     * @returns		A string represenation of the column name and row number
     */
    public function getExcelAlphanumericColumnRow($col, $row) {
        $name = $this->getExcelAlphaColumn($col);
        return $name.$row;
    }
	
    /**
     * Gets a string representation of a cell coordinate
     * 
     * @param      mixed	$col   The column name/number (can be either an int or string)
     * @param      int		$row   The row
     * @param      bool		$isCoordinate   (OPTIONAL) If this is false, the function expects
     * $col to be a numeric value. If this value is true
     * this funcition simply concatenates $col and $row
     * Default is false.
     *
     * @returns		A string representation of a cell coordinate
     */
    public function getCellCoord ($col, $row, $isCoordinate = false) {
        $cellCoord = 'A1';
        if ($isCoordinate) {
            $cellCoord = $col.$row;
        } else {
            if (is_numeric($col)) {
                $cellCoord = $this->getExcelAlphanumericColumnRow($col, $row);
            } else {
                throw new SpreadsheetProcessorException('Column ($col) must be a numeric value if $isCoordinate is false.');
            }
        }
        return $cellCoord;
    }
	
    /**
     * Open a file
     * 
     * @param      string				$fileName   The file to open
     * @param      PHPExcelWrapperType	$type   	(OPTIONAL) The type of file to open (CSV/Excel5/Excel2007). 
     *                                              Default is PHPExcelWrapperType::Excel5.
     */
    public function open($fileName, $type = PHPExcelWrapperType::Excel2007) {
        $this->fileName = $fileName;
        $this->type = $type;

        if ($this->isOpen) {
            $this->close();
        }

        $this->currentRow = 1; // Current row gets set to 1 (Excel is not 0 based)

        if ($this->type == PHPExcelWrapperType::CSV) {
            $this->handle = fopen($this->fileName, 'w+'); // Write/Read
        } else {
            if (file_exists($fileName)) {
                $inputFileType = \PHPExcel_IOFactory::identify($fileName);			
                $objReader = \PHPExcel_IOFactory::createReader($inputFileType);
                $objPHPExcel = $objReader->load($fileName);
                $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, $inputFileType);
                $this->phpXl = $objPHPExcel;
                $this->phpXlWriter = new \PHPExcel_Writer_Excel2007($this->phpXl);
                $this->phpXlReader = new \PHPExcel_Reader_Excel2007($this->phpXl);			
            } else {
                $this->phpXl = new \PHPExcel();
                $this->phpXl->setActiveSheetIndex(0);
                if ($this->type == PHPExcelWrapperType::Excel2007) {
                    $this->phpXlWriter = new \PHPExcel_Writer_Excel2007($this->phpXl);
                    $this->phpXlReader = new \PHPExcel_Reader_Excel2007($this->phpXl);
                } else {
                    $this->phpXlWriter = new \PHPExcel_Writer_Excel5($this->phpXl);
                    $this->phpXlReader = new \PHPExcel_Reader_Excel5($this->phpXl);
                }
            }
            $this->flush();
        }

        $this->isOpen = true;
    }
	
    /**
     * Sets the active worksheet
     * 
     * @param  int  $index   The worksheet number to set
     */
    public function setActiveWorksheet($index) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }
        $this->phpXl->setActiveSheetIndex($index);
        return true;
    }
	
    /**
     * Save the current data and writes it to disk
     */
    public function save() {
        return $this->flush();
    }
	
    /**
     * Set the columns in the Excel file to autofit the content
     * 
     * @param      int		$column	(OPTIONAL) The column to autofit. 
     *								Default is 0. (0 says all columns with content).
     */
    public function autoFit($column = 0) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return;
        }
		
        if ($column > 0) { // Set a specific column
            $colName = $this->getExcelAlphaColumn($column);
            $this->phpXl->getActiveSheet()->getColumnDimension($colName)->setAutoSize(true);
        } else { // Set ALL columns
            $lastColumn = $this->phpXl->getActiveSheet()->getHighestColumn(); // B
            $lastColumnIndex = PHPExcel_Cell::columnIndexFromString($lastColumn); // 2
            
            for ($i = 1; $i <= $lastColumnIndex; $i++) {
                $colName = $this->getExcelAlphaColumn($i);
                $this->phpXl->getActiveSheet()->getColumnDimension($colName)->setAutoSize(true);
            }
        }
        
        return true;
    }
	
    /**
     * Set the borders around cells in an Excel file
     * 
     * @param      int					$col   			(OPTIONAL) The column to set the borders around.
     * 													Default is 0. (0 says all columns)
     * @param      int					$row   			(OPTIONAL) The row to set the borders around.
     * 													Default is 0. (0 says all rows)
     * @param      PHPExcelBorderType	$borderSides   	(OPTIONAL) The sides to set the border on.
     * 													Default is PHPExcelBorderType::All.
     * @param      PHPExcelBorderStyle	$borderType   	(OPTIONAL) The border style to set.
     * 													Default is PHPExcelBorderStyle::BORDER_THIN.
     * @param      bool					$isCoordinate   (OPTIONAL) If this is false, the function expects
     *                                  			    $col to be a numeric value. If this value is true
     *			 										this funcition simply concatenates $col and $row
     *													Default is false.
     */
    public function setBorders (
        $col = 0,
        $row = 0,
        $borderSides = PHPExcelBorderType::All,
        $borderType = PHPExcelBorderStyle::BORDER_THIN,
        $isCoordinate = false
    ) {
        /*if ($this->type == PHPExcelWrapperType::CSV) { return; }
        $wholeRow = ($col == 0); $wholeCol = ($row == 0); 
        // getHighsetColumn returns letter (AZ), need to convert to num
        if($col == 0) { $col = $this->getExcelColumnFromAlpha($this->phpXl->getActiveSheet()->getHighestColumn()); }
        if ($row == 0) { $row = $this->phpXl->getActiveSheet()->getHighestRow(); }
        $cellCoordEnd = $this->getCellCoord($col, $row, $isCoordinate);
        $cellCoordStart = 'A1';
        if (!$wholeCol || !$wholeRow) { // Only fall in here if one of them is false (Which means don't do all cells)
                if ($wholeCol) { $cellCoordStart = $this->getCellCoord($col, 1, $isCoordinate); }
                if ($wholeRow) { $cellCoordStart = 'A'.$row; }
        }*/  //This code is broken. reverting to the $cellCoord method
        $style = array (
            'borders' => array (
                $borderSides => array (
                    'style' => $borderType
                )
            )
        );

        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }

        $cellCoord = $this->getCellCoord($col, $row, $isCoordinate);
        $this->phpXl->getActiveSheet()->getStyle($cellCoord)->applyFromArray($style);
        unset($style); // freeup the memory
        return true;
    }
	
    /**
     * Set the background color of a cell
     * 
     * @param      int		$col   			The column of the cell to set the back color
     * @param      int		$row   			The row of the cell to set the back color
     * @param      string	$rgb   			The HTML based RGB color (e.g. 'FF0000' is red)
     * @param      bool		$isCoordinate   (OPTIONAL) If this is false, the function expects
     *                                  	$col to be a numeric value. If this value is true
     *			 							this funcition simply concatenates $col and $row
     *										Default is false.
     */
    public function setCellBackColor($col, $row, $rgb, $isCoordinate = false) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }
        $cellCoord = $this->getCellCoord($col, $row, $isCoordinate);
        $this->phpXl->getActiveSheet()->getStyle($cellCoord)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
        $this->phpXl->getActiveSheet()->getStyle($cellCoord)->getFill()->getStartColor()->setARGB('FF'.$rgb);
        return true;
    }
	
    /**
     * Set the text color of a cell
     * 
     * @param      int		$col   			The column of the cell to set the text color
     * @param      int		$row   			The row of the cell to set the text color
     * @param      string	$rgb   			The HTML based RGB color (e.g. 'FF0000' is red)
     * @param      bool		$isCoordinate   (OPTIONAL) If this is false, the function expects
     *                                  	$col to be a numeric value. If this value is true
     *			 							this funcition simply concatenates $col and $row
     *										Default is false.
     */
    public function setCellTextColor($col, $row, $rgb, $isCoordinate = false) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }

        $cellCoord = $this->getCellCoord($col, $row, $isCoordinate);
        $this->phpXl->getActiveSheet()->getStyle($cellCoord)->getFont()->getColor()->setRGB($rgb);
        return true;
    }

    /**
     * Set the font name and size of a cell
     *
     * If $col and $row are set to 0 (their default values), then the entire
     * active sheet is set to the font family ans size
     * 
     * @param      string	$fontName   	The font family name to set (e.g. 'Arial', 'Calibri', etc.)
     * 						The font name must be a valid font name to set
     * @param      int		$fontSize   	The font size to set (e.g. 10, 12, etc.)
     * @param      int		$col            (OPTIONAL) The column to set the font on.
     *      					Default is 0. (0 means all columns)
     * @param      int		$row   		(OPTIONAL) The row to set the font on.
     * 						Default is 0. (0 means all rows)
     * @param      bool		$isCoordinate   (OPTIONAL) If this is false, the function expects
     *                                  	$col to be a numeric value. If this value is true
     *			 			this function simply concatenates $col and $row
     *						Default is false.
     */
    public function setCellFont($fontName, $fontSize, $col = 0, $row = 0, $isCoordinate = false) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }
        
        $wholeRow = ($col == 0); $wholeCol = ($row == 0);

        // getHighsetColumn returns letter (AZ), need to convert to num
        if ($col == 0) {
            $col = $this->getExcelColumnFromAlpha($this->phpXl->getActiveSheet()->getHighestColumn());
        }

        if ($row == 0) {
            $row = $this->phpXl->getActiveSheet()->getHighestRow();
        }

        $cellCoordEnd = $this->getCellCoord($col, $row, $isCoordinate);
        $cellCoordStart = 'A1';

        if (!$wholeCol || !$wholeRow) { // Only fall in here if one of them is false (Which means don't do all cells)
            if ($wholeCol) {
                $cellCoordStart = $this->getCellCoord($col, 1, $isCoordinate);
            }

            if ($wholeRow) {
                $cellCoordStart = 'A'.$row;
            }
        }

        $cellRange = $cellCoordStart.':'.$cellCoordEnd;

        if ($cellCoordStart == $cellCoordEnd || (!$wholeCol && !$wholeRow)) {
            $cellRange = $cellCoordEnd;
        }

        $this->phpXl->getActiveSheet()->getStyle($cellRange)->getFont()->setName($fontName);
        $this->phpXl->getActiveSheet()->getStyle($cellRange)->getFont()->setSize($fontSize);
        unset($style);  // freeup the memory
        return true;
    }
	
    /**
     * Sets the color of a column
     * 
     * @param      int			$col   The column to set the color to
     * @param      string		$rgb   The HTML based RGB color (e.g. 'FF0000' is red)
     */
    public function setColumnColor($col, $rgb) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }

        $lastRow = $this->phpXl->getActiveSheet()->getHighestRow();

        for ($i = 1; $i <= $lastRow; $i++) {
            $this->setCellBackColor($col, $i, $rgb);
        }
        
        return true;
    }
	
    /**
     * Set a column to certain font family and size
     * 
     * @param      int		$col   		The column to set the font on
     * @param      string	$fontName   	The font family name to set (e.g. 'Arial', 'Calibri', etc.)
     * 						The font name must be a valid font name to set
     * @param      int		$fontSize   	The font size to set (e.g. 10, 12, etc.)
     */
    public function setColumnFont($col, $fontName, $fontSize) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }
		
        $lastRow = $this->phpXl->getActiveSheet()->getHighestRow();
		
        for ($i = 1; $i <= $lastRow; $i++) {
            $this->setCellFont($fontName, $fontSize, $col, $i);
	}
        
        return true;
    }
	
    /**
     * Add a hyperlink to a cell
     *
     * When adding a hyperlink to a cell it does not color and underline
     * the cell as if you were in Excel, to emulate this, set $autoColor = true
     * 
     * @param      int		$col   		The column of the cell
     * @param      int		$row   		The row of the cell
     * @param      string	$link   	The link to set the cell to
     * @param      bool		$isCoordinate   (OPTIONAL) If this is false, the function expects
     *                             		$col to be a numeric value. If this value is true
     *			 			this funcition simply concatenates $col and $row
     *						Default is false.
     * @param      bool		$autoColor   	(OPTIONAL) True to emulate the coloring of a cell
     *						Default is false.
     */
    public function setHyperlink($col, $row, $link, $isCoordinate = false, $autoColor = false) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }

        $cellCoord = $this->getCellCoord($col, $row, $isCoordinate);
        $hyperlink = new \PHPExcel_Cell_Hyperlink($link, '');
        $this->phpXl->getActiveSheet()->setHyperlink($cellCoord, $hyperlink);

        if ($autoColor) {
            $this->setCellTextColor($col, $row, '0000FF', $isCoordinate);
        }
        
        return true;
    }
	
    /**
     * Set the current worksheet name
     *
     * This will set the name of the worksheet. You can see the name of
     * the worksheet at the bottom of the Excel window (normally on
     * a new worksheet it just says 'Sheet1')
     * 
     * @param      string	$name   The name of the sheet to set to
     */
    public function setWorksheetName($name) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }

        $this->phpXl->getActiveSheet()->setTitle($name);
        return true;
    }
	
    /**
     * Set an entire row to a certain color
     * 
     * @param      int		$row   The row to set the color to
     * @param      string	$rgb   The HTML based RGB color (e.g. 'FF0000' is red)
     */
    public function setRowColor($row, $rgb) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }
        
        $this->phpXl->getActiveSheet()->getStyle('A'.$row)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
        $this->phpXl->getActiveSheet()->getStyle('A'.$row)->getFill()->getStartColor()->setARGB('FF'.$rgb);
        $lastColumn = $this->phpXl->getActiveSheet()->getHighestColumn(); // B
        $this->phpXl->getActiveSheet()->duplicateStyle($this->phpXl->getActiveSheet()->getStyle('A'.$row), 'B'.$row.':'.$lastColumn.$row);
        return true;
    }
	
    /**
     * Sets the direction of text in a cell
     * 
     * @param      int                          $col   			The column of the cell to set the text direction
     * @param      int                          $row   			The row of the cell to set the text directoin
     * @param      PHPExcelTextDirection	$dir   			The PHPExcelTextDirection to set
     * @param      int				$angle   		The angle to set the text to
     * @param      bool				$isCoordinate           (OPTIONAL) If this is false, the function expects
     *                                                                  $col to be a numeric value. If this value is true
     *			 						this funcition simply concatenates $col and $row
     *									Default is false.
     */
    public function setCellTextDirection($col, $row, $dir, $angle, $isCoordinate = false) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }
        
        $cellCoord = $this->getCellCoord($col, $row, $isCoordinate);
        $angle = abs($angle);
    
        switch ($dir) {
            case PHPExcelTextDirection::Clockwise:
                $angle = -$angle;
                break;
            case PHPExcelTextDirection::CounterClockwise:
                // Do nothing since clockwise rotation is a positive value
                break;
            case PHPExcelTextDirection::Stacked:
                $angle = -165; // Stacked text ALWAYS has to be angle 165
                break;
            default:
                // Do we need to do anything here?? this will essentually be CounterClockwise
                break;
        }
        
        $this->phpXl->getActiveSheet()->getStyle($cellCoord)->getAlignment()->setTextRotation($angle);
        return true;
    }
	
    /**
     * Sets the font of a row
     * 
     * @param      int		$row   	        The row to set the font to
     * @param      string	$fontName   	The font family name to set (e.g. 'Arial', 'Calibri', etc.)
     * 						The font name must be a valid font name to set
     * @param      int		$fontSize   	The font size to set (e.g. 10, 12, etc.)
     */
    public function setRowFont($row, $fontName, $fontSize) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }
		
        $lastColumn = $this->getExcelColumnFromAlpha($this->phpXl->getActiveSheet()->getHighestColumn());
		
        for ($i = 1; $i <= $lastColumn; $i++) {
            $this->setCellFont($fontName, $fontSize, $i, $row);
        }
        
        return true;
    }
	
    /**
     * Write data to a specific cell
     * 
     * @param      int		$col            The column of the cell to write to
     * @param      int		$row   		The row of the cell to write to
     * @param      mixed	$data   	The data to write to the cell
     * @param      bool		$isCoordinate   (OPTIONAL) If this is false, the function expects
     *                                  	$col to be a numeric value. If this value is true
     *			 			this funcition simply concatenates $col and $row
     *						Default is false.
     */
    public function writeCell($col, $row, $data, $isCoordinate = false) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            return false;
        }

        if ($isCoordinate) {
            $this->phpXl->getActiveSheet()->setCellValue(($col.$row), $data);
        } else {
            $this->phpXl->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $data);
        }
        
        return true;
        //$this->flush();
    }
	
    /**
     * Writes a row of data an advances the current row pointer
     * 
     * @param      mixed	$data   The data to write (can be an array)
     */
    public function write($data) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            if (is_array($data)) {
                $splits = $data;
            } else {
                $splits = explode(',', $data);
            }
            
            return $this->writeRow($this->currentRow++, $splits);
        } else {
            return $this->writeRow($this->currentRow++, $data);
        }
    }
	
    /**
     * Writes data to a specific row
     * 
     * @param      int		$row   	The row to write the data to
     * @param      mixed	$data   The data to write
     */
    public function writeRow($row, $data) {
        if ($this->type == PHPExcelWrapperType::CSV) {
            $outValue = $data;
            
            if (is_array($data)) {
                $outValue = '';
                $count = count($data);
                
                for ($i = 0; $i < $count; $i++) {
                    $outValue .= $data[$i];

                    if ($i < $count - 1) {
                        $outValue .= ',';
                    }
                }
            }
            
            if (substr($outValue, (strlen($outValue) - 1), 1) != "\n") {
                $outValue .= "\n";
            }

            fputs($this->handle, $outValue);
        } else {
            if (is_array($data)) {
                $count = count($data);
                    
                for ($i = 0; $i < $count; $i++) {
                    $writeResult = $this->writeCell($i, $row, $data[$i]);
                }
            } else {
                $writeResult = $this->writeCell($col, $row, $data);
            }
        }
        
        return true;
    }
}
