<?php
namespace cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper;

/**
* PHPExcelTmpFilePath
* This stores filepaths that act as phpexcel's 'tmp' before
* user copies the report from temp to other folder
* custom .htaccess should be applied to each folder (for
* remote (sandbox) mode only)
*
* @category	CYMAPGTReporting
* @package	core.application.spreadsheet.SpreadsheetProcessor
* @copyright   Copyright (c) 2012 CYMAP-GT
*/
class PHPExcelTmpFilePath
{
    const TMP_EXCEL = "files/spreadsheet/";
    const TMP_PDF   = "files/other/pdf/";
    const TMP_HTML  = "files/other/html";
    const TMP_SWF   = "files/other/swf/";
    const TMP_TEXT  = "files/text/";
    const TMP_IMAGE = "files/images";
}
