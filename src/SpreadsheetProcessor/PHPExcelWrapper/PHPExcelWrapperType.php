<?php
namespace cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper;

/**
* PHPExcelWrapperType
* This stores codes for filetypes that we support
*
* @category	CYMAPGTReporting
* @package	core.application.spreadsheet.SpreadsheetProcessor
* @copyright    Copyright (c) 2012 CYMAP-GT
*/
//PHPExcel_Settings::setPdfRenderer(PHPExcel_Settings::PDF_RENDERER_MPDF,"PHPExcel/Shared/MPDF54/");
class PHPExcelWrapperType
{
    const Excel5    = 0;
    const Excel2007 = 1;
    const CSV       = 2;
    const PDF       = 3;
    const HTML      = 4;
}
