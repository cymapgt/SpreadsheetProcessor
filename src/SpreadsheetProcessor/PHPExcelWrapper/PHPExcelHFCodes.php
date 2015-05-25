<?php
namespace cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper;

/**
 * PHPExcelHFCodes
 * This enables use of OpenXML standards for header
 * codes. These can be used when declaring inline formatting
 * on a worksheets Header/Footer
 *
 * @category	CYMAPGTReporting
 * @package	core.application.spreadsheet.SpreadsheetProcessor
 * @copyright   Copyright (c) 2012 CYMAP-GT
 */
class PHPExcelHFCodes
{
    const HFCODE_LEFTSECTION         = '&L';
    const HFCODE_CURRENTPAGE         = '&P';
    const HFCODE_TOTALPAGES          = '&N';
    const HFCODE_FONTSIZE            = '&font size';
    const HFCODE_FONTCOLOR           = '&K';
    const HFCODE_STRIKETHROUGH       = '&S';
    const HFCODE_SUPERSCRIPT         = '&X';
    const HFCODE_SUBSCRIPT           = '&Y';
    const HFCODE_CENTERSECTION       = '&C';
    const HFCODE_DATE                = '&D';
    const HFCODE_TIME                = '&T';
    const HFCODE_PICTUREASBACKGROUND = '&G';
    const HFCODE_TEXTSINGLEUNDERLINE = '&U';
    const HFCODE_DOUBLEUNDERLINE     = '&E';
    const HFCODE_RIGHTSECTION        = '&R';
    const HFCODE_FILEPATH            = '&Z';
    const HFCODE_FILENAME            = '&F';
    const HFCODE_SHEETTABNAME        = '&A';
    const HFCODE_ADDTOPAGENO      	 = '&+';
    const HFCODE_SUBTRACTFROMPAGENO  = '&-';
    const HFCODE_FONTNAMEANDTYPE     = '&"font name,font type"';
    const HFCODE_BOLDFONTSTYLEA      = '&"-,Bold"';
    const HFCODE_BOLDFONTSTYLEB      = '&B';
    const HFCODE_REGULARFONTSTYLE    = '&"-,Regular"';
    const HFCODE_ITALICFONTSTYLEA    = '&"-,Italic"';
    const HFCODE_ITALICFONTSTYLEB    = '&I';
    const HFCODE_BOLDITALICFONTSTYLE = '&"-,Bold Italic"';
    const HFCODE_OUTLINESTYLE        = '&0';
    const HFCODE_SHADOWSTYLE         = '&H';
}
