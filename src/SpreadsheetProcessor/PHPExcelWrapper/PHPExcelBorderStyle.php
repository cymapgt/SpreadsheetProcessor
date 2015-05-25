<?php
namespace cymapgt\core\application\spreadsheet\SpreadsheetProcessor\PHPExcelWrapper;

// These are taken directly from PHPExcel/Style/Border.php
// They are duplicated becuase the point of this wrapper is to not have
// to include a whole bunch of other files in whatver script your using
class PHPExcelBorderStyle
{
    const BORDER_NONE               = 'none';
    const BORDER_DASHDOT            = 'dashDot';
    const BORDER_DASHDOTDOT         = 'dashDotDot';
    const BORDER_DASHED             = 'dashed';
    const BORDER_DOTTED             = 'dotted';
    const BORDER_DOUBLE             = 'double';
    const BORDER_HAIR               = 'hair';
    const BORDER_MEDIUM             = 'medium';
    const BORDER_MEDIUMDASHDOT      = 'mediumDashDot';
    const BORDER_MEDIUMDASHDOTDOT   = 'mediumDashDotDot';
    const BORDER_MEDIUMDASHED       = 'mediumDashed';
    const BORDER_SLANTDASHDOT       = 'slantDashDot';
    const BORDER_THICK              = 'thick';
    const BORDER_THIN               = 'thin';
}
