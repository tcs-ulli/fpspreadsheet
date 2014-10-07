{ Translatable strings for fpspreadsheet }

unit fpsStrings;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}


interface

resourcestring
  rsUnsupportedReadFormat = 'Tried to read a spreadsheet using an unsupported format';
  rsUnsupportedWriteFormat = 'Tried to write a spreadsheet using an unsupported format';
  rsNoValidSpreadsheetFile = '"%s" is not a valid spreadsheet file';
  rsUnknownSpreadsheetFormat = 'unknown format';
  rsMaxRowsExceeded = 'This workbook contains %d rows, but the selected ' +
    'file format does not support more than %d rows.';
  rsMaxColsExceeded = 'This workbook contains %d columns, but the selected ' +
    'file format does not support more than %d columns.';
  rsTooManyPaletteColors = 'This workbook contains more colors (%d) than ' +
    'supported by the file format (%d). The additional colors are replaced by '+
    'the best-matching palette colors.';
  rsInvalidFontIndex = 'Invalid font index';
  rsInvalidNumberFormat = 'Trying to use an incompatible number format.';
  rsInvalidDateTimeFormat = 'Trying to use an incompatible date/time format.';
  rsNoValidNumberFormatString = 'No valid number format string.';
  rsNoValidCellAddress = '"%s" is not a valid cell address.';
  rsNoValidCellRangeAddress = '"%s" is not a valid cell range address.';
  rsNoValidCellRangeOrCellAddress = '"%s" is not a valid cell or cell range address.';
  rsSpecifyNumberOfParams = 'Specify number of parameters for function %s';
  rsIncorrectParamCount = 'Funtion %s requires at least %d and at most %d parameters.';
  rsCircularReference = 'Circular reference found when calculating worksheet formulas';
  rsFileNotFound = 'File "%s" not found.';
  rsInvalidWorksheetName = '"%s" is not a valid worksheet name.';

  rsTRUE = 'TRUE';
  rsFALSE = 'FALSE';
  rsErrEmptyIntersection = '#NULL!';
  rsErrDivideByZero = '#DIV/0!';
  rsErrWrongType = '#VALUE!';
  rsErrIllegalRef = '#REF!';
  rsErrWrongName = '#NAME?';
  rsErrOverflow = '#NUM!';
  rsErrArgError = '#N/A';
  rsErrFormulaNotSupported = '<FORMULA?>';

{%H-}rsNoValidDateTimeFormatString = 'No valid date/time format string.';
{%H-}rsIllegalNumberFormat = 'Illegal number format.';


implementation

end.
