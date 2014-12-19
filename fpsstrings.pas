{ Translatable strings for fpspreadsheet }

unit fpsStrings;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}


interface

resourcestring
  rsExportFileIsRequired = 'Export file name is required';
  rsFPSExportDescription = 'Spreadsheet file';
  rsMultipleSheetsOnlyWithRestorePosition = 'Export to multiple sheets is possible '+
    'only if position is restored.';
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
  rsInvalidExtension = 'Attempting to save a spreadsheet by extension, ' +
    'but the extension %s is not valid.';
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
  rsDefectiveInternalStructure = 'Defective internal structure of %s file.';
  rsUnknownDataType = 'Unknown data type.';
  rsUnknownErrorType = 'Unknown error type.';
  rsTruncateTooLongCellText = 'Text value exceeds %d character limit in cell %s '+
    'and has been truncated.';
  rsColumnStyleNotFound = 'Column style not found.';
  rsRowStyleNotFound = 'Row style not found.';
  rsInvalidCharacterInCell = 'Invalid character(s) in cell %s.';
  rsUTF8TextExpectedButANSIFoundInCell = 'Expected UTF8 text but probably ANSI '+
      'text found in cell %s.';
  rsIndexInSSTOutOfRange = 'Index %d in SST out of range (0-%d).';
  rsAmbiguousDecThouSeparator = 'Assuming usage of decimal separator in "%s".';


  rsTRUE = 'TRUE';               // wp: Do we really want to translate these strings?
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
