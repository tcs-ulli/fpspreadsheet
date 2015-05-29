unit fpsTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpimage;

type
  {@@ File formats supported by fpspreadsheet }
  TsSpreadsheetFormat = (sfExcel2, sfExcel5, sfExcel8,
   sfOOXML, sfOpenDocument, sfCSV, sfWikiTable_Pipes, sfWikiTable_WikiMedia);

  {@@ Flag set during reading or writing of a workbook }
  TsReadWriteFlag = (rwfNormal, rwfRead, rwfWrite);

  {@@ Record collection limitations of a particular file format }
  TsSpreadsheetFormatLimitations = record
    MaxRowCount: Cardinal;
    MaxColCount: Cardinal;
    MaxPaletteSize: Integer;
  end;

const
  {@@ Default binary <b>Excel</b> file extension}
  STR_EXCEL_EXTENSION = '.xls';
  {@@ Default xml <b>Excel</b> file extension (>= Excel 2007) }
  STR_OOXML_EXCEL_EXTENSION = '.xlsx';
  {@@ Default <b>OpenDocument</b> spreadsheet file extension }
  STR_OPENDOCUMENT_CALC_EXTENSION = '.ods';
  {@@ Default extension of <b>comma-separated-values</b> file }
  STR_COMMA_SEPARATED_EXTENSION = '.csv';
  {@@ Default extension of <b>wikitable files</b> in <b>pipes</b> format}
  STR_WIKITABLE_PIPES_EXTENSION = '.wikitable_pipes';
  {@@ Default extension of <b>wikitable files</b> in <b>wikimedia</b> format }
  STR_WIKITABLE_WIKIMEDIA_EXTENSION = '.wikitable_wikimedia';

  {@@ Maximum count of worksheet columns}
  MAX_COL_COUNT = 65535;

  {@@ Name of the default font}
  DEFAULT_FONTNAME = 'Arial';
  {@@ Size of the default font}
  DEFAULT_FONTSIZE = 10;
  {@@ Index of the default font in workbook's font list }
  DEFAULT_FONTINDEX = 0;
  {@@ Index of the hyperlink font in workbook's font list }
  HYPERLINK_FONTINDEX = 1;
  {@@ Index of bold default font in workbook's font list }
  BOLD_FONTINDEX = 2;
  {@@ Index of italic default font in workbook's font list - not used directly }
  INTALIC_FONTINDEX = 3;

  {@@ Takes account of effect of cell margins on row height by adding this
      value to the nominal row height. Note that this is an empirical value
      and may be wrong. }
  ROW_HEIGHT_CORRECTION = 0.2;


type
                       (*
  {@@ Possible encodings for a non-unicode encoded text }
  TsEncoding = (
    seLatin1,
    seLatin2,
    seCyrillic,
    seGreek,
    seTurkish,
    seHebrew,
    seArabic,
    seUTF16
    );            *)

  {@@ Tokens to identify the <b>elements in an expanded formula</b>.

   NOTE: When adding or rearranging items
   * make sure that the subtypes TOperandTokens and TBasicOperationTokens
     are complete
   * make sure to keep the table "TokenIDs" in unit xlscommon in sync
  }
  TFEKind = (
    { Basic operands }
    fekCell, fekCellRef, fekCellRange, fekCellOffset, fekNum, fekInteger,
    fekString, fekBool, fekErr, fekMissingArg,
    { Basic operations }
    fekAdd, fekSub, fekMul, fekDiv, fekPercent, fekPower, fekUMinus, fekUPlus,
    fekConcat,  // string concatenation
    fekEqual, fekGreater, fekGreaterEqual, fekLess, fekLessEqual, fekNotEqual,
    fekParen,   // show parenthesis around expression node
    { Functions - they are identified by their name }
    fekFunc
  );

  {@@ These tokens identify operands in RPN formulas. }
  TOperandTokens = fekCell..fekMissingArg;

  {@@ These tokens identify basic operations in RPN formulas. }
  TBasicOperationTokens = fekAdd..fekParen;

type
  {@@ Flags to mark the address or a cell or a range of cells to be <b>absolute</b>
      or <b>relative</b>. They are used in the set TsRelFlags. }
  TsRelFlag = (rfRelRow, rfRelCol, rfRelRow2, rfRelCol2);

  {@@ Flags to mark the address of a cell or a range of cells to be <b>absolute</b>
      or <b>relative</b>. It is a set consisting of TsRelFlag elements. }
  TsRelFlags = set of TsRelFlag;

const
  {@@ Abbreviation of all-relative cell reference flags }
  rfAllRel = [rfRelRow, rfRelCol, rfRelRow2, rfRelCol2];

  {@@ Separator between worksheet name and cell (range) reference in an address }
  SHEETSEPARATOR = '!';

type
  {@@ Elements of an expanded formula.
    Note: If ElementKind is fekCellOffset, "Row" and "Col" have to be cast
          to signed integers! }
  TsFormulaElement = record
    ElementKind: TFEKind;
    Row, Row2: Cardinal;   // zero-based
    Col, Col2: Cardinal;   // zero-based
//    Param1, Param2: Word;  // Extra parameters
    DoubleValue: double;
    IntValue: Word;
    StringValue: String;
    RelFlags: TsRelFlags;  // store info on relative/absolute addresses
    FuncName: String;
    ParamsNum: Byte;
  end;

  {@@ RPN formula. Similar to the expanded formula, but in RPN notation.
      Simplifies the task of format writers which need RPN }
  TsRPNFormula = array of TsFormulaElement;

  {@@ Describes the <b>type of content</b> in a cell of a TsWorksheet }
  TCellContentType = (cctEmpty, cctFormula, cctNumber, cctUTF8String,
    cctDateTime, cctBool, cctError);

  {@@ The record TsComment describes a comment attached to a cell.
     @param   Row        (0-based) row index of the cell
     @param   Col        (0-based) column index of the cell
     @param   Text       Comment text }
  TsComment = record
    Row, Col: Cardinal;
    Text: String;
  end;

  {@@ Pointer to a TsComment record }
  PsComment = ^TsComment;
                    (*
  {@@ Specifies whether a hyperlink refers to an internal cell address
      within the current workbook, or a URI (file://, http://, mailto, etc). }
  TsHyperlinkKind = (hkNone, hkInternal, hkURI);
                      *)
  {@@ The record TsHyperlink contains info on a hyperlink in a cell
    @param   Row          Row index of the cell containing the hyperlink
    @param   Col          Column index of the cell containing the hyperlink
    @param   Target       Target of hyperlink: URI of file, web link, mail; or:
                          internal link (# followed by cell address)
    @param   Note         Text displayed as a popup hint by Excel }
  TsHyperlink = record
    Row, Col: Cardinal;
    Target: String;
    Tooltip: String;
  end;

  {@@ Pointer to a TsHyperlink record }
  PsHyperlink = ^TsHyperlink;

  {@@ Callback function, e.g. for iterating the internal AVL trees of the workbook/sheet}
  TsCallback = procedure (data, arg: Pointer) of object;

  {@@ Error code values }
  TsErrorValue = (
    errOK,                 // no error
    errEmptyIntersection,  // #NULL!
    errDivideByZero,       // #DIV/0!
    errWrongType,          // #VALUE!
    errIllegalRef,         // #REF!
    errWrongName,          // #NAME?
    errOverflow,           // #NUM!
    errArgError,           // #N/A
    // --- no Excel errors --
    errFormulaNotSupported
  );

  {@@ List of possible formatting fields }
  TsUsedFormattingField = (uffTextRotation, uffFont, {uffBold, }uffBorder,
    uffBackground, uffNumberFormat, uffWordWrap, uffHorAlign, uffVertAlign
  );
  { NOTE: "uffBackgroundColor" of older versions replaced by "uffBackground" }

  {@@ Describes which formatting fields are active }
  TsUsedFormattingFields = set of TsUsedFormattingField;

const
  { @@ Codes for curreny format according to FormatSettings.CurrencyFormat:
       "C" = currency symbol, "V" = currency value, "S" = space character
       For the negative value formats, we use also:
       "B" = bracket, "M" = Minus

       The order of these characters represents the order of these items.

       Example: 1000 dollars  --> "$1000"  for pCV,   or "1000 $"  for pVsC
               -1000 dollars --> "($1000)" for nbCVb, or "-$ 1000" for nMCSV

       Assignment taken from "sysstr.inc" }
  pcfDefault = -1;   // use value from Worksheet.FormatSettings.CurrencyFormat
  pcfCV      = 0;    // $1000
  pcfVC      = 1;    // 1000$
  pcfCSV     = 2;    // $ 1000
  pcfVSC     = 3;    // 1000 $

  ncfDefault = -1;   // use value from Worksheet.FormatSettings.NegCurrFormat
  ncfBCVB    = 0;    // ($1000)
  ncfMCV     = 1;    // -$1000
  ncfCMV     = 2;    // $-1000
  ncfCVM     = 3;    // $1000-
  ncfBVCB    = 4;    // (1000$)
  ncfMVC     = 5;    // -1000$
  ncfVMC     = 6;    // 1000-$
  ncfVCM     = 7;    // 1000$-
  ncfMVSC    = 8;    // -1000 $
  ncfMCSV    = 9;    // -$ 1000
  ncfVSCM    = 10;   // 1000 $-
  ncfCSVM    = 11;   // $ 1000-
  ncfCSMV    = 12;   // $ -1000
  ncfVMSC    = 13;   // 1000- $
  ncfBCSVB   = 14;   // ($ 1000)
  ncfBVSCB   = 15;   // (1000 $)

type
  {@@ Text rotation formatting. The text is rotated relative to the standard
      orientation, which is from left to right horizontal:
      <pre>
       --->
       ABC </pre>

      So 90 degrees clockwise means that the text will be:
      <pre>
       |  A
       |  B
       v  C </pre>

      And 90 degree counter clockwise will be:
      <pre>
       ^  C
       |  B
       |  A</pre>

      Due to limitations of the text mode the characters are not rotated here.
      There is, however, also a "stacked" variant which looks exactly like
      the 90-degrees-clockwise case.
  }
  TsTextRotation = (trHorizontal, rt90DegreeClockwiseRotation,
    rt90DegreeCounterClockwiseRotation, rtStacked);

  {@@ Indicates horizontal text alignment in cells }
  TsHorAlignment = (haDefault, haLeft, haCenter, haRight);

  {@@ Indicates vertical text alignment in cells }
  TsVertAlignment = (vaDefault, vaTop, vaCenter, vaBottom);

  {@@ Colors in fpspreadsheet are given as rgb values in little-endian notation
    (i.e. "r" is the low-value byte). The highest-value byte, if not zero,
    indicates special colors. }
  TsColor = DWord;

const
  {@@ These are some important rgb color volues.
  }
  {@@ rgb value of <b>black</b> color, BIFF2 palette index 0, BIFF8 index 8}
  scBlack = $00000000;
  {@@ rgb value of <b>white</b> color, BIFF2 palette index 1, BIFF8 index 9 }
  scWhite = $00FFFFFF;
  {@@ rgb value of <b>red</b> color, BIFF2 palette index 2, BIFF8 index 10 }
  scRed = $000000FF;
  {@@ rgb value of <b>green</b> color, BIFF2 palette index 3, BIFF8 index 11 }
  scGreen = $0000FF00;
  {@@ rgb value of <b>blue</b> color, BIFF2 palette index 4, BIFF8 indexes 12 and 39}
  scBlue = $00FF0000;
  {@@ rgb value of <b>yellow</b> color, BIFF2 palette index 5, BIFF8 indexes 13 and 34}
  scYellow = $0000FFFF;
  {@@ rgb value of <b>magenta</b> color, BIFF2 palette index 6, BIFF8 index 14 and 33}
  scMagenta = $00FF00FF;
  scPink = $00FE00FE;
  {@@ rgb value of <b>cyan</b> color, BIFF2 palette index 7, BIFF8 indexes 15}
  scCyan = $00FFFF00;
  scTurquoise = scCyan;
  {@@ rgb value of <b>dark red</b> color, BIFF8 indexes 16 and 35}
  scDarkRed = $00000080;
  {@@ rgb value of <b>dark green</b> color, BIFF8 index 17 }
  scDarkGreen = $00008000;
  {@@ rgb value of <b>dark blue</b> color }
  scDarkBlue = $008B0000;
  {@@ rgb value of <b>"navy"</b> color, BIFF8 palette indexes 18 and 32 }
  scNavy = $00800000;
  {@@ rgb value of <b>olive</b> color }
  scOlive = $00008080;
  {@@ rgb value of <b>purple</b> color, BIFF8 palette indexes 20 and 36 }
  scPurple = $00800080;
  {@@ rgb value of <b>teal</b> color, BIFF8 palette index 21 and 38 }
  scTeal = $00808000;
  {@@ rgb value of <b>silver</b> color }
  scSilver = $00C0C0C0;
  scGray25pct = scSilver;
  {@@ rgb value of <b>grey</b> color }
  scGray = $00808080;
  {@@ rgb value of <b>gray</b> color }
  scGrey = scGray;       // redefine to allow different spelling
  scGray50pct = scGray;
  {@@ rgb value of a <b>10% grey</b> color }
  scGray10pct = $00E6E6E6;
  {@@ rgb value of a <b>10% gray</b> color }
  scGrey10pct = scGray10pct;
  {@@ rgb value of a <b>20% grey</b> color }
  scGray20pct = $00CCCCCC;
  {@@ rgb value of a <b>20% gray</b> color }
  scGrey20pct = scGray20pct;
  {@@ rgb value of <b>periwinkle</b> color, BIFF8 palette index 24 }
  scPeriwinkle = $00FF9999;
  {@@ rgb value of <b>plum</b> color, BIFF8 palette indexes 25 and 61 }
  scPlum = $00663399;
  {@@ rgb value of <b>ivory</b> color, BIFF8 palette index 26 }
  scIvory = $00CCFFFF;
  {@@ rgb value of <b>light turquoise</b> color, BIFF8 palette indexes 27 and 41 }
  scLightTurquoise = $00FFFFCC;
  {@@ rgb value of <b>dark purple</b> color, BIFF8 palette index 28 }
  scDarkPurple = $00660066;
  {@@ rgb value of <b>coral</b> color, BIFF8 palette index 29 }
  scCoral = $008080FF;
  {@@ rgb value of <b>ocean blue</b> color, BIFF8 palette index 30 }
  scOceanBlue = $00CC6600;
  {@@ rgb value of <b>ice blue</b> color, BIFF8 palette index 31 }
  scIceBlue = $00FFCCCC;
  {@@ rgb value of <b>sky blue </b>color, BIFF8 palette index 40 }
  scSkyBlue = $00FFCC00;
  {@@ rgb value of <b>light green</b> color, BIFF8 palette index 42 }
  scLightGreen = $00CCFFCC;
  {@@ rgb value of <b>light yellow</b> color, BIFF8 palette index 43 }
  scLightYellow = $0099FFFF;
  {@@ rgb value of <b>pale blue</b> color, BIFF8 palette index 44 }
  scPaleBlue = $00FFCC99;
  {@@ rgb value of <b>rose</b> color, BIFF8 palette index 45 }
  scRose = $00CC99FF;
  {@@ rgb value of <b>lavander</b> color, BIFF8 palette index 46 }
  scLavander = $00FF99CC;
  {@@ rgb value of <b>tan</b> color, BIFF8 palette index 47 }
  scTan = $0099CCFF;
  {@@ rgb value of <b>light blue</b> color, BIFF8 palette index 48 }
  scLightBlue = $00FF6633;
  {@@ rgb value of <b>aqua</b> color, BIFF8 palette index 49 }
  scAqua = $00CCCC33;
  {@@ rgb value of <b>lime</b> color, BIFF8 palette index 50 }
  scLime = $0000CC99;
  {@@ rgb value of <b>golden</b> color, BIFF8 palette index 51 }
  scGold = $0000CCFF;
  {@@ rgb value of <b>light orange</b> color, BIFF8 palette index 52 }
  scLightOrange = $000099FF;
  {@@ rgb value of <b>orange</b> color, BIFF8 palette index 53 }
  scOrange = $000066FF;
  {@@ rgb value of <b>blue gray</b>, BIFF8 palette index 54 }
  scBlueGray = $00996666;
  scBlueGrey = scBlueGray;
  {@@ rgb value of <b>gray 40%</b>, BIFF8 palette index 55 }
  scGray40pct = $00969696;
  {@@ rgb value of <b>dark teal</b>, BIFF8 palette index 56 }
  scDarkTeal = $00663300;
  {@@ rgb value of <b>sea green</b>, BIFF8 palette index 57 }
  scSeaGreen = $00669933;
  {@@ rgb value of <b>very dark green</b>, BIFF8 palette index 58 }
  scVeryDarkGreen = $00003300;
  {@@ rgb value of <b>olive green</b> color, BIFF8 palette index 59 }
  scOliveGreen = $00003333;
  {@@ rgb value of <b>brown</b> color, BIFF8 palette index 60 }
  scBrown = $00003399;
  {@@ rgb value of <b>indigo</b> color, BIFF8 palette index 62 }
  scIndigo = $00993333;
  {@@ rgb value of <b>80% gray</b>, BIFF8 palette index 63 }
  scGray80pct = $00333333;
  scGrey80pct = scGray80pct;

//  {@@ rgb value of <b>orange</b> color }
//  scOrange = $0000A5FF;
  {@@ rgb value of <b>dark brown</b> color }
  scDarkBrown = $002D52A0;

//  {@@ rgb value of <b>brown</b> color }
//  scBrown = $003F85CD;
  {@@ rgb value of <b>beige</b> color }
  scBeige = $00DCF5F5;
  {@@ rgb value of <b>"wheat"</b> color (yellow-orange) }
  scWheat = $00B3DEF5;

  {@@ Identifier for not-defined color }
  scNotDefined = $40000000;
  {@@ Identifier for transparent color }
  scTransparent = $20000000;
  {@@ Identifier for palette index encoded into the TsColor }
  scPaletteIndexMask = $80000000;
  {@@ Mask for the rgb components contained in the TsColor }
  scRGBMask = $00FFFFFF;

type
  {@@ Font style (redefined to avoid usage of "Graphics" }
  TsFontStyle = (fssBold, fssItalic, fssStrikeOut, fssUnderline);

  {@@ Set of font styles }
  TsFontStyles = set of TsFontStyle;

  {@@ Font record used in fpspreadsheet. Contains the font name, the font size
      (in points), the font style, and the font color. }
  TsFont = class
    {@@ Name of the font face, such as 'Arial' or 'Times New Roman' }
    FontName: String;
    {@@ Size of the font in points }
    Size: Single;   // in "points"
    {@@ Font style, such as bold, italics etc. - see TsFontStyle}
    Style: TsFontStyles;
    {@@ Text color given as rgb value }
    Color: TsColor;
  end;

  {@@ Indicates the border for a cell. If included in the CellBorders set the
      corresponding border is drawn in the style defined by the CellBorderStyle. }
  TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth, cbDiagUp, cbDiagDown);

  {@@ Indicates the border for a cell }
  TsCellBorders = set of TsCellBorder;

  {@@ Line style (for cell borders) }
  TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair);

  {@@ The Cell border style reocrd contains the linestyle and color of a cell
      border. There is a CellBorderStyle for each border. }
  TsCellBorderStyle = record
    LineStyle: TsLineStyle;
    Color: TsColor;
  end;

  {@@ The cell border styles of each cell border are collected in this array. }
  TsCellBorderStyles = array[TsCellBorder] of TsCellBorderStyle;

  {@@ Border styles for each cell border used by default: a thin, black, solid line }
const
  DEFAULT_BORDERSTYLES: TsCellBorderStyles = (
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack)
  );

type
  {@@ Style of fill pattern for cell backgrounds }
  TsFillStyle = (fsNoFill, fsSolidFill,
    fsGray75, fsGray50, fsGray25, fsGray12, fsGray6,
    fsStripeHor, fsStripeVert, fsStripeDiagUp, fsStripeDiagDown,
    fsThinStripeHor, fsThinStripeVert, fsThinStripeDiagUp, fsThinStripeDiagDown,
    fsHatchDiag, fsThinHatchDiag, fsThickHatchDiag, fsThinHatchHor);

  {@@ Fill pattern record }
  TsFillPattern = record
    Style: TsFillStyle;  // pattern type
    FgColor: TsColor;    // pattern color
    BgColor: TsColor;    // background color
  end;

const
  {@@ Parameters for a non-filled cell background }
  EMPTY_FILL: TsFillPattern = (
    Style: fsNoFill;
    FgColor: scTransparent;
    BgColor: scTransparent;
  );

type
  {@@ Identifier for a compare operation }
  TsCompareOperation = (coNotUsed,
    coEqual, coNotEqual, coLess, coGreater, coLessEqual, coGreaterEqual
  );

  {@@ Number/cell formatting. Only uses a subset of the default formats,
      enough to be able to read/write date/time values.
      nfCustom allows to apply a format string directly. }
  TsNumberFormat = (
    // general-purpose for all numbers
    nfGeneral,
    // numbers
    nfFixed, nfFixedTh, nfExp, nfPercentage, nfFraction,
    // currency
    nfCurrency, nfCurrencyRed,
    // dates and times
    nfShortDateTime, nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfDayMonth, nfMonthYear, nfTimeInterval,
    // other (format string goes directly into the file)
    nfCustom);

  {@@ Tokens used by the elements of the number format parser }
  TsNumFormatToken = (
    nftGeneral,            // token for "general" number format
    nftText,               // must be quoted, stored in TextValue
    nftThSep,              // ',', replaced by FormatSettings.ThousandSeparator
    nftDecSep,             // '.', replaced by FormatSettings.DecimalSeparator
    nftYear,               // 'y' or 'Y', count stored in IntValue
    nftMonth,              // 'm' or 'M', count stored in IntValue
    nftDay,                // 'd' or 'D', count stored in IntValue
    nftHour,               // 'h' or 'H', count stored in IntValue
    nftMinute,             // 'n' or 'N' (or 'm'/'M'), count stored in IntValue
    nftSecond,             // 's' or 'S', count stored in IntValue
    nftMilliseconds,       // 'z', 'Z', '0', count stored in IntValue
    nftAMPM,               //
    nftMonthMinute,        // 'm'/'M' or 'n'/'N', meaning depending on context
    nftDateTimeSep,        // '/' or ':', replaced by value from FormatSettings, stored in TextValue
    nftSign,               // '+' or '-', stored in TextValue
    nftSignBracket,        // '(' or ')' for negative values, stored in TextValue
    nftIntOptDigit,        // '#', count stored in IntValue
    nftIntZeroDigit,       // '0', count stored in IntValue
    nftIntSpaceDigit,      // '?', count stored in IntValue
    nftIntTh,              // '#,##0' sequence for nfFixed, count of 0 stored in IntValue
    nftZeroDecs,           // '0' after dec sep, count stored in IntValue
    nftOptDecs,            // '#' after dec sep, count stored in IntValue
    nftSpaceDecs,          // '?' after dec sep, count stored in IntValue
    nftExpChar,            // 'e' or 'E', stored in TextValue
    nftExpSign,            // '+' or '-' in exponent
    nftExpDigits,          // '0' digits in exponent, count stored in IntValue
    nftPercent,            // '%' percent symbol
    nftFactor,             // thousand separators at end of format string, each one divides value by 1000
    nftFracSymbol,         // '/' fraction symbol
    nftFracNumOptDigit,    // '#' in numerator, count stored in IntValue
    nftFracNumSpaceDigit,  // '?' in numerator, count stored in IntValue
    nftFracNumZeroDigit,   // '0' in numerator, count stored in IntValue
    nftFracDenomOptDigit,  // '#' in denominator, count stored in IntValue
    nftFracDenomSpaceDigit,// '?' in denominator, count stored in IntValue
    nftFracDenomZeroDigit, // '0' in denominator, count stored in IntValue
    nftFracDenom,          // specified denominator, value stored in IntValue
    nftCurrSymbol,         // e.g., '"$"', stored in TextValue
    nftCountry,
    nftColor,              // e.g., '[red]', Color in IntValue
    nftCompareOp,
    nftCompareValue,
    nftSpace,
    nftEscaped,            // '\'
    nftRepeat,
    nftEmptyCharWidth,
    nftTextFormat);

  TsNumFormatElement = record
    Token: TsNumFormatToken;
    IntValue: Integer;
    FloatValue: Double;
    TextValue: String;
  end;

  TsNumFormatElements = array of TsNumFormatElement;

  TsNumFormatKind = (nfkPercent, nfkExp, nfkCurrency, nfkFraction,
    nfkDate, nfkTime, nfkTimeInterval, nfkHasColor, nfkHasThSep, nfkHasFactor);
  TsNumFormatKinds = set of TsNumFormatKind;

  TsNumFormatSection = record
    Elements: TsNumFormatElements;
    Kind: TsNumFormatKinds;
    NumFormat: TsNumberFormat;
    Decimals: Byte;
    Factor: Double;
    FracInt: Integer;
    FracNumerator: Integer;
    FracDenominator: Integer;
    CurrencySymbol: String;
    Color: TsColor;
  end;
  PsNumFormatSection = ^TsNumFormatSection;

  TsNumFormatSections = array of TsNumFormatSection;

  { TsNumFormatParams }

  TsNumFormatParams = class(TObject)
  private
  protected
    function GetNumFormat: TsNumberFormat; virtual;
    function GetNumFormatStr: String; virtual;
  public
    Sections: TsNumFormatSections;
    procedure DeleteElement(ASectionIndex, AElementIndex: Integer);
    procedure InsertElement(ASectionIndex, AElementIndex: Integer;
      AToken: TsNumFormatToken);
    function SectionsEqualTo(ASections: TsNumFormatSections): Boolean;
    procedure SetCurrSymbol(AValue: String);
    procedure SetDecimals(AValue: Byte);
    procedure SetNegativeRed(AEnable: Boolean);
    procedure SetThousandSep(AEnable: Boolean);
    property NumFormat: TsNumberFormat read GetNumFormat;
    property NumFormatStr: String read GetNumFormatStr;
  end;

  TsNumFormatParamsClass = class of TsNumFormatParams;

  {@@ Cell calculation state }
  TsCalcState = (csNotCalculated, csCalculating, csCalculated);

  {@@ Cell flag }
  TsCellFlag = (cfCalculating, cfCalculated, cfHasComment, cfHyperlink, cfMerged);

  {@@ Set of cell flags }
  TsCellFlags = set of TsCellFlag;

  {@@ Record combining a cell's row and column indexes }
  TsCellCoord = record
    Row, Col: Cardinal;
  end;

  {@@ Record combining row and column cornder indexes of a range of cells }
  TsCellRange = record
    Row1, Col1, Row2, Col2: Cardinal;
  end;
  PsCellRange = ^TsCellRange;

  {@@ Array with cell ranges }
  TsCellRangeArray = array of TsCellRange;

  {@@ Options for sorting }
  TsSortOption = (ssoDescending, ssoCaseInsensitive);
  {@@ Set of options for sorting }
  TsSortOptions = set of TsSortOption;

  {@@ Sort priority }
  TsSortPriority = (spNumAlpha, spAlphaNum);   // spNumAlpha: Number < Text

  {@@ Sort key: sorted column or row index and sort direction }
  TsSortKey = record
    ColRowIndex: Integer;
    Options: TsSortOptions;
  end;

  {@@ Array of sort keys for multiple sorting criteria }
  TsSortKeys = array of TsSortKey;

  {@@ Complete set of sorting parameters
    @param SortByCols  If true sorting is top-down, otherwise left-right
    @param Priority    Determines whether numbers are before or after text.
    @param SortKeys    Array of sorting indexes and sorting directions }
  TsSortParams = record
    SortByCols: Boolean;
    Priority: TsSortPriority;
    Keys: TsSortKeys;
  end;

  {@@ Record containing all details for cell formatting }
  TsCellFormat = record
    Name: String;
    ID: Integer;
    UsedFormattingFields: TsUsedFormattingFields;
    FontIndex: Integer;
    TextRotation: TsTextRotation;
    HorAlignment: TsHorAlignment;
    VertAlignment: TsVertAlignment;
    Border: TsCellBorders;
    BorderStyles: TsCelLBorderStyles;
    Background: TsFillPattern;
    NumberFormatIndex: Integer;
    // next two are deprecated...
    NumberFormat: TsNumberFormat;
    NumberFormatStr: String;
  end;

  {@@ Pointer to a format record }
  PsCellFormat = ^TsCellFormat;

  {@@ Specialized list for format records }
  TsCellFormatList = class(TFPList)
  private
    FAllowDuplicates: Boolean;
    function GetItem(AIndex: Integer): PsCellFormat;
    procedure SetItem(AIndex: Integer; const AValue: PsCellFormat);
  public
    constructor Create(AAllowDuplicates: Boolean);
    destructor Destroy; override;
    function Add(const AItem: TsCellFormat): Integer; overload;
    function Add(AItem: PsCellFormat): Integer; overload;
    procedure Clear;
    procedure Delete(AIndex: Integer);
    function FindIndexOfID(ID: Integer): Integer;
    function FindIndexOfName(AName: String): Integer;
    function IndexOf(const AItem: TsCellFormat): Integer; overload;
    property Items[AIndex: Integer]: PsCellFormat read GetItem write SetItem; default;
  end;

  {@@ Pointer to a TCell record }
  PCell = ^TCell;

  {@@ Cell structure for TsWorksheet
      The cell record contains information on the location of the cell (row and
      column index), on the value contained (number, date, text, ...), on
      formatting, etc.

      Never suppose that all *Value fields are valid,
      only one of the ContentTypes is valid. For other fields
      use TWorksheet.ReadAsUTF8Text and similar methods

      @see ReadAsUTF8Text }
  TCell = record
    { Location of the cell }
    Row: Cardinal; // zero-based
    Col: Cardinal; // zero-based
    Worksheet: Pointer;   // Must be cast to TsWorksheet when used  (avoids circular unit reference)
    { Status flags }
    Flags: TsCellFlags;
    { Index of format record in the workbook's FCellFormatList }
    FormatIndex: Integer;
    { Cell content }
    UTF8StringValue: String;   // Strings cannot be part of a variant record
    FormulaValue: String;
    case ContentType: TCellContentType of  // variant part must be at the end
      cctEmpty      : ();      // has no data at all
      cctFormula    : ();      // FormulaValue is outside the variant record
      cctNumber     : (Numbervalue: Double);
      cctUTF8String : ();      // UTF8StringValue is outside the variant record
      cctDateTime   : (DateTimevalue: TDateTime);
      cctBool       : (BoolValue: boolean);
      cctError      : (ErrorValue: TsErrorValue);
  end;

  TsPageOrientation = (spoPortrait, spoLandscape);

  TsPrintOption = (poPrintGridLines, poPrintHeaders, poPrintPagesByRows,
    poMonochrome, poDraftQuality, poPrintCellComments, poDefaultOrientation,
    poUseStartPageNumber, poCommentsAtEnd, poHorCentered, poVertCentered,
    poDifferentOddEven, poDifferentFirst, poFitPages);

  TsPrintOptions = set of TsPrintOption;

  TsPageLayout = record      // all lengths in mm
    Orientation: TsPageOrientation;
    PageWidth: Double;       // for "normal" orientation (mostly portrait)
    PageHeight: Double;
    LeftMargin: Double;
    RightMargin: Double;
    TopMargin: Double;
    BottomMargin: Double;
    HeaderMargin: Double;
    FooterMargin: Double;
    StartPageNumber: Integer;
    ScalingFactor: Integer;  // in percent
    FitWidthToPages: Integer;
    FitHeightToPages: Integer;
    Copies: Integer;
    Options: TsPrintOptions;
    { Headers and footers are in Excel syntax:
      - left/center/right sections begin with &L / &C / &R
      - page number: &P
      - page count: &N
      - current date: &D
      - current time:  &T
      - sheet name: &A
      - file name without path: &F
      - file path without file name: &Z
      - bold/italic/underlining/double underlining/strike out/shadowed/
        outlined/superscript/subscript on/off:
          &B / &I / &U / &E / &S / &H
          &O / &X / &Y
      There can be three headers/footers, for first ([0]) page and
      odd ([1])/even ([2]) page numbers.
      This is activated by Options poDifferentOddEven and poDifferentFirst.
      Array index 1 contains the strings if these options are not used. }
    Headers: array[0..2] of string;
    Footers: array[0..2] of string;
  end;

  PsPageLayout = ^TsPageLayout;

const
  {@@ Indexes to be used for the various headers and footers }
  HEADER_FOOTER_INDEX_FIRST   = 0;
  HEADER_FOOTER_INDEX_ODD     = 1;
  HEADER_FOOTER_INDEX_EVEN    = 2;
  HEADER_FOOTER_INDEX_ALL     = 1;

function BuildFormatStringFromSection(const ASection: TsNumFormatSection): String;


implementation

uses
  StrUtils;

{ TsCellFormatList }

constructor TsCellFormatList.Create(AAllowDuplicates: Boolean);
begin
  inherited Create;
  FAllowDuplicates := AAllowDuplicates;
end;

destructor TsCellFormatList.Destroy;
begin
  Clear;
  inherited;
end;

function TsCellFormatList.Add(const AItem: TsCellFormat): Integer;
var
  P: PsCellFormat;
begin
  if FAllowDuplicates then
    Result := -1
  else
    Result := IndexOf(AItem);
  if Result = -1 then begin
    New(P);
    P^.Name := AItem.Name;
    P^.ID := AItem.ID;
    P^.UsedFormattingFields := AItem.UsedFormattingFields;
    P^.FontIndex := AItem.FontIndex;
    P^.TextRotation := AItem.TextRotation;
    P^.HorAlignment := AItem.HorAlignment;
    P^.VertAlignment := AItem.VertAlignment;
    P^.Border := AItem.Border;
    P^.BorderStyles := AItem.BorderStyles;
    P^.Background := AItem.Background;
    P^.NumberFormatIndex := AItem.NumberFormatIndex;
    P^.NumberFormat := AItem.NumberFormat;
    P^.NumberFormatStr := AItem.NumberFormatStr;
    Result := inherited Add(P);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds a pointer to a FormatRecord to the list. Allows nil for the predefined
  formats which are not stored in the file.
-------------------------------------------------------------------------------}
function TsCellFormatList.Add(AItem: PsCellFormat): Integer;
begin
  if AItem = nil then
    Result := inherited Add(AItem)
  else
    Result := Add(AItem^);
end;

procedure TsCellFormatList.Clear;
var
  i: Integer;
begin
  for i:=Count-1 downto 0 do
    Delete(i);
  inherited;
end;

procedure TsCellFormatList.Delete(AIndex: Integer);
var
  P: PsCellFormat;
begin
  P := GetItem(AIndex);
  if P <> nil then
    Dispose(P);
  inherited Delete(AIndex);
end;

function TsCellFormatList.GetItem(AIndex: Integer): PsCellFormat;
begin
  Result := inherited Items[AIndex];
end;

function TsCellFormatList.FindIndexOfID(ID: Integer): Integer;
var
  P: PsCellFormat;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P <> nil) and (P^.ID = ID) then
      exit;
  end;
  Result := -1;
end;

function TsCellFormatList.FindIndexOfName(AName: String): Integer;
var
  P: PsCellFormat;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P <> nil) and (P^.Name = AName) then
      exit;
  end;
  Result := -1;
end;

function TsCellFormatList.IndexOf(const AItem: TsCellFormat): Integer;
var
  P: PsCellFormat;
  equ: Boolean;
  b: TsCellBorder;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P = nil) then continue;

    if (P^.UsedFormattingFields <> AItem.UsedFormattingFields) then continue;

    if (uffFont in AItem.UsedFormattingFields) then
      if (P^.FontIndex) <> (AItem.FontIndex) then continue;

    if (uffTextRotation in AItem.UsedFormattingFields) then
      if (P^.TextRotation <> AItem.TextRotation) then continue;

    if (uffHorAlign in AItem.UsedFormattingFields) then
      if (P^.HorAlignment <> AItem.HorAlignment) then continue;

    if (uffVertAlign in AItem.UsedFormattingFields) then
      if (P^.VertAlignment <> AItem.VertAlignment) then continue;

    if (uffBorder in AItem.UsedFormattingFields) then
      if (P^.Border <> AItem.Border) then continue;

    // Border styles can be set even if borders are not used --> don't check uffBorder!
    equ := true;
    for b in AItem.Border do begin
      if (P^.BorderStyles[b].LineStyle <> AItem.BorderStyles[b].LineStyle) or
         (P^.BorderStyles[b].Color <> Aitem.BorderStyles[b].Color)
      then begin
        equ := false;
        break;
      end;
    end;
    if not equ then continue;

    if (uffBackground in AItem.UsedFormattingFields) then begin
      if (P^.Background.Style <> AItem.Background.Style) then continue;
      if (P^.Background.BgColor <> AItem.Background.BgColor) then continue;
      if (P^.Background.FgColor <> AItem.Background.FgColor) then continue;
    end;

    if (uffNumberFormat in AItem.UsedFormattingFields) then begin
      if (P^.NumberFormatIndex <> AItem.NumberFormatIndex) then continue;
      if (P^.NumberFormat <> AItem.NumberFormat) then continue;
      if (P^.NumberFormatStr <> AItem.NumberFormatStr) then continue;
    end;

    // If we arrive here then the format records match.
    exit;
  end;

  // We get here if no record matches
  Result := -1;
end;

procedure TsCellFormatList.SetItem(AIndex: Integer; const AValue: PsCellFormat);
begin
  inherited Items[AIndex] := AValue;
end;


{ Creates a format string for the given number format section section.
  The format string is created according to Excel convention (which is used by
  ODS as well }
function BuildFormatStringFromSection(const ASection: TsNumFormatSection): String;
var
  element: TsNumFormatElement;
  i, n: Integer;
begin
  Result := '';

  for i := 0 to High(ASection.Elements)  do begin
    element := ASection.Elements[i];
    case element.Token of
      nftGeneral:
        Result := Result + 'General';
      nftIntOptDigit, nftOptDecs, nftFracNumOptDigit, nftFracDenomOptDigit:
        if element.IntValue > 0 then
          Result := Result + DupeString('#', element.IntValue);
      nftIntZeroDigit, nftZeroDecs, nftFracNumZeroDigit, nftFracDenomZeroDigit, nftExpDigits:
        if element.IntValue > 0 then
          Result := result + DupeString('0', element.IntValue);
      nftIntSpaceDigit, nftSpaceDecs, nftFracNumSpaceDigit, nftFracDenomSpaceDigit:
        if element.Intvalue > 0 then
          Result := result + DupeString('?', element.IntValue);
      nftFracDenom:
        Result := Result + IntToStr(element.IntValue);
      nftIntTh:
        case element.Intvalue of
          0: Result := Result + '#,###';
          1: Result := Result + '#,##0';
          2: Result := Result + '#,#00';
          3: Result := Result + '#,000';
        end;
      nftDecSep:
        Result := Result + '.';
      nftThSep:
        Result := Result + ',';
      nftFracSymbol:
        Result := Result + '/';
      nftPercent:
        Result := Result + '%';
      nftFactor:
        if element.IntValue <> 0 then
        begin
          n := element.IntValue;
          while (n > 0) do
          begin
            Result := Result + ',';
            dec(n);
          end;
        end;
      nftSpace:
        Result := Result + ' ';
      nftText:
        if element.TextValue <> '' then result := Result + '"' + element.TextValue + '"';
      nftYear:
        Result := Result + DupeString('Y', element.IntValue);
      nftMonth:
        Result := Result + DupeString('M', element.IntValue);
      nftDay:
        Result := Result + DupeString('D', element.IntValue);
      nftHour:
        if element.IntValue < 0
          then Result := Result + '[' + DupeString('h', -element.IntValue) + ']'
          else Result := Result + DupeString('h', element.IntValue);
      nftMinute:
        if element.IntValue < 0
          then Result := result + '[' + DupeString('m', -element.IntValue) + ']'
          else Result := Result + DupeString('m', element.IntValue);
      nftSecond:
        if element.IntValue < 0
          then Result := Result + '[' + DupeString('s', -element.IntValue) + ']'
          else Result := Result + DupeString('s', element.IntValue);
      nftMilliseconds:
        Result := Result + DupeString('0', element.IntValue);
      nftSign, nftSignBracket, nftExpChar, nftExpSign, nftAMPM, nftDateTimeSep:
        if element.TextValue <> '' then Result := Result + element.TextValue;
      nftCurrSymbol:
        if element.TextValue <> '' then
          Result := Result + '[$' + element.TextValue + ']';
      nftEscaped:
        if element.TextValue <> '' then
          Result := Result + '\' + element.TextValue;
      nftTextFormat:
        if element.TextValue <> '' then
          Result := Result + element.TextValue;
      nftRepeat:
        if element.TextValue <> '' then Result := Result + '*' + element.TextValue;
      nftColor:
        case element.IntValue of
          scBlack  : Result := '[black]';
          scWhite  : Result := '[white]';
          scRed    : Result := '[red]';
          scBlue   : Result := '[blue]';
          scGreen  : Result := '[green]';
          scYellow : Result := '[yellow]';
          scMagenta: Result := '[magenta]';
          scCyan   : Result := '[cyan]';
          else       Result := Format('[Color%d]', [element.IntValue]);
        end;
    end;
  end;
end;


{ TsNumFormatParams }

procedure TsNumFormatParams.DeleteElement(ASectionIndex, AElementIndex: Integer);
var
  i, n: Integer;
begin
  with Sections[ASectionIndex] do
  begin
    n := Length(Elements);
    for i:=AElementIndex+1 to n-1 do
      Elements[i-1] := Elements[i];
    SetLength(Elements, n-1);
  end;
end;


function TsNumFormatParams.GetNumFormat: TsNumberFormat;
begin
  Result := nfCustom;
  case Length(Sections) of
    0: Result := nfGeneral;
    1: Result := Sections[0].NumFormat;
    2: if (Sections[0].NumFormat = Sections[1].NumFormat) and
          (Sections[0].NumFormat in [nfCurrency, nfCurrencyRed])
       then
         Result := Sections[0].NumFormat;
    3: if (Sections[0].NumFormat = Sections[1].NumFormat) and
          (Sections[1].NumFormat = Sections[2].NumFormat) and
          (Sections[0].NumFormat in [nfCurrency, nfCurrencyRed])
       then
         Result := Sections[0].NumFormat;
  end;
end;

function TsNumFormatParams.GetNumFormatStr: String;
var
  i: Integer;
begin
  if Length(Sections) > 0 then begin
    Result := BuildFormatStringFromSection(Sections[0]);
    for i := 1 to High(Sections) do
      Result := Result + ';' + BuildFormatStringFromSection(Sections[i]);
  end else
    Result := '';
end;

procedure TsNumFormatParams.InsertElement(ASectionIndex, AElementIndex: Integer;
  AToken: TsNumFormatToken);
var
  i, n: Integer;
begin
  with Sections[ASectionIndex] do
  begin
    n := Length(Elements);
    SetLength(Elements, n+1);
    for i:=n-1 downto AElementIndex do
      Elements[i+1] := Elements[i];
    Elements[AElementIndex].Token := AToken;
  end;
end;

function TsNumFormatParams.SectionsEqualTo(ASections: TsNumFormatSections): Boolean;
var
  i, j: Integer;
begin
  Result := false;
  if Length(ASections) <> Length(Sections) then
    exit;
  for i := 0 to High(Sections) do begin
    if Length(Sections[i].Elements) <> Length(ASections[i].Elements) then
      exit;

    for j:=0 to High(Sections[i].Elements) do
    begin
      if Sections[i].Elements[j].Token <> ASections[i].Elements[j].Token then
        exit;

      if Sections[i].NumFormat <> ASections[i].NumFormat then
        exit;
      if Sections[i].Decimals <> ASections[i].Decimals then
        exit;
      {
      if Sections[i].Factor <> ASections[i].Factor then
        exit;
        }
      if Sections[i].FracInt <> ASections[i].FracInt then
        exit;
      if Sections[i].FracNumerator <> ASections[i].FracNumerator then
        exit;
      if Sections[i].FracDenominator <> ASections[i].FracDenominator then
        exit;
      if Sections[i].CurrencySymbol <> ASections[i].CurrencySymbol then
        exit;
      if Sections[i].Color <> ASections[i].Color then
        exit;

      case Sections[i].Elements[j].Token of
        nftText, nftThSep, nftDecSep, nftDateTimeSep,
        nftAMPM, nftSign, nftSignBracket,
        nftExpChar, nftExpSign, nftPercent, nftFracSymbol, nftCurrSymbol,
        nftCountry, nftSpace, nftEscaped, nftRepeat, nftEmptyCharWidth,
        nftTextFormat:
          if Sections[i].Elements[j].TextValue <> ASections[i].Elements[j].TextValue
            then exit;

        nftYear, nftMonth, nftDay,
        nftHour, nftMinute, nftSecond, nftMilliseconds,
        nftMonthMinute,
        nftIntOptDigit, nftIntZeroDigit, nftIntSpaceDigit, nftIntTh,
        nftZeroDecs, nftOptDecs, nftSpaceDecs, nftExpDigits, nftFactor,
        nftFracNumOptDigit, nftFracNumSpaceDigit, nftFracNumZeroDigit,
        nftFracDenomOptDigit, nftFracDenomSpaceDigit, nftFracDenomZeroDigit,
        nftColor:
          if Sections[i].Elements[j].IntValue <> ASections[i].Elements[j].IntValue
            then exit;

        nftCompareOp, nftCompareValue:
          if Sections[i].Elements[j].FloatValue <> ASections[i].Elements[j].FloatValue
            then exit;
      end;
    end;
  end;
  Result := true;
end;

procedure TsNumFormatParams.SetCurrSymbol(AValue: String);
var
  section: TsNumFormatSection;
  s, el: Integer;
begin
  for s:=0 to High(Sections) do
  begin
    section := Sections[s];
    if (nfkCurrency in section.Kind) then
    begin
      section.CurrencySymbol := AValue;
      for el := 0 to High(section.Elements) do
        if section.Elements[el].Token = nftCurrSymbol then
          section.Elements[el].Textvalue := AValue;
    end;
  end;
end;

procedure TsNumFormatParams.SetDecimals(AValue: byte);
var
  section: TsNumFormatSection;
  s, el: Integer;
begin
  for s := 0 to High(Sections) do
  begin
    section := Sections[s];
    if section.Kind * [nfkFraction, nfkDate, nfkTime] <> [] then
      Continue;
    section.Decimals := AValue;
    for el := High(section.Elements) downto 0 do
      case section.Elements[el].Token of
        nftZeroDecs:
          section.Elements[el].Intvalue := AValue;
        nftOptDecs, nftSpaceDecs:
          DeleteElement(s, el);
      end;
  end;
end;

procedure TsNumFormatParams.SetNegativeRed(AEnable: Boolean);
var
  el: Integer;
begin
  // Enable negative-value color
  if AEnable then
  begin
    if Length(Sections) = 1 then begin
      SetLength(Sections, 2);
      Sections[1] := Sections[0];
      InsertElement(1, 0, nftColor);
      Sections[1].Elements[0].Intvalue := scRed;
      InsertElement(1, 1, nftSign);
      Sections[1].Elements[1].TextValue := '-';
    end else
    begin
      if not (nfkHasColor in Sections[1].Kind) then
        InsertElement(1, 0, nftColor);
      for el := 0 to High(Sections[1].Elements) do
        if Sections[1].Elements[el].Token = nftColor then
          Sections[1].Elements[el].IntValue := scRed;
    end;
    Sections[1].Kind := Sections[1].Kind + [nfkHasColor];
    Sections[1].Color := scRed;
  end else
  // Disable negative-value color
  if Length(Sections) >= 2 then
  begin
    Sections[1].Kind := Sections[1].Kind - [nfkHasColor];
    Sections[1].Color := scBlack;
    for el := High(Sections[1].Elements) downto 0 do
      if Sections[1].Elements[el].Token = nftColor then
        DeleteElement(1, el);
  end;
end;

procedure TsNumFormatParams.SetThousandSep(AEnable: Boolean);
var
  section: TsNumFormatSection;
  s, el: Integer;
  replaced: Boolean;
begin
  for s := 0 to High(Sections) do
  begin
    section := Sections[s];
    replaced := false;
    for el := High(section.Elements) downto 0 do
    begin
      if AEnable then
      begin
        if section.Elements[el].Token in [nftIntOptDigit, nftIntSpaceDigit, nftIntZeroDigit] then
        begin
          if replaced then
            DeleteElement(s, el)
          else begin
            section.Elements[el].Token := nftIntTh;
            Include(section.Kind, nfkHasThSep);
            replaced := true;
          end;
        end;
      end else
      begin
        if section.Elements[el].Token = nftIntTh then begin
          section.Elements[el].Token := nftIntZeroDigit;
          Exclude(section.Kind, nfkHasThSep);
          break;
        end;
      end;
    end;
  end;
end;

end.

