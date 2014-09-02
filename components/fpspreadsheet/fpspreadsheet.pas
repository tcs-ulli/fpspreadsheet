{
fpspreadsheet.pas

Writes a spreadsheet document

AUTHORS: Felipe Monteiro de Carvalho
}
unit fpspreadsheet;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
 {$ifdef UNIX}{$ifndef DARWIN}{$ifndef FPS_DONT_USE_CLOCALE}
  clocale,
 {$endif}{$endif}{$endif}
  Classes, SysUtils, fpimage, AVL_Tree, avglvltree, lconvencoding;

type
  {@@ File formats suppored by fpspreadsheet }
  TsSpreadsheetFormat = (sfExcel2, {sfExcel3, sfExcel4,} sfExcel5, sfExcel8,
   sfOOXML, sfOpenDocument, sfCSV, sfWikiTable_Pipes, sfWikiTable_WikiMedia);

  {@@ Record collection limitations of a particular file format }
  TsSpreadsheetFormatLimitations = record
    MaxRowCount: Cardinal;
    MaxColCount: Cardinal;
    MaxPaletteSize: Cardinal;
  end;

const
  { Default extensions }
  STR_EXCEL_EXTENSION = '.xls';
  STR_OOXML_EXCEL_EXTENSION = '.xlsx';
  STR_OPENDOCUMENT_CALC_EXTENSION = '.ods';
  STR_COMMA_SEPARATED_EXTENSION = '.csv';
  STR_WIKITABLE_PIPES = '.wikitable_pipes';
  STR_WIKITABLE_WIKIMEDIA = '.wikitable_wikimedia';

  MAX_COL_COUNT = 65535;

type

  {@@ Possible encodings for a non-unicode encoded text }
  TsEncoding = (
    seLatin1,
    seLatin2,
    seCyrillic,
    seGreek,
    seTurkish,
    seHebrew,
    seArabic
    );
        (*
  {@@ Describes a formula

    Supported syntax:
    <pre>
    =A1+B1+C1/D2...  - Array with simple mathematical operations
    =SUM(A1:D1)      - SUM operation in a interval
    </pre>
  }
  TsFormula = record
    FormulaStr: string;
    DoubleValue: double;
  end;
          *)
  {@@ Tokens to identify the elements in an expanded formula.

   See http://www.techonthenet.com/excel/formulas/ for an explanation of
   meaning and parameters of each formula

   NOTE: When adding or rearranging items:
   - make sure that the subtypes TOperandTokens, TBasicOperationTokens and TFuncTokens
     are complete
   - make sure to keep the FEProps table in sync
   - make sure to keep the TokenID table
     in TsSpreadBIFFWriter.FormulaElementKindToExcelTokenID, unit xlscommon,
     in sync
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
  (*
  TFEKind = (
    { Basic operands }
    fekCell, fekCellRef, fekCellRange, fekCellOffset, fekNum, fekInteger,
    fekString, fekBool, fekErr, fekMissingArg,
    { Basic operations }
    fekAdd, fekSub, fekMul, fekDiv, fekPercent, fekPower, fekUMinus, fekUPlus,
    fekConcat,  // string concatenation
    fekEqual, fekGreater, fekGreaterEqual, fekLess, fekLessEqual, fekNotEqual,
    fekParen,
    { Built-in/Worksheet Functions}
    // math
    fekABS, fekACOS, fekACOSH, fekASIN, fekASINH, fekATAN, fekATANH,
    fekCOS, fekCOSH, fekDEGREES, fekEXP, fekINT, fekLN, fekLOG,
    fekLOG10, fekPI, fekRADIANS, fekRAND, fekROUND,
    fekSIGN, fekSIN, fekSINH, fekSQRT,
    fekTAN, fekTANH,
    // date/time
    fekDATE, fekDATEDIF, fekDATEVALUE, fekDAY, fekHOUR, fekMINUTE, fekMONTH,
    fekNOW, fekSECOND, fekTIME, fekTIMEVALUE, fekTODAY, fekWEEKDAY, fekYEAR,
    // statistical
    fekAVEDEV, fekAVERAGE, fekBETADIST, fekBETAINV, fekBINOMDIST, fekCHIDIST,
    fekCHIINV, fekCOUNT, fekCOUNTA, fekCOUNTBLANK, fekCOUNTIF,
    fekMAX, fekMEDIAN, fekMIN, fekPERMUT, fekPOISSON, fekPRODUCT,
    fekSTDEV, fekSTDEVP, fekSUM, fekSUMIF, fekSUMSQ, fekVAR, fekVARP,
    // financial
    fekFV, fekNPER, fekPMT, fekPV, fekRATE,
    // logical
    fekAND, fekFALSE, fekIF, fekNOT, fekOR, fekTRUE,
    // string
    fekCHAR, fekCODE, fekLEFT, fekLOWER, fekMID, fekPROPER, fekREPLACE, fekRIGHT,
    fekSUBSTITUTE, fekTRIM, fekUPPER,
    // lookup/reference
    fekCOLUMN, fekCOLUMNS, fekROW, fekROWS,
    // info
    fekCELLINFO, fekINFO, fekIsBLANK, fekIsERR, fekIsERROR,
    fekIsLOGICAL, fekIsNA, fekIsNONTEXT, fekIsNUMBER, fekIsRef, fekIsTEXT,
    fekValue,
    { Other operations }
    fekOpSUM {Unary sum operation. Note: CANNOT be used for summing sell contents; use fekSUM}
    );
    *)
  {@@ These tokens identify operands in RPN formulas. }
  TOperandTokens = fekCell..fekMissingArg;

  {@@ These tokens identify basic operations in RPN formulas. }
  TBasicOperationTokens = fekAdd..fekParen;

  (*
  {@@ These tokens identify spreadsheet functions in RPN formulas. }
  TFuncTokens = fekAbs..fekOpSum;
    *)

  {@@ Flags to mark the address or a cell or a range of cells to be absolute
      or relative. They are used in the set TsRelFlags. }
  TsRelFlag = (rfRelRow, rfRelCol, rfRelRow2, rfRelCol2);

  {@@ Flags to mark the address of a cell or a range of cells to be absolute
      or relative. It is a set consisting of TsRelFlag elements. }
  TsRelFlags = set of TsRelFlag;

  {@@ Elements of an expanded formula.
    Note: If ElementKind is fekCellOffset, "Row" and "Col" have to be cast
          to signed integers! }
  TsFormulaElement = record
    ElementKind: TFEKind;
    Row, Row2: Cardinal; // zero-based
    Col, Col2: Cardinal; // zero-based
    Param1, Param2: Word; // Extra parameters
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

  {@@ Describes the type of content in a cell of a TsWorksheet }
  TCellContentType = (cctEmpty, cctFormula, cctNumber, cctUTF8String,
    cctDateTime, cctBool, cctError);

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
  TsUsedFormattingField = (uffTextRotation, uffFont, uffBold, uffBorder,
    uffBackgroundColor, uffNumberFormat, uffWordWrap,
    uffHorAlign, uffVertAlign
  );

  {@@ Describes which formatting fields are active }
  TsUsedFormattingFields = set of TsUsedFormattingField;

  {@@ Number/cell formatting. Only uses a subset of the default formats,
      enough to be able to read/write date/time values.
      nfCustom allows to apply a format string directly. }
  TsNumberFormat = (
    // general-purpose for all numbers
    nfGeneral,
    // numbers
    nfFixed, nfFixedTh, nfExp, nfPercentage,
    // currency
    nfCurrency, nfCurrencyRed,
    // dates and times
    nfShortDateTime, {nfFmtDateTime, }nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfTimeInterval,
    // other (format string goes directly into the file)
    nfCustom);

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
  ccfMVC     = 5;    // -1000$
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
      the former case.
  }
  TsTextRotation = (trHorizontal, rt90DegreeClockwiseRotation,
    rt90DegreeCounterClockwiseRotation, rtStacked);

  {@@ Indicates horizontal text alignment in cells }
  TsHorAlignment = (haDefault, haLeft, haCenter, haRight);

  {@@ Indicates vertical text alignment in cells }
  TsVertAlignment = (vaDefault, vaTop, vaCenter, vaBottom);

  {@@
    Colors in fpspreadsheet are given as indices into a palette.
    Use the workbook's GetPaletteColor to determine the color rgb value as
    little-endian (with "r" being the low-value byte, in agreement with TColor).
    The data type for rgb values is TsColorValue. }
  TsColor = Word;

{@@
  These are some constants for color indices into the default palette.
  Note, however, that if a different palette is used there may be more colors,
  and the names of the color constants may no longer be correct.
}
const
  scBlack = $00;
  scWhite = $01;
  scRed = $02;
  scGreen = $03;
  scBlue = $04;
  scYellow = $05;
  scMagenta = $06;
  scCyan = $07;
  scDarkRed = $08;
  scDarkGreen = $09;
  scDarkBlue = $0A;    scNavy = $0A;
  scOlive = $0B;
  scPurple = $0C;
  scTeal = $0D;
  scSilver = $0E;
  scGrey = $0F;        scGray = $0F;       // redefine to allow different spelling
  scGrey10pct = $10;   scGray10pct = $10;
  scGrey20pct = $11;   scGray20pct = $11;
  scOrange = $12;
  scDarkbrown = $13;
  scBrown = $14;
  scBeige = $15;
  scWheat = $16;

  // not sure - but I think the mechanism with scRGBColor is not working...
  // Will be removed sooner or later...
  scRGBColor = $FFFF;

  scTransparent = $FFFE;
  scNotDefined = $FFFF;

type
  {@@ Data type for rgb color values }
  TsColorValue = DWord;

  {@@ Palette of color values. A "color value" is a DWord value containing
      rgb colors. }
  TsPalette = array[0..0] of TsColorValue;
  PsPalette = ^TsPalette;

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
    {@@ Text color given by the index into the workbook's color palette }
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
  {@@ Identifier for a compare operation }
  TsCompareOperation = (coNotUsed,
    coEqual, coNotEqual, coLess, coGreater, coLessEqual, coGreaterEqual
  );

  {@@ State flags while calculating formulas }
  TsCalcState = (csNotCalculated, csCalculating, csCalculated);

  {@@ Record combining a cell's row and column indexes }
  TsCellCoord = record
    Row, Col: Cardinal;
  end;

  {@@ Record combining row and column cornder indexes of a range of cells }
  TsCellRange = record
    Row1, Col1, Row2, Col2: Cardinal;
  end;

  {@@ Pointer to a TCell record }
  PCell = ^TCell;

  {@@ Cell structure for TsWorksheet
      The cell record contains information on the location of the cell (row and
      column index), on the value contained (number, date, text, ...), and on
      formatting.

      Never suppose that all *Value fields are valid,
      only one of the ContentTypes is valid. For other fields
      use TWorksheet.ReadAsUTF8Text and similar methods

      @see ReadAsUTF8Text }
  TCell = record
    Col: Cardinal; // zero-based
    Row: Cardinal; // zero-based
    ContentType: TCellContentType;
    { Possible values for the cells }
    FormulaValue: string;
//    RPNFormulaValue: TsRPNFormula;
    NumberValue: double;
    UTF8StringValue: ansistring;
    DateTimeValue: TDateTime;
    BoolValue: Boolean;
    ErrorValue: TsErrorValue;
    SharedFormulaBase: PCell;     // Cell containing the shared formula
    { Formatting fields }
    { When adding/deleting formatting fields don't forget to update CopyFormat! }
    UsedFormattingFields: TsUsedFormattingFields;
    FontIndex: Integer;
    TextRotation: TsTextRotation;
    HorAlignment: TsHorAlignment;
    VertAlignment: TsVertAlignment;
    Border: TsCellBorders;
    BorderStyles: TsCelLBorderStyles;
    BackgroundColor: TsColor;
    NumberFormat: TsNumberFormat;
    NumberFormatStr: String;
    RGBBackgroundColor: TFPColor; // only valid if BackgroundColor=scRGBCOLOR
    { Status flags }
    CalcState: TsCalcState;
  end;

const
  // Takes account of effect of cell margins on row height by adding this
  // value to the nominal row height. Note that this is an empirical value and may be wrong.
  ROW_HEIGHT_CORRECTION = 0.2;

type
  {@@ The record TRow contains information about a spreadsheet row:
    @param Row   The index of the row (beginning with 0)
    @param Height  The height of the row (expressed as lines count of the default font)
   Only rows with heights that cannot be derived from the font height have a
   row record. }
  TRow = record
    Row: Cardinal;
    Height: Single;  // in "lines"
  end;

  {@@ Pointer to a TRow record }
  PRow = ^TRow;

  {@@ The record TCol contains information about a spreadsheet column:
   @param Col    The index of the column (beginning with 0)
   @param Width  The width of the column (expressed in character count of the "0" character of the default font.
   Only columns with non-default widths have a column record. }
  TCol = record
    Col: Cardinal;
    Width: Single; // in "characters". Excel uses the width of char "0" in 1st font
  end;

  {@@ Pointer to a TCol record }
  PCol = ^TCol;

  {@@ WSorksheet user interface options:
    @param soShowGridLines  Show or hide the grid lines in the spreadsheet
    @param soShowHeaders    Show or hide the column or row headers of the spreadsheet
    @param soHasFrozenPanes If set a number of rows and columns of the spreadsheet
                            is fixed and does not scroll. The number is defined by
                            LeftPaneWidth and TopPaneHeight. }
  TsSheetOption = (soShowGridLines, soShowHeaders, soHasFrozenPanes);

  {@@ Set of user interface options
    @ see TsSheetOption }
  TsSheetOptions = set of TsSheetOption;

type

  TsCustomSpreadReader = class;
  TsCustomSpreadWriter = class;
  TsWorkbook = class;

  { TsWorksheet }

  {@@ This event fires whenever a cell value or cell formatting changes. It is
    handled by TsWorksheetGrid to update the grid. }
  TsCellEvent = procedure (Sender: TObject; ARow, ACol: Cardinal) of object;

  {@@ The worksheet contains a list of cells and provides a variety of methods
    to read or write data to the cells, or to change their formatting. }
  TsWorksheet = class
  private
    FWorkbook: TsWorkbook;
    FCells: TAvlTree; // Items are TCell
    FCurrentNode: TAVLTreeNode; // For GetFirstCell and GetNextCell
    FRows, FCols: TIndexedAVLTree; // This lists contain only rows or cols with styles different from default
    FLeftPaneWidth: Integer;
    FTopPaneHeight: Integer;
    FOptions: TsSheetOptions;
    FFirstRowIndex: Cardinal;
    FFirstColIndex: Cardinal;
    FLastRowIndex: Cardinal;
    FLastColIndex: Cardinal;
    FOnChangeCell: TsCellEvent;
    FOnChangeFont: TsCellEvent;

    { Setter/Getter }
    function GetFormatSettings: TFormatSettings;

    { Callback procedures called when iterating through all cells }
    procedure CalcFormulaCallback(data, arg: Pointer);
    procedure CalcStateCallback(data, arg: Pointer);
    procedure InsertColCallback(data, arg: Pointer);
    procedure InsertRowCallback(data, arg: Pointer);
    procedure RemoveCallback(data, arg: pointer);

  protected
    function CellUsedInFormula(ARow, ACol: Cardinal): Boolean;

    procedure ChangedCell(ARow, ACol: Cardinal);
    procedure ChangedFont(ARow, ACol: Cardinal);

  public
    {@@ Name of the sheet. In the popular spreadsheet applications this is
      displayed at the tab of the sheet. }
    Name: string;

    { Base methods }
    constructor Create;
    destructor Destroy; override;

    { Utils }
    class function CellPosToText(ARow, ACol: Cardinal): string;
    procedure RemoveAllCells;
    procedure UpdateCaches;

    { Reading of values }
    function  ReadAsUTF8Text(ARow, ACol: Cardinal): ansistring; overload;
    function  ReadAsUTF8Text(ACell: PCell): ansistring; overload;
    function  ReadAsNumber(ARow, ACol: Cardinal): Double; overload;
    function  ReadAsNumber(ACell: PCell): Double; overload;
    function  ReadAsDateTime(ARow, ACol: Cardinal; out AResult: TDateTime): Boolean; overload;
    function  ReadAsDateTime(ACell: PCell; out AResult: TDateTime): Boolean; overload;
    function  ReadFormulaAsString(ACell: PCell): String;
    function  ReadNumericValue(ACell: PCell; out AValue: Double): Boolean;

    { Reading of cell attributes }
    function GetNumberFormatAttributes(ACell: PCell; out ADecimals: Byte;
      out ACurrencySymbol: String): Boolean;
    function  ReadUsedFormatting(ARow, ACol: Cardinal): TsUsedFormattingFields;
    function  ReadBackgroundColor(ARow, ACol: Cardinal): TsColor;

    { Writing of values }
    function WriteBlank(ARow, ACol: Cardinal): PCell; overload;
    procedure WriteBlank(ACell: PCell); overload;

    function WriteBoolValue(ARow, ACol: Cardinal; AValue: Boolean): PCell; overload;
    procedure WriteBoolValue(ACell: PCell; AValue: Boolean); overload;

    function WriteCellValueAsString(ARow, ACol: Cardinal; AValue: String): PCell; overload;
    procedure WriteCellValueAsString(ACell: PCell; AValue: String); overload;

    function WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
      AFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = 2;
      ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
      ANegCurrFormat: Integer = -1): PCell; overload;
    procedure WriteCurrency(ACell: PCell; AValue: Double;
      AFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = -1;
      ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
      ANegCurrFormat: Integer = -1); overload;
    function WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
      AFormat: TsNumberFormat; AFormatString: String): PCell; overload;
    procedure WriteCurrency(ACell: PCell; AValue: Double;
      AFormat: TsNumberFormat; AFormatString: String); overload;

    function WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
      AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = ''): PCell; overload;
    procedure WriteDateTime(ACell: PCell; AValue: TDateTime;
      AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = ''); overload;
    function WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
      AFormatStr: String): PCell; overload;
    procedure WriteDateTime(ACell: PCell; AValue: TDateTime;
      AFormatStr: String); overload;

    function WriteErrorValue(ARow, ACol: Cardinal; AValue: TsErrorValue): PCell; overload;
    procedure WriteErrorValue(ACell: PCell; AValue: TsErrorValue); overload;

    function WriteFormula(ARow, ACol: Cardinal; AFormula: String): PCell; overload;
    procedure WriteFormula(ACell: PCell; AFormula: String); overload;

    function WriteNumber(ARow, ACol: Cardinal; ANumber: double): PCell; overload;
    procedure WriteNumber(ACell: PCell; ANumber: Double); overload;
    function WriteNumber(ARow, ACol: Cardinal; ANumber: double;
      AFormat: TsNumberFormat; ADecimals: Byte = 2): PCell; overload;
    procedure WriteNumber(ACell: PCell; ANumber: Double;
      AFormat: TsNumberFormat; ADecimals: Byte = 2); overload;
    function WriteNumber(ARow, ACol: Cardinal; ANumber: double;
      AFormat: TsNumberFormat; AFormatString: String): PCell; overload;
    procedure WriteNumber(ACell: PCell; ANumber: Double;
      AFormat: TsNumberFormat; AFormatString: String); overload;

    function WriteRPNFormula(ARow, ACol: Cardinal; AFormula: TsRPNFormula): PCell; overload;
    procedure WriteRPNFormula(ACell: PCell; AFormula: TsRPNFormula); overload;

    function WriteUTF8Text(ARow, ACol: Cardinal; AText: ansistring): PCell; overload;
    procedure WriteUTF8Text(ACell: PCell; AText: ansistring); overload;

    { Writing of cell attributes }
    function WriteBackgroundColor(ARow, ACol: Cardinal; AColor: TsColor): PCell; overload;
    procedure WriteBackgroundColor(ACell: PCell; AColor: TsColor); overload;

    function WriteBorderColor(ARow, ACol: Cardinal; ABorder: TsCellBorder; AColor: TsColor): PCell; overload;
    procedure WriteBorderColor(ACell: PCell; ABorder: TsCellBorder; AColor: TsColor); overload;
    function WriteBorderLineStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle): PCell; overload;
    procedure WriteBorderLineStyle(ACell: PCell; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle); overload;
    function WriteBorders(ARow, ACol: Cardinal; ABorders: TsCellBorders): PCell; overload;
    procedure WriteBorders(ACell: PCell; ABorders: TsCellBorders); overload;
    function WriteBorderStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      AStyle: TsCellBorderStyle): PCell; overload;
    procedure WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
      AStyle: TsCellBorderStyle); overload;
    function WriteBorderStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle; AColor: TsColor): PCell; overload;
    procedure WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle; AColor: TsColor); overload;
    function WriteBorderStyles(ARow, ACol: Cardinal; const AStyles: TsCellBorderStyles): PCell; overload;
    procedure WriteBorderStyles(ACell: PCell; const AStyles: TsCellBorderStyles); overload;

    function WriteDateTimeFormat(ARow, ACol: Cardinal; ANumberFormat: TsNumberFormat;
      const AFormatString: String = ''): PCell; overload;
    procedure WriteDateTimeFormat(ACell: PCell; ANumberFormat: TsNumberFormat;
      const AFormatString: String = ''); overload;

    function WriteDecimals(ARow, ACol: Cardinal; ADecimals: byte): PCell; overload;
    procedure WriteDecimals(ACell: PCell; ADecimals: Byte); overload;

    function  WriteFont(ARow, ACol: Cardinal; const AFontName: String;
      AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer; overload;
    function  WriteFont(ACell: PCell; const AFontName: String;
      AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer; overload;
    function WriteFont(ARow, ACol: Cardinal; AFontIndex: Integer): PCell; overload;
    procedure WriteFont(ACell: PCell; AFontIndex: Integer); overload;
    function WriteFontColor(ARow, ACol: Cardinal; AFontColor: TsColor): Integer; overload;
    function WriteFontColor(ACell: PCell; AFontColor: TsColor): Integer; overload;
    function WriteFontName(ARow, ACol: Cardinal; AFontName: String): Integer; overload;
    function WriteFontName(ACell: PCell; AFontName: String): Integer; overload;
    function WriteFontSize(ARow, ACol: Cardinal; ASize: Single): Integer; overload;
    function WriteFontSize(ACell: PCell; ASize: Single): Integer; overload;
    function WriteFontStyle(ARow, ACol: Cardinal; AStyle: TsFontStyles): Integer; overload;
    function WriteFontStyle(ACell: PCell; AStyle: TsFontStyles): Integer; overload;

    function WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment): PCell; overload;
    procedure WriteHorAlignment(ACell: PCell; AValue: TsHorAlignment); overload;

    function WriteNumberFormat(ARow, ACol: Cardinal; ANumberFormat: TsNumberFormat;
      const AFormatString: String = ''): PCell; overload;
    procedure WriteNumberFormat(ACell: PCell; ANumberFormat: TsNumberFormat;
      const AFormatString: String = ''); overload;
    function WriteNumberFormat(ARow, ACol: Cardinal; ANumberFormat: TsNumberFormat;
      ADecimals: Integer; ACurrencySymbol: String = ''; APosCurrFormat: Integer = -1;
      ANegCurrFormat: Integer = -1): PCell; overload;
    procedure WriteNumberFormat(ACell: PCell; ANumberFormat: TsNumberFormat;
      ADecimals: Integer; ACurrencySymbol: String = '';
      APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1); overload;

    function WriteTextRotation(ARow, ACol: Cardinal; ARotation: TsTextRotation): PCell; overload;
    procedure WriteTextRotation(ACell: PCell; ARotation: TsTextRotation); overload;

    procedure WriteUsedFormatting(ARow, ACol: Cardinal; AUsedFormatting: TsUsedFormattingFields);

    function WriteVertAlignment(ARow, ACol: Cardinal; AValue: TsVertAlignment): PCell; overload;
    procedure WriteVertAlignment(ACell: PCell; AValue: TsVertAlignment); overload;

    function WriteWordwrap(ARow, ACol: Cardinal; AValue: boolean): PCell; overload;
    procedure WriteWordwrap(ACell: PCell; AValue: boolean); overload;

    { Formulas }
    function BuildRPNFormula(ACell: PCell): TsRPNFormula;
    procedure CalcFormula(ACell: PCell);
    procedure CalcFormulas;
    function ConvertRPNFormulaToStringFormula(const AFormula: TsRPNFormula): String;
    function UseSharedFormula(ARow, ACol: Cardinal; ASharedFormulaBase: PCell): PCell; overload;
    procedure UseSharedFormula(ACellRangeStr: String; ASharedFormulaBase: PCell); overload;

    { Data manipulation methods - For Cells }
    procedure CopyCell(AFromRow, AFromCol, AToRow, AToCol: Cardinal; AFromWorksheet: TsWorksheet);
    procedure CopyFormat(AFormat: PCell; AToRow, AToCol: Cardinal); overload;
    procedure CopyFormat(AFromCell, AToCell: PCell); overload;
    function  FindCell(ARow, ACol: Cardinal): PCell; overload;
    function  FindCell(AddressStr: String): PCell; overload;
    function  GetCell(ARow, ACol: Cardinal): PCell; overload;
    function  GetCell(AddressStr: String): PCell; overload;
    function  GetCellCount: Cardinal;
    function  GetFirstCell(): PCell;
    function  GetNextCell(): PCell;
    function  GetFirstCellOfRow(ARow: Cardinal): PCell;
    function  GetLastCellOfRow(ARow: Cardinal): PCell;
    function  GetFirstColIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastColIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastColNumber: Cardinal; deprecated 'Use GetLastColIndex';
    function  GetFirstRowIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastRowIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastRowNumber: Cardinal; deprecated 'Use GetLastRowIndex';

    { Data manipulation methods - For Rows and Cols }
    function  CalcAutoRowHeight(ARow: Cardinal): Single;
    function  FindRow(ARow: Cardinal): PRow;
    function  FindCol(ACol: Cardinal): PCol;
    function  GetCellCountInRow(ARow: Cardinal): Cardinal;
    function  GetCellCountInCol(ACol: Cardinal): Cardinal;
    function  GetRow(ARow: Cardinal): PRow;
    function  GetRowHeight(ARow: Cardinal): Single;
    function  GetCol(ACol: Cardinal): PCol;
    function  GetColWidth(ACol: Cardinal): Single;
    procedure InsertCol(ACol: Cardinal);
    procedure InsertRow(ARow: Cardinal);
    procedure RemoveAllRows;
    procedure RemoveAllCols;
    procedure WriteRowInfo(ARow: Cardinal; AData: TRow);
    procedure WriteRowHeight(ARow: Cardinal; AHeight: Single);
    procedure WriteColInfo(ACol: Cardinal; AData: TCol);
    procedure WriteColWidth(ACol: Cardinal; AWidth: Single);

    { Properties }

    {@@ List of cells of the worksheet. Only cells with contents or with formatting
        are listed }
    property  Cells: TAVLTree read FCells;
    {@@ List of all column records of the worksheet having a non-standard column width }
    property  Cols: TIndexedAVLTree read FCols;
    {@@ FormatSettings for localization of some formatting strings }
    property  FormatSettings: TFormatSettings read GetFormatSettings;
    {@@ List of all row records of the worksheet having a non-standard row height }
    property  Rows: TIndexedAVLTree read FRows;
    {@@ Workbook to which the worksheet belongs }
    property  Workbook: TsWorkbook read FWorkbook;

    // These are properties to interface to TsWorksheetGrid
    {@@ Parameters controlling visibility of grid lines and row/column headers,
        usage of frozen panes etc. }
    property  Options: TsSheetOptions read FOptions write FOptions;
    {@@ Number of frozen columns which do not scroll }
    property  LeftPaneWidth: Integer read FLeftPaneWidth write FLeftPaneWidth;
    {@@ Number of frozen rows which do not scroll }
    property  TopPaneHeight: Integer read FTopPaneHeight write FTopPaneHeight;
    {@@ Event fired when cell contents or formatting changes }
    property  OnChangeCell: TsCellEvent read FOnChangeCell write FOnChangeCell;
    {@@ Event fired when the font size in a cell changes }
    property  OnChangeFont: TsCellEvent read FOnChangeFont write FOnChangeFont;
  end;

  {@@
    Option flags for the workbook

    @param  boVirtualMode      If in virtual mode date are not taken from cells
                               when a spreadsheet is written to file, but are
                               provided by means of the event OnWriteCellData.
                               Similarly, when data are read they are not added as
                               cells but passed the the event OnReadCellData;
    @param  boBufStream        When this option is set a buffered stream is used
                               for writing (a memory stream swapping to disk) or
                               reading (a file stream pre-reading chunks of data
                               to memory)
    @param  boAutoCalc         Automatically recalculate rpn formulas whenever
                               a cell value changes.
    @param  boCalcBeforeSaving Calculates formulas before saving the file.
                               Otherwise there are no results when the file is
                               loaded back by fpspreadsheet.
    @param  boReadFormulas     Allows to turn off reading of rpn formulas; this is
                               a precaution since formulas not correctly
                               implemented by fpspreadsheet could crash the
                               reading operation. }
  TsWorkbookOption = (boVirtualMode, boBufStream, boAutoCalc, boCalcBeforeSaving,
    boReadFormulas);

  {@@
    Set of options flags for the workbook }
  TsWorkbookOptions = set of TsWorkbookOption;

  {@@
    Event fired when writing a file in virtual mode. The event handler has to
    pass data ("AValue") and formatting style to be copied from a template
    cell ("AStyleCell") to the writer }
  TsWorkbookWriteCellDataEvent = procedure(Sender: TObject; ARow, ACol: Cardinal;
    var AValue: variant; var AStyleCell: PCell) of object;

  {@@
    Event fired when reading a file in virtual mode. Read data are provided in
    the "ADataCell" (which is not added to the worksheet in virtual mode). }
  TsWorkbookReadCellDataEvent = procedure(Sender: TObject; ARow, ACol: Cardinal;
    const ADataCell: PCell) of object;

  {@@
    The workbook contains the worksheets and provides methods for reading from
    and writing to file.
  }
  TsWorkbook = class
  private
    { Internal data }
    FWorksheets: TFPList;
    FEncoding: TsEncoding;
    FFormat: TsSpreadsheetFormat;
    FFontList: TFPList;
    FBuiltinFontCount: Integer;
    FPalette: array of TsColorValue;
    FDefaultColWidth: Single;   // in "characters". Excel uses the width of char "0" in 1st font
    FDefaultRowHeight: Single;  // in "character heights", i.e. line count
    FVirtualColCount: Cardinal;
    FVirtualRowCount: Cardinal;
    FWriting: Boolean;
    FCalculationLock: Integer;
    FOptions: TsWorkbookOptions;
    FOnWriteCellData: TsWorkbookWriteCellDataEvent;
    FOnReadCellData: TsWorkbookReadCellDataEvent;
    FFileName: String;
    FLog: TStringList;

    { Setter/Getter }
    function GetErrorMsg: String;
    procedure SetVirtualColCount(AValue: Cardinal);
    procedure SetVirtualRowCount(AValue: Cardinal);

    { Internal methods }
    procedure GetLastRowColIndex(out ALastRow, ALastCol: Cardinal);
    procedure PrepareBeforeReading;
    procedure PrepareBeforeSaving;
    procedure ReCalc;
    procedure RemoveWorksheetsCallback(data, arg: pointer);
    procedure UpdateCaches;

  public
    {@@ A copy of SysUtil's DefaultFormatSettings to provide some kind of
      localization to some formatting strings. Can be modified before
      loading/writing files }
    FormatSettings: TFormatSettings;

    { Base methods }
    constructor Create;
    destructor Destroy; override;
    class function GetFormatFromFileName(const AFileName: TFileName; out SheetType: TsSpreadsheetFormat): Boolean;
    function  CreateSpreadReader(AFormat: TsSpreadsheetFormat): TsCustomSpreadReader;
    function  CreateSpreadWriter(AFormat: TsSpreadsheetFormat): TsCustomSpreadWriter;
    procedure ReadFromFile(AFileName: string; AFormat: TsSpreadsheetFormat); overload;
    procedure ReadFromFile(AFileName: string); overload;
    procedure ReadFromFileIgnoringExtension(AFileName: string);
    procedure ReadFromStream(AStream: TStream; AFormat: TsSpreadsheetFormat);
    procedure WriteToFile(const AFileName: string;
      const AFormat: TsSpreadsheetFormat;
      const AOverwriteExisting: Boolean = False); overload;
    procedure WriteToFile(const AFileName: String; const AOverwriteExisting: Boolean = False); overload;
    procedure WriteToStream(AStream: TStream; AFormat: TsSpreadsheetFormat);

    { Worksheet list handling methods }
    function  AddWorksheet(AName: string): TsWorksheet;
    function  GetFirstWorksheet: TsWorksheet;
    function  GetWorksheetByIndex(AIndex: Cardinal): TsWorksheet;
    function  GetWorksheetByName(AName: String): TsWorksheet;
    function  GetWorksheetCount: Cardinal;
    procedure RemoveAllWorksheets;

    { Font handling }
    function AddFont(const AFontName: String; ASize: Single;
      AStyle: TsFontStyles; AColor: TsColor): Integer; overload;
    function AddFont(const AFont: TsFont): Integer; overload;
    procedure CopyFontList(ASource: TFPList);
    procedure DeleteFont(AFontIndex: Integer);
    function FindFont(const AFontName: String; ASize: Single;
      AStyle: TsFontStyles; AColor: TsColor): Integer;
    function GetDefaultFont: TsFont;
    function GetDefaultFontSize: Single;
    function GetFont(AIndex: Integer): TsFont;
    function GetFontAsString(AIndex: Integer): String;
    function GetFontCount: Integer;
    procedure InitFonts;
    procedure RemoveAllFonts;
    procedure SetDefaultFont(const AFontName: String; ASize: Single);

    { Color handling }
    function AddColorToPalette(AColorValue: TsColorValue): TsColor;
    function FindClosestColor(AColorValue: TsColorValue;
      AMaxPaletteCount: Integer): TsColor;
    function FPSColorToHexString(AColor: TsColor; ARGBColor: TFPColor): String;
    function GetColorName(AColorIndex: TsColor): string;
    function GetPaletteColor(AColorIndex: TsColor): TsColorValue;
    function GetPaletteColorAsHTMLStr(AColorIndex: TsColor): String;
    procedure SetPaletteColor(AColorIndex: TsColor; AColorValue: TsColorValue);
    function GetPaletteSize: Integer;
    procedure UseDefaultPalette;
    procedure UsePalette(APalette: PsPalette; APaletteCount: Word;
      ABigEndian: Boolean = false);
    function UsesColor(AColorIndex: TsColor): Boolean;

    { Error messages }
    procedure AddErrorMsg(const AMsg: String); overload;
    procedure AddErrorMsg(const AMsg: String; const Args: array of const); overload;
    procedure ClearErrorList;

    {@@ The default column width given in "character units" (width of the
      character "0" in the default font) }
    property DefaultColWidth: Single read FDefaultColWidth;
    {@@ The default row height is given in "line count" (height of the
      default font }
    property DefaultRowHeight: Single read FDefaultRowHeight;
    {@@ This property is only used for formats which don't support unicode
      and support a single encoding for the whole document, like Excel 2 to 5 }
    property Encoding: TsEncoding read FEncoding write FEncoding;
    {@@ Retrieves error messages collected during reading/writing }
    property ErrorMsg: String read GetErrorMsg;
    {@@ Filename of the saved workbook }
    property FileName: String read FFileName;
    {@@ Identifies the file format which was detected when reading the file }
    property FileFormat: TsSpreadsheetFormat read FFormat;
    property VirtualColCount: cardinal read FVirtualColCount write SetVirtualColCount;
    property VirtualRowCount: cardinal read FVirtualRowCount write SetVirtualRowCount;
    property Options: TsWorkbookOptions read FOptions write FOptions;
    {@@ This event allows to provide external cell data for writing to file,
      standard cells are ignored. Intended for converting large database files
      to a spreadsheet format. Requires Option boVirtualMode to be set. }
    property OnWriteCellData: TsWorkbookWriteCellDataEvent read FOnWriteCellData write FOnWriteCellData;
    {@@ This event accepts cell data while reading a spreadsheet file. Data are
      not encorporated in a spreadsheet, they are just passed through to the
      event handler for processing. Requires option boVirtualMode to be set. }
    property OnReadCellData: TsWorkbookReadCellDataEvent read FOnReadCellData write FOnReadCellData;
  end;

  {@@ Contents of a number format record }
  TsNumFormatData = class
  public
    {@@ Excel refers to a number format by means of the format "index". }
    Index: Integer;
    {@@ OpenDocument refers to a number format by means of the format "name". }
    Name: String;
    {@@ Identifier of a built-in number format, see TsNumberFormat }
    NumFormat: TsNumberFormat;
    {@@ String of format codes, such as '#,##0.00', or 'hh:nn'. }
    FormatString: string;
  end;

  {@@ Specialized list for number format items }
  TsCustomNumFormatList = class(TFPList)
  private
    function GetItem(AIndex: Integer): TsNumFormatData;
    procedure SetItem(AIndex: Integer; AValue: TsNumFormatData);
  protected
    {@@ Workbook from which the number formats are collected in the list. It is
     mainly needed to get access to the FormatSettings for easy localization of some
     formatting strings. }
    FWorkbook: TsWorkbook;
    {@@ Identifies the first number format item that is written to the file. Items
     having a smaller index are not written. }
    FFirstFormatIndexInFile: Integer;
    {@@ Identifies the index of the next Excel number format item to be written.
     Needed for auto-creating of the user-defined Excel number format indexes }
    FNextFormatIndex: Integer;
    procedure AddBuiltinFormats; virtual;
    procedure RemoveFormat(AIndex: Integer);

  public
    constructor Create(AWorkbook: TsWorkbook);
    destructor Destroy; override;
    function AddFormat(AFormatCell: PCell): Integer; overload;
    function AddFormat(AFormatIndex: Integer; AFormatName, AFormatString: String;
      ANumFormat: TsNumberFormat): Integer; overload;
    function AddFormat(AFormatIndex: Integer; AFormatString: String;
      ANumFormat: TsNumberFormat): Integer; overload;
    function AddFormat(AFormatName, AFormatString: String;
      ANumFormat: TsNumberFormat): Integer; overload;
    function AddFormat(AFormatString: String; ANumFormat: TsNumberFormat): Integer; overload;
    procedure AnalyzeAndAdd(AFormatIndex: Integer; AFormatString: String);
    procedure Clear;
    procedure ConvertAfterReading(AFormatIndex: Integer; var AFormatString: String;
      var ANumFormat: TsNumberFormat); virtual;
    procedure ConvertBeforeWriting(var AFormatString: String;
      var ANumFormat: TsNumberFormat); virtual;
    procedure Delete(AIndex: Integer);
    function Find(ANumFormat: TsNumberFormat; AFormatString: String): Integer; overload;
    function Find(AFormatString: String): Integer; overload;
    function FindByIndex(AFormatIndex: Integer): Integer;
    function FindByName(AFormatName: String): Integer;
    function FindFormatOf(AFormatCell: PCell): integer; virtual;
    function FormatStringForWriting(AIndex: Integer): String; virtual;
    procedure Sort;

    {@@ Workbook from which the number formats are collected in the list. It is
     mainly needed to get access to the FormatSettings for easy localization of some
     formatting strings. }
    property Workbook: TsWorkbook read FWorkbook;
    {@@ Identifies the first number format item that is written to the file. Items
     having a smaller index are not written. }
    property FirstFormatIndexInFile: Integer read FFirstFormatIndexInFile;
    {@@ Number format items contained in the list }
    property Items[AIndex: Integer]: TsNumFormatData read GetItem write SetItem; default;
  end;


  { TsCustomSpreadReaderWriter }

  {@@ Common ancestor of the spreadsheet reader and writer classes providing
      shared data and methods. }
  TsCustomSpreadReaderWriter = class
  protected
    {@@ Instance of the workbook which is currently being read. }
    FWorkbook: TsWorkbook;
    {@@ Instance of the worksheet which is currently being read. }
    FWorksheet: TsWorksheet;
    {@@ Limitations for the specific data file format }
    FLimitations: TsSpreadsheetFormatLimitations;
  protected
    {@@ List of number formats found in the file }
    FNumFormatList: TsCustomNumFormatList;
    procedure CreateNumFormatList; virtual;
  public
    constructor Create(AWorkbook: TsWorkbook); virtual; // to allow descendents to override it
    destructor Destroy; override;
    function Limitations: TsSpreadsheetFormatLimitations;
    {@@ Instance of the workbook which is currently being read/written. }
    property Workbook: TsWorkbook read FWorkbook;
    {@@ List of number formats found in the workbook. }
    property NumFormatList: TsCustomNumFormatList read FNumFormatList;
  end;

  { TsCustomSpreadReader }

  {@@ TsSpreadReader class reference type }
  TsSpreadReaderClass = class of TsCustomSpreadReader;

  {@@
    Custom reader of spreadsheet files. "Custom" means that it provides only
    the basic functionality. The main implementation is done in derived classes
    for each individual file format.
  }
  TsCustomSpreadReader = class(TsCustomSpreadReaderWriter)
  protected
    {@@ Temporary cell for virtual mode}
    FVirtualCell: TCell;
    {@@ Stores if the reader is in virtual mode }
    FIsVirtualMode: Boolean;
    { Record reading methods }
    {@@ Abstract method for reading a blank cell. Must be overridden by descendent classes. }
    procedure ReadBlank(AStream: TStream); virtual; abstract;
    {@@ Abstract method for reading a formula cell. Must be overridden by descendent classes. }
    procedure ReadFormula(AStream: TStream); virtual; abstract;
    {@@ Abstract method for reading a text cell. Must be overridden by descendent classes. }
    procedure ReadLabel(AStream: TStream); virtual; abstract;
    {@@ Abstract method for reading a number cell. Must be overridden by descendent classes. }
    procedure ReadNumber(AStream: TStream); virtual; abstract;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    { General writing methods }
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); virtual;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); virtual;
    procedure ReadFromStrings(AStrings: TStrings; AData: TsWorkbook); virtual;
  end;


  { TsCustomSpreadWriter }

  {@@ TsSpreadWriter class reference type }
  TsSpreadWriterClass = class of TsCustomSpreadWriter;

  {@@ Callback function when iterating cells while accessing a stream }
  TCellsCallback = procedure (ACell: PCell; AStream: TStream) of object;

  {@@
    Custom writer of spreadsheet files. "Custom" means that it provides only
    the basic functionality. The main implementation is done in derived classes
    for each individual file format. }
  TsCustomSpreadWriter = class(TsCustomSpreadReaderWriter)
  protected
    { Helper routines }
    procedure AddDefaultFormats(); virtual;
    procedure CheckLimitations;
    function  FindFormattingInList(AFormat: PCell): Integer;
    procedure FixCellColors(ACell: PCell);
    function  FixColor(AColor: TsColor): TsColor; virtual;
    procedure FixFormat(ACell: PCell); virtual;
    procedure FixRelativeReferences(ACell: PCell; var AElement: TsFormulaElement); virtual;
    procedure GetSheetDimensions(AWorksheet: TsWorksheet;
      out AFirstRow, ALastRow, AFirstCol, ALastCol: Cardinal); virtual;
    procedure ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
    procedure ListAllFormattingStyles; virtual;
    procedure ListAllNumFormatsCallback(ACell: PCell; AStream: TStream);
    procedure ListAllNumFormats; virtual;
    { Helpers for writing }
    procedure WriteCellCallback(ACell: PCell; AStream: TStream);
    procedure WriteCellsToStream(AStream: TStream; ACells: TAVLTree);
    { Record writing methods }
    {@@ Abstract method for writing a blank cell. Must be overridden by descendent classes. }
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal; ACell: PCell); virtual; abstract;
    {@@ Abstract method for writing a date/time value to a cell. Must be overridden by descendent classes. }
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell); virtual; abstract;
    {@@ Abstract method for writing a formula to a cell. Must be overridden by descendent classes. }
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal; ACell: PCell); virtual;
    (*
    {@@ Abstract method for writing an RPN formula to a cell. Must be overridden by descendent classes. }
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal;
      const AFormula: TsRPNFormula; ACell: PCell); virtual;
      *)
    {@@ Abstract method for writing a string to a cell. Must be overridden by descendent classes. }
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal; const AValue: string; ACell: PCell); virtual; abstract;
    {@@ Abstract method for writing a number value to a cell. Must be overridden by descendent classes. }
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double; ACell: PCell); virtual; abstract;

  public
    {@@ An array with cells which are models for the used styles
    In this array the Row property holds the index to the corresponding XF field  }
    FFormattingStyles: array of TCell;
    {@@ Indicates which should be the next XF (style) index when filling the FFormattingStyles array }
    NextXFIndex: Integer;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    { General writing methods }
    procedure IterateThroughCells(AStream: TStream; ACells: TAVLTree; ACallback: TCellsCallback);
    procedure WriteToFile(const AFileName: string; const AOverwriteExisting: Boolean = False); virtual;
    procedure WriteToStream(AStream: TStream); virtual;
    procedure WriteToStrings(AStrings: TStrings); virtual;
  end;

  {@@ List of registered formats }
  TsSpreadFormatData = record
    ReaderClass: TsSpreadReaderClass;
    WriterClass: TsSpreadWriterClass;
    Format: TsSpreadsheetFormat;
  end;

  { Simple creation an RPNFormula array to be used in fpspreadsheet. }

  {@@ Helper record for simplification of RPN formula creation }
  PRPNItem = ^TRPNItem;
  TRPNItem = record
    FE: TsFormulaElement;
    Next: PRPNItem;
  end;

  function CreateRPNFormula(AItem: PRPNItem; AReverse: Boolean = false): TsRPNFormula;
  procedure DestroyRPNFormula(AItem: PRPNItem);

  function RPNBool(AValue: Boolean;
    ANext: PRPNItem): PRPNItem;
  function RPNCellValue(ACellAddress: String;
    ANext: PRPNItem): PRPNItem; overload;
  function RPNCellValue(ARow, ACol: Integer; AFlags: TsRelFlags;
    ANext: PRPNItem): PRPNItem; overload;
  function RPNCellRef(ACellAddress: String;
    ANext: PRPNItem): PRPNItem; overload;
  function RPNCellRef(ARow, ACol: Integer; AFlags: TsRelFlags;
    ANext: PRPNItem): PRPNItem; overload;
  function RPNCellRange(ACellRangeAddress: String;
    ANext: PRPNItem): PRPNItem; overload;
  function RPNCellRange(ARow, ACol, ARow2, ACol2: Integer; AFlags: TsRelFlags;
    ANext: PRPNItem): PRPNItem; overload;
  function RPNCellOffset(ARowOffset, AColOffset: Integer; AFlags: TsRelFlags;
    ANext: PRPNItem): PRPNItem;
  function RPNErr(AErrCode: Byte; ANext: PRPNItem): PRPNItem;
  function RPNInteger(AValue: Word; ANext: PRPNItem): PRPNItem;
  function RPNMissingArg(ANext: PRPNItem): PRPNItem;
  function RPNNumber(AValue: Double; ANext: PRPNItem): PRPNItem;
  function RPNParenthesis(ANext: PRPNItem): PRPNItem;
  function RPNString(AValue: String; ANext: PRPNItem): PRPNItem;
  function RPNFunc(AToken: TFEKind; ANext: PRPNItem): PRPNItem; overload;
  function RPNFunc(AFuncName: String; ANext: PRPNItem): PRPNItem; overload;
  function RPNFunc(AFuncName: String; ANumParams: Byte; ANext: PRPNItem): PRPNItem; overload;

//  function FixedParamCount(AElementKind: TFEKind): Boolean;

var
  GsSpreadFormats: array of TsSpreadFormatData;

//procedure RegisterFormulaFunc(AFormulaKind: TFEKind; AFunc: pointer);

procedure RegisterSpreadFormat( AReaderClass: TsSpreadReaderClass;
  AWriterClass: TsSpreadWriterClass; AFormat: TsSpreadsheetFormat);

procedure CopyCellFormat(AFromCell, AToCell: PCell);
function GetFileFormatName(AFormat: TsSpreadsheetFormat): String;
procedure MakeLEPalette(APalette: PsPalette; APaletteSize: Integer);
function SameCellBorders(ACell1, ACell2: PCell): Boolean;

procedure InitCell(out ACell: TCell); overload;
procedure InitCell(ARow, ACol: Cardinal; out ACell: TCell); overload;

function HasFormula(ACell: PCell): Boolean;


implementation

uses
  Math, StrUtils, TypInfo,
  fpsStreams, fpsUtils, fpsNumFormatParser, fpsExprParser, fpsFunc;

{ Translatable strings }
resourcestring
  lpUnsupportedReadFormat = 'Tried to read a spreadsheet using an unsupported format';
  lpUnsupportedWriteFormat = 'Tried to write a spreadsheet using an unsupported format';
  lpNoValidSpreadsheetFile = '"%s" is not a valid spreadsheet file';
  lpUnknownSpreadsheetFormat = 'unknown format';
  lpMaxRowsExceeded = 'This workbook contains %d rows, but the selected ' +
    'file format does not support more than %d rows.';
  lpMaxColsExceeded = 'This workbook contains %d columns, but the selected ' +
    'file format does not support more than %d columns.';
  lpTooManyPaletteColors = 'This workbook contains more colors (%d) than ' +
    'supported by the file format (%d). The additional colors are replaced by '+
    'the best-matching palette colors.';
  lpInvalidFontIndex = 'Invalid font index';
  lpInvalidNumberFormat = 'Trying to use an incompatible number format.';
  lpInvalidDateTimeFormat = 'Trying to use an incompatible date/time format.';
  lpNoValidNumberFormatString = 'No valid number format string.';
  lpNoValidCellAddress = '"%s" is not a valid cell address.';
  lpNoValidCellRangeAddress = '"%s" is not a valid cell range address.';
  lpNoValidCellRangeOrCellAddress = '"%s" is not a valid cell or cell range address.';
  lpSpecifyNumberOfParams = 'Specify number of parameters for function %s';
  lpIncorrectParamCount = 'Funtion %s requires at least %d and at most %d parameters.';
  lpCircularReference = 'Circular reference found when calculating worksheet formulas';
  lpTRUE = 'TRUE';
  lpFALSE = 'FALSE';
  lpErrEmptyIntersection = '#NULL!';
  lpErrDivideByZero = '#DIV/0!';
  lpErrWrongType = '#VALUE!';
  lpErrIllegalRef = '#REF!';
  lpErrWrongName = '#NAME?';
  lpErrOverflow = '#NUM!';
  lpErrArgError = '#N/A';
  lpErrFormulaNotSupported = '<FORMULA?>';

{%H-}lpNoValidDateTimeFormatString = 'No valid date/time format string.';
{%H-}lpIllegalNumberFormat = 'Illegal number format.';


const
  { These are reserved system colors by Microsoft
    0x0040 - Default foreground color - window text color in the sheet display.
    0x0041 - Default background color - window background color in the sheet
             display and is the default background color for a cell.
    0x004D - Default chart foreground color - window text color in the
             chart display.
    0x004E - Default chart background color - window background color in the
             chart display.
    0x004F - Chart neutral color which is black, an RGB value of (0,0,0).
    0x0051 - ToolTip text color - automatic font color for comments.
    0x7FFF - Font automatic color - window text color. }

  // Color indexes of reserved system colors
  DEF_FOREGROUND_COLOR = $0040;
  DEF_BACKGROUND_COLOR = $0041;
  DEF_CHART_FOREGROUND_COLOR = $004D;
  DEF_CHART_BACKGROUND_COLOR = $004E;
  DEF_CHART_NEUTRAL_COLOR = $004F;
  DEF_TOOLTIP_TEXT_COLOR = $0051;
  DEF_FONT_AUTOMATIC_COLOR = $7FFF;

  // Color rgb values of reserved system colors
  DEF_FOREGROUND_COLORVALUE = $000000;
  DEF_BACKGROUND_COLORVALUE = $FFFFFF;
  DEF_CHART_FOREGROUND_COLORVALUE = $000000;
  DEF_CHART_BACKGROUND_COLORVALUE = $FFFFFF;
  DEF_CHART_NEUTRAL_COLORVALUE = $FFFFFF;
  DEF_TOOLTIP_TEXT_COLORVALUE = $000000;
  DEF_FONT_AUTOMATIC_COLORVALUE = $000000;

var
  {@@ RGB colors RGB in "big-endian" notation (red at left). The values are inverted
    at initialization to be little-endian at run-time!
    The indices into this palette are named as scXXXX color constants. }
  DEFAULT_PALETTE: array[$00..$16] of TsColorValue = (
    $000000,  // $00: black
    $FFFFFF,  // $01: white
    $FF0000,  // $02: red
    $00FF00,  // $03: green
    $0000FF,  // $04: blue
    $FFFF00,  // $05: yellow
    $FF00FF,  // $06: magenta
    $00FFFF,  // $07: cyan
    $800000,  // $08: dark red
    $008000,  // $09: dark green
    $000080,  // $0A: dark blue
    $808000,  // $0B: olive
    $800080,  // $0C: purple
    $008080,  // $0D: teal
    $C0C0C0,  // $0E: silver
    $808080,  // $0F: gray
    $E6E6E6,  // $10: gray 10%
    $CCCCCC,  // $11: gray 20%
    $FFA500,  // $12: orange
    $A0522D,  // $13: dark brown
    $CD853F,  // $14: brown
    $F5F5DC,  // $15: beige
    $F5DEB3   // $16: wheat
  );

  {@@ Names of the colors of the DEFAULT_PALETTE }
  DEFAULT_COLORNAMES: array[$00..$16] of string = (
    'black',      // 0
    'white',      // 1
    'red',        // 2
    'green',      // 3
    'blue',       // 4
    'yellow',     // 5
    'magenta',    // 6
    'cyan',       // 7
    'dark red',   // 8
    'dark green', // 9
    'dark blue',  // $0A
    'olive',      // $0B
    'purple',     // $0C
    'teal',       // $0D
    'silver',     // $0E
    'gray',       // $0F
    'gray 10%',   // $10
    'gray 20%',   // $11
    'orange',     // $12
    'dark brown', // $13
    'brown',      // $14
    'beige',      // $15
    'wheat'       // $16
  );

(*
{ Properties of formula elements }

type
  {@@ Properties of formula elements:
    @param  Symbol     Symbol used in the formula
    @param  MinParams  Minimum count of parameters used in this function
    @param  MaxParams  Maximum count of parameters used in this function
    @param  Func      Function to be calculated }
  TFEProp = record
    Symbol: String;
    MinParams, MaxParams: Byte;
    Func: TsFormulaFunc;
  end;

var
  FEProps: array[TFEKind] of TFEProp = (                                        // functions marked by ( * )
  { Operands }                                                                  // are only partially supported
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCell
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellRef
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellRange
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellOffset
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellNum
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellInteger
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellString
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellBool
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellErr
    (Symbol:'';          MinParams:Byte(-1); MaxParams:Byte(-1); Func:nil),     // fekCellMissingArg
  { Basic operations }
    (Symbol:'+';         MinParams:2; MaxParams:2;  Func:fpsAdd),               // fekAdd
    (Symbol:'-';         MinParams:2; MaxParams:2;  Func:fpsSub),               // fekSub
    (Symbol:'*';         MinParams:2; MaxParams:2;  Func:fpsMul),               // fekMul
    (Symbol:'/';         MinParams:2; MaxParams:2;  Func:fpsDiv),               // fekDiv
    (Symbol:'%';         MinParams:1; MaxParams:1;  Func:fpsPercent),           // fekPercent
    (Symbol:'^';         MinParams:2; MaxParams:2;  Func:fpsPower),             // fekPower
    (Symbol:'-';         MinParams:1; MaxParams:1;  Func:fpsUMinus),            // fekUMinus
    (Symbol:'+';         MinParams:1; MaxParams:1;  Func:fpsUPlus),             // fekUPlus
    (Symbol:'&';         MinParams:2; MaxParams:2;  Func:fpsConcat),            // fekConcat (string concatenation)
    (Symbol:'=';         MinParams:2; MaxParams:2;  Func:fpsEqual),             // fekEqual
    (Symbol:'>';         MinParams:2; MaxParams:2;  Func:fpsGreater),           // fekGreater
    (Symbol:'>=';        MinParams:2; MaxParams:2;  Func:fpsGreaterEqual),      // fekGreaterEqual
    (Symbol:'<';         MinParams:2; MaxParams:2;  Func:fpsLess),              // fekLess
    (Symbol:'<=';        MinParams:2; MaxParams:2;  Func:fpsLessEqual),         // fekLessEqual
    (Symbol:'<>';        MinParams:2; MaxParams:2;  Func:fpsNotEqual),          // fekNotEqual
    (Symbol:'';          MinParams:1; MaxParams:1;  Func:nil),                  // fekParen   -- no need to calculate!
  { math }
    (Symbol:'ABS';       MinParams:1; MaxParams:1;  Func:fpsABS),               // fekABS
    (Symbol:'ACOS';      MinParams:1; MaxParams:1;  Func:fpsACOS),              // fekACOS
    (Symbol:'ACOSH';     MinParams:1; MaxParams:1;  Func:fpsACOSH),             // fekACOSH
    (Symbol:'ASIN';      MinParams:1; MaxParams:1;  Func:fpsASIN),              // fekASIN
    (Symbol:'ASINH';     MinParams:1; MaxParams:1;  Func:fpsASINH),             // fekASINH
    (Symbol:'ATAN';      MinParams:1; MaxParams:1;  Func:fpsATAN),              // fekATAN
    (Symbol:'ATANH';     MinParams:1; MaxParams:1;  Func:fpsATANH),             // fekATANH,
    (Symbol:'COS';       MinParams:1; MaxParams:1;  Func:fpsCOS),               // fekCOS
    (Symbol:'COSH';      MinParams:1; MaxParams:1;  Func:fpsCOSH),              // fekCOSH
    (Symbol:'DEGREES';   MinParams:1; MaxParams:1;  Func:fpsDEGREES),           // fekDEGREES
    (Symbol:'EXP';       MinParams:1; MaxParams:1;  Func:fpsEXP),               // fekEXP
    (Symbol:'INT';       MinParams:1; MaxParams:1;  Func:fpsINT),               // fekINT
    (Symbol:'LN';        MinParams:1; MaxParams:1;  Func:fpsLN),                // fekLN
    (Symbol:'LOG';       MinParams:1; MaxParams:2;  Func:fpsLOG),               // fekLOG,
    (Symbol:'LOG10';     MinParams:1; MaxParams:1;  Func:fpsLOG10),             // fekLOG10
    (Symbol:'PI';        MinParams:0; MaxParams:0;  Func:fpsPI),                // fekPI
    (Symbol:'RADIANS';   MinParams:1; MaxParams:1;  Func:fpsRADIANS),           // fekRADIANS
    (Symbol:'RAND';      MinParams:0; MaxParams:0;  Func:fpsRAND),              // fekRAND
    (Symbol:'ROUND';     MinParams:2; MaxParams:2;  Func:fpsROUND),             // fekROUND,
    (Symbol:'SIGN';      MinParams:1; MaxParams:1;  Func:fpsSIGN),              // fekSIGN
    (Symbol:'SIN';       MinParams:1; MaxParams:1;  Func:fpsSIN),               // fekSIN
    (Symbol:'SINH';      MinParams:1; MaxParams:1;  Func:fpsSINH),              // fekSINH
    (Symbol:'SQRT';      MinParams:1; MaxParams:1;  Func:fpsSQRT),              // fekSQRT,
    (Symbol:'TAN';       MinParams:1; MaxParams:1;  Func:fpsTAN),               // fekTAN
    (Symbol:'TANH';      MinParams:1; MaxParams:1;  Func:fpsTANH),              // fekTANH,
  { date/time }
    (Symbol:'DATE';      MinParams:3; MaxParams:3;  Func:fpsDATE),              // fekDATE
    (Symbol:'DATEDIF';   MinParams:3; MaxParams:3;  Func:fpsDATEDIF),           // fekDATEDIF    ( * )
    (Symbol:'DATEVALUE'; MinParams:1; MaxParams:1;  Func:fpsDATEVALUE),         // fekDATEVALUE
    (Symbol:'DAY';       MinParams:1; MaxParams:1;  Func:fpsDAY),               // fekDAY
    (Symbol:'HOUR';      MinParams:1; MaxParams:1;  Func:fpsHOUR),              // fekHOUR
    (Symbol:'MINUTE';    MinParams:1; MaxParams:1;  Func:fpsMINUTE),            // fekMINUTE
    (Symbol:'MONTH';     MinParams:1; MaxParams:1;  Func:fpsMONTH),             // fekMONTH
    (Symbol:'NOW';       MinParams:0; MaxParams:0;  Func:fpsNOW),               // fekNOW
    (Symbol:'SECOND';    MinParams:1; MaxParams:1;  Func:fpsSECOND),            // fekSECOND
    (Symbol:'TIME';      MinParams:3; MaxParams:3;  Func:fpsTIME),              // fekTIME
    (Symbol:'TIMEVALUE'; MinParams:1; MaxParams:1;  Func:fpsTIMEVALUE),         // fekTIMEVALUE
    (Symbol:'TODAY';     MinParams:0; MaxParams:0;  Func:fpsTODAY),             // fekTODAY
    (Symbol:'WEEKDAY';   MinParams:1; MaxParams:2;  Func:fpsWEEKDAY),           // fekWEEKDAY
    (Symbol:'YEAR';      MinParams:1; MaxParams:1;  Func:fpsYEAR),              // fekYEAR
  { statistical }
    (Symbol:'AVEDEV';    MinParams:1; MaxParams:30; Func:fpsAVEDEV),            // fekAVEDEV
    (Symbol:'AVERAGE';   MinParams:1; MaxParams:30; Func:fpsAVERAGE),           // fekAVERAGE
    (Symbol:'BETADIST';  MinParams:3; MaxParams:5;  Func:nil),   // fekBETADIST
    (Symbol:'BETAINV';   MinParams:3; MaxParams:5;  Func:nil),   // fekBETAINV
    (Symbol:'BINOMDIST'; MinParams:4; MaxParams:4;  Func:nil),   // fekBINOMDIST
    (Symbol:'CHIDIST';   MinParams:2; MaxParams:2;  Func:nil),   // fekCHIDIST
    (Symbol:'CHIINV';    MinParams:2; MaxParams:2;  Func:nil),   // fekCHIINV
    (Symbol:'COUNT';     MinParams:0; MaxParams:30; Func:fpsCOUNT),             // fekCOUNT
    (Symbol:'COUNTA';    MinParams:0; MaxParams:30; Func:fpsCOUNTA),            // fekCOUNTA
    (Symbol:'COUNTBLANK';MinParams:1; MaxParams:1;  Func:fpsCOUNTBLANK),        // fekCOUNTBLANK
    (Symbol:'COUNTIF';   MinParams:2; MaxParams:2;  Func:fpsCOUNTIF),           // fekCOUNTIF
    (Symbol:'MAX';       MinParams:1; MaxParams:30; Func:fpsMAX),               // fekMAX
    (Symbol:'MEDIAN';    MinParams:1; MaxParams:30; Func:nil),  // fekMEDIAN
    (Symbol:'MIN';       MinParams:1; MaxParams:30; Func:fpsMIN),               // fekMIN
    (Symbol:'PERMUT';    MinParams:2; MaxParams:2;  Func:nil),   // fekPERMUT
    (Symbol:'POISSON';   MinParams:3; MaxParams:3;  Func:nil),   // fekPOISSON
    (Symbol:'PRODUCT';   MinParams:0; MaxParams:30; Func:fpsPRODUCT),           // fekPRODUCT
    (Symbol:'STDEV';     MinParams:1; MaxParams:30; Func:fpsSTDEV),             // fekSTDEV
    (Symbol:'STDEVP';    MinParams:1; MaxParams:30; Func:fpsSTDEVP),            // fekSTDEVP
    (Symbol:'SUM';       MinParams:0; MaxParams:30; Func:fpsSUM),               // fekSUM
    (Symbol:'SUMIF';     MinParams:2; MaxParams:3;  Func:fpsSUMIF),             // fekSUMIF
    (Symbol:'SUMSQ';     MinParams:0; MaxParams:30; Func:fpsSUMSQ),             // fekSUMSQ
    (Symbol:'VAR';       MinParams:1; MaxParams:30; Func:fpsVAR),               // fekVAR
    (Symbol:'VARP';      MinParams:1; MaxParams:30; Func:fpsVARP),              // fekVARP
  { financial }
    (Symbol:'FV';        MinParams:3; MaxParams:5;  Func:nil),   // fekFV
    (Symbol:'NPER';      MinParams:3; MaxParams:5;  Func:nil),   // fekNPER
    (Symbol:'PMT';       MinParams:3; MaxParams:5;  Func:nil),   // fekPMT
    (Symbol:'PV';        MinParams:3; MaxParams:5;  Func:nil),   // fekPV
    (Symbol:'RATE';      MinParams:3; MaxParams:6;  Func:nil),   // fekRATE
  { logical }
    (Symbol:'AND';       MinParams:0; MaxParams:30; Func:fpsAND),               // fekAND
    (Symbol:'FALSE';     MinParams:0; MaxParams:0;  Func:fpsFALSE),             // fekFALSE
    (Symbol:'IF';        MinParams:2; MaxParams:3;  Func:fpsIF),                // fekIF
    (Symbol:'NOT';       MinParams:1; MaxParams:1;  Func:fpsNOT),               // fekNOT
    (Symbol:'OR';        MinParams:1; MaxParams:30; Func:fpsOR),                // fekOR
    (Symbol:'TRUE';      MinParams:0; MaxParams:0;  Func:fpsTRUE),              // fekTRUE
  {  string }
    (Symbol:'CHAR';      MinParams:1; MaxParams:1;  Func:fpsCHAR),              // fekCHAR
    (Symbol:'CODE';      MinParams:1; MaxParams:1;  Func:fpsCODE),              // fekCODE
    (Symbol:'LEFT';      MinParams:1; MaxParams:2;  Func:fpsLEFT),              // fekLEFT
    (Symbol:'LOWER';     MinParams:1; MaxParams:1;  Func:fpsLOWER),             // fekLOWER
    (Symbol:'MID';       MinParams:3; MaxParams:3;  Func:fpsMID),               // fekMID
    (Symbol:'PROPER';    MinParams:1; MaxParams:1;  Func:nil),   // fekPROPER
    (Symbol:'REPLACE';   MinParams:4; MaxParams:4;  Func:fpsREPLACE),           // fekREPLACE
    (Symbol:'RIGHT';     MinParams:1; MaxParams:2;  Func:fpsRIGHT),             // fekRIGHT
    (Symbol:'SUBSTITUTE';MinParams:3; MaxParams:4;  Func:fpsSUBSTITUTE),        // fekSUBSTITUTE ( * )
    (Symbol:'TRIM';      MinParams:1; MaxParams:1;  Func:fpsTRIM),              // fekTRIM
    (Symbol:'UPPER';     MinParams:1; MaxParams:1;  Func:fpsUPPER),             // fekUPPER
  {  lookup/reference }
    (Symbol:'COLUMN';    MinParams:0; MaxParams:1;  Func:fpsCOLUMN),            // fekCOLUMN
    (Symbol:'COLUMNS';   MinParams:1; MaxParams:1;  Func:fpsCOLUMNS),           // fekCOLUMNS
    (Symbol:'ROW';       MinParams:0; MaxParams:1;  Func:fpsROW),               // fekROW
    (Symbol:'ROWS';      MinParams:1; MaxParams:1;  Func:fpsROWS),              // fekROWS
  { info }
    (Symbol:'CELL';      MinParams:1; MaxParams:2;  Func:fpsCELLINFO),          // fekCELLINFO  ( * )
    (Symbol:'INFO';      MinParams:1; MaxParams:1;  Func:fpsINFO),              // fekINFO      ( * )
    (Symbol:'ISBLANK';   MinParams:1; MaxParams:1;  Func:fpsISBLANK),           // fekIsBLANK
    (Symbol:'ISERR';     MinParams:1; MaxParams:1;  Func:fpsISERR),             // fekIsERR
    (Symbol:'ISERROR';   MinParams:1; MaxParams:1;  Func:fpsISERROR),           // fekIsERROR
    (Symbol:'ISLOGICAL'; MinParams:1; MaxParams:1;  Func:fpsISLOGICAL),         // fekIsLOGICAL
    (Symbol:'ISNA';      MinParams:1; MaxParams:1;  Func:fpsISNA),              // fekIsNA
    (Symbol:'ISNONTEXT'; MinParams:1; MaxParams:1;  Func:fpsISNONTEXT),         // fekIsNONTEXT
    (Symbol:'ISNUMBER';  MinParams:1; MaxParams:1;  Func:fpsISNUMBER),          // fekIsNUMBER
    (Symbol:'ISREF';     MinParams:1; MaxParams:1;  Func:fpsISREF),             // fekIsRef
    (Symbol:'ISTEXT';    MinParams:1; MaxParams:1;  Func:fpsISTEXT),            // fekIsTEXT
    (Symbol:'VALUE';     MinParams:1; MaxParams:1;  Func:fpsVALUE),             // fekValue
  { Other operations }
    (Symbol:'SUM';       MinParams:1; MaxParams:1;  Func:nil)    // fekOpSUM (Unary sum operation). Note: CANNOT be used for summing sell contents; use fekSUM}
  );

{@@
  Registers a function used when calculating a formula.
  This feature allows to extend the built-in functions directly available in
  fpspreadsheet.

  @param  AFormulaKind   Identifier of the formula element
  @param  AFunc          Function to be executed when the identifier is met
                         in an rpn formula. The function declaration MUST
                         follow the structure given by TsFormulaFunc.
}
procedure RegisterFormulaFunc(AFormulaKind: TFEKind; AFunc: Pointer);
begin
  FEProps[AFormulaKind].Func := TsFormulaFunc(AFunc);
end;
  *)

{@@
  Registers a new reader/writer pair for a given spreadsheet file format
}
procedure RegisterSpreadFormat(
  AReaderClass: TsSpreadReaderClass;
  AWriterClass: TsSpreadWriterClass;
  AFormat: TsSpreadsheetFormat);
var
  len: Integer;
begin
  len := Length(GsSpreadFormats);
  SetLength(GsSpreadFormats, len + 1);
  
  GsSpreadFormats[len].ReaderClass := AReaderClass;
  GsSpreadFormats[len].WriterClass := AWriterClass;
  GsSpreadFormats[len].Format := AFormat;
end;

{@@
  Returns the name of the given spreadsheet file format.

  @param   AFormat  Identifier of the file format
  @return  'BIFF2', 'BIFF3', 'BIFF4', 'BIFF5', 'BIFF8', 'OOXML', 'Open Document',
           'CSV, 'WikiTable Pipes', or 'WikiTable WikiMedia"
}
function GetFileFormatName(AFormat: TsSpreadsheetFormat): string;
begin
  case AFormat of
    sfExcel2              : Result := 'BIFF2';
    {
    sfExcel3              : Result := 'BIFF3';
    sfExcel4              : Result := 'BIFF4';
    }
    sfExcel5              : Result := 'BIFF5';
    sfExcel8              : Result := 'BIFF8';
    sfooxml               : Result := 'OOXML';
    sfOpenDocument        : Result := 'Open Document';
    sfCSV                 : Result := 'CSV';
    sfWikiTable_Pipes     : Result := 'WikiTable Pipes';
    sfWikiTable_WikiMedia : Result := 'WikiTable WikiMedia';
    else                    Result := lpUnknownSpreadsheetFormat;
  end;
end;


{@@
  If a palette is coded as big-endian (e.g. by copying the rgb values from
  the OpenOffice doc) the palette values can be converted by means of this
  procedure to little-endian which is required internally by TsWorkbook.

  @param APalette     Pointer to the palette to be converted. After conversion,
                      its color values are replaced.
  @param APaletteSize Number of colors contained in the palette
}
procedure MakeLEPalette(APalette: PsPalette; APaletteSize: Integer);
var
  i: Integer;
begin
 {$PUSH}{$R-}
  for i := 0 to APaletteSize-1 do
    APalette^[i] := LongRGBToExcelPhysical(APalette^[i])
 {$POP}
end;

{@@
  Copies the format of a cell to another one.

  @param  AFromCell   cell from which the format is to be copied
  @param  AToCell     cell to which the format is to be copied
}
procedure CopyCellFormat(AFromCell, AToCell: PCell);
begin
  Assert(AFromCell <> nil);
  Assert(AToCell <> nil);

  AToCell^.UsedFormattingFields := AFromCell^.UsedFormattingFields;
  AToCell^.BackgroundColor := AFromCell^.BackgroundColor;
  AToCell^.Border := AFromCell^.Border;
  AToCell^.BorderStyles := AFromCell^.BorderStyles;
  AToCell^.FontIndex := AFromCell^.FontIndex;
  AToCell^.HorAlignment := AFromCell^.HorAlignment;
  AToCell^.VertAlignment := AFromCell^.VertAlignment;
  AToCell^.TextRotation := AFromCell^.TextRotation;
  AToCell^.NumberFormat := AFromCell^.NumberFormat;
  AToCell^.NumberFormatStr := AFromCell^.NumberFormatStr;
end;

{@@
  Checks whether two cells have same border attributes }
function SameCellBorders(ACell1, ACell2: PCell): Boolean;

  function NoBorder(ACell: PCell): Boolean;
  begin
    Result := (ACell = nil) or
      not (uffBorder in ACell^.UsedFormattingFields) or
      (ACell^.Border = []);
  end;

var
  nobrdr1, nobrdr2: Boolean;
  cb: TsCellBorder;
begin
  nobrdr1 := NoBorder(ACell1);
  nobrdr2 := NoBorder(ACell2);
  if (nobrdr1 and nobrdr2) then
    Result := true
  else
  if (nobrdr1 and (not nobrdr2) ) or ( (not nobrdr1) and nobrdr2) then
    Result := false
  else begin
    Result := false;
    if ACell1^.Border <> ACell2^.Border then
      exit;
    for cb in TsCellBorder do begin
      if ACell1^.BorderStyles[cb].LineStyle <> ACell2^.BorderStyles[cb].LineStyle then
        exit;
      if ACell1^.BorderStyles[cb].Color <> ACell2^.BorderStyles[cb].Color then
        exit;
    end;
    Result := true;
  end;
end;

{@@
  Initalizes a new cell
  @return  New cell record
}
procedure InitCell(out ACell: TCell);
begin
  ACell.FormulaValue := '';
  ACell.UTF8StringValue := '';
  ACell.NumberFormatStr := '';
  FillChar(ACell, SizeOf(ACell), 0);
end;

{@@
  Initalizes a new cell and presets the row and column fields of the cell record
  to the parameters passesd to the procedure.

  @param  ARow   Row index of the new cell
  @param  ACol   Column index of the new cell
  @return New cell record with row and column fields preset to passed parameters.
}
procedure InitCell(ARow, ACol: Cardinal; out ACell: TCell);
begin
  InitCell(ACell);
  ACell.Row := ARow;
  ACell.Col := ACol;
end;

{@@
  Returns TRUE if the cell contains a formula (direct or shared, does not matter).
}
function HasFormula(ACell: PCell): Boolean;
begin
  Result := Assigned(ACell) and (
    (ACell^.SharedFormulaBase <> nil) or (Length(ACell^.FormulaValue) > 0)
  );
end;



{ TsWorksheet }

{@@
  Helper method for clearing the records in a spreadsheet.
}
procedure TsWorksheet.RemoveCallback(data, arg: pointer);
begin
  Unused(arg);
  (*
  { The strings and dyn arrays must be reset to nil content manually, because
    FreeMem only frees the record mem, without checking its content }
  PCell(data).UTF8StringValue := '';
  PCell(data).NumberFormatStr := '';
  SetLength(PCell(data).RPNFormulaValue, 0);
//  FreeMem(data);
*)
  Dispose(PCell(data));
end;

function CompareCells(Item1, Item2: Pointer): Integer;
begin
  result := PCell(Item1).Row - PCell(Item2).Row;
  if Result = 0 then
    Result := PCell(Item1).Col - PCell(Item2).Col;
end;

function CompareRows(Item1, Item2: Pointer): Integer;
begin
  result := PRow(Item1).Row - PRow(Item2).Row;
end;

function CompareCols(Item1, Item2: Pointer): Integer;
begin
  result := PCol(Item1).Col - PCol(Item2).Col;
end;

{@@
  Constructor of the TsWorksheet class.
}
constructor TsWorksheet.Create;
begin
  inherited Create;

  FCells := TAVLTree.Create(@CompareCells);
  FRows := TIndexedAVLTree.Create(@CompareRows);
  FCols := TIndexedAVLTree.Create(@CompareCols);

  FFirstRowIndex := $FFFFFFFF;
  FFirstColIndex := $FFFFFFFF;
  FLastRowIndex := 0;
  FLastColIndex := 0;

  FOptions := [soShowGridLines, soShowHeaders];
end;

{@@
  Destructor of the TsWorksheet class.
}
destructor TsWorksheet.Destroy;
begin
  RemoveAllCells;
  RemoveAllRows;
  RemoveAllCols;

  FCells.Free;
  FRows.Free;
  FCols.Free;

  inherited Destroy;
end;

{@@
  Helper function which constructs an rpn formula from the cell's string
  formula. This is needed, for example, when writing a formula to xls biff
  file format.
}
function TsWorksheet.BuildRPNFormula(ACell: PCell): TsRPNFormula;
var
  parser: TsSpreadsheetParser;
begin
  if not HasFormula(ACell) then begin
    SetLength(Result, 0);
    exit;
  end;
  parser := TsSpreadsheetParser.Create(self);
  try
    parser.Expression := ACell^.FormulaValue;
    Result := parser.RPNFormula;
  finally
    parser.Free;
  end;
end;

{@@
  Helper method for calculation of the formulas in a spreadsheet.
}
procedure TsWorksheet.CalcFormulaCallback(data, arg: pointer);
var
  cell: PCell;
begin
  Unused(arg);
  cell := PCell(data);

  // Empty cell or error cell --> nothing to do
  if (cell = nil) or (cell^.ContentType = cctError) then
    exit;

  if HasFormula(cell) or HasFormula(cell^.SharedFormulaBase) then
    CalcFormula(cell);
end;

{@@
  Calculates the formula in a cell
  Should not be called by itself because the result may depend on other cells
  which may have not yet been calculated. It is better to call CalcFormulas
  instead.

  @param  ACell  Cell containing the formula.
}
procedure TsWorksheet.CalcFormula(ACell: PCell);
var
  parser: TsSpreadsheetParser;
  res: TsExpressionResult;
  formula: String;
begin
  ACell^.CalcState := csCalculating;

  parser := TsSpreadsheetParser.Create(self);
  try
    if ACell^.SharedFormulaBase = nil then
      formula := ACell^.FormulaValue
    else
      formula := ACell^.SharedFormulaBase^.FormulaValue;
    parser.Expression := formula;
    parser.EvaluateExpression(res);
    case res.ResultType of
      rtEmpty    : WriteBlank(ACell);
      rtError    : WriteErrorValue(ACell, res.ResError);
      rtInteger  : WriteNumber(ACell, res.ResInteger);
      rtFloat    : WriteNumber(ACell, res.ResFloat);
      rtDateTime : WriteDateTime(ACell, res.ResDateTime);
      rtString   : WriteUTF8Text(ACell, res.ResString);
      rtBoolean  : WriteBoolValue(ACell, res.ResBoolean);
    end;
  finally
    parser.Free;
  end;

  ACell^.CalcState := csCalculated;
end;

{@@
  Calculates all rpn formulas of the worksheet.
}
procedure TsWorksheet.CalcFormulas;
var
  node: TAVLTreeNode;
begin
  // prevent infinite loop due to triggering of formula calculation whenever
  // a cell changes during execution of CalcFormulas.
  inc(FWorkbook.FCalculationLock);
  try
    // Step 1 - mark all formula cells as "not calculated"
    node := FCells.FindLowest;
    while Assigned(node) do begin
      CalcStateCallback(node.Data, nil);
      node := FCells.FindSuccessor(node);
    end;

    // Step 2 - calculate cells. If a not-calculated cell is found it is
    // calculated and then marked as such.
    node := FCells.FindLowest;
    while Assigned(Node) do begin
      CalcFormulaCallback(Node.Data, nil);
      node := FCells.FindSuccessor(node);
    end;
  finally
    dec(FWorkbook.FCalculationLock);
  end;
end;

{@@
  Helper method marking all cells with formulas as "not calculated". This flag
  is needed for recursive calculation of the entire worksheet.
}
procedure TsWorksheet.CalcStateCallback(data, arg: Pointer);
var
  cell: PCell;
begin
  Unused(arg);
  cell := PCell(data);

  if HasFormula(cell) then
    cell^.CalcState := csNotCalculated;
end;

{@@
  Converts a FPSpreadsheet cell position, which is Row, Col in numbers
  and zero based - e.g. 0,0 - to a textual representation which is [Col][Row],
  where the Col is in letters and the row is in 1-based numbers - e.g. A1 }
class function TsWorksheet.CellPosToText(ARow, ACol: Cardinal): string;
begin
  Result := GetCellString(ARow, ACol, [rfRelCol, rfRelRow]);
end;

{@@
  Checks entire workbook, whether this cell is used in any formula.

  @param   ARow  Row index of the cell considered
  @param   ACol  Column index of the cell considered
  @return  TRUE if the cell is used in a formula, FALSE if not
}
function TsWorksheet.CellUsedInFormula(ARow, ACol: Cardinal): Boolean;
var
  cell: PCell;
  cellNode: TAVLTreeNode;
  fe: TsFormulaElement;
  i: Integer;
  rpnFormula: TsRPNFormula;
begin
  cellNode := FCells.FindLowest;
  while Assigned(cellNode) do begin
    cell := PCell(cellNode.Data);
    if HasFormula(cell) then begin
      rpnFormula := BuildRPNFormula(cell);
      for i := 0 to Length(rpnFormula)-1 do
      begin
        fe := rpnFormula[i];
        case fe.ElementKind of
          fekCell, fekCellRef:
            if (fe.Row = ARow) and (fe.Col = ACol) then
            begin
              Result := true;
              exit;
            end;
          fekCellRange:
            if (fe.Row <= ARow) and (ARow <= fe.Row2) and
               (fe.Col <= ACol) and (ACol <= fe.Col2) then
            begin
              Result := true;
              exit;
            end;
        end;
      end;
    end;
    cellNode := FCells.FindSuccessor(cellNode);
  end;
  SetLength(rpnFormula, 0);
end;

{@@
  Is called whenever a cell value or formatting has changed. Fires an event
  "OnChangeCell". This is handled by TsWorksheetGrid to update the grid cell.

  @param  ARow   Row index of the cell which has been changed
  @param  ACol   Column index of the cell which has been changed
}
procedure TsWorksheet.ChangedCell(ARow, ACol: Cardinal);
begin
  if (FWorkbook.FCalculationLock = 0) and (boAutoCalc in FWorkbook.Options) then
  begin
    if CellUsedInFormula(ARow, ACol) then
      CalcFormulas;
  end;
  if Assigned(FOnChangeCell) then
    FOnChangeCell(Self, ARow, ACol);
end;

{@@
  Is called whenever a font height changes. Fires an even "OnChangeFont"
  which is handled by TsWorksheetGrid to update the row heights.

  @param  ARow  Row index of the cell for which the font height has changed
  @param  ACol  Column index of the cell for which the font height has changed.
}
procedure TsWorksheet.ChangedFont(ARow, ACol: Cardinal);
begin
  if Assigned(FonChangeFont) then FOnChangeFont(Self, ARow, ACol);
end;

{@@
  Copies a cell. The source cell can be located in a different worksheet, while
  the destination cell must be in the same worksheet which calls the methode.

  @param AFromRow  Row index of the source cell
  @param AFromCol  Column index of the source cell
  @param AToRow    Row index of the destination cell
  @param AToCol    Column index of the destination cell
  @param AFromWorksheet  Worksheet containing the source cell.
}
procedure TsWorksheet.CopyCell(AFromRow, AFromCol, AToRow, AToCol: Cardinal;
  AFromWorksheet: TsWorksheet);
var
  lSrcCell, lDestCell: PCell;
begin
  lSrcCell := AFromWorksheet.FindCell(AFromRow, AFromCol);
  lDestCell := GetCell(AToRow, AToCol);
  lDestCell^ := lSrcCell^;
  lDestCell^.Row := AToRow;
  lDestCell^.Col := AToCol;
  ChangedCell(AToRow, AToCol);
  ChangedFont(AToRow, AToCol);
end;

{@@
  Copies all format parameters from the format cell to another cell.

  @param AFromCell  Pointer to source cell
  @param AToCell    Pointer to destination cell
}
procedure TsWorksheet.CopyFormat(AFromCell, AToCell: PCell);
begin
  if (AFromCell = nil) or (AToCell = nil) then
    exit;

  CopyCellFormat(AFromCell, AToCell);
  ChangedCell(AToCell^.Row, AToCell^.Col);
  ChangedFont(AToCell^.Row, AToCell^.Col);
end;

{@@
  Copies all format parameters from a given cell to another cell identified
  by its row/column indexes.

  @param  AFormat  Pointer to the source cell from which the format is copied.
  @param  AToRow   Row index of the destination cell
  @param  AToCol   Column index of the destination cell
}
procedure TsWorksheet.CopyFormat(AFormat: PCell; AToRow, AToCol: Cardinal);
begin
  CopyFormat(AFormat, GetCell(AToRow, AToCol));
end;

{@@
  Tries to locate a Cell in the list of already written Cells

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @return Pointer to the cell if found, or nil if not found
  @see    TCell
}
function TsWorksheet.FindCell(ARow, ACol: Cardinal): PCell;
var
  LCell: TCell;
  AVLNode: TAVLTreeNode;
begin
  Result := nil;

  LCell.Row := ARow;
  LCell.Col := ACol;
  AVLNode := FCells.Find(@LCell);
  if Assigned(AVLNode) then
    result := PCell(AVLNode.Data);
end;

{@@
  Tries to locate a Cell in the list of already written Cells

  @param  AddressStr  Address of the cell in Excel A1 notation
  @return Pointer to the cell if found, or nil if not found
  @see    TCell
}
function TsWorksheet.FindCell(AddressStr: String): PCell;
var
  r, c: Cardinal;
begin
  if ParseCellString(AddressStr, r, c) then
    Result := FindCell(r, c)
  else
    Result := nil;
end;

{@@
  Obtains an allocated cell at the desired location.

  If the Cell already exists, a pointer to it will be returned.

  If not, then new memory for the cell will be allocated, a pointer to it
  will be returned and it will be added to the list of cells.

  @param  ARow      Row index of the cell
  @param  ACol      Column index of the cell

  @return A pointer to the cell at the desired location.

  @see    TCell
}
function TsWorksheet.GetCell(ARow, ACol: Cardinal): PCell;
begin
  Result := FindCell(ARow, ACol);
  
  if (Result = nil) then
  begin
    New(Result);
    FillChar(Result^, SizeOf(TCell), #0);

    Result^.Row := ARow;
    Result^.Col := ACol;
    Result^.ContentType := cctEmpty;
    Result^.BorderStyles := DEFAULT_BORDERSTYLES;

    Cells.Add(Result);
    if FFirstColIndex = $FFFFFFFF then FFirstColIndex := GetFirstColIndex(true)
      else FFirstColIndex := Min(FFirstColIndex, ACol);
    if FFirstRowIndex = $FFFFFFFF then FFirstRowIndex := GetFirstRowIndex(true)
      else FFirstRowIndex := Min(FFirstRowIndex, ARow);
    if FLastColIndex = 0 then FLastColIndex := GetLastColIndex(true)
      else FLastColIndex := Max(FLastColIndex, ACol);
    if FLastRowIndex = 0 then FLastRowIndex := GetLastRowIndex(true)
      else FLastRowIndex := Max(FLastRowIndex, ARow);
  end;
end;

{@@
  Obtains an allocated cell at the desired location.

  If the Cell already exists, a pointer to it will be returned.

  If not, then new memory for the cell will be allocated, a pointer to it
  will be returned and it will be added to the list of cells.

  @param  AddressStr  Address of the cell in Excel A1 notation (an exception is
                      raised in case on an invalid cell address).
  @return A pointer to the cell at the desired location.

  @see    TCell
}
function TsWorksheet.GetCell(AddressStr: String): PCell;
var
  r, c: Cardinal;
begin
  if ParseCellString(AddressStr, r, c) then
    Result := GetCell(r, c)
  else
    raise Exception.CreateFmt(lpNoValidCellAddress, [AddressStr]);
end;

{@@
  Returns the number of cells in the worksheet with contents.

  This routine is used together with GetFirstCell and GetNextCell
  to iterate througth all cells in a worksheet efficiently.

  @return The number of cells with contents in the worksheet

  @see    TCell
  @see    GetFirstCell
  @see    GetNextCell
}
function TsWorksheet.GetCellCount: Cardinal;
begin
  Result := FCells.Count;
end;

{@@
  Determines some number format attributes (decimal places, currency symbol) of
  a cell

  @param  ACell            Pointer to the cell under investigation
  @param  ADecimals        Number of decimal places that can be extracted from
                           the formatting string, e.g. in case of '0.000' this
                           would be 3.
  @param  ACurrencySymbol  String representing the currency symbol extracted from
                           the formatting string.

  @return true if the the format string could be analyzed successfully, false if not
}
function TsWorksheet.GetNumberFormatAttributes(ACell: PCell; out ADecimals: byte;
  out ACurrencySymbol: String): Boolean;
var
  parser: TsNumFormatParser;
  nf: TsNumberFormat;
begin
  Result := false;
  if ACell <> nil then
  begin
    parser := TsNumFormatParser.Create(FWorkbook, ACell^.NumberFormatStr);
    try
      if parser.Status = psOK then begin
        nf := parser.NumFormat;
        if (nf = nfGeneral) or IsDateTimeFormat(nf) then
        begin
          ADecimals := 2;
          ACurrencySymbol := '?';
        end
        else
        begin
          ADecimals := parser.Decimals;
          ACurrencySymbol := parser.CurrencySymbol;
        end;
        Result := true;
      end;
    finally
      parser.Free;
    end;
  end;
end;

{@@
  Returns the first Cell.

  Use together with GetCellCount and GetNextCell
  to iterate througth all cells in a worksheet efficiently.

  @return The first cell if any exists, nil otherwise

  @see    TCell
  @see    GetCellCount
  @see    GetNextCell
}
function TsWorksheet.GetFirstCell(): PCell;
begin
  FCurrentNode := FCells.FindLowest();
  if FCurrentNode <> nil then
    Result := PCell(FCurrentNode.Data)
  else Result := nil;
end;

{@@
  Returns the next Cell.

  Should always be used either after GetFirstCell or
  after GetNextCell.

  Use together with GetCellCount and GetFirstCell
  to iterate througth all cells in a worksheet efficiently.

  @return The first cell if any exists, nil otherwise

  @see    TCell
  @see    GetCellCount
  @see    GetFirstCell
}
function TsWorksheet.GetNextCell(): PCell;
begin
  FCurrentNode := FCells.FindSuccessor(FCurrentNode);
  if FCurrentNode <> nil then
    Result := PCell(FCurrentNode.Data)
  else Result := nil;
end;

{@@
  Returns the 0-based index of the first column with a cell with contents.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @param  AForceCalculation  The index of the first column is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the first column.
  @see GetCellCount
}
function TsWorksheet.GetFirstColIndex(AForceCalculation: Boolean = false): Cardinal;
var
  AVLNode: TAVLTreeNode;
  i: Integer;
begin
  if AForceCalculation then
  begin
    Result := $FFFFFFFF;
    // Traverse the tree from lowest to highest.
    // Since tree primary sort order is on row lowest col could exist anywhere.
    AVLNode := FCells.FindLowest;
    while Assigned(AVLNode) do
    begin
      Result := Math.Min(Result, PCell(AVLNode.Data)^.Col);
      AVLNode := FCells.FindSuccessor(AVLNode);
    end;
   // In addition, there may be column records defining the column width even
    // without content
    for i:=0 to FCols.Count-1 do
      if FCols[i] <> nil then
        Result := Math.Min(Result, PCol(FCols[i])^.Col);
    // Store the result
    FFirstColIndex := Result;
  end
  else
  begin
    Result := FFirstColIndex;
    if Result = $FFFFFFFF then
      Result := GetFirstColIndex(true);
  end;
end;

{@@
  Returns the 0-based index of the last column with a cell with contents.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @param  AForceCalculation  The index of the last column is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the last column.
  @see GetCellCount
}
function TsWorksheet.GetLastColIndex(AForceCalculation: Boolean = false): Cardinal;
var
  AVLNode: TAVLTreeNode;
  i: Integer;
begin
  if AForceCalculation then
  begin
    Result := 0;
    // Traverse the tree from lowest to highest.
    // Since tree primary sort order is on Row
    // highest Col could exist anywhere.
    AVLNode := FCells.FindLowest;
    While Assigned(AVLNode) do
    begin
      Result := Math.Max(Result, PCell(AVLNode.Data)^.Col);
      AVLNode := FCells.FindSuccessor(AVLNode);
    end;
   // In addition, there may be column records defining the column width even
    // without content
    for i:=0 to FCols.Count-1 do
      if FCols[i] <> nil then
        Result := Math.Max(Result, PCol(FCols[i])^.Col);
    // Store the result
    FLastColIndex := Result;
  end
  else
    Result := FLastColIndex;
end;

{@@
  Deprecated, use GetLastColIndex instead

  @see GetLastColIndex
}
function TsWorksheet.GetLastColNumber: Cardinal;
begin
  Result := GetLastColIndex;
end;

{@@
  Finds the first cell with contents in a given row

  @param  ARow  Index of the row considered
  @return       Pointer to the first cell in this row, or nil if the row is empty.
}
function TsWorksheet.GetFirstCellOfRow(ARow: Cardinal): PCell;
var
  c, n: Cardinal;
begin
  n := GetLastColIndex;
  c := 0;
  Result := FindCell(ARow, c);
  while (result = nil) and (c < n) do
  begin
    inc(c);
    result := FindCell(ARow, c);
  end;
end;

{@@
  Finds the last cell with contents in a given row

  @param  ARow  Index of the row considered
  @return       Pointer to the last cell in this row, or nil if the row is empty.
}
function TsWorksheet.GetLastCellOfRow(ARow: Cardinal): PCell;
var
  c, n: Cardinal;
begin
  n := GetLastColIndex;
  c := n;
  Result := FindCell(ARow, c);
  while (Result = nil) and (c > 0) do
  begin
    dec(c);
    Result := FindCell(ARow, c);
  end;
end;

{@@
  Returns the 0-based index of the first row with a cell with contents.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @param  AForceCalculation  The index of the first row is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the first row.
  @see GetCellCount
}
function TsWorksheet.GetFirstRowIndex(AForceCalculation: Boolean = false): Cardinal;
var
  AVLNode: TAVLTreeNode;
  i: Integer;
begin
  if AForceCalculation then
  begin
    Result := $FFFFFFFF;
    AVLNode := FCells.FindLowest;
    if Assigned(AVLNode) then
      Result := PCell(AVLNode.Data).Row;
    // In addition, there may be row records even for rows without cells.
    for i:=0 to FRows.Count-1 do
      if FRows[i] <> nil then
        Result := Math.Min(Result, PRow(FRows[i])^.Row);
    // Store result
    FFirstRowIndex := Result;
  end
  else
  begin
    Result := FFirstRowIndex;
    if Result = $FFFFFFFF then
      Result := GetFirstRowIndex(true);
  end;
end;

{@@
  Returns the 0-based index of the last row with a cell with contents.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @param  AForceCalculation  The index of the last row is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the last row.
  @see GetCellCount
}
function TsWorksheet.GetLastRowIndex(AForceCalculation: Boolean = false): Cardinal;
var
  AVLNode: TAVLTreeNode;
  i: Integer;
begin
  if AForceCalculation then
  begin
    Result := 0;
    AVLNode := FCells.FindHighest;
    if Assigned(AVLNode) then
      Result := PCell(AVLNode.Data).Row;
    // In addition, there may be row records even for empty rows.
    for i:=0 to FRows.Count-1 do
      if FRows[i] <> nil then
        Result := Math.Max(Result, PRow(FRows[i])^.Row);
    // Store result
    FLastRowIndex := Result;
  end
  else
    Result := FLastRowIndex
end;

{@@
  Deprecated, use GetLastColIndex instead

  @see GetLastColIndex
}
function TsWorksheet.GetLastRowNumber: Cardinal;
begin
  Result := GetLastRowIndex;
end;

{@@
  Reads the contents of a cell and returns an user readable text
  representing the contents of the cell.

  The resulting ansistring is UTF-8 encoded.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @return The text representation of the cell
}
function TsWorksheet.ReadAsUTF8Text(ARow, ACol: Cardinal): ansistring;
begin
  Result := ReadAsUTF8Text(GetCell(ARow, ACol));
end;

{@@
  Reads the contents of a cell and returns an user readable text
  representing the contents of the cell.

  The resulting ansistring is UTF-8 encoded.

  @param  ACell     Pointer to the cell
  @return The text representation of the cell
}
function TsWorksheet.ReadAsUTF8Text(ACell: PCell): ansistring;

  function FloatToStrNoNaN(const Value: Double;
    ANumberFormat: TsNumberFormat; ANumberFormatStr: string): ansistring;
  var
    fs: TFormatSettings;
  begin
    fs := FWorkbook.FormatSettings;
    if IsNan(Value) then
      Result := ''
    else
    if (ANumberFormat = nfGeneral) or (ANumberFormatStr = '') then
      Result := FloatToStr(Value, fs)
    else
    if (ANumberFormat = nfPercentage) then
      Result := FormatFloat(ANumberFormatStr, Value*100, fs)
    else
      Result := FormatFloat(ANumberFormatStr, Value, fs)
  end;

  function DateTimeToStrNoNaN(const Value: Double;
    ANumberFormat: TsNumberFormat; ANumberFormatStr: String): ansistring;
  var
    fmtp, fmtn, fmt0: String;
  begin
    Result := '';
    if not IsNaN(Value) then
    begin
      if (ANumberFormat = nfGeneral) then
      begin
        if frac(Value) = 0 then                 // date only
          ANumberFormatStr := Workbook.FormatSettings.ShortDateFormat
        else if trunc(Value) = 0 then           // time only
          ANumberFormatStr := Workbook.FormatSettings.LongTimeFormat
        else
          ANumberFormatStr := 'cc'
      end else
      if ANumberFormatStr = '' then
        ANumberFormatStr := BuildDateTimeFormatString(ANumberFormat,
          Workbook.FormatSettings, ANumberFormatStr);

      // Saw strange cases in ods where date/time formats contained pos/neg/zero parts.
      // Split to be on the safe side.
      SplitFormatString(ANumberFormatStr, fmtp, fmtn, fmt0);
      if (Value > 0) or ((Value = 0) and (fmt0 = '')) or ((Value < 0) and (fmtn = '')) then
        Result := FormatDateTime(fmtp, Value, [fdoInterval])
      else
      if (Value < 0) then
        Result := FormatDateTime(fmtn, Value, [fdoInterval])
      else
      if (Value = 0) then
        Result := FormatDateTime(fmt0, Value, [fdoInterval]);
    end;
  end;

begin
  Result := '';
  if ACell = nil then
    Exit;

  with ACell^ do
    case ContentType of
      cctNumber:
        Result := FloatToStrNoNaN(NumberValue, NumberFormat, NumberFormatStr);
      cctUTF8String:
        Result := UTF8StringValue;
      cctDateTime:
        Result := DateTimeToStrNoNaN(DateTimeValue, NumberFormat, NumberFormatStr);
      cctBool:
        Result := StrUtils.IfThen(BoolValue, lpTRUE, lpFALSE);
      cctError:
        case TsErrorValue(ErrorValue) of
          errEmptyIntersection  : Result := lpErrEmptyIntersection;
          errDivideByZero       : Result := lpErrDivideByZero;
          errWrongType          : Result := lpErrWrongType;
          errIllegalRef         : Result := lpErrIllegalRef;
          errWrongName          : Result := lpErrWrongName;
          errOverflow           : Result := lpErrOverflow;
          errArgError           : Result := lpErrArgError;
          errFormulaNotSupported: Result := lpErrFormulaNotSupported;
        end;
      else
        Result := '';
    end;
end;

{@@
  Returns the value of a cell as a number.

  If the cell contains a date/time value its serial value is returned
  (as FPC TDateTime).

  If the cell contains a text value it is attempted to convert it to a number.

  If the cell is empty or its contents cannot be represented as a number the
  value 0.0 is returned.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @return Floating-point value representing the cell contents, or 0.0 if cell
          does not exist or its contents cannot be converted to a number.
}
function TsWorksheet.ReadAsNumber(ARow, ACol: Cardinal): Double;
begin
  Result := ReadAsNumber(FindCell(ARow, ACol));
end;

{@@
  Returns the value of a cell as a number.

  If the cell contains a date/time value its serial value is returned
  (as FPC TDateTime).

  If the cell contains a text value it is attempted to convert it to a number.

  If the cell is empty or its contents cannot be represented as a number the
  value 0.0 is returned.

  @param  ACell     Pointer to the cell
  @return Floating-point value representing the cell contents, or 0.0 if cell
          does not exist or its contents cannot be converted to a number.
}
function TsWorksheet.ReadAsNumber(ACell: PCell): Double;
begin
  Result := 0.0;
  if ACell = nil then
    exit;

  case ACell^.ContentType of
    cctDateTime   : Result := ACell^.DateTimeValue; //this is in FPC TDateTime format, not Excel
    cctNumber     : Result := ACell^.NumberValue;
    cctUTF8String : if not TryStrToFloat(ACell^.UTF8StringValue, Result) then Result := 0.0;
    cctBool       : if ACell^.BoolValue then Result := 1.0 else Result := 0.0;
  end;
end;

{@@
  Reads the contents of a cell and returns the date/time value of the cell.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AResult   Date/time value of the cell (or 0.0, if no date/time cell)
  @return True if the cell is a datetime value, false otherwise
}
function TsWorksheet.ReadAsDateTime(ARow, ACol: Cardinal; out AResult: TDateTime): Boolean;
begin
  Result := ReadAsDateTime(FindCell(ARow, ACol), AResult);
end;

{@@
  Reads the contents of a cell and returns the date/time value of the cell.

  @param  ACell     Pointer to the cell
  @param  AResult   Date/time value of the cell (or 0.0, if no date/time cell)
  @return True if the cell is a datetime value, false otherwise
}
function TsWorksheet.ReadAsDateTime(ACell: PCell; out AResult: TDateTime): Boolean;
begin
  if (ACell = nil) or (ACell^.ContentType <> cctDateTime) then
  begin
    AResult := 0;
    Result := False;
    Exit;
  end;

  AResult := ACell^.DateTimeValue;
  Result := True;
end;

{@@ If a cell contains a formula (string formula or RPN formula) the formula
  is returned as a string in Excel syntax.

  @param   ACell Pointer to the cell considered
  @return  Formula string in Excel syntax (does not contain a leading "=")
}
function TsWorksheet.ReadFormulaAsString(ACell: PCell): String;
begin
  Result := '';
  if ACell = nil then
    exit;
  if HasFormula(ACell) then
    Result := ACell^.FormulaValue;
end;

{@@
  Returns to numeric equivalent of the cell contents. This is the NumberValue
  of a number cell, the DateTimeValue of a date/time cell, the ordinal BoolValue
  of a boolean cell, or the string converted to a number of a string cell.
  All other cases return NaN.

  @param   ACell   Cell to be considered
  @param   AValue  (output) extracted numeric value
  @return  True if conversion to number is successful, otherwise false }
function TsWorksheet.ReadNumericValue(ACell: PCell; out AValue: Double): Boolean;
begin
  AValue := NaN;
  if ACell <> nil then begin
    Result := True;
    case ACell^.ContentType of
      cctNumber:
        AValue := ACell^.NumberValue;
      cctDateTime:
        AValue := ACell^.DateTimeValue;
      cctBool:
        AValue := ord(ACell^.BoolValue);
      else
        if (ACell^.ContentType <> cctUTF8String) or
           not TryStrToFloat(ACell^.UTF8StringValue, AValue) or
           not TryStrToDateTime(ACell^.UTF8StringValue, AValue)
        then
          Result := False;
      end;
  end else
    Result := False;
end;

{@@
  Converts an RPN formula (as read from an xls biff file, for example) to a
  string formula.

  @param    AFormula  Array of rpn formula tokens
  @return   Formula string in Excel syntax (without leading "=")
}
function TsWorksheet.ConvertRPNFormulaToStringFormula(const AFormula: TsRPNFormula): String;
var
  parser: TsSpreadsheetParser;
begin
  Result := '';

  parser := TsSpreadsheetParser.Create(self);
  try
    parser.RPNFormula := AFormula;
    Result := parser.Expression;
  finally
    parser.Free;
  end;
end;

{@@
  Reads the set of used formatting fields of a cell.

  Each cell contains a set of "used formatting fields". Formatting is applied
  only if the corresponding element is contained in the set.

  @param  ARow    Row index of the considered cell
  @param  ACol    Column index of the considered cell
  @return Set of elements used in formatting the cell
}
function TsWorksheet.ReadUsedFormatting(ARow, ACol: Cardinal): TsUsedFormattingFields;
var
  ACell: PCell;
begin
  ACell := FindCell(ARow, ACol);

  if ACell = nil then
  begin
    Result := [];
    Exit;
  end;

  Result := ACell^.UsedFormattingFields;
end;

{@@
  Returns the background color of a cell as index into the workbook's color palette.

  @param ARow  Row index of the cell
  @param ACol  Column index of the cell
  @return Index of the cell background color into the workbook's color palette
}
function TsWorksheet.ReadBackgroundColor(ARow, ACol: Cardinal): TsColor;
var
  ACell: PCell;
begin
  ACell := FindCell(ARow, ACol);

  if ACell = nil then
  begin
    Result := scWhite;
    Exit;
  end;

  Result := ACell^.BackgroundColor;
end;

{@@
  Clears the list of cells and releases their memory.
}
procedure TsWorksheet.RemoveAllCells;
var
  Node: TAVLTreeNode;
begin
  Node:=FCells.FindLowest;
  while Assigned(Node) do begin
    RemoveCallback(Node.Data,nil);
    Node.Data:=nil;
    Node:=FCells.FindSuccessor(Node);
  end;
  FCells.Clear;
end;

{@@
  Helper method to update internal caching variables
}
procedure TsWorksheet.UpdateCaches;
begin
  FFirstColIndex := GetFirstColIndex(true);
  FFirstRowIndex := GetFirstRowIndex(true);
  FLastColIndex := GetLastColIndex(true);
  FLastRowIndex := GetLastRowIndex(true);
end;

{@@
  Creates a link to a shared formula

  @param  ARow                Row of the cell
  @param  ACol                Column index of the cell
  @param  ASharedFormulaBase  Cell containing the shared formula token array

  Note:   An exception is raised if the cell already contains a formula (and is
          different from the ASharedFormulaBase cell).
}
function TsWorksheet.UseSharedFormula(ARow, ACol: Cardinal;
  ASharedFormulaBase: PCell): PCell;
begin
  if ASharedFormulaBase = nil then begin
    Result := nil;
    exit;
  end;
  Result := GetCell(ARow, ACol);
  Result.SharedFormulaBase := ASharedFormulaBase;
  if HasFormula(Result) and
     ((ASharedFormulaBase.Row <> ARow) and (ASharedFormulaBase.Col <> ACol))
  then
    raise Exception.CreateFmt('Cell %s uses a shared formula, but contains an own formula.',
      [GetCellString(ARow, ACol)]);
end;

{@@
  Uses the formula defined in cell "ASharedFormulaBase" as a shared formula in
  all cells of the given cell range.

  @param ACellRangeStr       Range of cells which will use the shared formula.
                             The range is given as a string in Excel notation,
                             such as A1:B5, or A1
  @param ASharedFormulaBase  Cell containing the formula to be shared
}
procedure TsWorksheet.UseSharedFormula(ACellRangeStr: String; ASharedFormulaBase: PCell);
var
  r, c, r1, c1, r2, c2: Cardinal;
  ok: Boolean;
begin
  if pos(':', ACellRangeStr) = 0 then
  begin
    ok := ParseCellString(ACellRangeStr, r1, c1);
    r2 := r1;
    c2 := c1;
  end else
    ok := ParseCellRangeString(ACellRangeStr, r1, c1, r2, c2);
  if ok then
    for r := r1 to r2 do
      for c := c1 to c2 do
        UseSharedFormula(r, c, ASharedFormulaBase);
end;

{@@
  Writes UTF-8 encoded text to a cell.

  On formats that don't support unicode, the text will be converted
  to ISO Latin 1.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AText     The text to be written encoded in utf-8
  @return Pointer to cell created or used
}
function TsWorksheet.WriteUTF8Text(ARow, ACol: Cardinal; AText: ansistring): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteUTF8Text(Result, AText);
end;

{@@
  Writes UTF-8 encoded text to a cell.

  On formats that don't support unicode, the text will be converted
  to ISO Latin 1.

  @param  ACell     Poiner to the cell
  @param  AText     The text to be written encoded in utf-8
}
procedure TsWorksheet.WriteUTF8Text(ACell: PCell; AText: ansistring);
begin
  if ACell = nil then
    exit;
  ACell^.ContentType := cctUTF8String;
  ACell^.UTF8StringValue := AText;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Writes a floating-point number to a cell. Does not change number format.

  @param  ARow      Cell row index
  @param  ACol      Cell column index
  @param  ANumber   Number to be written
  @return Pointer to cell created or used
}
function TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: double): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumber(Result, ANumber);
end;

{@@
  Writes a floating-point number to a cell. Does not change number format.

  @param  ARow      Cell row index
  @param  ACol      Cell column index
  @param  ANumber   Number to be written
}
procedure TsWorksheet.WriteNumber(ACell: PCell; ANumber: double);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := ANumber;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a floating-point number to a cell

  @param  ARow      Cell row index
  @param  ACol      Cell column index
  @param  ANumber   Number to be written
  @param  AFormat   Identifier for a built-in number format, e.g. nfFixed (optional)
  @param  ADecimals Number of decimal places used for formatting (optional)
  @return Pointer to cell created or used
  @see    TsNumberFormat
}
function TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: double;
  AFormat: TsNumberFormat; ADecimals: Byte = 2): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumber(Result, ANumber, AFormat, ADecimals);
end;

{@@
  Writes a floating-point number to a cell

  @param  ACell     Pointer to the cell
  @param  ANumber   Number to be written
  @param  AFormat   Identifier for a built-in number format, e.g. nfFixed (optional)
  @param  ADecimals Number of decimal places used for formatting (optional)
  @see TsNumberFormat
}
procedure TsWorksheet.WriteNumber(ACell: PCell; ANumber: Double;
  AFormat: TsNumberFormat; ADecimals: Byte = 2);
begin
  if IsDateTimeFormat(AFormat) or IsCurrencyFormat(AFormat) then
    raise Exception.Create(lpInvalidNumberFormat);

  if ACell <> nil then begin
    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := ANumber;
    ACell^.NumberFormat := AFormat;

    if AFormat <> nfGeneral then begin
      Include(ACell^.UsedFormattingFields, uffNumberFormat);
      ACell^.NumberFormatStr := BuildNumberFormatString(ACell^.NumberFormat,
        Workbook.FormatSettings, ADecimals);
    end else begin
      Exclude(ACell^.UsedFormattingFields, uffNumberFormat);
      ACell^.NumberFormatStr := '';
    end;

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a floating point number to the cell and uses a custom number format
  specified by the format string.
  Note that fpspreadsheet may not be able to detect the formatting when reading
  the file.

  @param  ARow           Cell row index
  @param  ACol           Cell column index
  @param  ANumber        Number to be written
  @param  AFormat        Format identifier (nfCustom)
  @param  AFormatString  String of formatting codes (such as 'dd/mmm'
  @return Pointer to cell created or used
}
function TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: Double;
  AFormat: TsNumberFormat; AFormatString: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumber(Result, ANumber, AFormat, AFormatString);
end;

{@@
  Writes a floating point number to the cell and uses a custom number format
  specified by the format string.
  Note that fpspreadsheet may not be able to detect the formatting when reading
  the file.

  @param  ACell          Pointer to the cell considered
  @param  ANumber        Number to be written
  @param  AFormat        Format identifier (nfCustom)
  @param  AFormatString  String of formatting codes (such as 'dd/mmm'
}
procedure TsWorksheet.WriteNumber(ACell: PCell; ANumber: Double;
  AFormat: TsNumberFormat; AFormatString: String);
var
  parser: TsNumFormatParser;
begin
  if ACell <> nil then begin
    parser := TsNumFormatParser.Create(Workbook, AFormatString);
    try
      // Format string ok?
      if parser.Status <> psOK then
        raise Exception.Create(lpNoValidNumberFormatString);
      // Make sure that we do not write a date/time value here
      if parser.IsDateTimeFormat
        then raise Exception.Create(lpInvalidNumberFormat);
      // If format string matches a built-in format use its format identifier,
      // All this is considered when calling Builtin_NumFormat of the parser.
    finally
      parser.Free;
    end;

    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := ANumber;
    ACell^.NumberFormat := AFormat;
    if AFormat <> nfGeneral then begin
      Include(ACell^.UsedFormattingFields, uffNumberFormat);
      ACell^.NumberFormatStr := AFormatString;
    end else begin
      Exclude(ACell^.UsedFormattingFields, uffNumberFormat);
      ACell^.NumberFormatStr := '';
    end;

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes as empty cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @return Pointer to the cell
  Note:   Empty cells are useful when, for example, a border line extends
          along a range of cells including empty cells.
}
function TsWorksheet.WriteBlank(ARow, ACol: Cardinal): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBlank(Result);
end;

{@@
  Writes as empty cell

  @param  ACel      Pointer to the cell
  Note:   Empty cells are useful when, for example, a border line extends
          along a range of cells including empty cells.
}
procedure TsWorksheet.WriteBlank(ACell: PCell);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctEmpty;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes as boolean cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The boolean value
  @return Pointer to the cell
}
function TsWorksheet.WriteBoolValue(ARow, ACol: Cardinal; AValue: Boolean): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBoolValue(Result, AValue);
end;

{@@
  Writes as boolean cell

  @param  ACell      Pointer to the cell
  @param  AValue     The boolean value
}
procedure TsWorksheet.WriteBoolValue(ACell: PCell; AValue: Boolean);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctBool;
    ACell^.BoolValue := AValue;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes data defined as a string into a cell. Depending on the structure of the
  string, the worksheet tries to guess whether it is a number, a date/time or
  a text and calls the corresponding writing method.

  @param  ARow    Row index of the cell
  @param  ACol    Column index of the cell
  @param  AValue  Value to be written into the cell given as a string. Depending
                  on the structure of the string, however, the value is written
                  as a number, a date/time or a text.
  @return Pointer to the cell
}
function TsWorksheet.WriteCellValueAsString(ARow, ACol: Cardinal;
  AValue: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteCellValueAsString(Result, AValue);
end;

{@@
  Writes data defined as a string into a cell. Depending on the structure of the
  string, the worksheet tries to guess whether it is a number, a date/time or
  a text and calls the corresponding writing method.

  @param  ACell   Poiner to the cell
  @param  AValue  Value to be written into the cell given as a string. Depending
                  on the structure of the string, however, the value is written
                  as a number, a date/time or a text.
}
procedure TsWorksheet.WriteCellValueAsString(ACell: PCell; AValue: String);
var
  isPercent: Boolean;
  number: Double;
begin
  if ACell = nil then
    exit;

  if AValue = '' then begin
    WriteBlank(ACell^.Row, ACell^.Col);
    exit;
  end;

  isPercent := Pos('%', AValue) = Length(AValue);
  if isPercent then Delete(AValue, Length(AValue), 1);

  if TryStrToFloat(AValue, number) then begin
    if isPercent then
      WriteNumber(ACell, number/100, nfPercentage)
    else begin
      if IsDateTimeFormat(ACell^.NumberFormat) then begin
        ACell^.NumberFormat := nfGeneral;
        ACell^.NumberFormatStr := '';
      end;
      WriteNumber(ACell, number, ACell^.NumberFormat, ACell^.NumberFormatStr);
    end;
    exit;
  end;

  if TryStrToDateTime(AValue, number) then begin
    if number < 1.0 then begin    // this is a time alone
      if not IsTimeFormat(ACell^.NumberFormat) then begin
        ACell^.NumberFormat := nfLongTime;
        ACell^.NumberFormatStr := '';
      end;
    end else
    if frac(number) = 0.0 then begin  // this is a date alone
      if not (ACell^.NumberFormat in [nfShortDate, nfLongDate, nfShortDateTime])
      then begin
        ACell^.NumberFormat := nfShortDate;
        ACell^.NumberFormatStr := '';
      end;
    end else begin
      if not IsDateTimeFormat(ACell^.NumberFormat) then begin
        ACell^.NumberFormat := nfShortDateTime;
        ACell^.NumberFormatStr := '';
      end;
    end;
    WriteDateTime(ACell, number, ACell^.NumberFormat, ACell^.NumberFormatStr);
    exit;
  end;

  WriteUTF8Text(ACell, AValue);
end;

{@@
  Writes a currency value to a given cell. Its number format can be provided
  optionally by specifying various parameters.

  @param ARow            Cell row index
  @param ACol            Cell column index
  @param AValue          Number value to be written
  @param AFormat         Format identifier, must be nfCurrency, or nfCurrencyRed.
  @param ADecimals       Number of decimal places
  @param APosCurrFormat  Code specifying the order of value, currency symbol
                         and spaces (see pcfXXXX constants)
  @param ANegCurrFormat  Code specifying the order of value, currency symbol,
                         spaces, and how negative values are shown
                         (see ncfXXXX constants)
  @param ACurrencySymbol String to be shown as currency, such as '$', or 'EUR'.
                         In case of '?' the currency symbol defined in the
                         workbook's FormatSettings is used.
  @return Pointer to the cell
}
function TsWorksheet.WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
  AFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = 2;
  ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
  ANegCurrFormat: Integer = -1): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteCurrency(Result, AValue, AFormat, ADecimals, ACurrencySymbol,
    APosCurrFormat, ANegCurrFormat);
end;

{@@
  Writes a currency value to a given cell. Its number format can be provided
  optionally by specifying various parameters.

  @param ACell           Pointer to the cell considered
  @param AValue          Number value to be written
  @param AFormat         Format identifier, must be nfCurrency or nfCurrencyRed.
  @param ADecimals       Number of decimal places
  @param APosCurrFormat  Code specifying the order of value, currency symbol
                         and spaces (see pcfXXXX constants)
  @param ANegCurrFormat  Code specifying the order of value, currency symbol,
                         spaces, and how negative values are shown
                         (see ncfXXXX constants)
  @param ACurrencySymbol String to be shown as currency, such as '$', or 'EUR'.
                         In case of '?' the currency symbol defined in the
                         workbook's FormatSettings is used.
}
procedure TsWorksheet.WriteCurrency(ACell: PCell; AValue: Double;
  AFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = -1;
  ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
  ANegCurrFormat: Integer = -1);
var
  fmt: String;
begin
  if ADecimals = -1 then
    ADecimals := Workbook.FormatSettings.CurrencyDecimals;
  if APosCurrFormat = -1 then
    APosCurrFormat := Workbook.FormatSettings.CurrencyFormat;
  if ANegCurrFormat = -1 then
    ANegCurrFormat := Workbook.FormatSettings.NegCurrFormat;
  if ACurrencySymbol = '?' then
    ACurrencySymbol := AnsiToUTF8(Workbook.FormatSettings.CurrencyString);

  fmt := BuildCurrencyFormatString(
    nfdDefault,
    AFormat,
    Workbook.FormatSettings,
    ADecimals,
    APosCurrFormat, ANegCurrFormat,
    ACurrencySymbol);

  WriteCurrency(ACell, AValue, AFormat, fmt);
end;

{@@
  Writes a currency value to a given cell. Its number format is specified by
  means of a format string.

  @param ARow            Cell row index
  @param ACol            Cell column index
  @param AValue          Number value to be written
  @param AFormat         Format identifier, must be nfCurrency or nfCurrencyRed.
  @param AFormatString   String of formatting codes, including currency symbol.
                         Can contain sections for different formatting of positive
                         and negative number. Example: '"EUR" #,##0.00;("EUR" #,##0.00)'
  @return Pointer to the cell
}
function TsWorksheet.WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
  AFormat: TsNumberFormat; AFormatString: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteCurrency(Result, AValue, AFormat, AFormatString);
end;

{@@
  Writes a currency value to a given cell. Its number format is specified by
  means of a format string.

  @param ACell           Pointer to the cell considered
  @param AValue          Number value to be written
  @param AFormat         Format identifier, must be nfCurrency or nfCurrencyRed.
  @param AFormatString   String of formatting codes, including currency symbol.
                         Can contain sections for different formatting of positive
                         and negative number. Example: '"EUR" #,##0.00;("EUR" #,##0.00)'
}
procedure TsWorksheet.WriteCurrency(ACell: PCell; AValue: Double;
  AFormat: TsNumberFormat; AFormatString: String);
begin
  if (ACell <> nil) and IsCurrencyFormat(AFormat) then begin
    Include(ACell^.UsedFormattingFields, uffNumberFormat);
    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := AValue;
    ACell^.NumberFormat := AFormat;
    ACell^.NumberFormatStr := AFormatString;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a date/time value to a cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The date/time/datetime to be written
  @param  AFormat    The format specifier, e.g. nfShortDate (optional)
                     If not specified format is not changed.
  @param  AFormatStr Format string, used only for nfCustom or nfTimeInterval.
  @return Pointer to the cell

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
}
function TsWorksheet.WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
  AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTime(Result, AValue, AFormat, AFormatStr);
end;

{@@
  Writes a date/time value to a cell

  @param  ACell      Pointer to the cell considered
  @param  AValue     The date/time/datetime to be written
  @param  AFormat    The format specifier, e.g. nfShortDate (optional)
                     If not specified format is not changed.
  @param  AFormatStr Format string, used only for nfCustom or nfTimeInterval.

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
}
procedure TsWorksheet.WriteDateTime(ACell: PCell; AValue: TDateTime;
  AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = '');
var
  parser: TsNumFormatParser;
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctDateTime;
    ACell^.DateTimeValue := AValue;

    // Date/time is actually a number field in Excel.
    // To make sure it gets saved correctly, set a date format (instead of General).
    // The user can choose another date format if he wants to

    if AFormat = nfGeneral then begin
      if trunc(AValue) = 0 then         // time only
        AFormat := nfLongTime
      else if frac(AValue) = 0.0 then   // date only
        AFormat := nfShortDate;
    end;

    if AFormatStr = '' then
      AFormatStr := BuildDateTimeFormatString(AFormat, Workbook.FormatSettings, AFormatStr)
    else
    if AFormat = nfTimeInterval then
      AFormatStr := AddIntervalBrackets(AFormatStr);

    // Check whether the formatstring is for date/times.
    if AFormatStr <> '' then begin
      parser := TsNumFormatParser.Create(Workbook, AFormatStr);
      try
        // Format string ok?
        if parser.Status <> psOK then
          raise Exception.Create(lpNoValidNumberFormatString);
        // Make sure that we do not use a number format for date/times values.
        if not parser.IsDateTimeFormat
          then raise Exception.Create(lpInvalidDateTimeFormat);
        // Avoid possible duplication of standard formats
        if AFormat = nfCustom then
          AFormat := parser.NumFormat;
      finally
        parser.Free;
      end;
    end;

    Include(ACell^.UsedFormattingFields, uffNumberFormat);
    ACell^.NumberFormat := AFormat;
    ACell^.NumberFormatStr := AFormatStr;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a date/time value to a cell

  @param  ARow       The row index of the cell
  @param  ACol       The column index of the cell
  @param  AValue     The date/time/datetime to be written
  @param  AFormatStr Format string (the format identifier nfCustom is used to classify the format).
  @return Pointer to the cell

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
}
function TsWorksheet.WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
  AFormatStr: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTime(Result, AValue, AFormatStr);
end;

{@@
  Writes a date/time value to a cell

  @param  ACell      Pointer to the cell considered
  @param  AValue     The date/time/datetime to be written
  @param  AFormatStr Format string (the format identifier nfCustom is used to classify the format).

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
}
procedure TsWorksheet.WriteDateTime(ACell: PCell; AValue: TDateTime;
  AFormatStr: String);
begin
  WriteDateTime(ACell, AValue, nfCustom, AFormatStr);
end;

{@@
  Adds a date/time format to the formatting of a cell

  @param  ARow          The row of the cell
  @param  ACol          The column of the cell
  @param  ANumberFormat Identifier of the format to be applied (nfXXXX constant)
  @param  AFormatString optional string of formatting codes. Is only considered
                        if ANumberFormat is nfCustom.
  @return Pointer to the cell

  @see    TsNumberFormat
}
function TsWorksheet.WriteDateTimeFormat(ARow, ACol: Cardinal;
  ANumberFormat: TsNumberFormat; const AFormatString: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTimeFormat(Result, ANumberFormat, AFormatString);
end;

{@@
  Adds a date/time format to the formatting of a cell

  @param  ACell         Pointer to the cell considered
  @param  ANumberFormat Identifier of the format to be applied (nxXXXX constant)
  @param  AFormatString optional string of formatting codes. Is only considered
                        if ANumberFormat is nfCustom.

  @see    TsNumberFormat
}
procedure TsWorksheet.WriteDateTimeFormat(ACell: PCell;
  ANumberFormat: TsNumberFormat; const AFormatString: String = '');
begin
  if ACell = nil then
    exit;

  if not ((ANumberFormat in [nfGeneral, nfCustom]) or IsDateTimeFormat(ANumberFormat)) then
    raise Exception.Create('WriteDateTimeFormat can only be called with date/time formats.');

  ACell^.NumberFormat := ANumberFormat;
  if (ANumberFormat <> nfGeneral) then begin
    Include(ACell^.UsedFormattingFields, uffNumberFormat);
    if (AFormatString = '') then
      ACell^.NumberFormatStr := BuildDateTimeFormatString(ANumberFormat, Workbook.FormatSettings)
    else
      ACell^.NumberFormatStr := AFormatString;
  end else begin
    Exclude(ACell^.UsedFormattingFields, uffNumberFormat);
    ACell^.NumberFormatStr := '';
  end;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Formats the number in a cell to show a given count of decimal places.
  Is ignored for non-decimal formats (such as most date/time formats).

  @param  ARow       Row indows of the cell considered
  @param  ACol       Column indows of the cell considered
  @param  ADecimals  Number of decimal places to be displayed
  @return Pointer to the cell
  @see    TsNumberFormat
}
function TsWorksheet.WriteDecimals(ARow, ACol: Cardinal; ADecimals: Byte): PCell;
begin
  Result := FindCell(ARow, ACol);
  WriteDecimals(Result, ADecimals);
end;

{@@
  Formats the number in a cell to show a given count of decimal places.
  Is ignored for non-decimal formats (such as most date/time formats).

  @param  ACell      Pointer to the cell considered
  @param  ADecimals  Number of decimal places to be displayed
  @see    TsNumberFormat
}
procedure TsWorksheet.WriteDecimals(ACell: PCell; ADecimals: Byte);
var
  parser: TsNumFormatParser;
begin
  if (ACell <> nil) and (ACell^.ContentType = cctNumber) and (ACell^.NumberFormat <> nfCustom)
  then begin
    parser := TsNumFormatParser.Create(Workbook, ACell^.NumberFormatStr);
    try
      parser.Decimals := ADecimals;
      ACell^.NumberFormatStr := parser.FormatString[nfdDefault];
    finally
      parser.Free;
    end;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes an error value to a cell.

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The error code value
  @return Pointer to the cell

  @see TsErrorValue
}
function TsWorksheet.WriteErrorValue(ARow, ACol: Cardinal; AValue: TsErrorValue): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteErrorValue(Result, AValue);
end;

{@@
  Writes an error value to a cell.

  @param  ACol       Pointer to the cell considered
  @param  AValue     The error code value

  @see TsErrorValue
}
procedure TsWorksheet.WriteErrorValue(ACell: PCell; AValue: TsErrorValue);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctError;
    ACell^.ErrorValue := AValue;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a formula to a given cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AFormula  The formula string to be written. A leading "=" will be removed.
  @return Pointer to the cell
}
function TsWorksheet.WriteFormula(ARow, ACol: Cardinal; AFormula: string): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteFormula(Result, AFormula);
end;

{@@
  Writes a formula to a given cell

  @param  ACell     Pointer to the cell
  @param  AFormula  Formula string to be written. A leading '=' will be removed.
}
procedure TsWorksheet.WriteFormula(ACell: PCell; AFormula: string);
begin
  if ACell = nil then
    exit;
  ACell^.ContentType := cctFormula;
  if (AFormula <> '') and (AFormula[1] = '=') then
    ACell^.FormulaValue := Copy(AFormula, 2, Length(AFormula))
  else
    ACell^.FormulaValue := AFormula;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Adds a number format to the formatting of a cell

  @param  ARow            The row of the cell
  @param  ACol            The column of the cell
  @param  ANumberFormat   Identifier of the format to be applied
  @param  ADecimals       Number of decimal places
  @param  ACurrencySymbol optional currency symbol in case of nfCurrency
  @param  APosCurrFormat  optional identifier for positive currencies
  @param  ANegCurrFormat  optional identifier for negative currencies
  @return Pointer to the cell

  @see    TsNumberFormat
}
function TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumberFormat: TsNumberFormat; ADecimals: Integer; ACurrencySymbol: String = '';
  APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumberFormat(Result, ANumberFormat, ADecimals, ACurrencySymbol,
    APosCurrFormat, ANegCurrFormat);
end;

{@@
  Adds a number format to the formatting of a cell

  @param  ARow            The row of the cell
  @param  ACol            The column of the cell
  @param  ANumberFormat   Identifier of the format to be applied
  @param  ADecimals       Number of decimal places
  @param  ACurrencySymbol optional currency symbol in case of nfCurrency
  @param  APosCurrFormat  optional identifier for positive currencies
  @param  ANegCurrFormat  optional identifier for negative currencies

  @see    TsNumberFormat
}
procedure TsWorksheet.WriteNumberFormat(ACell: PCell;
  ANumberFormat: TsNumberFormat; ADecimals: Integer; ACurrencySymbol: String = '';
  APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1);
begin
  if ACell = nil then
    exit;

  ACell^.NumberFormat := ANumberFormat;
  if ANumberFormat <> nfGeneral then begin
    Include(ACell^.UsedFormattingFields, uffNumberFormat);
    if ANumberFormat in [nfCurrency, nfCurrencyRed] then
      ACell^.NumberFormatStr := BuildCurrencyFormatString(nfdDefault, ANumberFormat,
        Workbook.FormatSettings, ADecimals,
        APosCurrFormat, ANegCurrFormat, ACurrencySymbol)
    else
      ACell^.NumberFormatStr := BuildNumberFormatString(ANumberFormat,
        Workbook.FormatSettings, ADecimals);
  end else begin
    Exclude(ACell^.UsedFormattingFields, uffNumberFormat);
    ACell^.NumberFormatStr := '';
  end;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Adds number format to the formatting of a cell

  @param  ARow          The row of the cell
  @param  ACol          The column of the cell
  @param  ANumberFormat Identifier of the format to be applied
  @param  AFormatString optional string of formatting codes. Is only considered
                        if ANumberFormat is nfCustom.
  @return Pointer to the cell

  @see    TsNumberFormat
}
function TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumberFormat: TsNumberFormat; const AFormatString: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumberFormat(Result, ANumberFormat, AFormatString);
end;

{@@
  Adds a number format to the formatting of a cell

  @param  ACell         Pointer to the cell considered
  @param  ANumberFormat Identifier of the format to be applied
  @param  AFormatString optional string of formatting codes. Is only considered
                        if ANumberFormat is nfCustom.

  @see    TsNumberFormat
}
procedure TsWorksheet.WriteNumberFormat(ACell: PCell;
  ANumberFormat: TsNumberFormat; const AFormatString: String = '');
begin
  if ACell = nil then
    exit;

  ACell^.NumberFormat := ANumberFormat;
  if ANumberFormat <> nfGeneral then begin
    Include(ACell^.UsedFormattingFields, uffNumberFormat);
    if (AFormatString = '') then
      ACell^.NumberFormatStr := BuildNumberFormatString(ANumberFormat, Workbook.FormatSettings)
    else
      ACell^.NumberFormatStr := AFormatString;
  end else begin
    Exclude(ACell^.UsedFormattingFields, uffNumberFormat);
    ACell^.NumberFormatStr := '';
  end;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Writes an RPN formula to a cell. An RPN formula is an array of tokens
  describing the calculation to be performed.

  @param  ARow          Row indows of the cell considered
  @param  ACol          Column index of the cell
  @param  AFormula      Array of TsFormulaElements. The array can be created by
                        using "CreateRPNFormla".
  @return Pointer to the cell

  @see    TsNumberFormat
  @see    TsFormulaElements
  @see    CreateRPNFormula
}
function TsWorksheet.WriteRPNFormula(ARow, ACol: Cardinal;
  AFormula: TsRPNFormula): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteRPNFormula(Result, AFormula);
end;

{@@
  Writes an RPN formula to a cell. An RPN formula is an array of tokens
  describing the calculation to be performed. In addition,the RPN formula is
  converted to a string formula.

  @param  ACell         Pointer to the cell
  @param  AFormula      Array of TsFormulaElements. The array can be created by
                        using "CreateRPNFormla".

  @see    TsNumberFormat
  @see    TsFormulaElements
  @see    CreateRPNFormula
}
procedure TsWorksheet.WriteRPNFormula(ACell: PCell; AFormula: TsRPNFormula);
begin
  if ACell = nil then
    exit;

  ACell^.ContentType := cctFormula;
  ACell^.FormulaValue := ConvertRPNFormulaToStringFormula(AFormula);

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Adds font specification to the formatting of a cell. Looks in the workbook's
  FontList and creates an new entry if the font is not used so far. Returns the
  index of the font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontName   Name of the font
  @param  AFontSize   Size of the font, in points
  @param  AFontStyle  Set with font style attributes
                      (don't use those of unit "graphics" !)
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFont(ARow, ACol: Cardinal; const AFontName: String;
  AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer;
begin
  Result := WriteFont(GetCell(ARow, ACol), AFontName, AFontSize, AFontStyle, AFontColor);
end;

{@@
  Adds font specification to the formatting of a cell. Looks in the workbook's
  FontList and creates an new entry if the font is not used so far. Returns the
  index of the font in the font list.

  @param  ACell       Pointer to the cell considered
  @param  AFontName   Name of the font
  @param  AFontSize   Size of the font, in points
  @param  AFontStyle  Set with font style attributes
                      (don't use those of unit "graphics" !)
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFont(ACell: PCell; const AFontName: String;
  AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer;
begin
  if ACell = nil then begin
    Result := -1;
    Exit;
  end;

  Include(ACell^.UsedFormattingFields, uffFont);
  Result := FWorkbook.FindFont(AFontName, AFontSize, AFontStyle, AFontColor);
  if Result = -1 then
    result := FWorkbook.AddFont(AFontName, AFontSize, AFontStyle, AFontColor);
  ACell^.FontIndex := Result;
  ChangedFont(ACell^.Row, ACell^.Col);
end;

{@@
  Applies a font to the formatting of a cell. The font is determined by its
  index in the workbook's font list:

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontIndex  Index of the font in the workbook's font list
  @return Pointer to the cell
}
function TsWorksheet.WriteFont(ARow, ACol: Cardinal; AFontIndex: Integer): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteFont(Result, AFontIndex);
end;

{@@
  Applies a font to the formatting of a cell. The font is determined by its
  index in the workbook's font list:

  @param  ACell       Pointer to the cell considered
  @param  AFontIndex  Index of the font in the workbook's font list
}
procedure TsWorksheet.WriteFont(ACell: PCell; AFontIndex: Integer);
begin
  if ACell = nil then
    exit;

  if (AFontIndex < 0) or (AFontIndex >= Workbook.GetFontCount) or (AFontIndex = 4) then
    // note: Font index 4 is not defined in BIFF
    raise Exception.Create(lpInvalidFontIndex);

  Include(ACell^.UsedFormattingFields, uffFont);
  ACell^.FontIndex := AFontIndex;
  ChangedFont(ACell^.Row, ACell^.Col);
end;

{@@
  Replaces the text color used in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontColor  Index into the workbook's color palette identifying the
                      new text color.
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFontColor(ARow, ACol: Cardinal; AFontColor: TsColor): Integer;
begin
  Result := WriteFontColor(GetCell(ARow, ACol), AFontColor);
end;

{@@
  Replaces the text color used in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell
  @param  AFontColor  Index into the workbook's color palette identifying the
                      new text color.
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFontColor(ACell: PCell; AFontColor: TsColor): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := Workbook.GetFont(ACell^.FontIndex);
  Result := WriteFont(ACell, fnt.FontName, fnt.Size, fnt.Style, AFontColor);
end;

{@@
  Replaces the font used in formatting of a cell considering only the font face
  and leaving font size, style and color unchanged. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontName   Name of the new font to be used
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFontName(ARow, ACol: Cardinal; AFontName: String): Integer;
begin
  result := WriteFontName(GetCell(ARow, ACol), AFontName);
end;

{@@
  Replaces the font used in formatting of a cell considering only the font face
  and leaving font size, style and color unchanged. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell
  @param  AFontName   Name of the new font to be used
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFontName(ACell: PCell; AFontName: String): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := Workbook.GetFont(ACell^.FontIndex);
  result := WriteFont(ACell, AFontName, fnt.Size, fnt.Style, fnt.Color);
end;

{@@
  Replaces the font size in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  ASize       Size of the font to be used (in points).
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFontSize(ARow, ACol: Cardinal; ASize: Single): Integer;
begin
  Result := WriteFontSize(GetCell(ARow, ACol), ASize);
end;

{@@
  Replaces the font size in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell
  @param  ASize       Size of the font to be used (in points).
  @return Index of the font in the workbook's font list.
}
function TsWorksheet.WriteFontSize(ACell: PCell; ASize: Single): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := Workbook.GetFont(ACell^.FontIndex);
  Result := WriteFont(ACell, fnt.FontName, ASize, fnt.Style, fnt.Color);
end;

{@@
  Replaces the font style (bold, italic, etc) in formatting of a cell.
  Looks in the workbook's font list if this modified font has already been used.
  If not a new font entry is created.
  Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AStyle      New font style to be used
  @return Index of the font in the workbook's font list.

  @see TsFontStyle
}
function TsWorksheet.WriteFontStyle(ARow, ACol: Cardinal;
  AStyle: TsFontStyles): Integer;
begin
  Result := WriteFontStyle(GetCell(ARow, ACol), AStyle);
end;

{@@
  Replaces the font style (bold, italic, etc) in formatting of a cell.
  Looks in the workbook's font list if this modified font has already been used.
  If not a new font entry is created.
  Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell considered
  @param  AStyle      New font style to be used
  @return Index of the font in the workbook's font list.

  @see TsFontStyle
}
function TsWorksheet.WriteFontStyle(ACell: PCell; AStyle: TsFontStyles): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := -1;
    exit;
  end;

  fnt := Workbook.GetFont(ACell^.FontIndex);
  Result := WriteFont(ACell, fnt.FontName, fnt.Size, AStyle, fnt.Color);
end;

{@@
  Adds text rotation to the formatting of a cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  ARotation How to rotate the text
  @return Pointer to cell

  @see    TsTextRotation
}
function TsWorksheet.WriteTextRotation(ARow, ACol: Cardinal;
  ARotation: TsTextRotation): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteTextRotation(Result, ARotation);
end;

{@@
  Adds text rotation to the formatting of a cell

  @param  ACell      Pointer to the cell
  @param  ARotation  How to rotate the text

  @see    TsTextRotation
}
procedure TsWorksheet.WriteTextRotation(ACell: PCell; ARotation: TsTextRotation);
begin
  if ACell = nil then
    exit;
  Include(ACell^.UsedFormattingFields, uffTextRotation);
  ACell^.TextRotation := ARotation;
  ChangedFont(ACell^.Row, ACell^.Col);
end;

{@@
  Directly modifies the used formatting fields of a cell.
  Only formatting corresponding to items included in this set is executed.

  @param  ARow            The row of the cell
  @param  ACol            The column of the cell
  @param  AUsedFormatting set of the used formatting fields

  @see    TsUsedFormattingFields
  @see    TCell
}
procedure TsWorksheet.WriteUsedFormatting(ARow, ACol: Cardinal;
  AUsedFormatting: TsUsedFormattingFields);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.UsedFormattingFields := AUsedFormatting;
  ChangedCell(ARow, ACol);
end;

{@@
  Sets the color of a background color of a cell.

  @param  ARow       Row index of the cell
  @param  ACol       Column index of the cell
  @param  AColor     Index of the new background color into the workbook's
                     color palette. Use the color index scTransparent to
                     erase an existing background color.
  @return Pointer to cell
}
function TsWorksheet.WriteBackgroundColor(ARow, ACol: Cardinal;
  AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBackgroundColor(Result, AColor);
end;

{@@
  Sets the color of a background color of a cell.

  @param  ACell      Pointer to cell
  @param  AColor     Index of the new background color into the workbook's
                     color palette. Use the color index scTransparent to
                     erase an existing background color.
}
procedure TsWorksheet.WriteBackgroundColor(ACell: PCell; AColor: TsColor);
begin
  if ACell <> nil then begin
    if AColor = scTransparent then
      Exclude(ACell^.UsedFormattingFields, uffBackgroundColor)
    else begin
      Include(ACell^.UsedFormattingFields, uffBackgroundColor);
      ACell^.BackgroundColor := AColor;
    end;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Sets the color of a cell border line.
  Note: the border must be included in Borders set in order to be shown!

  @param  ARow       Row index of the cell
  @param  ACol       Column index of the cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  AColor     Index of the new border color into the workbook's
                     color palette.
  @return Pointer to cell
  }
function TsWorksheet.WriteBorderColor(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderColor(Result, ABorder, AColor);
end;

{@@
  Sets the color of a cell border line.
  Note: the border must be included in Borders set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  AColor     Index of the new border color into the workbook's
                     color palette.
  }
procedure TsWorksheet.WriteBorderColor(ACell: PCell; ABorder: TsCellBorder;
  AColor: TsColor);
begin
  if ACell <> nil then begin
    ACell^.BorderStyles[ABorder].Color := AColor;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Sets the linestyle of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ARow       Row index of the cell
  @param  ACol       Column index of the cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  ALineStyle Identifier of the new line style to be applied.
  @return Pointer to cell

  @see    TsLineStyle
}
function TsWorksheet.WriteBorderLineStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderLineStyle(Result, ABorder, ALineStyle);
end;

{@@
  Sets the linestyle of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  ALineStyle Identifier of the new line style to be applied.

  @see    TsLineStyle
}
procedure TsWorksheet.WriteBorderLineStyle(ACell: PCell;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle);
begin
  if ACell <> nil then begin
    ACell^.BorderStyles[ABorder].LineStyle := ALineStyle;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Shows the cell borders included in the set ABorders. No border lines are drawn
  for those not included.

  The borders are drawn using the "BorderStyles" assigned to the cell.

  @param  ARow      Row index of the cell
  @param  ACol      Column index of the cell
  @param  ABorders  Set with elements to identify the border(s) to will be shown
  @return Pointer to cell
  @see    TsCellBorder
}
function TsWorksheet.WriteBorders(ARow, ACol: Cardinal; ABorders: TsCellBorders): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorders(Result, ABorders);
end;

{@@
  Shows the cell borders included in the set ABorders. No border lines are drawn
  for those not included.

  The borders are drawn using the "BorderStyles" assigned to the cell.

  @param  ACell     Pointer to cell
  @param  ABorders  Set with elements to identify the border(s) to will be shown
  @see    TsCellBorder
}
procedure TsWorksheet.WriteBorders(ACell: PCell; ABorders: TsCellBorders);
begin
  if ACell <> nil then begin
    Include(ACell^.UsedFormattingFields, uffBorder);
    ACell^.Border := ABorders;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Sets the style of a cell border, i.e. line style and line color.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ARow       Row index of the cell considered
  @param  ACol       Column index of the cell considered
  @param  ABorder    Identifies the border to be modified (left/top/right/bottom)
  @param  AStyle     record of parameters controlling how the border line is drawn
                     (line style, line color)
  @result Pointer to cell
}
function TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AStyle: TsCellBorderStyle): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyle(Result, ABorder, AStyle);
end;

{@@
  Sets the style of a cell border, i.e. line style and line color.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Identifies the border to be modified (left/top/right/bottom)
  @param  AStyle     record of parameters controlling how the border line is drawn
                     (line style, line color)
}
procedure TsWorksheet.WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
  AStyle: TsCellBorderStyle);
begin
  if ACell <> nil then begin
    ACell^.BorderStyles[ABorder] := AStyle;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Sets line style and line color of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ARow       Row index of the considered cell
  @param  ACol       Column index of the considered cell
  @param  ABorder    Identifier of the border to be modified
  @param  ALineStyle Identifier for the new line style of the border
  @param  AColor     Palette index for the color of the border line
  @return Pointer to cell

  @see WriteBorderStyles
}
function TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLinestyle; AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyle(Result, ABorder, ALineStyle, AColor);
end;

{@@
  Sets line style and line color of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Identifier of the border to be modified
  @param  ALineStyle Identifier for the new line style of the border
  @param  AColor     Palette index for the color of the border line

  @see WriteBorderStyles
}
procedure TsWorksheet.WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
  ALineStyle: TsLinestyle; AColor: TsColor);
begin
  if ACell <> nil then begin
    ACell^.BorderStyles[ABorder].LineStyle := ALineStyle;
    ACell^.BorderStyles[ABorder].Color := AColor;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Sets the style of all cell border of a cell, i.e. line style and line color.
  Note: Only those borders included in the "Borders" set are shown!

  @param  ARow    Row index of the considered cell
  @param  ACol    Column index of the considered cell
  @param  AStyles Array of CellBorderStyles for each cell border.
  @return Pointer to cell

  @see WriteBorderStyle
}
function TsWorksheet.WriteBorderStyles(ARow, ACol: Cardinal;
  const AStyles: TsCellBorderStyles): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyles(Result, AStyles);
end;

{@@
  Sets the style of all cell border of a cell, i.e. line style and line color.
  Note: Only those borders included in the "Borders" set are shown!

  @param  ACell   Pointer to cell
  @param  ACol    Column index of the considered cell
  @param  AStyles Array of CellBorderStyles for each cell border.

  @see WriteBorderStyle
}
procedure TsWorksheet.WriteBorderStyles(ACell: PCell;
  const AStyles: TsCellBorderStyles);
var
  b: TsCellBorder;
begin
  if Assigned(ACell) then begin
    for b in TsCellBorder do ACell^.BorderStyles[b] := AStyles[b];
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Defines the horizontal alignment of text in a cell.

  @param ARow    Row index of the cell considered
  @param ACol    Column index of the cell considered
  @param AValue  Parameter for horizontal text alignment (haDefault, vaLeft, haCenter, haRight)
                 By default, texts are left-aligned, numbers and dates are right-aligned.
  @return Pointer to cell
}
function TsWorksheet.WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteHorAlignment(Result, AValue);
end;

{@@
  Defines the horizontal alignment of text in a cell.

  @param ACell   Pointer to the cell considered
  @param AValue  Parameter for horizontal text alignment (haDefault, vaLeft, haCenter, haRight)
                 By default, texts are left-aligned, numbers and dates are right-aligned.
}
procedure TsWorksheet.WriteHorAlignment(ACell: PCell; AValue: TsHorAlignment);
begin
  if ACell = nil then
    exit;
  Include(ACell^.UsedFormattingFields, uffHorAlign);
  ACell^.HorAlignment := AValue;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Defines the vertical alignment of text in a cell.

  @param ARow    Row index of the cell considered
  @param ACol    Column index of the cell considered
  @param AValue  Parameter for vertical text alignment (vaDefault, vaTop, vaCenter, vaBottom)
                 By default, texts are bottom-aligned.
  @return Pointer to cell
}
function TsWorksheet.WriteVertAlignment(ARow, ACol: Cardinal; AValue: TsVertAlignment): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteVertAlignment(Result, AValue);
end;

{@@
  Defines the vertical alignment of text in a cell.

  @param ACell   Poiner to the cell considered
  @param AValue  Parameter for vertical text alignment (vaDefault, vaTop, vaCenter, vaBottom)
                 By default, texts are bottom-aligned.
}
procedure TsWorksheet.WriteVertAlignment(ACell: PCell; AValue: TsVertAlignment);
begin
  if ACell = nil then
    exit;
  Include(ACell^.UsedFormattingFields, uffVertAlign);
  ACell^.VertAlignment := AValue;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@
  Enables or disables the word-wrapping feature for a cell.

  @param  ARow    Row index of the cell considered
  @param  ACol    Column index of the cell considered
  @param  AValue  true = word-wrapping enabled, false = disabled.
  @return Pointer to cell
}
function TsWorksheet.WriteWordwrap(ARow, ACol: Cardinal; AValue: boolean): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteWordWrap(Result, AValue);
end;

{@@
  Enables or disables the word-wrapping feature for a cell.

  @param ACel    Pointer to the cell considered
  @param AValue  true = word-wrapping enabled, false = disabled.
}
procedure TsWorksheet.WriteWordwrap(ACell: PCell; AValue: boolean);
begin
  if ACell = nil then
    exit;
  if AValue then
    Include(ACell^.UsedFormattingFields, uffWordwrap)
  else
    Exclude(ACell^.UsedFormattingFields, uffWordwrap);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

function TsWorksheet.GetFormatSettings: TFormatSettings;
begin
  Result := FWorkbook.FormatSettings;
end;

{@@
  Calculates the optimum height of a given row. Depends on the font size
  of the individual cells in the row.

  @param  ARow   Index of the row to be considered
  @return Row height in line count of the default font.
}
function TsWorksheet.CalcAutoRowHeight(ARow: Cardinal): Single;
var
  cell: PCell;
  col: Integer;
  h0: Single;
begin
  Result := 0;
  h0 := Workbook.GetDefaultFontSize;
  for col := GetFirstColIndex to GetLastColIndex do begin
    cell := FindCell(ARow, col);
    if cell <> nil then
      Result := Max(Result, Workbook.GetFont(cell^.FontIndex).Size / h0);
  end;
end;

{@@
 Checks if a row record exists for the given row index and returns a pointer
 to the row record, or nil if not found

 @param  ARow   Index of the row looked for
 @return        Pointer to the row record with this row index, or nil if not found }
function TsWorksheet.FindRow(ARow: Cardinal): PRow;
var
  LElement: TRow;
  AVLNode: TAVGLVLTreeNode;
begin
  Result := nil;
  LElement.Row := ARow;
  AVLNode := FRows.Find(@LElement);
  if Assigned(AVLNode) then
    result := PRow(AVLNode.Data);
end;

{@@
 Checks if a column record exists for the given column index and returns a pointer
 to the TCol record, or nil if not found

 @param  ACol   Index of the column looked for
 @return        Pointer to the column record with this column index, or nil if not found }
function TsWorksheet.FindCol(ACol: Cardinal): PCol;
var
  LElement: TCol;
  AVLNode: TAVGLVLTreeNode;
begin
  Result := nil;
  LElement.Col := ACol;
  AVLNode := FCols.Find(@LElement);
  if Assigned(AVLNode) then
    result := PCol(AVLNode.Data);
end;

{@@
 Checks if a row record exists for the given row index and creates it if not found.

 @param  ARow   Index of the row looked for
 @return        Pointer to the row record with this row index. It can safely be
                assumed that this row record exists. }
function TsWorksheet.GetRow(ARow: Cardinal): PRow;
begin
  Result := FindRow(ARow);
  if (Result = nil) then begin
    Result := GetMem(SizeOf(TRow));
    FillChar(Result^, SizeOf(TRow), #0);
    Result^.Row := ARow;
    FRows.Add(Result);
    if FLastRowIndex = 0 then
      FLastRowIndex := GetLastRowIndex(true)
    else
      FLastRowIndex := Max(FLastRowIndex, ARow);
  end;
end;

{@@
 Checks if a column record exists for the given column index and creates it
 if not found.

 @param  ACol   Index of the column looked for
 @return        Pointer to the TCol record with this column index. It can safely be
                assumed that this column record exists. }
function TsWorksheet.GetCol(ACol: Cardinal): PCol;
begin
  Result := FindCol(ACol);
  if (Result = nil) then begin
    Result := GetMem(SizeOf(TCol));
    FillChar(Result^, SizeOf(TCol), #0);
    Result^.Col := ACol;
    FCols.Add(Result);
    if FFirstColIndex = 0
      then FFirstColIndex := GetFirstColIndex(true)
      else FFirstColIndex := Min(FFirstColIndex, ACol);
    if FLastColIndex = 0
      then FLastColIndex := GetLastColIndex(true)
      else FLastColIndex := Max(FLastColIndex, ACol);
  end;
end;

{@@
  Counts how many cells exist in the given column. Blank cells do contribute
  to the sum, as well as formatted cells.

  @param  ACol  Index of the column considered
  @return Count of cells with value or format in this column }
function TsWorksheet.GetCellCountInCol(ACol: Cardinal): Cardinal;
var
  cell: PCell;
  r: Cardinal;
  row: PRow;
begin
  Result := 0;
  for r := 0 to GetLastRowIndex do begin
    cell := FindCell(r, ACol);
    if cell <> nil then
      inc(Result)
    else begin
      row := FindRow(r);
      if row <> nil then inc(Result);
    end;
  end;
end;

{@@
  Counts how many cells exist in the given row. Blank cells do contribute
  to the sum, as well as formatted cell.s

  @param  ARow  Index of the row considered
  @return Count of cells with value or format in this row
}
function TsWorksheet.GetCellCountInRow(ARow: Cardinal): Cardinal;
var
  cell: PCell;
  c: Cardinal;
  col: PCol;
begin
  Result := 0;
  for c := 0 to GetLastColIndex do begin
    cell := FindCell(ARow, c);
    if cell <> nil then
      inc(Result)
    else begin
      col := FindCol(c);
      if col <> nil then inc(Result);
    end;
  end;
end;

{@@
  Returns the width of the given column. If there is no column record then
  the default column width is returned.

  @param  ACol  Index of the column considered
  @return Width of the column (in count of "0" characters of the default font)
}
function TsWorksheet.GetColWidth(ACol: Cardinal): Single;
var
  col: PCol;
begin
  col := FindCol(ACol);
  if col <> nil then
    Result := col^.Width
  else
    Result := FWorkbook.DefaultColWidth;
end;

{@@
  Returns the height of the given row. If there is no row record then the
  default row height is returned

  @param  ARow  Index of the row considered
  @return Height of the row (in line count of the default font).
}
function TsWorksheet.GetRowHeight(ARow: Cardinal): Single;
var
  row: PRow;
begin
  row := FindRow(ARow);
  if row <> nil then
    Result := row^.Height
  else
    //Result := CalcAutoRowHeight(ARow);
    Result := FWorkbook.DefaultRowHeight;
end;

procedure TsWorksheet.InsertColCallback(data, arg: Pointer);
var
  cell: PCell;
  col: PtrInt;
  fe: TsFormulaElement;
  formula: TsRPNFormula;
  i: Integer;
begin
  col := PtrInt(arg);
  cell := PCell(data);

  // Update column index of moved cells
  if cell^.Col >= col then
    inc(cell^.Col);

  // Update formulas
  // (1) create an rpn formula
  formula := BuildRPNFormula(cell);
  // (2) update cell addresses affected by the insertion of a column
  for i:=0 to Length(formula)-1 do begin
    fe := Formula[i];   // "fe" means "formula element"
    case fe.ElementKind of
      fekCell, fekCellRef:
        if fe.Col >= col then inc(fe.Col);
      fekCellRange:
        begin
          if fe.Col >= col then inc(fe.Col);
          if fe.Col2 >= col then inc(fe.Col2);
        end;
    end;
  end;
  // (3) convert rpn formula back to string formula
  cell^.FormulaValue := ConvertRPNFormulaToStringFormula(formula);
end;

{@@
  Inserts a column BEFORE the index specified. Cells with greater column indexes are
  moved one column up. Cell references in rpn formulas are considered as well.
  However, lacking a parser for string formulas, references in string formulas
  are not changed which may lead to incorrect operation!

  @param   ACol   Index of the column before which a new column is inserted.
}
procedure TsWorksheet.InsertCol(ACol: Cardinal);
var
  cellnode: TAVLTreeNode;
  col: PCol;
  i: Integer;
begin
  // Update column index of cell records
  cellnode := FCells.FindLowest;
  while Assigned(cellnode) do begin
    InsertColCallback(cellnode.Data, pointer(PtrInt(ACol)));
    cellnode := FCells.FindSuccessor(cellnode);
  end;

  // Update column index of column records
  for i:=0 to FCols.Count-1 do begin
    col := PCol(FCols.Items[i]);
    if col^.Col >= ACol then inc(col^.Col);
  end;

  // Update last column index
  inc(FLastColIndex);

  ChangedCell(0, ACol);
end;

procedure TsWorksheet.InsertRowCallback(data, arg: Pointer);
var
  cell: PCell;
  row: PtrInt;
  i: Integer;
  fe: TsFormulaElement;
  formula: TsRPNFormula;
begin
  row := PtrInt(arg);
  cell := PCell(data);

  // Update row index of moved cells
  if cell^.Row >= row then
    inc(cell^.Row);

  // Update rpn formulas
  // (1) create an rpn formula
  formula := BuildRPNFormula(cell);
  // (2) update cell addresses affected by the insertion of a column
  for i:=0 to Length(formula)-1 do begin
    fe := formula[i];   // "fe" means "formula element"
    case fe.ElementKind of
      fekCell, fekCellRef:
        if fe.Row >= row then inc(fe.Row);
      fekCellRange:
        begin
          if fe.Row >= row then inc(fe.Row);
          if fe.Row2 >= row then inc(fe.Row2);
        end;
    end;
  end;
  // (3) convert rpn formula back to string formula
  cell^.FormulaValue := ConvertRPNFormulaToStringFormula(formula);
end;

{@@
  Inserts a row BEFORE the row specified. Cells with greater row indexes are
  moved one row up. Cell references in rpn formulas are considered as well.
  However, lacking a parser for string formulas, references in string formulas
  are not changed which may lead to incorrect operation!

  @param   ARow   Index of the row before which a new row is inserted.
}
procedure TsWorksheet.InsertRow(ARow: Cardinal);
var
  row: PRow;
  cellnode: TAVLTreeNode;
  i: Integer;
begin
  // Update row index of cell records
  cellnode := FCells.FindLowest;
  while Assigned(cellnode) do begin
    InsertRowCallback(cellnode.Data, pointer(PtrInt(ARow)));
    cellnode := FCells.FindSuccessor(cellnode);
  end;

  // Update row index of row records
  for i:=0 to FRows.Count-1 do begin
    row := PRow(FRows.Items[i]);
    if row^.Row >= ARow then inc(row^.Row);
  end;

  // Update last row index
  inc(FLastRowIndex);

  ChangedCell(ARow, 0);
end;

{@@
  Removes all row records from the worksheet and frees the occupied memory.
  Note: Cells are retained.
}
procedure TsWorksheet.RemoveAllRows;
var
  Node: Pointer;
  i: Integer;
begin
  for i := FRows.Count-1 downto 0 do begin
    Node := FRows.Items[i];
    FreeMem(Node, SizeOf(TRow));
  end;
  FRows.Clear;
end;

{@@
  Removes all column records from the worksheet and frees the occupied memory.
  Note: Cells are retained.
}
procedure TsWorksheet.RemoveAllCols;
var
  Node: Pointer;
  i: Integer;
begin
  for i := FCols.Count-1 downto 0 do begin
    Node := FCols.Items[i];
    FreeMem(Node, SizeOf(TCol));
  end;
  FCols.Clear;
end;

{@@
  Writes a row record for the row at a given index to the spreadsheet.
  Currently the row record contains only the row height (and the row index, of course).

  Creates a new row record if it does not yet exist.

  @param  ARow   Index of the row record which will be created or modified
  @param  AData  Data to be written.
}
procedure TsWorksheet.WriteRowInfo(ARow: Cardinal; AData: TRow);
var
  AElement: PRow;
begin
  AElement := GetRow(ARow);
  AElement^.Height := AData.Height;
end;

{@@
  Sets the row height for a given row. Creates a new row record if it
  does not yet exist.

  @param  ARow     Index of the row to be considered
  @param  AHeight  Row height to be assigned to the row. The row height is
                   expressed as the line count of the default font size.
}
procedure TsWorksheet.WriteRowHeight(ARow: Cardinal; AHeight: Single);
var
  AElement: PRow;
begin
  AElement := GetRow(ARow);
  AElement^.Height := AHeight;
end;

{@@
  Writes a column record for the column at a given index to the spreadsheet.
  Currently the column record contains only the column width (and the column
  index, of course).

  Creates a new column record if it does not yet exist.

  @param  ACol   Index of the column record which will be created or modified
  @param  AData  Data to be written (essentially column width).
}
procedure TsWorksheet.WriteColInfo(ACol: Cardinal; AData: TCol);
var
  AElement: PCol;
begin
  AElement := GetCol(ACol);
  AElement^.Width := AData.Width;
end;

{@@
  Sets the column width for a given column. Creates a new column record if it
  does not yet exist.

  @param  ACol     Index of the column to be considered
  @param  AWidth   Width to be assigned to the column. The column width is
                   expressed as the count of "0" characters of the default font.
}
procedure TsWorksheet.WriteColWidth(ACol: Cardinal; AWidth: Single);
var
  AElement: PCol;
begin
  AElement := GetCol(ACol);
  AElement^.Width := AWidth;
end;


{ TsWorkbook }

{@@
  Helper method called before reading the workbook. Clears the error log.
}
procedure TsWorkbook.PrepareBeforeReading;
begin
  // Clear error log
  FLog.Clear;
end;

{@@
  Helper method called before saving the workbook. Clears the error log, and
  calculates the formulas in all worksheets if workbook option soCalcBeforeSaving
  is set.
}
procedure TsWorkbook.PrepareBeforeSaving;
var
  sheet: TsWorksheet;
begin
  // Clear error log
  FLog.Clear;

  // Updates fist/last column/row index
  UpdateCaches;

  // Calculated formulas (if requested)
  if (boCalcBeforeSaving in FOptions) then
    for sheet in FWorksheets do
      sheet.CalcFormulas;
end;

{@@
  Recalculates rpn formulas in all worksheets
}
procedure TsWorkbook.Recalc;
var
  sheet: TsWorksheet;
begin
  for sheet in FWorksheets do
    sheet.CalcFormulas;
end;

{@@
  Helper method for clearing the spreadsheet list.
}
procedure TsWorkbook.RemoveWorksheetsCallback(data, arg: pointer);
begin
  Unused(arg);
  TsWorksheet(data).Free;
end;

{@@
  Helper method to update internal caching variables
}
procedure TsWorkbook.UpdateCaches;
var
  sheet: TsWorksheet;
begin
  for sheet in FWorksheets do
    sheet.UpdateCaches;
end;

{@@
  Constructor of the workbook class. Among others, it initializes the built-in
  fonts, defines the default font, and sets up the FormatSettings for localization
  of some number formats.
}
constructor TsWorkbook.Create;
begin
  inherited Create;
  FWorksheets := TFPList.Create;
  FLog := TStringList.Create;
  FFormat := sfExcel8;
  FDefaultColWidth := 12;
  FDefaultRowHeight := 1;
  FormatSettings := DefaultFormatSettings;
  FormatSettings.ShortDateFormat := MakeShortDateFormat(FormatSettings.ShortDateFormat);
  FormatSettings.LongDateFormat := MakeLongDateFormat(FormatSettings.ShortDateFormat);
  FFontList := TFPList.Create;
  SetDefaultFont('Arial', 10.0);
  InitFonts;
end;

{@@
  Destructor of the workbook class
}
destructor TsWorkbook.Destroy;
begin
  RemoveAllWorksheets;
  RemoveAllFonts;

  FWorksheets.Free;
  FFontList.Free;

  FLog.Free;
  inherited Destroy;
end;


{@@
  Helper method for determining the spreadsheet type from the file type extension

  @param   AFileName   Name of the file to be considered
  @param   SheetType   File format found from analysis of the extension (output)
  @return  True if the file matches any of the known formats, false otherwise
}
class function TsWorkbook.GetFormatFromFileName(const AFileName: TFileName;
  out SheetType: TsSpreadsheetFormat): Boolean;
var
  suffix: String;
begin
  Result := True;
  suffix := Lowercase(ExtractFileExt(AFileName));
  if suffix = STR_EXCEL_EXTENSION then SheetType := sfExcel8
  else if suffix = STR_OOXML_EXCEL_EXTENSION then SheetType := sfOOXML
  else if suffix = STR_OPENDOCUMENT_CALC_EXTENSION then SheetType := sfOpenDocument
  else if suffix = STR_COMMA_SEPARATED_EXTENSION then SheetType := sfCSV
  else if suffix = STR_WIKITABLE_PIPES then SheetType := sfWikiTable_Pipes
  else if suffix = STR_WIKITABLE_WIKIMEDIA then SheetType := sfWikiTable_WikiMedia
  else Result := False;
end;

{@@
  Convenience method which creates the correct reader object for a given
  spreadsheet format.

  @param  AFormat  File format which is assumed when reading a document into
                   to workbook. An exception is raised when the document has
                   a different format.

  @return An instance of a TsCustomSpreadReader descendent which is able to
          read thi given file format.
}
function TsWorkbook.CreateSpreadReader(AFormat: TsSpreadsheetFormat): TsCustomSpreadReader;
var
  i: Integer;
begin
  Result := nil;

  for i := 0 to Length(GsSpreadFormats) - 1 do
    if GsSpreadFormats[i].Format = AFormat then
    begin
      Result := GsSpreadFormats[i].ReaderClass.Create(self);
      Break;
    end;

  if Result = nil then raise Exception.Create(lpUnsupportedReadFormat);
end;

{@@
  Convenience method which creates the correct writer object for a given
  spreadsheet format.

  @param  AFormat  File format to be used for writing the workbook

  @return An instance of a TsCustomSpreadWriter descendent which is able to
          write the given file format.
}
function TsWorkbook.CreateSpreadWriter(AFormat: TsSpreadsheetFormat): TsCustomSpreadWriter;
var
  i: Integer;
begin
  Result := nil;

  for i := 0 to Length(GsSpreadFormats) - 1 do
    if GsSpreadFormats[i].Format = AFormat then
    begin
      Result := GsSpreadFormats[i].WriterClass.Create(self);
      Break;
    end;
    
  if Result = nil then raise Exception.Create(lpUnsupportedWriteFormat);
end;

{@@
  Determines the maximum index of used columns and rows in all sheets of this
  workbook. Respects VirtualMode.
  Is needed to disable saving when limitations of the format is exceeded. }
procedure TsWorkbook.GetLastRowColIndex(out ALastRow, ALastCol: Cardinal);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  if (boVirtualMode in Options) then begin
    ALastRow := FVirtualRowCount - 1;
    ALastCol := FVirtualColCount - 1;
  end else begin
    ALastRow := 0;
    ALastCol := 0;
    for i:=0 to GetWorksheetCount-1 do begin
      sheet := GetWorksheetByIndex(i);
      ALastRow := Max(ALastRow, sheet.GetLastRowIndex);
      ALastCol := Max(ALastCol, sheet.GetLastColIndex);
    end;
  end;
end;

{@@
  Reads the document from a file. It is assumed to have a given file format.

  @param  AFileName  Name of the file to be read
  @param  AFormat    File format assumed
}
procedure TsWorkbook.ReadFromFile(AFileName: string;
  AFormat: TsSpreadsheetFormat);
var
  AReader: TsCustomSpreadReader;
begin
  AReader := CreateSpreadReader(AFormat);
  try
    FFileName := AFileName;
    PrepareBeforeReading;
    AReader.ReadFromFile(AFileName, Self);
    UpdateCaches;
    if (boAutoCalc in Options) then
      Recalc;
    FFormat := AFormat;
  finally
    AReader.Free;
  end;
end;

{@@
  Reads the document from a file. This method will try to guess the format from
  the extension. In the case of the ambiguous xls extension, it will simply
  assume that it is BIFF8. Note that it could be BIFF2 or 5 as well.
}
procedure TsWorkbook.ReadFromFile(AFileName: string); overload;
var
  SheetType: TsSpreadsheetFormat;
  valid: Boolean;
  lException: Exception = nil;
begin
  valid := GetFormatFromFileName(AFileName, SheetType);
  if valid then
  begin
    if SheetType = sfExcel8 then
    begin
      while True do
      begin
        try
          ReadFromFile(AFileName, SheetType);
          valid := True;
        except
          on E: Exception do
          begin
            if SheetType = sfExcel8 then lException := E;
            valid := False
          end;
        end;
        if valid or (SheetType = sfExcel2) then Break;
        SheetType := Pred(SheetType);
      end;

      // A failed attempt to read a file should bring an exception, so re-raise
      // the exception if necessary. We re-raise the exception brought by Excel 8,
      // since this is the most common format
      if (not valid) and (lException <> nil) then raise lException;
    end
    else
      ReadFromFile(AFileName, SheetType);
  end else
    raise Exception.CreateFmt(lpNoValidSpreadsheetFile, [AFileName]);
end;

{@@
  Reads the document from a file, but ignores the extension.
}
procedure TsWorkbook.ReadFromFileIgnoringExtension(AFileName: string);
var
  SheetType: TsSpreadsheetFormat;
  lException: Exception;
begin
  lException := pointer(1);  // Must not be nil initially
  SheetType := sfExcel8;
  while (SheetType in [sfExcel2..sfExcel8, sfOpenDocument, sfOOXML]) and (lException <> nil) do
  begin
    try
      Dec(SheetType);
      ReadFromFile(AFileName, SheetType);
      lException := nil;
    except
      on E: Exception do { do nothing } ;
    end;
    if lException = nil then Break;
  end;
end;

{@@
  Reads the document from a seekable stream.

  @param  AStream  Stream being read
  @param  AFormat  File format assumed.
}
procedure TsWorkbook.ReadFromStream(AStream: TStream;
  AFormat: TsSpreadsheetFormat);
var
  AReader: TsCustomSpreadReader;
begin
  AReader := CreateSpreadReader(AFormat);
  try
    PrepareBeforeReading;
    AReader.ReadFromStream(AStream, Self);
    UpdateCaches;
    if (boAutoCalc in Options) then
      Recalc;
  finally
    AReader.Free;
  end;
end;

procedure TsWorkbook.SetVirtualColCount(AValue: Cardinal);
begin
  if FWriting then exit;
  FVirtualColCount := AValue;
end;

procedure TsWorkbook.SetVirtualRowCount(AValue: Cardinal);
begin
  if FWriting then exit;
  FVirtualRowCount := AValue;
end;

{@@
  Writes the document to a file. If the file doesn't exist, it will be created.

  @param  AFileName  Name of the file to be written
  @param  AFormat    The file will be written in this file format
  @param  AOverwriteExisting  If the file is already existing it will be
                     overwritten in case of AOverwriteExisting = true.
                     If false an exception will be raised.
}
procedure TsWorkbook.WriteToFile(const AFileName: string;
 const AFormat: TsSpreadsheetFormat; const AOverwriteExisting: Boolean = False);
var
  AWriter: TsCustomSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);
  try
    FFileName := AFileName;
    PrepareBeforeSaving;
    AWriter.CheckLimitations;
    FWriting := true;
    AWriter.WriteToFile(AFileName, AOverwriteExisting);
  finally
    FWriting := false;
    AWriter.Free;
  end;
end;

{@@
  Writes the document to file based on the extension.
  If this was an earlier sfExcel type file, it will be upgraded to sfExcel8.

  @param  AFileName  Name of the destination file
  @param  AOverwriteExisting  If the file already exists it will be overwritten
                     of AOverwriteExisting is true. In case of false, an
                     exception will be raised.
}
procedure TsWorkbook.WriteToFile(const AFileName: String;
  const AOverwriteExisting: Boolean);
var
  SheetType: TsSpreadsheetFormat;
  valid: Boolean;
begin
  valid := GetFormatFromFileName(AFileName, SheetType);
  if valid then WriteToFile(AFileName, SheetType, AOverwriteExisting)
  else raise Exception.Create(Format(
    '[TsWorkbook.WriteToFile] Attempted to save a spreadsheet by extension, but the extension %s is invalid.', [ExtractFileExt(AFileName)]));
end;

{@@
  Writes the document to a stream

  @param  AStream   Instance of the stream being written to
  @param  AFormat   File format being written.
}
procedure TsWorkbook.WriteToStream(AStream: TStream; AFormat: TsSpreadsheetFormat);
var
  AWriter: TsCustomSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);
  try
    PrepareBeforeSaving;
    AWriter.CheckLimitations;
    FWriting := true;
    AWriter.WriteToStream(AStream);
  finally
    FWriting := false;
    AWriter.Free;
  end;
end;

{@@
  Adds a new worksheet to the workbook

  It is added to the end of the list of worksheets

  @param  AName     The name of the new worksheet
  @return The instance of the newly created worksheet
  @see    TsWorksheet
}
function TsWorkbook.AddWorksheet(AName: string): TsWorksheet;
begin
  Result := TsWorksheet.Create;

  Result.Name := AName;
  Result.FWorkbook := Self;

  FWorksheets.Add(Pointer(Result));
end;

{@@
  Quick helper routine which returns the first worksheet

  @return A TsWorksheet instance if at least one is present.
          nil otherwise.

  @see    TsWorkbook.GetWorksheetByIndex
  @see    TsWorkbook.GetWorksheetByName
  @see    TsWorksheet
}
function TsWorkbook.GetFirstWorksheet: TsWorksheet;
begin
  Result := TsWorksheet(FWorksheets.First);
end;

{@@
  Gets the worksheet with a given index

  The index is zero-based, so the first worksheet
  added has index 0, the second 1, etc.

  @param  AIndex    The index of the worksheet (0-based)

  @return A TsWorksheet instance if one is present at that index.
          nil otherwise.

  @see    TsWorkbook.GetFirstWorksheet
  @see    TsWorkbook.GetWorksheetByName
  @see    TsWorksheet
}
function TsWorkbook.GetWorksheetByIndex(AIndex: Cardinal): TsWorksheet;
begin
  if (integer(AIndex) < FWorksheets.Count) and (integer(AIndex)>=0) then
    Result := TsWorksheet(FWorksheets.Items[AIndex])
  else
    Result := nil;
end;

{@@
  Gets the worksheet with a given worksheet name

  @param  AName    The name of the worksheet

  @return A TsWorksheet instance if one is found with that name,
          nil otherwise.

  @see    TsWorkbook.GetFirstWorksheet
  @see    TsWorkbook.GetWorksheetByIndex
  @see    TsWorksheet
}
function TsWorkbook.GetWorksheetByName(AName: String): TsWorksheet;
var
  i:integer;
begin
  Result := nil;
  for i:=0 to FWorksheets.Count-1 do
  begin
    if TsWorkSheet(FWorkSheets.Items[i]).Name=AName then
    begin
      Result := TsWorksheet(FWorksheets.Items[i]);
      exit;
    end;
  end;
end;

{@@
  The number of worksheets on the workbook

  @see    TsWorksheet
}
function TsWorkbook.GetWorksheetCount: Cardinal;
begin
  Result := FWorksheets.Count;
end;

{@@
  Clears the list of Worksheets and releases their memory.
}
procedure TsWorkbook.RemoveAllWorksheets;
begin
  FWorksheets.ForEachCall(RemoveWorksheetsCallback, nil);
end;
                            (*
{@@
  Sets the selected flag for the sheet with the given index.
  Excel requires one sheet to be selected, otherwise strange things happen when
  the file is loaded into Excel (cannot print, hanging instance of Excel - see
  bug 0026386).

  @param  AIndex  Index of the worksheet to be selected
}
procedure TsWorkbook.SelectWorksheet(AIndex: Integer);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  for i:=0 to FWorksheets.Count-1 do begin
    sheet := TsWorksheet(FWorksheets.Items[i]);
    if i = AIndex then
      sheet.Options := sheet.Options + [soSelected]
    else
      sheet.Options := sheet.Options - [soSelected];
  end;
end;
                              *)

{ Font handling }

{@@
  Adds a font to the font list. Returns the index in the font list.

  @param AFontName  Name of the font (like 'Arial')
  @param ASize      Size of the font in points
  @param AStyle     Style of the font, a combination of TsFontStyle elements
  @param AColor     Color of the font, given by its index into the workbook's palette.
  @return           Index of the font in the workbook's font list
}
function TsWorkbook.AddFont(const AFontName: String; ASize: Single;
  AStyle: TsFontStyles; AColor: TsColor): Integer;
var
  fnt: TsFont;
begin
  fnt := TsFont.Create;
  fnt.FontName := AFontName;
  fnt.Size := ASize;
  fnt.Style := AStyle;
  fnt.Color := AColor;
  Result := AddFont(fnt);
end;

{@@
  Adds a font to the font list. Returns the index in the font list.

  @param AFont      TsFont record containing all font parameters
  @return           Index of the font in the workbook's font list
}
function TsWorkbook.AddFont(const AFont: TsFont): Integer;
begin
  // Font index 4 does not exist in BIFF. Avoid that a real font gets this index.
  if FFontList.Count = 4 then
    FFontList.Add(nil);
  result := FFontList.Add(AFont);
end;

{@@
  Copies a font list to the workbook's font list

  @param   ASource   Font list to be copied
}
procedure TsWorkbook.CopyFontList(ASource: TFPList);
var
  fnt: TsFont;
  i: Integer;
begin
  RemoveAllFonts;
  for i:=0 to ASource.Count-1 do begin
    fnt := TsFont(ASource.Items[i]);
    AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
  end;
end;

{@@
  Deletes a font.
  Use with caution because this will screw up the font assignment to cells.
  The only legal reason to call this method is from a reader of a file format
  in which the missing font #4 of BIFF does exist.
}
procedure TsWorkbook.DeleteFont(AFontIndex: Integer);
var
  fnt: TsFont;
begin
  if AFontIndex < FFontList.Count then begin
    fnt := TsFont(FFontList.Items[AFontIndex]);
    if fnt <> nil then fnt.Free;
    FFontList.Delete(AFontIndex);
  end;
end;

{@@
  Checks whether the font with the given specification is already contained in
  the font list. Returns the index, or -1 if not found.

  @param AFontName  Name of the font (like 'Arial')
  @param ASize      Size of the font in points
  @param AStyle     Style of the font, a combination of TsFontStyle elements
  @param AColor     Color of the font, given by its index into the workbook's palette.
  @return           Index of the font in the font list, or -1 if not found.
}
function TsWorkbook.FindFont(const AFontName: String; ASize: Single;
  AStyle: TsFontStyles; AColor: TsColor): Integer;
var
  fnt: TsFont;
begin
  for Result := 0 to FFontList.Count-1 do begin
    fnt := TsFont(FFontList.Items[Result]);
    if (fnt <> nil) and
       SameText(AFontName, fnt.FontName) and
      (abs(ASize - fnt.Size) < 0.001) and   // careful when comparing floating point numbers
      (AStyle = fnt.Style) and
      (AColor = fnt.Color)    // Take care of limited palette size!
    then
      exit;
  end;
  Result := -1;
end;

{@@
  Initializes the font list. In case of BIFF format, adds 5 fonts:

    0: default font
    1: like default font, but bold
    2: like default font, but italic
    3: like default font, but underlined
    4: empty (due to a restriction of Excel)
    5: like default font, but bold and italic
}
procedure TsWorkbook.InitFonts;
var
  fntName: String;
  fntSize: Single;
begin
  // Memorize old default font
  with TsFont(FFontList.Items[0]) do begin
    fntName := FontName;
    fntSize := Size;
  end;

  // Remove current font list
  RemoveAllFonts;

  // Build new font list
  SetDefaultFont(fntName, fntSize);                      // Default font (FONT0)
  AddFont(fntName, fntSize, [fssBold], scBlack);         // FONT1 for uffBold

  AddFont(fntName, fntSize, [fssItalic], scBlack);       // FONT2 (Italic)
  AddFont(fntName, fntSize, [fssUnderline], scBlack);    // FONT3 (fUnderline)
  // FONT4 which does not exist in BIFF is added automatically with nil as place-holder
  AddFont(fntName, fntSize, [fssBold, fssItalic], scBlack); // FONT5 (bold & italic)

  FBuiltinFontCount := FFontList.Count;
end;

{@@
  Clears the list of fonts and releases their memory.
}
procedure TsWorkbook.RemoveAllFonts;
var
  i: Integer;
  fnt: TsFont;
begin
  for i:=FFontList.Count-1 downto 0 do begin
    fnt := TsFont(FFontList.Items[i]);
    fnt.Free;
    FFontList.Delete(i);
  end;
  FBuiltinFontCount := 0;
end;

{@@
  Defines the default font. This is the font with index 0 in the FontList.
  The next built-in fonts will have the same font name and size
}
procedure TsWorkbook.SetDefaultFont(const AFontName: String; ASize: Single);
var
  i: Integer;
begin
  if FFontList.Count = 0 then
    AddFont(AFontName, ASize, [], scBlack)
  else
  for i:=0 to FBuiltinFontCount-1 do begin
    if (i <> 4) and (i < FFontList.Count) then
      with TsFont(FFontList[i]) do begin
        FontName := AFontName;
        Size := ASize;
      end;
  end;
end;

{@@
  Returns the default font. This is the first font (index 0) in the font list
}
function TsWorkbook.GetDefaultFont: TsFont;
begin
  Result := GetFont(0);
end;

{@@
  Returns the point size of the default font
}
function TsWorkbook.GetDefaultFontSize: Single;
begin
  Result := GetFont(0).Size;
end;

{@@
  Returns the font with the given index.

  @param  AIndex   Index of the font to be considered
  @return Record containing all parameters of the font (or nil if not found).
}
function TsWorkbook.GetFont(AIndex: Integer): TsFont;
begin
  if (AIndex >= 0) and (AIndex < FFontList.Count) then
    Result := FFontList.Items[AIndex]
  else
    Result := nil;
end;

{@@
  Returns a string which identifies the font with a given index.

  @param  AIndex    Index of the font
  @return String with font name, font size etc.
}
function TsWorkbook.GetFontAsString(AIndex: Integer): String;
var
  fnt: TsFont;
begin
  fnt := GetFont(AIndex);
  if fnt <> nil then begin
    Result := Format('%s, size %.1f, color %s', [
      fnt.FontName, fnt.Size, GetColorName(fnt.Color)]);
    if (fssBold in fnt.Style) then Result := Result + ', bold';
    if (fssItalic in fnt.Style) then Result := Result + ', italic';
    if (fssUnderline in fnt.Style) then Result := Result + ', underline';
    if (fssStrikeout in fnt.Style) then result := Result + ', strikeout';
  end else
    Result := '';
end;

{@@
  Returns the count of registered fonts
}
function TsWorkbook.GetFontCount: Integer;
begin
  Result := FFontList.Count;
end;

{@@
  Adds a color to the palette and returns its palette index, but only if the
  color does not already exist - in this case, it returns the index of the
  existing color entry.
  The color must in little-endian notation (like TColor of the graphics units)

  @param  AColorValue    Number containing the rgb code of the color to be added
  @return  Index of the new (or already existing) color item
}
function TsWorkbook.AddColorToPalette(AColorValue: TsColorValue): TsColor;
var
  n: Integer;
begin
  n := Length(FPalette);

  // Look look for the color. Is it already in the existing palette?
  if n > 0 then
    for Result := 0 to n-1 do
      if FPalette[Result] = AColorValue then
        exit;

  // No --> Add it to the palette.

  // Do not overwrite Excel's built-in system colors
  case n of
    DEF_FOREGROUND_COLOR:
      begin
        SetLength(FPalette, n+3);
        FPalette[n] := DEF_FOREGROUND_COLORVALUE;
        FPalette[n+1] := DEF_BACKGROUND_COLORVALUE;
        FPalette[n+2] := AColorValue;
      end;
    DEF_BACKGROUND_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_BACKGROUND_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    DEF_CHART_FOREGROUND_COLOR:
      begin
        SetLength(FPalette, n+4);
        FPalette[n] := DEF_CHART_FOREGROUND_COLORVALUE;
        FPalette[n+1] := DEF_CHART_BACKGROUND_COLORVALUE;
        FPalette[n+2] := DEF_CHART_NEUTRAL_COLORVALUE;
        FPalette[n+3] := AColorValue;
      end;
    DEF_CHART_BACKGROUND_COLOR:
      begin
        SetLength(FPalette, n+3);
        FPalette[n] := DEF_CHART_BACKGROUND_COLORVALUE;
        FPalette[n+1] := DEF_CHART_NEUTRAL_COLORVALUE;
        FPalette[n+2] := AColorValue;
      end;
    DEF_CHART_NEUTRAL_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_CHART_NEUTRAL_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    DEF_TOOLTIP_TEXT_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_TOOLTIP_TEXT_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    DEF_FONT_AUTOMATIC_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_FONT_AUTOMATIC_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    else
      begin
        SetLength(FPalette, n+1);
        FPalette[n] := AColorValue;
      end;
  end;
  Result := Length(FPalette) - 1;
end;

{@@
  Adds a (simple) error message to an internal list

  @param   AMsg   Error text to be stored in the list
}
procedure TsWorkbook.AddErrorMsg(const AMsg: String);
begin
  FLog.Add(AMsg);
end;

{@@
  Adds an error message composed by means of format codes to an internal list

  @param   AMsg   Error text to be stored in the list
  @param   Args   Array of arguments to be used by the Format() function
}
procedure TsWorkbook.AddErrorMsg(const AMsg: String; const Args: Array of const);
begin
  FLog.Add(Format(AMsg, Args));
end;

{@@
  Clears the internal list with error messages
}
procedure TsWorkbook.ClearErrorList;
begin
  FLog.Clear;
end;

{@@
  Getter to retrieve the error messages collected during reading/writing
}
function TsWorkbook.GetErrorMsg: String;
begin
  Result := FLog.Text;
end;

{@@
  Finds the palette color index which points to a color that is closest to a
  given color. "Close" means here smallest length of the rgb-difference vector.

  @param   AColorValue       Rgb color value to be considered
  @param   AMaxPaletteCount  Number of palette entries considered. Example:
                             BIFF5/BIFF8 can write only 64 colors, i.e
                             AMaxPaletteCount = 64
  @return  Palette index of the color closest to AColorValue
}
function TsWorkbook.FindClosestColor(AColorValue: TsColorValue;
  AMaxPaletteCount: Integer): TsColor;
type
  TRGBA = record r,g,b, a: Byte end;
var
  rgb: TRGBA;
  rgb0: TRGBA absolute AColorValue;
  dist: Double;
  minDist: Double;
  i: Integer;
  n: Integer;
begin
  Result := scNotDefined;
  minDist := 1E108;
  n := Min(Length(FPalette), AMaxPaletteCount);
  for i:=0 to n-1 do begin
    rgb := TRGBA(GetPaletteColor(i));
    dist := sqr(rgb.r - rgb0.r) + sqr(rgb.g - rgb0.g) + sqr(rgb.b - rgb0.b);
    if dist < minDist then begin
      Result := i;
      minDist := dist;
    end;
  end;
end;

{@@
  Converts a fpspreadsheet color into into a string RRGGBB.
  Note that colors are written to xls files as ABGR (where A is 0).
  if the color is scRGBColor the color value is taken from the argument
  ARGBColor, otherwise from the palette entry for the color index.
}
function TsWorkbook.FPSColorToHexString(AColor: TsColor;
  ARGBColor: TFPColor): string;
type
  TRgba = packed record Red, Green, Blue, A: Byte end;
var
  colorvalue: TsColorValue;
  r,g,b: Byte;
begin
  if AColor = scRGBColor then begin
    r := ARGBColor.Red div $100;
    g := ARGBColor.Green div $100;
    b := ARGBColor.Blue div $100;
  end else begin
    colorvalue := GetPaletteColor(AColor);
    r := TRgba(colorvalue).Red;
    g := TRgba(colorvalue).Green;
    b := TRgba(colorvalue).Blue;
  end;
  Result := Format('%.2x%.2x%.2x', [r, g, b]);
end;

{@@
  Returns the name of the color pointed to by the given color index.
  If the name is not known the hex string is returned as RRGGBB.

  @param   AColorIndex   Palette index of the color considered
  @return  String identifying the color (a color name or, if unknown, a string showing the rgb components
}
function TsWorkbook.GetColorName(AColorIndex: TsColor): string;
var
  i: Integer;
  c: TsColorValue;
begin
  // Get color rgb value
  c := GetPaletteColor(AColorIndex);

  // Find color value in default palette
  for i:=0 to High(DEFAULT_PALETTE) do
    if DEFAULT_PALETTE[i] = c then begin
      // if found: get the color name from the default color names array
      Result := DEFAULT_COLORNAMES[i];
      exit;
    end;

  // if not found: construct a string from rgb byte values.
  Result := FPSColorToHexString(AColorIndex, colBlack);
end;

{@@
  Reads the rgb color for the given index from the current palette. Can be
  type-cast to TColor for usage in GUI applications.

  @param  AColorIndex  Index of the color considered
  @return A number containing the rgb components in little-endian notation.
}
function TsWorkbook.GetPaletteColor(AColorIndex: TsColor): TsColorValue;
begin
  if (AColorIndex >= 0) and (AColorIndex < GetPaletteSize) then
  begin
    if ((FPalette = nil) or (Length(FPalette) = 0)) then
      Result := DEFAULT_PALETTE[AColorIndex]
    else
      Result := FPalette[AColorIndex];
  end
  else
    Result := $000000;  // "black" as default
  {

  case AColorIndex of
    $0040: Result := DEF_FOREGROUND_COLORVALUE;
    $0041: Result := DEF_BACKGROUND_COLORVALUE;
    $004D: Result := DEF_CHART_FOREGROUND_COLORVALUE;
    $004E: Result := DEF_CHART_BACKGROUND_COLORVALUE;
    $004F: Result := DEF_CHART_NEUTRAL_COLORVALUE;
    $0051: Result := DEF_TOOLTIP_TEXT_COLORVALUE;
    $7FFF: Result := DEF_FONT_AUTOMATIC_COLORVALUE;
    else   if (AColorIndex >= 0) and (AColorIndex < GetPaletteSize) then
           begin
             if ((FPalette = nil) or (Length(FPalette) = 0)) then
               Result := DEFAULT_PALETTE[AColorIndex]
             else
               Result := FPalette[AColorIndex];
           end
           else
             Result := $000000;  // "black" as default
  end;
  }
end;

{@@
  Converts the palette color of the given index to a string that can be used
  in HTML code. For ODS.

  @param  AColorIndex Index of the color considered
  @return A HTML-compatible string identifying the color. "Red", for example, is returned as '#FF0000';
}
function TsWorkbook.GetPaletteColorAsHTMLStr(AColorIndex: TsColor): String;
begin
  Result := ColorToHTMLColorStr(GetPaletteColor(AColorIndex));
end;

{@@
  Replaces a color value of the current palette by a new value. The color must
  be given as ABGR (little-endian), with A=0).

  @param  AColorIndex   Palette index of the color to be replaced
  @param  AColorValue   Number containing the rgb components of the new color
}
procedure TsWorkbook.SetPaletteColor(AColorIndex: TsColor; AColorValue: TsColorValue);
begin
  if (AColorIndex >= 0) and (AColorIndex < GetPaletteSize) then begin
    if ((FPalette = nil) or (Length(FPalette) = 0)) then
      DEFAULT_PALETTE[AColorIndex] := AColorValue
    else
      FPalette[AColorIndex] := AColorValue;
  end;
end;

{@@
  Returns the count of palette colors
}
function TsWorkbook.GetPaletteSize: Integer;
begin
  if (FPalette = nil) or (Length(FPalette) = 0) then
    Result := High(DEFAULT_PALETTE) + 1
  else
    Result := Length(FPalette);
end;

{@@
  Instructs the workbook to take colors from the default palette. Is called
  from ODS reader because ODS does not have a palette. Without a palette the
  color constants (scRed etc.) would not be correct any more.
}
procedure TsWorkbook.UseDefaultPalette;
begin
  UsePalette(@DEFAULT_PALETTE, Length(DEFAULT_PALETTE), false);
end;

{@@
  Instructs the Workbook to take colors from the palette pointed to by the parameter APalette
  This palette is only used for writing. When reading the palette found in the
  file is used.

  @param  APalette      Pointer to the array of TsColorValue numbers which will
                        become the new palette
  @param  APaletteCount Count of numbers in the source palette
  @param  ABigEnding    If true, indicates that the source palette is in
                        big-endian notation. The methods inverts the rgb
                        components to little-endian which is used by fpspreadsheet
                        internally.
}
procedure TsWorkbook.UsePalette(APalette: PsPalette; APaletteCount: Word;
  ABigEndian: Boolean);
var
  i: Integer;
begin
  if APaletteCount > 64 then
    raise Exception.Create('Due to Excel-compatibility, palettes cannot have more then 64 colors.');

 {$IFOPT R+}
  {$DEFINE RNGCHECK}
 {$ENDIF}
  SetLength(FPalette, APaletteCount);
  if ABigEndian then
    for i:=0 to APaletteCount-1 do
     {$IFDEF RNGCHECK}
      {$R-}
     {$ENDIF}
      FPalette[i] := LongRGBToExcelPhysical(APalette^[i])
     {$IFDEF RNGCHECK}
      {$R+}
     {$ENDIF}
  else
    for i:=0 to APaletteCount-1 do
     {$IFDEF RNGCHECK}
      {$R-}
     {$ENDIF}
      FPalette[i] := APalette^[i];
     {$IFDEF RNGCHECK}
      {$R+}
     {$ENDIF}
end;

{@@
  Checks whether a given color is used somewhere within the entire workbook

  @param  AColorIndex   Palette index of the color
  @result True if the color is used by at least one cell, false if not.
}
function TsWorkbook.UsesColor(AColorIndex: TsColor): Boolean;
var
  Node: TAVLTreeNode;
  sheet: TsWorksheet;
  cell: PCell;
  i: Integer;
  fnt: TsFont;
  b: TsCellBorder;
begin
  Result := true;
  for i:=0 to GetWorksheetCount-1 do begin
    sheet := GetWorksheetByIndex(i);
    Node := sheet.Cells.FindLowest;
    while Assigned(Node) do
    begin
      cell := PCell(Node.Data);
      if (uffBackgroundColor in cell^.UsedFormattingFields) then
        if cell^.BackgroundColor = AColorIndex then exit;
      if (uffBorder in cell^.UsedFormattingFields) then
        for b in TsCellBorders do
          if cell^.BorderStyles[b].Color = AColorIndex then
            exit;
      if (uffFont in cell^.UsedFormattingFields) then
      begin
        fnt := GetFont(cell^.FontIndex);
        if fnt.Color = AColorIndex then
          exit;
      end;
      Node := sheet.Cells.FindSuccessor(Node);
    end;
  end;
  Result := false;
end;

{ TsCustomNumFormatList }

{@@ Constructor of the number format list.
  @param AWorkbook The workbook is needed to get access to its "FormatSettings"
                   for localization of some formatting strings. }
constructor TsCustomNumFormatList.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  AddBuiltinFormats;
end;

{@@ Destructor of the number format list: clears the list and destroys the
    format items }
destructor TsCustomNumFormatList.Destroy;
begin
  Clear;
  inherited Destroy;
end;

{@@ Adds a number format described by the Excel format index, the ODF format
    name, the format string, and the built-in format identifier to the list
    and returns the index of the new item.
  @param AFormatIndex  Format index to be used by Excel
  @param AFormatName   Format name to be used by OpenDocument
  @param AFormatString String of formatting codes
  @param ANumFormat    Identifier for built-in number format
  @return              List index of the new item }
function TsCustomNumFormatList.AddFormat(AFormatIndex: Integer;
  AFormatName, AFormatString: String; ANumFormat: TsNumberFormat): Integer;
var
  item: TsNumFormatData;
begin
  item := TsNumFormatData.Create;
  item.Index := AFormatIndex;
  item.Name := AFormatName;
  item.NumFormat := ANumFormat;
  item.FormatString := AFormatString;
  Result := inherited Add(item);
end;

{@@ Adds a number format described by the Excel format index, the format string,
    and the built-in format identifier to the list and returns the index of
    the new item in the format list. To be used when writing an Excel file.
  @param AFormatIndex  Format index to be used by Excel
  @param AFormatString String of formatting codes
  @param ANumFormat    Identifier for built-in number format
  @return              Index of the new item in the format list }
function TsCustomNumFormatList.AddFormat(AFormatIndex: Integer;
  AFormatString: String; ANumFormat: TsNumberFormat): integer;
begin
  Result := AddFormat(AFormatIndex, '', AFormatString, ANumFormat);
end;

{@@ Adds a number format described by the ODF format name, the format string,
    and the built-in format identifier to the list and returns the index of
    the new item in the format list. To be used when writing an ODS file.
  @param AFormatName   Format name to be used by OpenDocument
  @param AFormatString String of formatting codes
  @param ANumFormat    Identifier for built-in number format
  @return              Index of the new item in the format list }
function TsCustomNumFormatList.AddFormat(AFormatName, AFormatString: String;
  ANumFormat: TsNumberFormat): Integer;
begin
  if (AFormatString = '') and (ANumFormat <> nfGeneral) then begin
    Result := 0;
    exit;
  end;
  Result := AddFormat(FNextFormatIndex, AFormatName, AFormatString, ANumFormat);
  inc(FNextFormatIndex);
end;

{@@ Adds a number format described by the format string, and the built-in
    format identifier to the format list and returns the index of the new
    item in the list. The Excel format index and ODS format name are auto-generated.
  @param AFormatString String of formatting codes
  @param ANumFormat    Identifier for built-in number format
  @return              Index of the new item in the list }
function TsCustomNumFormatList.AddFormat(AFormatString: String;
  ANumFormat: TsNumberFormat): Integer;
begin
  Result := AddFormat('', AFormatString, ANumFormat);
end;

{@@ Adds the number format used by a given cell to the list.
  @param AFormatCell Pointer to a cell providing the format to be stored in the list }
function TsCustomNumFormatList.AddFormat(AFormatCell: PCell): Integer;
begin
  if AFormatCell = nil then
    raise Exception.Create('TsCustomNumFormat.Add: No nil pointers please');

  if Count = 0 then
    raise Exception.Create('TsCustomNumFormatList: Error in program logics: You must provide built-in formats first.');

  Result := AddFormat(FNextFormatIndex,
    AFormatCell^.NumberFormatStr,
    AFormatCell^.NumberFormat
  );

  inc(FNextFormatIndex);
end;

{@@
  Adds the builtin format items to the list. The formats must be specified in
  a way that is compatible with fpc syntax.

  Conversion of the formatstrings to the syntax used in the destination file
  can be done by calling "ConvertAfterReadung" bzw. "ConvertBeforeWriting".
  "AddBuiltInFormats" must be called before user items are added.

  Must specify FFirstFormatIndexInFile (BIFF5-8, e.g. doesn't save formats <164)
  and must initialize the index of the first user format (FNextFormatIndex)
  which is automatically incremented when adding user formats.

  In TsCustomNumFormatList nothing is added. }
procedure TsCustomNumFormatList.AddBuiltinFormats;
begin
  // must be overridden - see xlscommon as an example.
end;

{@@ Called from the reader when a format item has been read from an Excel file.
  Determines the number format type, format string etc and converts the
  format string to fpc syntax which is used directly for getting the cell text.
 @param AFormatIndex Excel index of the number format read from the file
 @param AFormatString String of formatting codes as read fromt the file. }
procedure TsCustomNumFormatList.AnalyzeAndAdd(AFormatIndex: Integer;
  AFormatString: String);
var
  nf: TsNumberFormat = nfGeneral;
begin
  if FindByIndex(AFormatIndex) > -1 then
    exit;

  // Analyze & convert the format string, extract infos for internal formatting
  ConvertAfterReading(AFormatIndex, AFormatString, nf);

  // Add the new item
  AddFormat(AFormatIndex, AFormatString, nf);
end;

{@@ Clears the number format list and frees memory occupied by the format items. }
procedure TsCustomNumFormatList.Clear;
var
  i: Integer;
begin
  for i:=0 to Count-1 do RemoveFormat(i);
  inherited Clear;
end;

{@@
  Takes the format string as it is read from the file and extracts the
  built-in number format identifier out of it for use by fpc.
  The method also converts the format string to a form that can be used
  by fpc's FormatDateTime and FormatFloat.

  The method should be overridden in a class that knows knows more about the
  details of the spreadsheet file format.

  @param AFormatIndex   Excel index of the number format read
  @param AFormatString  string of formatting codes extracted from the file data
  @param ANumFormat     identifier for built-in fpspreadsheet format extracted
                        from the file data }
procedure TsCustomNumFormatList.ConvertAfterReading(AFormatIndex: Integer;
  var AFormatString: String; var ANumFormat: TsNumberFormat);
var
  parser: TsNumFormatParser;
  fmt: String;
  lFormatData: TsNumFormatData;
  i: Integer;
begin
  i := FindByIndex(AFormatIndex);
  if i > 0 then begin
    lFormatData := Items[i];
    fmt := lFormatData.FormatString;
  end else
    fmt := AFormatString;

  // Analyzes the format string and tries to convert it to fpSpreadsheet format.
  parser := TsNumFormatParser.Create(Workbook, fmt);
  try
    if parser.Status = psOK then begin
      ANumFormat := parser.NumFormat;
      AFormatString := parser.FormatString[nfdDefault];
    end else begin
      //  Show an error here?
    end;
  finally
    parser.Free;
  end;
end;

{@@
  Is called before collecting all number formats of the spreadsheet and before
  writing them to file. Its purpose is to convert the format string as used by fpc
  to a format compatible with the spreadsheet file format.
  Nothing is changed in the TsCustomNumFormatList, the method needs to be
  overridden by a descendant class which known more about the details of the
  destination file format.

  Needs to be overridden by a class knowing more about the destination file
  format.

  @param AFormatString String of formatting codes. On input in fpc syntax. Is
                       overwritten on output by format string compatible with
                       the destination file.
  @param ANumFormat    Identifier for built-in fpspreadsheet number format }
procedure TsCustomNumFormatList.ConvertBeforeWriting(var AFormatString: String;
  var ANumFormat: TsNumberFormat);
begin
  Unused(AFormatString, ANumFormat);
  // nothing to do here. But see, e.g., xlscommon.TsBIFFNumFormatList
end;


{@@ Deletes a format item from the list, and makes sure that its memory is
    released.

    @param  AIndex   List index of the item to be deleted. }
procedure TsCustomNumFormatList.Delete(AIndex: Integer);
begin
  RemoveFormat(AIndex);
  Delete(AIndex);
end;

{@@ Seeks a format item with the given properties and returns its list index,
  or -1 if not found.

 @param ANumFormat    Built-in format identifier
 @param AFormatString String of formatting codes
 @return              Index of the format item in the format list, or -1 if not found. }
function TsCustomNumFormatList.Find(ANumFormat: TsNumberFormat;
  AFormatString: String): Integer;
var
  item: TsNumFormatData;
begin
  for Result := Count-1 downto 0 do begin
    item := Items[Result];
    if (item <> nil) and (item.NumFormat = ANumFormat) and (item.FormatString = AFormatString)
      then exit;
  end;
  Result := -1;
end;

{@@ Finds the item with the given format string and returns its index in the
  format list, or -1 if not found.

  @param  AFormatString  string of formatting codes to be searched in the list.
  @return Index of the format item in the format list, or -1 if not found. }
function TsCustomNumFormatList.Find(AFormatString: String): integer;
var
  item: TsNumFormatData;
begin
  { We search backwards to find user-defined items first. They usually are
    more appropriate than built-in items. }
  for Result := Count-1 downto 0 do begin
    item := Items[Result];
    if item.FormatString = AFormatString then
      exit;
  end;
  Result := -1;
end;

{@@ Finds the item with the given Excel format index and returns its index in
  the format list, or -1 if not found.
  Is used by BIFF file formats.

  @param  AFormatIndex  Excel format index to the searched
  @return Index of the format item in the format list, or -1 if not found. }
function TsCustomNumFormatList.FindByIndex(AFormatIndex: Integer): integer;
var
  item: TsNumFormatData;
begin
  for Result := 0 to Count-1 do begin
    item := Items[Result];
    if item.Index = AFormatIndex then
      exit;
  end;
  Result := -1;
end;

{@@
  Finds the item with the given ODS format name and returns its index in
  the format list (or -1, if not found)
  To be used by OpenDocument file format.

  @param  AFormatName  Format name as used by OpenDocument to identify a number format
  @return Index of the format item in the list, or -1 if not found }
function TsCustomNumFormatList.FindByName(AFormatName: String): integer;
var
  item: TsNumFormatData;
begin
  for Result := 0 to Count-1 do begin
    item := Items[Result];
    if item.Name = AFormatName then
      exit;
  end;
  Result := -1;
end;

{@@
  Determines whether the format attributed to the given cell is already
  contained in the list and returns its list index, or -1 if not found

  @param  AFormatCell Pointer to a spreadsheet cell having the number format that is looked for.
  @return Index of the format item in the list, or -1 if not found. }
function TsCustomNumFormatList.FindFormatOf(AFormatCell: PCell): integer;
begin
  if AFormatCell = nil then
    Result := -1
  else
    Result := Find(AFormatCell^.NumberFormat, AFormatCell^.NumberFormatStr);
end;

{@@
  Determines the format string to be written into the spreadsheet file. Calls
  ConvertBeforeWriting in order to convert the fpc format strings to the dialect
  used in the file.

  @param AIndex  Index of the format item under consideration.
  @return        String of formatting codes that will be written to the file. }
function TsCustomNumFormatList.FormatStringForWriting(AIndex: Integer): String;
var
  item: TsNumFormatdata;
  nf: TsNumberFormat;
begin
  item := Items[AIndex];
  if item <> nil then begin
    Result := item.FormatString;
    nf := item.NumFormat;
    ConvertBeforeWriting(Result, nf);
  end else
    Result := '';
end;

function TsCustomNumFormatList.GetItem(AIndex: Integer): TsNumFormatData;
begin
  Result := TsNumFormatData(inherited Items[AIndex]);
end;

{@@
  Deletes the memory occupied by the formatting data, but keeps an empty item in
  the list to retain the indexes of following items.

  @param AIndex The number format item at this index will be removed. }
procedure TsCustomNumFormatList.RemoveFormat(AIndex: Integer);
var
  item: TsNumFormatData;
begin
  item := GetItem(AIndex);
  if item <> nil then begin
    item.Free;
    SetItem(AIndex, nil);
  end;
end;

procedure TsCustomNumFormatList.SetItem(AIndex: Integer; AValue: TsNumFormatData);
begin
  inherited Items[AIndex] := AValue;
end;

function CompareNumFormatData(Item1, Item2: Pointer): Integer;
begin
  Result := CompareValue(TsNumFormatData(Item1).Index, TsNumFormatData(Item2).Index);
end;

{@@ Sorts the format data items in ascending order of the Excel format indexes. }
procedure TsCustomNumFormatList.Sort;
begin
  inherited Sort(@CompareNumFormatData);
end;


{ TsCustomSpreadReaderWriter }

{@@
  Constructor of the reader/writer. Has the workbook to be read/written as a
  parameter to apply the localization information found in its FormatSettings.
  Creates an internal instance of the number format list according to the
  file format being read/written.

  @param AWorkbook  Workbook into which the file is being read or from with the
                    file is written. This parameter is passed from the workbook
                    which creates the reader/writer. }
constructor TsCustomSpreadReaderWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  { A good starting point valid for many formats ... }
  FLimitations.MaxColCount := 256;
  FLimitations.MaxRowCount := 65536;
  FLimitations.MaxPaletteSize := $FFFFFFFF;
  // Number formats
  CreateNumFormatList;
end;

{@@
  Destructor of the reader. Destroys the internal number format list and the
  error log list. }
destructor TsCustomSpreadReaderWriter.Destroy;
begin
  FNumFormatList.Free;
  inherited Destroy;
end;

{@@
  Creates an instance of the number format list which contains prototypes of
  all number formats found in the workbook (when writing) or in the file (when
  reading).

  The method has to be overridden because the descendants know the special
  requirements of the file format. }
procedure TsCustomSpreadReaderWriter.CreateNumFormatList;
begin
  // nothing to do here
end;

{@@
  Returns a record containing limitations of the specific file format of the
  writer.
}
function TsCustomSpreadReaderWriter.Limitations: TsSpreadsheetFormatLimitations;
begin
  Result := FLimitations;
end;

{ TsCustomSpreadReader }

{@@
  Constructor of the reader. Has the workbook to be read as a parameter to
  apply the localization information found in its FormatSettings.
  Creates an internal instance of the number format list according to the
  file format being read.

  @param AWorkbook  Workbook into which the file is being read. This parameter
                    is passed from the workbook which creates the reader. }
constructor TsCustomSpreadReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FIsVirtualMode := (boVirtualMode in FWorkbook.Options) and
    Assigned(FWorkbook.OnReadCellData);
end;

{@@
  Default file reading method.

  Opens the file and calls ReadFromStream

  @param  AFileName The input file name.
  @param  AData     The Workbook to be filled with information from the file.
  @see    TsWorkbook
}
procedure TsCustomSpreadReader.ReadFromFile(AFileName: string; AData: TsWorkbook);
{
var
  fs, ms: TStream;
begin
  fs := TFileStream.Create(AFileName, fmOpenRead);
  ms := TMemoryStream.Create;
  try
    ms.CopyFrom(fs, fs.Size);
    ms.Position := 0;
    ReadFromStream(ms, AData);
  finally
    ms.Free;
    fs.Free;
  end;
end;
 }
var
  InputFile: TStream;
begin
  if (boBufStream in Workbook.Options) then
    InputFile := TBufStream.Create(AFileName, fmOpenRead)
  else
    InputFile := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadFromStream(InputFile, AData);
  finally
    InputFile.Free;
  end;
end;

{@@
  This routine has the purpose to read the workbook data from the stream.
  It should be overriden in descendent classes.

  Its basic implementation here assumes that the stream is a TStringStream and
  the data are provided by calling ReadFromStrings. This mechanism is valid
  for wikitables.

  @param  AStream   Stream containing the workbook data
  @param  AData     Workbook which is filled by the data from the stream.
}
procedure TsCustomSpreadReader.ReadFromStream(AStream: TStream; AData: TsWorkbook);
var
  AStringStream: TStringStream;
  AStrings: TStringList;
begin
  AStringStream := TStringStream.Create('');
  AStrings := TStringList.Create;
  try
    AStringStream.CopyFrom(AStream, AStream.Size);
    AStringStream.Seek(0, soFromBeginning);
    AStrings.Text := AStringStream.DataString;
    ReadFromStrings(AStrings, AData);
  finally
    AStringStream.Free;
    AStrings.Free;
  end;
end;

{@@
  Reads workbook data from a string list. This abstract implementation does
  nothing and raises an exception. Must be overridden, like for wikitables.
}
procedure TsCustomSpreadReader.ReadFromStrings(AStrings: TStrings;
  AData: TsWorkbook);
begin
  Unused(AStrings, AData);
  raise Exception.Create(lpUnsupportedReadFormat);
end;

{ TsCustomSpreadWriter }

{@@ Constructor of the writer. Has the workbook to be written as a parameter to
  apply the localization information found in its FormatSettings.
  Creates an internal number format list to collect unique samples of all the
  number formats found in the workbook.

  @param AWorkbook  Workbook which is to be written to file/stream.
                    This parameter is passed from the workbook which creates the
                    writer.
}
constructor TsCustomSpreadWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
end;

{@@
  Checks if the formatting style of a cell is in the list of manually added
  FFormattingStyles and returns its index, or -1 if it isn't

  @param  AFormat  Cell containing the formatting styles which are seeked in the
                   FFormattingStyles array.
}
function TsCustomSpreadWriter.FindFormattingInList(AFormat: PCell): Integer;
var
  i, n: Integer;
  b: TsCellBorder;
  equ: Boolean;
begin
  Result := -1;

  n := Length(FFormattingStyles);
  for i := n - 1 downto 0 do begin
    if (FFormattingStyles[i].UsedFormattingFields <> AFormat^.UsedFormattingFields) then Continue;

    if uffHorAlign in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].HorAlignment <> AFormat^.HorAlignment) then Continue;

    if uffVertAlign in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].VertAlignment <> AFormat^.VertAlignment) then Continue;

    if uffTextRotation in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].TextRotation <> AFormat^.TextRotation) then Continue;

    if uffBorder in AFormat^.UsedFormattingFields then begin
      if (FFormattingStyles[i].Border <> AFormat^.Border) then Continue;
      equ := true;
      for b in TsCellBorder do begin
        if FFormattingStyles[i].BorderStyles[b].LineStyle <> AFormat^.BorderStyles[b].LineStyle
        then begin
          equ := false;
          Break;
        end;
        if FFormattingStyles[i].BorderStyles[b].Color <> FixColor(AFormat^.BorderStyles[b].Color)
        then begin
          equ := false;
          Break;
        end;
      end;
      if not equ then Continue;
    end;

    if uffBackgroundColor in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].BackgroundColor <> FixColor(AFormat^.BackgroundColor)) then Continue;

    if uffNumberFormat in AFormat^.UsedFormattingFields then begin
      if (FFormattingStyles[i].NumberFormat <> AFormat^.NumberFormat) then Continue;
      if (FFormattingStyles[i].NumberFormatStr <> AFormat^.NumberFormatStr) then Continue;
    end;

    if uffFont in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].FontIndex <> AFormat^.FontIndex) then Continue;

    // If we arrived here it means that the styles match
    Exit(i);
  end;
end;

{@@
  Makes sure that all colors used in a given cell belong to the workbook's
  color palette.
}
procedure TsCustomSpreadWriter.FixCellColors(ACell: PCell);
var
  b: TsCellBorder;
begin
  if ACell = nil then
    exit;

  ACell^.BackgroundColor := FixColor(ACell^.BackgroundColor);

  for b in TsCellBorders do
    ACell^.BorderStyles[b].Color := FixColor(ACell^.BorderStyles[b].Color);

  // Font color is not corrected here because this would affect other writers.
  // Font color is handled immediately before writing.
end;

{@@
  If a color index is greater then the maximum palette color count this
  color is replaced by the closest palette color.

  The present implementation does not change the color. Must be overridden by
  writers of formats with limited palette sizes.

  @param  AColor   Color palette index to be checked
  @return Closest color to AColor. If AColor belongs to the palette it must
          be returned unchanged. }
function TsCustomSpreadWriter.FixColor(AColor: TsColor): TsColor;
begin
  Result := AColor;
end;

{@@
  If formatting features of a cell are not supported by the destination file
  format of the writer, here is the place to apply replacements.
  Must be overridden by descendants, nothin happens here. See BIFF2.

  @param  ACell  Pointer to the cell being investigated. Note that this cell
                 does not belong to the workbook, but is a cell of the
                 FFormattingStyles array.
}
procedure TsCustomSpreadWriter.FixFormat(ACell: PCell);
begin
  Unused(ACell);
  // to be overridden
end;

{@@
  Adjusts relative references in the formula element to the position of cell
  and the shared formula base. }
procedure TsCustomSpreadWriter.FixRelativeReferences(ACell: PCell;
  var AElement: TsFormulaElement);
var
  rowOffset: Integer;
  colOffset: Integer;
begin
  if (ACell = nil) or (ACell^.SharedFormulaBase = nil) then
    exit;

  case AElement.ElementKind of
    fekCell:
      begin
        rowOffset := AElement.Row - ACell^.SharedFormulaBase^.Row;
        colOffset := AElement.Col - ACell^.SharedFormulaBase^.Col;
        if (rfRelRow in AElement.RelFlags) then
          AElement.Row := Integer(ACell^.Row) + rowOffset;
        if (rfRelCol in AElement.RelFlags) then
          AElement.Col := Integer(ACell^.Col) + colOffset;
      end;
    fekCellRange:
      begin
        rowOffset := AElement.Row - ACell^.SharedFormulaBase^.Row;
        colOffset := AElement.Col - ACell^.SharedFormulaBase^.Col;
        if (rfRelRow in AElement.RelFlags) then
          AElement.Row := Integer(ACell^.Row) + rowOffset;
        if (rfRelCol in AElement.RelFlags) then
          AElement.Col := Integer(ACell^.Col) + colOffset;
        if (rfRelRow2 in AElement.RelFlags) then
          AElement.Row2 := Integer(ACell^.Row) + rowOffset;
        if (rfRelCol2 in AElement.RelFlags) then
          AElement.Col2 := Integer(ACell^.Col) + colOffset;
      end;
  end;
end;

{@@
  Determines the size of the worksheet to be written. VirtualMode is respected.
  Is called when the writer needs the size for output. Column and row count
  limitations are repsected as well.

  @param   AWorksheet  Worksheet to be written
  @param   AFirsRow    Index of first row to be written
  @param   ALastRow    Index of last row
  @param   AFirstCol   Index of first column to be written
  @param   ALastCol    Index of last column to be written
}
procedure TsCustomSpreadWriter.GetSheetDimensions(AWorksheet: TsWorksheet;
  out AFirstRow, ALastRow, AFirstCol, ALastCol: Cardinal);
begin
  if (boVirtualMode in AWorksheet.Workbook.Options) then begin
    AFirstRow := 0;
    AFirstCol := 0;
    ALastRow := AWorksheet.Workbook.VirtualRowCount-1;
    ALastCol := AWorksheet.Workbook.VirtualColCount-1;
  end else begin
    Workbook.UpdateCaches;
    AFirstRow := AWorksheet.GetFirstRowIndex;
    AFirstCol := AWorksheet.GetFirstColIndex;
    ALastRow := AWorksheet.GetLastRowIndex;
    ALastCol := AWorksheet.GetLastColIndex;
  end;
  if AFirstCol >= Limitations.MaxColCount then
    AFirstCol := Limitations.MaxColCount-1;
  if AFirstRow >= Limitations.MaxRowCount then
    AFirstRow := Limitations.MaxRowCount-1;
  if ALastCol >= Limitations.MaxColCount then
    ALastCol := Limitations.MaxColCount-1;
  if ALastRow >= Limitations.MaxRowCount then
    ALastRow := Limitations.MaxRowCount-1;
end;

{@@
  Each descendent should define its own default formats, if any.
  Always add the normal, unformatted style first to speed things up.

  To be overridden by descendants.
}
procedure TsCustomSpreadWriter.AddDefaultFormats();
begin
  SetLength(FFormattingStyles, 0);
  NextXFIndex := 0;
end;

{@@
  Checks limitations of the writer, e.g max row/column count
}
procedure TsCustomSpreadWriter.CheckLimitations;
var
  lastCol, lastRow: Cardinal;
  i, n: Integer;
begin
  Workbook.GetLastRowColIndex(lastRow, lastCol);

  // Check row count
  if lastRow >= FLimitations.MaxRowCount then
    Workbook.AddErrorMsg(lpMaxRowsExceeded, [lastRow+1, FLimitations.MaxRowCount]);

  // Check column count
  if lastCol >= FLimitations.MaxColCount then
    Workbook.AddErrorMsg(lpMaxColsExceeded, [lastCol+1, FLimitations.MaxColCount]);

  // Check color count.
  n := Workbook.GetPaletteSize;
  if n > FLimitations.MaxPaletteSize then
    for i:= FLimitations.MaxPaletteSize to n-1 do
      if Workbook.UsesColor(i) then
      begin
        Workbook.AddErrorMsg(lpTooManyPaletteColors, [n, FLimitations.MaxPaletteSize]);
        break;
      end;
end;

{@@
  Callback function for collecting all formatting styles found in the worksheet.

  @param  ACell    Pointer to the worksheet cell being tested whether its format
                   already has been found in the array FFormattingStyles.
  @param  AStream  Stream to which the workbook is written
}
procedure TsCustomSpreadWriter.ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
var
  Len: Integer;
begin
  Unused(AStream);

  FixFormat(ACell);

  if ACell^.UsedFormattingFields = [] then Exit;
  if FindFormattingInList(ACell) <> -1 then Exit;

  Len := Length(FFormattingStyles);
  SetLength(FFormattingStyles, Len+1);
  FFormattingStyles[Len] := ACell^;

  // Make sure that all colors of the formatting style cell are used in the workbook's
  // palette.
  FixCellColors(@FFormattingStyles[Len]);

  // We store the index of the XF record that will be assigned to this style in
  // the "row" of the style. Will be needed when writing the XF record.
  FFormattingStyles[Len].Row := NextXFIndex;
  Inc(NextXFIndex);
end;

{@@
  This method collects all formatting styles found in the worksheet and
  stores unique prototypes in the array FFormattingStyles.
}
procedure TsCustomSpreadWriter.ListAllFormattingStyles;
var
  i: Integer;
begin
  SetLength(FFormattingStyles, 0);

  // Add default styles which are required to be there by the destination file
  AddDefaultFormats();

  // Iterate through all cells and collect the individual styles
  for i := 0 to Workbook.GetWorksheetCount - 1 do
    IterateThroughCells(nil, Workbook.GetWorksheetByIndex(i).Cells, ListAllFormattingStylesCallback);
end;

{@@
  Adds the number format of the given cell to the NumFormatList, but only if
  it does not yet exist in the list.
}
procedure TsCustomSpreadWriter.ListAllNumFormatsCallback(ACell: PCell; AStream: TStream);
var
  fmt: string;
  nf: TsNumberFormat;
begin
  Unused(AStream);

  if ACell^.NumberFormat = nfGeneral then
    exit;

  // The builtin format list is in fpc dialect.
  fmt := ACell^.NumberFormatStr;
  nf := ACell^.NumberFormat;

  // Seek the format string in the current number format list.
  // If not found add the format to the list.
  if FNumFormatList.Find(nf, fmt) = -1 then
    FNumFormatList.AddFormat(fmt, nf);
end;

{@@
  Iterates through all cells and collects the number formats in
  FNumFormatList (without duplicates).
  The index of the list item is needed for the field FormatIndex of the XF record.
  At the time when the method is called the formats are still in fpc dialect. }
procedure TsCustomSpreadWriter.ListAllNumFormats;
var
  i: Integer;
begin
  for i:=0 to Workbook.GetWorksheetCount-1 do
    IterateThroughCells(nil, Workbook.GetWorksheetByIndex(i).Cells, ListAllNumFormatsCallback);
  NumFormatList.Sort;
end;

{@@
  Helper function for the spreadsheet writers. Writes the cell value to the
  stream. Calls the WriteNumber method of the worksheet for writing a number,
  the WriteDateTime method for writing a date/time etc.

  @param  ACell   Pointer to the worksheet cell being written
  @param  AStream Stream to which data are written

  @see    TsCustomSpreadWriter.WriteCellsToStream
}
procedure TsCustomSpreadWriter.WriteCellCallback(ACell: PCell; AStream: TStream);
begin
  if HasFormula(ACell) then
    WriteFormula(AStream, ACell^.Row, ACell^.Col, ACell)
  else
  {
  if Length(ACell^.RPNFormulaValue) > 0 then
    // A non-calculated RPN formula has ContentType cctUTF8Formula, but after
    // calculation it has the content type of the result. Both cases have in
    // common that there is a non-vanishing array of rpn tokens which has to
    // be written to file.
    WriteRPNFormula(AStream, ACell^.Row, ACell^.Col, ACell^.RPNFormulaValue, ACell)
  else
  }
  case ACell.ContentType of
    cctEmpty      : WriteBlank(AStream, ACell^.Row, ACell^.Col, ACell);
    cctDateTime   : WriteDateTime(AStream, ACell^.Row, ACell^.Col, ACell^.DateTimeValue, ACell);
    cctNumber     : WriteNumber(AStream, ACell^.Row, ACell^.Col, ACell^.NumberValue, ACell);
    cctUTF8String : WriteLabel(AStream, ACell^.Row, ACell^.Col, ACell^.UTF8StringValue, ACell);
   // cctFormula    : WriteFormula(AStream, ACell^.Row, ACell^.Col, ACell^.FormulaValue, ACell);
  end;
end;

{@@
  Helper function for the spreadsheet writers.

  Iterates all cells on a list, calling the appropriate write method for them.

  @param  AStream The output stream.
  @param  ACells  List of cells to be writeen
}
procedure TsCustomSpreadWriter.WriteCellsToStream(AStream: TStream; ACells: TAVLTree);
begin
  IterateThroughCells(AStream, ACells, WriteCellCallback);
end;

{@@
  A generic method to iterate through all cells in a worksheet and call a callback
  routine for each cell.

  @param  AStream    The output stream, passed to the callback routine.
  @param  ACells     List of cells to be iterated
  @param  ACallback  Callback routine; it requires as arguments a pointer to the
                     cell as well as the destination stream.
}
procedure TsCustomSpreadWriter.IterateThroughCells(AStream: TStream; ACells: TAVLTree; ACallback: TCellsCallback);
var
  AVLNode: TAVLTreeNode;
begin
  AVLNode := ACells.FindLowest;
  While Assigned(AVLNode) do
  begin
    ACallback(PCell(AVLNode.Data), AStream);
    AVLNode := ACells.FindSuccessor(AVLNode);
  end;
end;

{@@
  Default file writting method.

  Opens the file and calls WriteToStream
  The workbook written is the one specified in the constructor of the writer.

  @param  AFileName           The output file name.
  @param  AOverwriteExisting  If the file already exists it will be replaced.

  @see    TsWorkbook
}
procedure TsCustomSpreadWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean = False);
var
  OutputFile: TStream;
  lMode: Word;
begin
  if AOverwriteExisting then lMode := fmCreate or fmOpenWrite
  else lMode := fmCreate;

  if (boBufStream in Workbook.Options) then
    OutputFile := TBufStream.Create(AFileName, lMode)
  else
    OutputFile := TFileStream.Create(AFileName, lMode);
  try
    WriteToStream(OutputFile);
  finally
    OutputFile.Free;
  end;
end;

{@@
  This routine has the purpose to write the workbook to a stream.
  Present implementation writes to a stringlists by means of WriteToStrings;
  this behavior is required for wikitables.
  Must be overriden in descendent classes for all other cases.

  @param  AStream   Stream to which the workbook is written
}
procedure TsCustomSpreadWriter.WriteToStream(AStream: TStream);
var
  lStringList: TStringList;
begin
  lStringList := TStringList.Create;
  try
    WriteToStrings(lStringList);
    lStringList.SaveToStream(AStream);
  finally
    lStringList.Free;
  end;
end;

{@@
  Writes the worksheet to a list of strings. Not implemented here, needs to
  be overridden by descendants. See wikitables.
}
procedure TsCustomSpreadWriter.WriteToStrings(AStrings: TStrings);
begin
  Unused(AStrings);
  raise Exception.Create(lpUnsupportedWriteFormat);
end;

{@@
  Basic method which is called when writing a formula to a stream. The formula
  is already stored in the cell fields.
  Present implementation does nothing. Needs to be overridden by descendants.

  @param   AStream   Stream to be written
  @param   ARow      Row index of the cell containing the formula
  @param   ACol      Column index of the cell containing the formula
  @param   ACell     Pointer to the cell containing the formula and being written
                     to the stream
}
procedure TsCustomSpreadWriter.WriteFormula(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
end;
(*
{@@
  Basic method which is called when writing an RPN formula to a stream.
  Present implementation does nothing. Needs to be overridden by descendants.

  RPN formula are used by the BIFF file format.

  @param   AStream   Stream to be written
  @param   ARow      Row index of the cell containing the formula
  @param   ACol      Column index of the cell containing the formula
  @param   AFormula  RPN formula given as an array of RPN tokens
  @param   ACell     Pointer to the cell containing the formula and being written
                     to the stream
}
procedure TsCustomSpreadWriter.WriteRPNFormula(AStream: TStream;
  const ARow, ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell);
begin
  Unused(AStream, ARow, ACol);
  Unused(AFormula, ACell);
  // Silently dump the formula; child classes should implement their own support
end;
*)


{******************************************************************************}
{                   Simplified creation of RPN formulas                        }
{******************************************************************************}

{@@
  Creates a pointer to a new RPN item. This represents an element in the array
  of token of an RPN formula.

  @return  Pointer to the RPN item
}
function NewRPNItem: PRPNItem;
begin
  New(Result);
  FillChar(Result^.FE, SizeOf(Result^.FE), 0);
  Result^.FE.StringValue := '';
end;

{@@
  Destroys an RPN item
}
procedure DisposeRPNItem(AItem: PRPNItem);
begin
  if AItem <> nil then
    Dispose(AItem);
end;

{@@
  Creates a boolean value entry in the RPN array.

  @param  AValue   Boolean value to be stored in the RPN item
  @next   ANext    Pointer to the next RPN item in the list
}
function RPNBool(AValue: Boolean; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekBool;
  if AValue then Result^.FE.DoubleValue := 1.0 else Result^.FE.DoubleValue := 0.0;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a cell value, specifed by its
  address, e.g. 'A1'. Takes care of absolute and relative cell addresses.

  @param  ACellAddress   Adress of the cell given in Excel A1 notation
  @param  ANext          Pointer to the next RPN item in the list
}
function RPNCellValue(ACellAddress: String; ANext: PRPNItem): PRPNItem;
var
  r,c: Cardinal;
  flags: TsRelFlags;
begin
  if not ParseCellString(ACellAddress, r, c, flags) then
    raise Exception.CreateFmt('"%s" is not a valid cell address.', [ACellAddress]);
  Result := RPNCellValue(r,c, flags, ANext);
end;

{@@
  Creates an entry in the RPN array for a cell value, specifed by its
  row and column index and a flag containing information on relative addresses.

  @param  ARow     Row index of the cell
  @param  ACol     Column index of the cell
  @param  AFlags   Flags specifying absolute or relative cell addresses
  @param  ANext    Pointer to the next RPN item in the list
}
function RPNCellValue(ARow, ACol: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCell;
  Result^.FE.Row := ARow;
  Result^.FE.Col := ACol;
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a cell reference, specifed by its
  address, e.g. 'A1'. Takes care of absolute and relative cell addresses.
  "Cell reference" means that all properties of the cell can be handled.
  Note that most Excel formulas with cells require the cell value only
  (--> RPNCellValue)

  @param  ACellAddress   Adress of the cell given in Excel A1 notation
  @param  ANext          Pointer to the next RPN item in the list
}
function RPNCellRef(ACellAddress: String; ANext: PRPNItem): PRPNItem;
var
  r,c: Cardinal;
  flags: TsRelFlags;
begin
  if not ParseCellString(ACellAddress, r, c, flags) then
    raise Exception.CreateFmt(lpNoValidCellAddress, [ACellAddress]);
  Result := RPNCellRef(r,c, flags, ANext);
end;

{@@
  Creates an entry in the RPN array for a cell reference, specifed by its
  row and column index and flags containing information on relative addresses.
  "Cell reference" means that all properties of the cell can be handled.
  Note that most Excel formulas with cells require the cell value only
  (--> RPNCellValue)

  @param  ARow     Row index of the cell
  @param  ACol     Column index of the cell
  @param  AFlags   Flags specifying absolute or relative cell addresses
  @param  ANext    Pointer to the next RPN item in the list
}
function RPNCellRef(ARow, ACol: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCellRef;
  Result^.FE.Row := ARow;
  Result^.FE.Col := ACol;
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a range of cells, specified by an
  Excel-style address, e.g. A1:G5. As in Excel, use a $ sign to indicate
  absolute addresses.

  @param  ACellRangeAddress   Adress of the cell range given in Excel notation, such as A1:G5
  @param  ANext               Pointer to the next RPN item in the list
}
function RPNCellRange(ACellRangeAddress: String; ANext: PRPNItem): PRPNItem;
var
  r1,c1, r2,c2: Cardinal;
  flags: TsRelFlags;
begin
  if not ParseCellRangeString(ACellRangeAddress, r1,c1, r2,c2, flags) then
    raise Exception.CreateFmt(lpNoValidCellRangeAddress, [ACellRangeAddress]);
  Result := RPNCellRange(r1,c1, r2,c2, flags, ANext);
end;

{@@
  Creates an entry in the RPN array for a range of cells, specified by the
  row/column indexes of the top/left and bottom/right corners of the block.
  The flags indicate relative indexes.

  @param  ARow     Row index of the top/left cell
  @param  ACol     Column index of the top/left cell
  @param  ARow2    Row index of the bottom/right cell
  @param  ACol2    Column index of the bottom/right cell
  @param  AFlags   Flags specifying absolute or relative cell addresses
  @param  ANext    Pointer to the next RPN item in the list
}
function RPNCellRange(ARow, ACol, ARow2, ACol2: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCellRange;
  Result^.FE.Row := ARow;
  Result^.FE.Col := ACol;
  Result^.FE.Row2 := ARow2;
  Result^.FE.Col2 := ACol2;
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a relative cell reference as used in
  shared formulas. The given parameters indicate the relativ offset between
  the current cell coordinates and a reference rell.

  @param  ARowOffset  Offset between current row and the row of a reference cell
  @param  AColOffset  Offset between current column and the column of a reference cell
  @param  AFlags      Flags specifying absolute or relative cell addresses
  @param  ANext       Pointer to the next RPN item in the list
}
function RPNCellOffset(ARowOffset, AColOffset: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCellOffset;
  Result^.FE.Row := Cardinal(ARowOffset);
  Result^.FE.Col := Cardinal(AColOffset);
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array with an error value.

  @param  AErrCode  Error code to be inserted (see TsErrorValue
  @param  ANext     Pointer to the next RPN item in the list
  @see TsErrorValue
}
function RPNErr(AErrCode: Byte; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekErr;
  Result^.FE.IntValue := AErrCode;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a 2-byte unsigned integer

  @param  AValue  Integer value to be inserted into the formula
  @param  ANext   Pointer to the next RPN item in the list
}
function RPNInteger(AValue: Word; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekInteger;
  Result^.FE.IntValue := AValue;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a missing argument in of function call.
  Use this in a formula to indicate a missing argument

  @param ANext  Pointer to the next RPN item in the list.
}
function RPNMissingArg(ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekMissingArg;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a floating point number.

  @param  AValue  Number value to be inserted into the formula
  @param  ANext   Pointer to the next RPN item in the list
}
function RPNNumber(AValue: Double; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekNum;
  Result^.FE.DoubleValue := AValue;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array which puts the current operator in parenthesis.
  For display purposes only, does not affect calculation.

  @param  ANext   Pointer to the next RPN item in the list
}
function RPNParenthesis(ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekParen;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a string.

  @param  AValue  String to be inserted into the formula
  @param  ANext   Pointer to the next RPN item in the list
}
function RPNString(AValue: String; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekString;
  Result^.FE.StringValue := AValue;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for an Excel function or operation
  specified by its TokenID (--> TFEKind). Note that array elements for all
  needed parameters must have been created before.

  @param  AToken  Formula element indicating the function to be executed,
                  see the TFEKind enumeration for possible values.
  @param  ANext   Pointer to the next RPN item in the list

  @see TFEKind
}
function RPNFunc(AToken: TFEKind; ANext: PRPNItem): PRPNItem;
begin
  {
  if FEProps[AToken].MinParams <> FEProps[AToken].MaxParams then
    raise Exception.CreateFmt(lpSpecifyNumberOfParams, [FEProps[AToken].Symbol]);
   }
  Result := NewRPNItem;
  Result^.FE.ElementKind := AToken;
  Result^.Fe.FuncName := '';
  Result^.Next := ANext;
end;


{@@
  Creates an entry in the RPN array for an Excel function or operation
  specified by its TokenID (--> TFEKind). Note that array elements for all
  needed parameters must have been created before.

  @param  AToken  Formula element indicating the function to be executed,
                  see the TFEKind enumeration for possible values.
  @param  ANext   Pointer to the next RPN item in the list

  @see TFEKind
}
function RPNFunc(AFuncName: String; ANext: PRPNItem): PRPNItem;
begin
  {
  if FEProps[AToken].MinParams <> FEProps[AToken].MaxParams then
    raise Exception.CreateFmt(lpSpecifyNumberOfParams, [FEProps[AToken].Symbol]);
   }
  Result := RPNFunc(AFuncName, 255, ANext); //FEProps[AToken].MinParams, ANext);
end;

{@@
  Creates an entry in the RPN array for an Excel function or operation
  specified by its TokenID (--> TFEKind). Specify the number of parameters used.
  They must have been created before.

  @param  AToken     Formula element indicating the function to be executed,
                     see the TFEKind enumeration for possible values.
  @param  ANumParams Number of arguments used in the formula. If -1 then the
                     fixed number of arguments known from the function definiton
                     is used.
  @param  ANext      Pointer to the next RPN item in the list

  @see TFEKind
}
function RPNFunc(AFuncName: String; ANumParams: Byte; ANext: PRPNItem): PRPNItem;
begin
   {
  if (ANumParams > -1) then
    if (ANumParams < FEProps[AToken].MinParams) or (ANumParams > FEProps[AToken].MaxParams) then
      raise Exception.CreateFmt(lpIncorrectParamCount, [
        FEProps[AToken].Symbol, FEProps[AToken].MinParams, FEProps[AToken].MaxParams
      ]);
  }
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekFunc;
  Result^.Fe.FuncName := AFuncName;
  Result^.FE.ParamsNum := ANumParams;
  Result^.Next := ANext;
end;
             (*
{@@
  Returns if the function defined by the token requires a fixed number of parameter.

  @param AElementKind  Identifier of the formula function considered
}
function FixedParamCount(AElementKind: TFEKind): Boolean;
begin
  Result := (FEProps[AElementKind].MinParams = FEProps[AElementKind].MaxParams)
        and (FEProps[AElementKind].MinParams >= 0);
end;
               *)
{@@
  Creates an RPN formula by a single call using nested RPN items.

  For each formula element, use one of the RPNxxxx functions implemented here.
  They are designed to be nested into each other. Terminate the chain by using nil.

  @param  AItem     Pointer to the first RPN item representing the formula.
                    Each item contains a pointer to the next item in the list.
                    The list is terminated by nil.
  @param  AReverse  If true the first rpn item in the chained list becomes the
                    last item in the token array. This feature is needed for
                    reading an xls file.

  @example
    The RPN formula for the string expression "$A1+2" can be created as follows:
    <pre>
      var
        f: TsRPNFormula;
      begin
        f := CreateRPNFormula(
          RPNCellValue('$A1',
          RPNNumber(2,
          RPNFunc(fekAdd,
          nil))));
    </pre>
}
function CreateRPNFormula(AItem: PRPNItem; AReverse: Boolean = false): TsRPNFormula;
var
  item: PRPNItem;
  nextitem: PRPNItem;
  n: Integer;
begin
  // Determine count of RPN elements
  n := 0;
  item := AItem;
  while item <> nil do begin
    inc(n);
    item := item^.Next;
  end;

  // Set array length of TsRPNFormula result
  SetLength(Result, n);

  // Copy FormulaElements to result and free temporary RPNItems
  item := AItem;
  if AReverse then n := Length(Result)-1 else n := 0;
  while item <> nil do begin
    nextitem := item^.Next;
    Result[n] := item^.FE;
    if AReverse then dec(n) else inc(n);
    DisposeRPNItem(item);
    item := nextitem;
  end;
end;

{@@
  Destroys the RPN formula starting with the given RPN item.

  @param  AItem  Pointer to the first RPN items representing the formula.
                 Each item contains a pointer to the next item in the list.
                 The list is terminated by nil.
}
procedure DestroyRPNFormula(AItem: PRPNItem);
var
  nextitem: PRPNItem;
begin
  while AItem <> nil do begin
    nextitem := AItem^.Next;
    DisposeRPNItem(AItem);
    AItem := nextitem;
  end;
end;


initialization
  MakeLEPalette(@DEFAULT_PALETTE, Length(DEFAULT_PALETTE));

finalization
  SetLength(GsSpreadFormats, 0);

end.

{ Strategy for handling of number formats:

Problem:
For number formats, fpspreadsheet uses a syntax which is slightly different from
the syntax that Excel uses in the xls files. Moreover, the file syntax can be
different from file type to file type (biff2, for example, allows only a few
predefined formats, while the number of allowed formats is unlimited (?) for
biff8.

Number format handling in fpspreadsheet is implemented with the following
concept in mind:

- Formats written into TsWorksheet cells always follow the fpspreadsheet syntax.

- For writing, the writer creates a TsNumFormatList which stores all formats
  in file syntax.
  - The built-in formats of the file types are coded in the fpc syntax.
  - The method "ConvertBeforeWriting" converts the cell formats from the
    fpspreadsheet to the file syntax.

- For reading, the reader creates another TsNumFormatList.
  - The built-in formats of the file types are coded again in fpc syntax.
  - After reading, the formats are converted to fpc syntax by means of
    "ConvertAfterReading".

- Format conversion is done internally by means of the TsNumFormatParser.
}

