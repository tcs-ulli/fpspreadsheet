{
fpspreadsheet.pas

Writes an spreadsheet document

AUTHORS: Felipe Monteiro de Carvalho
}
unit fpspreadsheet;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils, fpimage, AVL_Tree, avglvltree, lconvencoding;

type
  TsSpreadsheetFormat = (sfExcel2, sfExcel3, sfExcel4, sfExcel5, sfExcel8,
   sfOOXML, sfOpenDocument, sfCSV, sfWikiTable_Pipes, sfWikiTable_WikiMedia);

const
  { Default extensions }
  STR_EXCEL_EXTENSION = '.xls';
  STR_OOXML_EXCEL_EXTENSION = '.xlsx';
  STR_OPENDOCUMENT_CALC_EXTENSION = '.ods';
  STR_COMMA_SEPARATED_EXTENSION = '.csv';
  STR_WIKITABLE_PIPES = '.wikitable_pipes';
  STR_WIKITABLE_WIKIMEDIA = '.wikitable_wikimedia';

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

  {@@ Describes a formula

    Supported syntax:

    =A1+B1+C1/D2...  - Array with simple mathematical operations

    =SUM(A1:D1)      - SUM operation in a interval
  }

  TsFormula = record
    FormulaStr: string;
    DoubleValue: double;
  end;

  {@@ Expanded formula. Used by backend modules. Provides more information than the text only

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
    fekCell, fekCellRef, fekCellRange, fekNum, fekInteger, fekString, fekBool,
    fekErr, fekMissingArg,
    { Basic operations }
    fekAdd, fekSub, fekDiv, fekMul, fekPercent, fekPower, fekUMinus, fekUPlus,
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

  TOperandTokens = fekCell..fekMissingArg;
  TBasicOperationTokens = fekAdd..fekParen;
  TFuncTokens = fekAbs..fekOpSum;

  TsRelFlag = (rfRelRow, rfRelCol, rfRelRow2, rfRelCol2);
  TsRelFlags = set of TsRelFlag;

  TsFormulaElement = record
    ElementKind: TFEKind;
    Row, Row2: Word; // zero-based
    Col, Col2: Word; // zero-based
    Param1, Param2: Word; // Extra parameters
    DoubleValue: double;
    IntValue: Word;
    StringValue: String;
    RelFlags: TsRelFlags;  // store info on relative/absolute addresses
    ParamsNum: Byte;
  end;

  TsExpandedFormula = array of TsFormulaElement;

  {@@ RPN formula. Similar to the expanded formula, but in RPN notation.
      Simplifies the task of format writers which need RPN }
  TsRPNFormula = array of TsFormulaElement;

  {@@ Describes the type of content of a cell on a TsWorksheet }
  TCellContentType = (cctEmpty, cctFormula, cctRPNFormula, cctNumber,
    cctUTF8String, cctDateTime, cctBool, cctError);

  {@@ Error code values }
  TErrorValue = (
    errEmptyIntersection,  // #NULL!
    errDivideByZero,       // #DIV/0!
    errWrongType,          // #VALUE!
    errIllegalRef,         // #REF!
    errWrongName,          // #NAME?
    errOverflow,           // #NUM!
    errArgNotAvail,        // #N/A
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
    nfFixed, nfFixedTh, nfExp, nfSci, nfPercentage,
    // currency
    nfCurrency, nfCurrencyRed, nfAccounting, nfAccountingRed,
    // dates and times
    nfShortDateTime, nfFmtDateTime, nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfTimeInterval,
    // other (format string goes directly into the file)
    nfCustom);

  {@@ Text rotation formatting. The text is rotated relative to the standard
      orientation, which is from left to right horizontal:
       --->
       ABC

      So 90 degrees clockwise means that the text will be:
       |  A
       |  B
       v  C

      And 90 degree counter clockwise will be:

       ^  C
       |  B
       |  A

      Due to limitations of the text mode the characters are not rotated here.
      There is, however, also a "stacked" variant which looks exactly like
      the former case.
  }
  TsTextRotation = (trHorizontal, rt90DegreeClockwiseRotation,
    rt90DegreeCounterClockwiseRotation, rtStacked);

  {@@ Indicates horizontal and vertical text alignment in cells }
  TsHorAlignment = (haDefault, haLeft, haCenter, haRight);
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
  scGrey = $0F;        scGray = $0F;       // redefine to allow different kinds of writing
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

  scNotDefined = $FFFF;

type
  {@@ Data type for rgb color values }
  TsColorValue = DWord;

  {@@ Palette of color values }
  TsPalette = array[0..0] of TsColorValue;
  PsPalette = ^TsPalette;

  {@@ Font style (redefined to avoid usage of "Graphics" }
  TsFontStyle = (fssBold, fssItalic, fssStrikeOut, fssUnderline);
  TsFontStyles = set of TsFontStyle;

  {@@ Font }
  TsFont = class
    FontName: String;
    Size: Single;   // in "points"
    Style: TsFontStyles;
    Color: TsColor;
  end;

  {@@ Indicates the border for a cell }
  TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth);

  {@@ Indicates the border for a cell }
  TsCellBorders = set of TsCellBorder;

  {@@ Line style (for cell borders) }
  TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair);

  {@@ Cell border style }
  TsCellBorderStyle = record
    LineStyle: TsLineStyle;
    Color: TsColor;
  end;

  TsCellBorderStyles = array[TsCellBorder] of TsCellBorderStyle;

const
  DEFAULT_BORDERSTYLES: TsCellBorderStyles = (
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack)
  );

type
  {@@ Cell structure for TsWorksheet

      Never suppose that all *Value fields are valid,
      only one of the ContentTypes is valid. For other fields
      use TWorksheet.ReadAsUTF8Text and similar methods

      @see TWorksheet.ReadAsUTF8Text
  }

  TCell = record
    Col: Cardinal; // zero-based
    Row: Cardinal; // zero-based
    ContentType: TCellContentType;
    { Possible values for the cells }
    FormulaValue: TsFormula;
    RPNFormulaValue: TsRPNFormula;
    NumberValue: double;
    UTF8StringValue: ansistring;
    DateTimeValue: TDateTime;
    BoolValue: Boolean;
    StatusValue: Byte;
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
    Decimals: Byte;
    CurrencySymbol: String;
    RGBBackgroundColor: TFPColor; // only valid if BackgroundColor=scRGBCOLOR
  end;

  PCell = ^TCell;

  TRow = record
    Row: Cardinal;
    Height: Single;       // in millimeters
  end;

  PRow = ^TRow;

  TCol = record
    Col: Cardinal;
    Width: Single; // in "characters". Excel uses the with of char "0" in 1st font
  end;

  PCol = ^TCol;

  TsSheetOption = (soShowGridLines, soShowHeaders, soHasFrozenPanes, soSelected);
  TsSheetOptions = set of TsSheetOption;

type

  TsCustomSpreadReader = class;
  TsCustomSpreadWriter = class;
  TsWorkbook = class;


  { TsWorksheet }

  TsCellEvent = procedure (Sender: TObject; ARow, ACol: Cardinal) of object;

  TsWorksheet = class
  private
    FWorkbook: TsWorkbook;
    FCells: TAvlTree; // Items are TCell
    FCurrentNode: TAVLTreeNode; // For GetFirstCell and GetNextCell
    FRows, FCols: TIndexedAVLTree; // This lists contain only rows or cols with styles different from the standard
    FLeftPaneWidth: Integer;
    FTopPaneHeight: Integer;
    FOptions: TsSheetOptions;
    FOnChangeCell: TsCellEvent;
    FOnChangeFont: TsCellEvent;
    procedure RemoveCallback(data, arg: pointer);

  protected
    procedure ChangedCell(ARow, ACol: Cardinal);
    procedure ChangedFont(ARow, ACol: Cardinal);

  public
    Name: string;

    { Base methods }
    constructor Create;
    destructor Destroy; override;

    { Utils }
    class function CellPosToText(ARow, ACol: Cardinal): string;

    { Data manipulation methods - For Cells }
    procedure CopyCell(AFromRow, AFromCol, AToRow, AToCol: Cardinal; AFromWorksheet: TsWorksheet);
    procedure CopyFormat(AFormat: PCell; AToRow, AToCol: Cardinal); overload;
    procedure CopyFormat(AFromCell, AToCell: PCell); overload;
    function  FindCell(ARow, ACol: Cardinal): PCell;
    function  GetCell(ARow, ACol: Cardinal): PCell;
    function  GetCellCount: Cardinal;
    function  GetFirstCell(): PCell;
    function  GetNextCell(): PCell;
    function  GetFirstCellOfRow(ARow: Cardinal): PCell;
    function  GetLastCellOfRow(ARow: Cardinal): PCell;
    function  GetLastColNumber: Cardinal;
    function  GetLastRowNumber: Cardinal;
    function  ReadAsUTF8Text(ARow, ACol: Cardinal): ansistring; overload;
    function  ReadAsUTF8Text(ACell: PCell): ansistring; overload;
    function  ReadAsNumber(ARow, ACol: Cardinal): Double;
    function  ReadAsDateTime(ARow, ACol: Cardinal; out AResult: TDateTime): Boolean;
    function  ReadRPNFormulaAsString(ACell: PCell): String;
    function  ReadUsedFormatting(ARow, ACol: Cardinal): TsUsedFormattingFields;
    function  ReadBackgroundColor(ARow, ACol: Cardinal): TsColor;

    procedure RemoveAllCells;

    { Writing of values }
    procedure WriteBlank(ARow, ACol: Cardinal);
    procedure WriteBoolValue(ARow, ACol: Cardinal; AValue: Boolean);
    procedure WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
      AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = ''); overload;
    procedure WriteDateTime(ACell: PCell; AValue: TDateTime;
      AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = ''); overload;
    procedure WriteErrorValue(ARow, ACol: Cardinal; AValue: TErrorValue); overload;
    procedure WriteErrorValue(ACell: PCell; AValue: TErrorValue); overload;
    procedure WriteFormula(ARow, ACol: Cardinal; AFormula: TsFormula);
    procedure WriteNumber(ARow, ACol: Cardinal; ANumber: double;
      AFormat: TsNumberFormat = nfGeneral; ADecimals: Byte = 2;
      ACurrencySymbol: String = ''); overload;
    procedure WriteNumber(ACell: PCell; ANumber: Double; AFormat: TsNumberFormat = nfGeneral;
      ADecimals: Byte = 2; ACurrencySymbol: String = ''); overload;
    procedure WriteNumber(ARow, ACol: Cardinal; ANumber: double;
      AFormatString: String); overload;
    procedure WriteRPNFormula(ARow, ACol: Cardinal; AFormula: TsRPNFormula);
    procedure WriteUTF8Text(ARow, ACol: Cardinal; AText: ansistring);

    { Writing of cell attributes }
    procedure WriteBackgroundColor(ARow, ACol: Cardinal; AColor: TsColor);

    procedure WriteBorderColor(ARow, ACol: Cardinal; ABorder: TsCellBorder; AColor: TsColor);
    procedure WriteBorderLineStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle);
    procedure WriteBorders(ARow, ACol: Cardinal; ABorders: TsCellBorders);
    procedure WriteBorderStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      AStyle: TsCellBorderStyle); overload;
    procedure WriteBorderStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle; AColor: TsColor); overload;
    procedure WriteBorderStyles(ARow, ACol: Cardinal; const AStyles: TsCellBorderStyles);

    procedure WriteDecimals(ARow, ACol: Cardinal; ADecimals: byte); overload;
    procedure WriteDecimals(ACell: PCell; ADecimals: Byte); overload;

    function  WriteFont(ARow, ACol: Cardinal; const AFontName: String;
      AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer; overload;
    procedure WriteFont(ARow, ACol: Cardinal; AFontIndex: Integer); overload;
    function WriteFontColor(ARow, ACol: Cardinal; AFontColor: TsColor): Integer;
    function WriteFontName(ARow, ACol: Cardinal; AFontName: String): Integer;
    function WriteFontSize(ARow, ACol: Cardinal; ASize: Single): Integer;
    function WriteFontStyle(ARow, ACol: Cardinal; AStyle: TsFontStyles): Integer;

    procedure WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment);

    procedure WriteNumberFormat(ARow, ACol: Cardinal; ANumberFormat: TsNumberFormat;
      const AFormatString: String = '');

    procedure WriteTextRotation(ARow, ACol: Cardinal; ARotation: TsTextRotation);

    procedure WriteUsedFormatting(ARow, ACol: Cardinal; AUsedFormatting: TsUsedFormattingFields);

    procedure WriteVertAlignment(ARow, ACol: Cardinal; AValue: TsVertAlignment);

    procedure WriteWordwrap(ARow, ACol: Cardinal; AValue: boolean);

    { Data manipulation methods - For Rows and Cols }
    function  FindRow(ARow: Cardinal): PRow;
    function  FindCol(ACol: Cardinal): PCol;
    function  GetRow(ARow: Cardinal): PRow;
    function  GetCol(ACol: Cardinal): PCol;
    procedure RemoveAllRows;
    procedure RemoveAllCols;
    procedure WriteRowInfo(ARow: Cardinal; AData: TRow);
    procedure WriteRowHeight(ARow: Cardinal; AHeight: Single);
    procedure WriteColInfo(ACol: Cardinal; AData: TCol);
    procedure WriteColWidth(ACol: Cardinal; AWidth: Single);

    { Properties }
    property  Cells: TAVLTree read FCells;
    property  Cols: TIndexedAVLTree read FCols;
    property  Rows: TIndexedAVLTree read FRows;
    property  Workbook: TsWorkbook read FWorkbook;

    // These are properties to interface to fpspreadsheetgrid.
    property  Options: TsSheetOptions read FOptions write FOptions;
    property  LeftPaneWidth: Integer read FLeftPaneWidth write FLeftPaneWidth;
    property  TopPaneHeight: Integer read FTopPaneHeight write FTopPaneHeight;
    property  OnChangeCell: TsCellEvent read FOnChangeCell write FOnChangeCell;
    property  OnChangeFont: TsCellEvent read FOnChangeFont write FOnChangeFont;
  end;


  { TsWorkbook }

  TsWorkbook = class
  private
    { Internal data }
    FWorksheets: TFPList;
    FEncoding: TsEncoding;
    FFormat: TsSpreadsheetFormat;
    FFontList: TFPList;
    FBuiltinFontCount: Integer;
    FPalette: array of TsColorValue;
    FReadFormulas: Boolean;
    { Internal methods }
    procedure RemoveWorksheetsCallback(data, arg: pointer);
  public
    FormatSettings: TFormatSettings;
    { Base methods }
    constructor Create;
    destructor Destroy; override;
    class function GetFormatFromFileName(const AFileName: TFileName; var SheetType: TsSpreadsheetFormat): Boolean;
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
    function FindFont(const AFontName: String; ASize: Single;
      AStyle: TsFontStyles; AColor: TsColor): Integer;
    function GetFont(AIndex: Integer): TsFont;
    function GetFontCount: Integer;
    procedure InitFonts;
    procedure RemoveAllFonts;
    procedure SetDefaultFont(const AFontName: String; ASize: Single);
    { Color handling }
    function FPSColorToHexString(AColor: TsColor; ARGBColor: TFPColor): String;
    function GetColorName(AColorIndex: TsColor): string;
    function GetPaletteColor(AColorIndex: TsColor): TsColorValue;
    procedure SetPaletteColor(AColorIndex: TsColor; AColorValue: TsColorValue);
    function GetPaletteSize: Integer;
    procedure UsePalette(APalette: PsPalette; APaletteCount: Word;
      ABigEndian: Boolean = false);
    {@@ This property is only used for formats which don't support unicode
      and support a single encoding for the whole document, like Excel 2 to 5 }
    property Encoding: TsEncoding read FEncoding write FEncoding;
    property FileFormat: TsSpreadsheetFormat read FFormat;
    property ReadFormulas: Boolean read FReadFormulas write FReadFormulas;
  end;


  {@@ Contents of the format record  }

  TsNumFormatData = class
  public
    Index: Integer;
    NumFormat: TsNumberFormat;
    Decimals: Byte;
    CurrencySymbol: String;
    FormatString: string;
  end;

  {@@ Specialized list for number format items }

  TsCustomNumFormatList = class(TFPList)
  private
    function GetItem(AIndex: Integer): TsNumFormatData;
    procedure SetItem(AIndex: Integer; AValue: TsNumFormatData);
  protected
    FWorkbook: TsWorkbook;
    FFirstFormatIndexInFile: Integer;
    FNextFormatIndex: Integer;
    procedure AddBuiltinFormats; virtual;
    procedure RemoveFormat(AIndex: Integer);
  public
    constructor Create(AWorkbook: TsWorkbook);
    destructor Destroy; override;
    function AddFormat(AFormatCell: PCell): Integer; overload;
    function AddFormat(AFormatIndex: Integer; AFormatString: String;
      ANumFormat: TsNumberFormat; ADecimals: Byte = 0;
      ACurrencySymbol: String = ''): Integer; overload;
    function AddFormat(AFormatString: String; ANumFormat: TsNumberFormat;
      ADecimals: Byte = 0; ACurrencySymbol: String = ''): Integer; overload;
    procedure AnalyzeAndAdd(AFormatIndex: Integer; AFormatString: String);
    procedure Clear;
    procedure ConvertAfterReading(AFormatIndex: Integer; var AFormatString: String;
      var ANumFormat: TsNumberFormat; var ADecimals: Byte;
      var ACurrencySymbol: String); virtual;
    procedure ConvertBeforeWriting(var AFormatString: String;
      var ANumFormat: TsNumberFormat; var ADecimals: Byte;
      var ACurrencySymbol: String); virtual;
    procedure Delete(AIndex: Integer);
    function Find(ANumFormat: TsNumberFormat; AFormatString: String;
      ADecimals: Byte; ACurrencySymbol: String): Integer; overload;
    function Find(AFormatIndex: Integer): Integer; overload;
    function Find(AFormatString: String): Integer; overload;
    function FindFormatOf(AFormatCell: PCell): integer; virtual;
    function FormatStringForWriting(AIndex: Integer): String; virtual;
    procedure Sort;

    property Workbook: TsWorkbook read FWorkbook;
    property FirstFormatIndexInFile: Integer read FFirstFormatIndexInFile;
    property Items[AIndex: Integer]: TsNumFormatData read GetItem write SetItem; default;
  end;

  {@@ TsSpreadReader class reference type }

  TsSpreadReaderClass = class of TsCustomSpreadReader;
  
  { TsCustomSpreadReader }

  TsCustomSpreadReader = class
  protected
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FNumFormatList: TsCustomNumFormatList;
    procedure CreateNumFormatList; virtual;
    { Record reading methods }
    procedure ReadBlank(AStream: TStream); virtual; abstract;
    procedure ReadFormula(AStream: TStream); virtual; abstract;
    procedure ReadLabel(AStream: TStream); virtual; abstract;
    procedure ReadNumber(AStream: TStream); virtual; abstract;
  public
    constructor Create(AWorkbook: TsWorkbook); virtual; // To allow descendents to override it
    destructor Destroy; override;
    { General writing methods }
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); virtual;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); virtual;
    procedure ReadFromStrings(AStrings: TStrings; AData: TsWorkbook); virtual;
    property Workbook: TsWorkbook read FWorkbook;
    property NumFormatList: TsCustomNumFormatList read FNumFormatList;
  end;

  {@@ TsSpreadWriter class reference type }

  TsSpreadWriterClass = class of TsCustomSpreadWriter;

  TCellsCallback = procedure (ACell: PCell; AStream: TStream) of object;

  { TsCustomSpreadWriter }

  TsCustomSpreadWriter = class
  private
    FWorkbook: TsWorkbook;
  protected
    FNumFormatList: TsCustomNumFormatList;
    { Helper routines }
    procedure AddDefaultFormats(); virtual;
    procedure CreateNumFormatList; virtual;
    function  ExpandFormula(AFormula: TsFormula): TsExpandedFormula;
    function  FindFormattingInList(AFormat: PCell): Integer;
    procedure FixFormat(ACell: PCell); virtual;
    procedure ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
    procedure ListAllFormattingStyles; virtual;
    procedure ListAllNumFormatsCallback(ACell: PCell; AStream: TStream);
    procedure ListAllNumFormats; virtual;
    { Helpers for writing }
    procedure WriteCellCallback(ACell: PCell; AStream: TStream);
    procedure WriteCellsToStream(AStream: TStream; ACells: TAVLTree);
    { Record writing methods }
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal; ACell: PCell); virtual; abstract;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell); virtual; abstract;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal; const AFormula: TsFormula; ACell: PCell); virtual;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell); virtual;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal; const AValue: string; ACell: PCell); virtual; abstract;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double; ACell: PCell); virtual; abstract;
  public
    {@@
    An array with cells which are models for the used styles
    In this array the Row property holds the Index to the corresponding XF field
    }
    FFormattingStyles: array of TCell;
    NextXFIndex: Integer; // Indicates which should be the next XF (Style) Index when filling the styles list
    constructor Create(AWorkbook: TsWorkbook); virtual; // To allow descendents to override it
    destructor Destroy; override;
    { General writing methods }
    procedure IterateThroughCells(AStream: TStream; ACells: TAVLTree; ACallback: TCellsCallback);
    procedure WriteToFile(const AFileName: string; const AOverwriteExisting: Boolean = False); virtual;
    procedure WriteToStream(AStream: TStream); virtual;
    procedure WriteToStrings(AStrings: TStrings); virtual;
    property Workbook: TsWorkbook read FWorkbook;
    property NumFormatList: TsCustomNumFormatList read FNumFormatList;
  end;

  {@@ List of registered formats }

  TsSpreadFormatData = record
    ReaderClass: TsSpreadReaderClass;
    WriterClass: TsSpreadWriterClass;
    Format: TsSpreadsheetFormat;
  end;

  {@@ Helper for simplification of RPN formula creation }
  PRPNItem = ^TRPNItem;
  TRPNItem = record
    FE: TsFormulaElement;
    Next: PRPNItem;
  end;

  {@@
    Simple creation an RPNFormula array to be used in fpspreadsheet.
    For each formula element, use one of the RPNxxxx functions implemented here.
    They are designed to be nested into each other. Terminate the chain by
    using nil.

    Example:
    The RPN formula for the string expression "$A1+2" can be created as follows:

      var
        f: TsRPNFormula;

        f := CreateRPNFormula(
          RPNCellValue('A1',
          RPNNumber(2,
          RPNFunc(fekAdd,
          nil))));
  }

  function CreateRPNFormula(AItem: PRPNItem): TsRPNFormula;
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
  function RPNErr(AErrCode: Byte; ANext: PRPNItem): PRPNItem;
  function RPNInteger(AValue: Word; ANext: PRPNItem): PRPNItem;
  function RPNMissingArg(ANext: PRPNItem): PRPNItem;
  function RPNNumber(AValue: Double; ANext: PRPNItem): PRPNItem;
  function RPNParenthesis(ANext: PRPNItem): PRPNItem;
  function RPNString(AValue: String; ANext: PRPNItem): PRPNItem;
  function RPNFunc(AToken: TFEKind; ANext: PRPNItem): PRPNItem; overload;
  function RPNFunc(AToken: TFEKind; ANumParams: Byte; ANext: PRPNItem): PRPNItem; overload;

  function FixedParamCount(AElementKind: TFEKind): Boolean;

var
  GsSpreadFormats: array of TsSpreadFormatData;

procedure RegisterSpreadFormat(
  AReaderClass: TsSpreadReaderClass;
  AWriterClass: TsSpreadWriterClass;
  AFormat: TsSpreadsheetFormat);

procedure CopyCellFormat(AFromCell, AToCell: PCell);
function GetFileFormatName(AFormat: TsSpreadsheetFormat): String;
procedure MakeLEPalette(APalette: PsPalette; APaletteSize: Integer);

implementation

uses
  Math, StrUtils, TypInfo, fpsUtils, fpsNumFormatParser;

{ Translatable strings }
resourcestring
  lpUnsupportedReadFormat = 'Tried to read a spreadsheet using an unsupported format';
  lpUnsupportedWriteFormat = 'Tried to write a spreadsheet using an unsupported format';
  lpNoValidSpreadsheetFile = '"%s" is not a valid spreadsheet file';
  lpUnknownSpreadsheetFormat = 'unknown format';
  lpInvalidFontIndex = 'Invalid font index';
  lpInvalidNumberFormat = 'Trying to use an incompatible number format.';
  lpNoValidNumberFormatString = 'No valid number format string.';
  lpNoValidDateTimeFormatString = 'No valid date/time format string.';
  lpIllegalNumberFormat = 'Illegal number format.';
  lpSpecifyNumberOfParams = 'Specify number of parameters for function %s';
  lpIncorrectParamCount = 'Funtion %s requires at least %d and at most %d parameters.';
  lpTRUE = 'TRUE';
  lpFALSE = 'FALSE';
  lpErrEmptyIntersection = '#NULL!';
  lpErrDivideByZero = '#DIV/0!';
  lpErrWrongType = '#VALUE!';
  lpErrIllegalRef = '#REF!';
  lpErrWrongName = '#NAME?';
  lpErrOverflow = '#NUM!';
  lpErrArgNotAvail = '#N/A';
  lpErrFormulaNotSupported = '<FORMULA?>';

var
  {@@
    Colors in RGB in "big-endian" notation (red at left). The values are inverted
    at initialization to be little-endian at run-time!
    The indices into this palette are named as scXXXX color constants.
  }
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


{ Properties of formula elements }

type
  TFEProp = record Symbol: String; MinParams, MaxParams: Byte; end;

const
  FEProps: array[TFEKind] of TFEProp = (
  { Operands }
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCell
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellRef
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellRange
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellNum
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellInteger
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellString
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellBool
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellErr
    (Symbol:'';          MinParams:-1; MaxParams:-1), // fekCellMissingArg
  { Basic operations }
    (Symbol:'+';         MinParams:2; MaxParams:2),   // fekAdd
    (Symbol:'-';         MinParams:2; MaxParams:2),   // fekSub
    (Symbol:'*';         MinParams:2; MaxParams:2),   // fekDiv
    (Symbol:'/';         MinParams:2; MaxParams:2),   // fekMul
    (Symbol:'%';         MinParams:1; MaxParams:1),   // fekPercent
    (Symbol:'^';         MinParams:2; MaxParams:2),   // fekPower
    (Symbol:'-';         MinParams:1; MaxParams:1),   // fekUMinus
    (Symbol:'+';         MinParams:1; MaxParams:1),   // fekUPlus
    (Symbol:'&';         MinParams:2; MaxParams:2),   // fekConcat (string concatenation)
    (Symbol:'=';         MinParams:2; MaxParams:2),   // fekEqual
    (Symbol:'>';         MinParams:2; MaxParams:2),   // fekGreater
    (Symbol:'>=';        MinParams:2; MaxParams:2),   // fekGreaterEqual
    (Symbol:'<';         MinParams:2; MaxParams:2),   // fekLess
    (Symbol:'<=';        MinParams:2; MaxParams:2),   // fekLessEqual
    (Symbol:'<>';        MinParams:2; MaxParams:2),   // fekNotEqual
    (Symbol:'';          MinParams:1; MaxParams:1),   // fekParen
  { math }
    (Symbol:'ABS';       MinParams:1; MaxParams:1),   // fekABS
    (Symbol:'ACOS';      MinParams:1; MaxParams:1),   // fekACOS
    (Symbol:'ACOSH';     MinParams:1; MaxParams:1),   // fekACOSH
    (Symbol:'ASIN';      MinParams:1; MaxParams:1),   // fekASIN
    (Symbol:'ASINH';     MinParams:1; MaxParams:1),   // fekASINH
    (Symbol:'ATAN';      MinParams:1; MaxParams:1),   // fekATAN
    (Symbol:'ATANH';     MinParams:1; MaxParams:1),   // fekATANH,
    (Symbol:'COS';       MinParams:1; MaxParams:1),   // fekCOS
    (Symbol:'COSH';      MinParams:1; MaxParams:1),   // fekCOSH
    (Symbol:'DEGREES';   MinParams:1; MaxParams:1),   // fekDEGREES
    (Symbol:'EXP';       MinParams:1; MaxParams:1),   // fekEXP
    (Symbol:'INT';       MinParams:1; MaxParams:1),   // fekINT
    (Symbol:'LN';        MinParams:1; MaxParams:1),   // fekLN
    (Symbol:'LOG';       MinParams:1; MaxParams:2),   // fekLOG,
    (Symbol:'LOG10';     MinParams:1; MaxParams:1),   // fekLOG10
    (Symbol:'PI';        MinParams:0; MaxParams:0),   // fekPI
    (Symbol:'RADIANS';   MinParams:1; MaxParams:1),   // fekRADIANS
    (Symbol:'RAND';      MinParams:0; MaxParams:0),   // fekRAND
    (Symbol:'ROUND';     MinParams:2; MaxParams:2),   // fekROUND,
    (Symbol:'SIGN';      MinParams:1; MaxParams:1),   // fekSIGN
    (Symbol:'SIN';       MinParams:1; MaxParams:1),   // fekSIN
    (Symbol:'SINH';      MinParams:1; MaxParams:1),   // fekSINH
    (Symbol:'SQRT';      MinParams:1; MaxParams:1),   // fekSQRT,
    (Symbol:'TAN';       MinParams:1; MaxParams:1),   // fekTAN
    (Symbol:'TANH';      MinParams:1; MaxParams:1),   // fekTANH,
  { date/time }
    (Symbol:'DATE';      MinParams:3; MaxParams:3),   // fekDATE
    (Symbol:'DATEDIF';   MinParams:3; MaxParams:3),   // fekDATEDIF
    (Symbol:'DATEVALUE'; MinParams:1; MaxParams:1),   // fekDATEVALUE
    (Symbol:'DAY';       MinParams:1; MaxParams:1),   // fekDAY
    (Symbol:'HOUR';      MinParams:1; MaxParams:1),   // fekHOUR
    (Symbol:'MINUTE';    MinParams:1; MaxParams:1),   // fekMINUTE
    (Symbol:'MONTH';     MinParams:1; MaxParams:1),   // fekMONTH
    (Symbol:'NOW';       MinParams:0; MaxParams:0),   // fekNOW
    (Symbol:'SECOND';    MinParams:1; MaxParams:1),   // fekSECOND
    (Symbol:'TIME';      MinParams:3; MaxParams:3),   // fekTIME
    (Symbol:'TIMEVALUE'; MinParams:1; MaxParams:1),   // fekTIMEVALUE
    (Symbol:'TODAY';     MinParams:0; MaxParams:0),   // fekTODAY
    (Symbol:'WEEKDAY';   MinParams:1; MaxParams:2),   // fekWEEKDAY
    (Symbol:'YEAR';      MinParams:1; MaxParams:1),   // fekYEAR
  { statistical }
    (Symbol:'AVEDEV';    MinParams:1; MaxParams:30),  // fekAVEDEV
    (Symbol:'AVERAGE';   MinParams:1; MaxParams:30),  // fekAVERAGE
    (Symbol:'BETADIST';  MinParams:3; MaxParams:5),   // fekBETADIST
    (Symbol:'BETAINV';   MinParams:3; MaxParams:5),   // fekBETAINV
    (Symbol:'BINOMDIST'; MinParams:4; MaxParams:4),   // fekBINOMDIST
    (Symbol:'CHIDIST';   MinParams:2; MaxParams:2),   // fekCHIDIST
    (Symbol:'CHIINV';    MinParams:2; MaxParams:2),   // fekCHIINV
    (Symbol:'COUNT';     MinParams:0; MaxParams:30),  // fekCOUNT
    (Symbol:'COUNTA';    MinParams:0; MaxParams:30),  // fekCOUNTA
    (Symbol:'COUNTBLANK';MinParams:1; MaxParams:1),   // fekCOUNTBLANK
    (Symbol:'COUNTIF';   MinParams:2; MaxParams:2),   // fekCOUNTIF
    (Symbol:'MAX';       MinParams:1; MaxParams:30),  // fekMAX
    (Symbol:'MEDIAN';    MinParams:1; MaxParams:30),  // fekMEDIAN
    (Symbol:'MIN';       MinParams:1; MaxParams:30),  // fekMIN
    (Symbol:'PERMUT';    MinParams:2; MaxParams:2),   // fekPERMUT
    (Symbol:'POISSON';   MinParams:3; MaxParams:3),   // fekPOISSON
    (Symbol:'PRODUCT';   MinParams:0; MaxParams:30),  // fekPRODUCT
    (Symbol:'STDEV';     MinParams:1; MaxParams:30),  // fekSTDEV
    (Symbol:'STDEVP';    MinParams:1; MaxParams:30),  // fekSTDEVP
    (Symbol:'SUM';       MinParams:0; MaxParams:30),  // fekSUM
    (Symbol:'SUMIF';     MinParams:2; MaxParams:3),   // fekSUMIF
    (Symbol:'SUMSQ';     MinParams:0; MaxParams:30),  // fekSUMSQ
    (Symbol:'VAR';       MinParams:1; MaxParams:30),  // fekVAR
    (Symbol:'VARP';      MinParams:1; MaxParams:30),  // fekVARP
  { financial }
    (Symbol:'FV';        MinParams:3; MaxParams:5),   // fekFV
    (Symbol:'NPER';      MinParams:3; MaxParams:5),   // fekNPER
    (Symbol:'PMT';       MinParams:3; MaxParams:5),   // fekPMT
    (Symbol:'PV';        MinParams:3; MaxParams:5),   // fekPV
    (Symbol:'RATE';      MinParams:3; MaxParams:6),   // fekRATE
  { logical }
    (Symbol:'AND';       MinParams:0; MaxParams:30),  // fekAND
    (Symbol:'FALSE';     MinParams:0; MaxParams:0),   // fekFALSE
    (Symbol:'IF';        MinParams:2; MaxParams:3),   // fekIF
    (Symbol:'NOT';       MinParams:1; MaxParams:1),   // fekNOT
    (Symbol:'OR';        MinParams:1; MaxParams:30),  // fekOR
    (Symbol:'TRUE';      MinParams:0; MaxParams:0),   // fekTRUE
  {  string }
    (Symbol:'CHAR';      MinParams:1; MaxParams:1),   // fekCHAR
    (Symbol:'CODE';      MinParams:1; MaxParams:1),   // fekCODE
    (Symbol:'LEFT';      MinParams:1; MaxParams:2),   // fekLEFT
    (Symbol:'LOWER';     MinParams:1; MaxParams:1),   // fekLOWER
    (Symbol:'MID';       MinParams:3; MaxParams:3),   // fekMID
    (Symbol:'PROPER';    MinParams:1; MaxParams:1),   // fekPROPER
    (Symbol:'REPLACE';   MinParams:4; MaxParams:4),   // fekREPLACE
    (Symbol:'RIGHT';     MinParams:1; MaxParams:2),   // fekRIGHT
    (Symbol:'SUBSTITUTE';MinParams:3; MaxParams:4),   // fekSUBSTITUTE
    (Symbol:'TRIM';      MinParams:1; MaxParams:1),   // fekTRIM
    (Symbol:'UPPER';     MinParams:1; MaxParams:1),   // fekUPPER
  {  lookup/reference }
    (Symbol:'COLUMN';    MinParams:0; MaxParams:1),   // fekCOLUMN
    (Symbol:'COLUMNS';   MinParams:1; MaxParams:1),   // fekCOLUMNS
    (Symbol:'ROW';       MinParams:0; MaxParams:1),   // fekROW
    (Symbol:'ROWS';      MinParams:1; MaxParams:1),   // fekROWS
  { info }
    (Symbol:'CELL';      MinParams:1; MaxParams:2),   // fekCELLINFO
    (Symbol:'INFO';      MinParams:1; MaxParams:1),   // fekINFO
    (Symbol:'ISBLANK';   MinParams:1; MaxParams:1),   // fekIsBLANK
    (Symbol:'ISERR';     MinParams:1; MaxParams:1),   // fekIsERR
    (Symbol:'ISERROR';   MinParams:1; MaxParams:1),   // fekIsERROR
    (Symbol:'ISLOGICAL'; MinParams:1; MaxParams:1),   // fekIsLOGICAL
    (Symbol:'ISNA';      MinParams:1; MaxParams:1),   // fekIsNA
    (Symbol:'ISNONTEXT'; MinParams:1; MaxParams:1),   // fekIsNONTEXT
    (Symbol:'ISNUMBER';  MinParams:1; MaxParams:1),   // fekIsNUMBER
    (Symbol:'ISREF';     MinParams:1; MaxParams:1),   // fekIsRef
    (Symbol:'ISTEXT';    MinParams:1; MaxParams:1),   // fekIsTEXT
    (Symbol:'VALUE';     MinParams:1; MaxParams:1),   // fekValue
  { Other operations }
    (Symbol:'SUM';       MinParams:1; MaxParams:1)    // fekOpSUM (Unary sum operation). Note: CANNOT be used for summing sell contents; use fekSUM}
  );

{@@
  Registers a new reader/writer pair for a format
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
  Returns the name of the given file format.
}
function GetFileFormatName(AFormat: TsSpreadsheetFormat): string;
begin
  case AFormat of
    sfExcel2              : Result := 'BIFF2';
    sfExcel3              : Result := 'BIFF3';
    sfExcel4              : Result := 'BIFF4';
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
}
procedure MakeLEPalette(APalette: PsPalette; APaletteSize: Integer);
var
  i: Integer;
begin
  for i := 0 to APaletteSize-1 do
   {$IFDEF RNGCHECK}
    {$R-}
   {$ENDIF}
    APalette^[i] := LongRGBToExcelPhysical(APalette^[i])
   {$IFDEF RNGCHECK}
    {$R+}
   {$ENDIF}
end;

{@@
  Copies the format of a cell to another one.
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
  AToCell^.Decimals := AFromCell^.Decimals;
  AToCell^.CurrencySymbol := AFromCell^.CurrencySymbol;
end;


{ TsWorksheet }

{@@
  Helper method for clearing the records in a spreadsheet.
}
procedure TsWorksheet.RemoveCallback(data, arg: pointer);
begin
  { The strings and dyn arrays must be reset to nil content manually, because
    FreeMem only frees the record mem, without checking its content }
  PCell(data).UTF8StringValue := '';
  PCell(data).NumberFormatStr := '';
  SetLength(PCell(data).RPNFormulaValue, 0);
  FreeMem(data);
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
  Constructor.
}
constructor TsWorksheet.Create;
begin
  inherited Create;

  FCells := TAVLTree.Create(@CompareCells);
  FRows := TIndexedAVLTree.Create(@CompareRows);
  FCols := TIndexedAVLTree.Create(@CompareCols);

  FOptions := [soShowGridLines, soShowHeaders];
end;

{@@
  Destructor.
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

{@@ Converts a FPSpreadsheet cell position, which is Row, Col in numbers
 and zero based, to a textual representation which is [Col][Row],
 being that the Col is in letters and the row is in 1-based numbers }
class function TsWorksheet.CellPosToText(ARow, ACol: Cardinal): string;
var
  lStr: string;
begin
  lStr := '';
  if ACol < 26 then lStr := Char(ACol+65);

  Result := Format('%s%d', [lStr, ARow+1]);
end;

{ Is called whenever a cell value or formatting has changed. }
procedure TsWorksheet.ChangedCell(ARow, ACol: Cardinal);
begin
  if Assigned(FOnChangeCell) then FOnChangeCell(Self, ARow, ACol);
end;

{ Is called whenever a font height changes. Event can be caught by the grid
  to update the row height. }
procedure TsWorksheet.ChangedFont(ARow, ACol: Cardinal);
begin
  if Assigned(FonChangeFont) then FOnChangeFont(Self, ARow, ACol);
end;

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
  {
  lCurStr := AFromWorksheet.ReadAsUTF8Text(AFromRow, AFromCol);
  lCurUsedFormatting := AFromWorksheet.ReadUsedFormatting(AFromRow, AFromCol);
  lCurColor := AFromWorksheet.ReadBackgroundColor(AFromRow, AFromCol);
  WriteUTF8Text(AToRow, AToCol, lCurStr);
  WriteUsedFormatting(AToRow, AToCol, lCurUsedFormatting);
  if uffBackgroundColor in lCurUsedFormatting then
  begin
    WriteBackgroundColor(AToRow, AToCol, lCurColor);
  end;
  }
end;

{@@
  Copies all format parameters from the format cell to another cell.
}
procedure TsWorksheet.CopyFormat(AFromCell, AToCell: PCell);
begin
  if (AFromCell = nil) or (AToCell = nil) then
    exit;

  CopyCellFormat(AFromCell, AToCell);
  ChangedCell(AToCell^.Row, AToCell^.Col);
  ChangedFont(AToCell^.Row, AToCell^.Col);
end;

procedure TsWorksheet.CopyFormat(AFormat: PCell; AToRow, AToCol: Cardinal);
begin
  CopyFormat(AFormat, GetCell(AToRow, AToCol));
end;

{@@
  Tries to locate a Cell in the list of already
  written Cells

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell

  @return Nil if no existing cell was found,
          otherwise a pointer to the desired Cell

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
  Obtains an allocated cell at the desired location.

  If the Cell already exists, a pointer to it will
  be returned.

  If not, then new memory for the cell will be allocated,
  a pointer to it will be returned and it will be added
  to the list of Cells.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell

  @return A pointer to the Cell on the desired location.

  @see    TCell
}
function TsWorksheet.GetCell(ARow, ACol: Cardinal): PCell;
begin
  Result := FindCell(ARow, ACol);
  
  if (Result = nil) then
  begin
    Result := GetMem(SizeOf(TCell));
    FillChar(Result^, SizeOf(TCell), #0);

    Result^.Row := ARow;
    Result^.Col := ACol;
    Result^.BorderStyles := DEFAULT_BORDERSTYLES;

    Cells.Add(Result);
  end;
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
  Returns the 0-based number of the last column with a cell with contents.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @see GetCellCount
}
function TsWorksheet.GetLastColNumber: Cardinal;
var
  AVLNode: TAVLTreeNode;
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
end;

function TsWorksheet.GetFirstCellOfRow(ARow: Cardinal): PCell;
var
  c, n: Cardinal;
begin
  n := GetLastColNumber;
  c := 0;
  Result := FindCell(ARow, c);
  while (result = nil) and (c < n) do begin
    inc(c);
    result := FindCell(ARow, c);
  end;
end;

function TsWorksheet.GetLastCellOfRow(ARow: Cardinal): PCell;
var
  c, n: Cardinal;
begin
  n := GetLastColNumber;
  c := n;
  Result := FindCell(ARow, c);
  while (Result = nil) and (c > 0) do begin
    dec(c);
    Result := FindCell(ARow, c);
  end;
end;

{@@
  Returns the 0-based number of the last row with a cell with contents.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @see GetCellCount
}
function TsWorksheet.GetLastRowNumber: Cardinal;
var
  AVLNode: TAVLTreeNode;
begin
  Result := 0;

  AVLNode := FCells.FindHighest;
  if Assigned(AVLNode) then
    Result := PCell(AVLNode.Data).Row;
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

function TsWorksheet.ReadAsUTF8Text(ACell: PCell): ansistring;

  function FloatToStrNoNaN(const Value: Double;
    ANumberFormat: TsNumberFormat; ANumberFormatStr: string; ADecimals: byte): ansistring;
  var
    fs: TFormatSettings;
    left, right: String;
    i: Integer;
  begin
    fs := FWorkbook.FormatSettings;
    if IsNan(Value) then
      Result := ''
    else
    if ANumberFormat = nfSci then
      Result := SciFloat(Value, ADecimals)
    else
    if (ANumberFormat = nfGeneral) or (ANumberFormatStr = '') then
      Result := FloatToStr(Value, fs)
    else
    if (ANumberFormat = nfPercentage) then
      Result := FormatFloat(ANumberFormatStr, Value*100, fs)
    else
    if (ANumberFormat in [nfAccounting, nfAccountingRed]) then
      case SplitAccountingFormatString(ANumberFormatStr, Sign(Value), left, right) of
        0: Result := FormatFloat(ANumberFormatStr, Value, fs);
        1: Result := FormatFloat(left, abs(Value), fs) + ' '  + Right;
        2: Result := Left + ' ' + FormatFloat(right, abs(Value), fs);
      end
    else
      Result := FormatFloat(ANumberFormatStr, Value, fs)
  end;

  function DateTimeToStrNoNaN(const Value: Double;
    ANumberFormat: TsNumberFormat; ANumberFormatStr: String; ADecimals: Word): ansistring;
  begin
    Result := '';
    if not IsNaN(Value) then begin
      if ANumberFormatStr = '' then
        ANumberFormatStr := BuildDateTimeFormatString(ANumberFormat,
          Workbook.FormatSettings, ANumberFormatStr);
      Result := FormatDateTime(ANumberFormatStr, Value, [fdoInterval]);
    end;
  end;

begin
  Result := '';
  if ACell = nil then
    Exit;

  with ACell^ do
    case ContentType of
      cctNumber:
        Result := FloatToStrNoNaN(NumberValue, NumberFormat, NumberFormatStr, Decimals);
      cctUTF8String:
        Result := UTF8StringValue;
      cctDateTime:
        Result := DateTimeToStrNoNaN(DateTimeValue, NumberFormat, NumberFormatStr, Decimals);
      cctBool:
        Result := IfThen(BoolValue, lpTRUE, lpFALSE);
      cctError:
        case TErrorValue(StatusValue and $0F) of
          errEmptyIntersection  : Result := lpErrEmptyIntersection;
          errDivideByZero       : Result := lpErrDivideByZero;
          errWrongType          : Result := lpErrWrongType;
          errIllegalRef         : Result := lpErrIllegalRef;
          errWrongName          : Result := lpErrWrongName;
          errOverflow           : Result := lpErrOverflow;
          errArgNotAvail        : Result := lpErrArgNotAvail;
          errFormulaNotSupported: Result := lpErrFormulaNotSupported;
        end;
      else
        Result := '';
    end;
end;

function TsWorksheet.ReadAsNumber(ARow, ACol: Cardinal): Double;
var
  ACell: PCell;
  Str: string;
begin
  Result := 0.0;
  ACell := FindCell(ARow, ACol);
  if ACell = nil then
    exit;

  case ACell^.ContentType of
    cctDateTime   : Result := ACell^.DateTimeValue; //this is in FPC TDateTime format, not Excel
    cctNumber     : Result := ACell^.NumberValue;
    cctUTF8String : TryStrToFloat(ACell^.UTF8StringValue, Result);
  end;
end;

{@@
  Reads the contents of a cell and returns the date/time value of the cell.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell

  @return True if the cell is a datetime value, false otherwise
}
function TsWorksheet.ReadAsDateTime(ARow, ACol: Cardinal; out AResult: TDateTime): Boolean;
var
  ACell: PCell;
  Str: string;
begin
  ACell := FindCell(ARow, ACol);

  if (ACell = nil) or (ACell^.ContentType <> cctDateTime) then
  begin
    AResult := 0;
    Result := False;
    Exit;
  end;

  AResult := ACell^.DateTimeValue;
  Result := True;
end;

function TsWorksheet.ReadRPNFormulaAsString(ACell: PCell): String;
var
  formula: TsRPNFormula;
  elem: TsFormulaElement;
  i, j: Integer;
  L: TStringList;
  s: String;
  ptr: Pointer;
  fek: TFEKind;

  procedure Store(s: String);
  begin
    L.Clear;
    L.Add(s);
  end;

begin
  Result := '';
  if ACell = nil then
    exit;

  L := TStringList.Create;
  try
    for i:=0 to Length(ACell^.RPNFormulaValue)-1 do begin
      elem := ACell^.RPNFormulaValue[i];
      ptr := Pointer(elem.ElementKind);
      case elem.ElementKind of
        fekNum:
          L.AddObject(Format('%g', [elem.DoubleValue]), ptr);
        fekInteger:
          L.AddObject(IntToStr(elem.IntValue), ptr);
        fekString:
          L.AddObject('"' + elem.StringValue + '"', ptr);
        fekBool:
          L.AddObject(IfThen(elem.DoubleValue=0, 'TRUE', 'FALSE'), ptr);
        fekCell,
        fekCellRef:
          L.AddObject(GetCellString(elem.Row, elem.Col, elem.RelFlags), ptr);
        fekCellRange:
          L.AddObject(GetCellRangeString(elem.Row, elem.Col, elem.Row2, elem.Col2, elem.RelFlags), ptr);
        // Operations:
        fekAdd         : L.AddObject('+', ptr);
        fekSub         : L.AddObject('-', ptr);
        fekMul         : L.AddObject('*', ptr);
        fekDiv         : L.AddObject('/', ptr);
        fekPower       : L.AddObject('^', ptr);
        fekConcat      : L.AddObject('&', ptr);
        fekParen       : L.AddObject('', ptr);
        fekEqual       : L.AddObject('=', ptr);
        fekNotEqual    : L.AddObject('<>', ptr);
        fekLess        : L.AddObject('<', ptr);
        fekLessEqual   : L.AddObject('<=', ptr);
        fekGreater     : L.AddObject('>', ptr);
        fekGreaterEqual: L.AddObject('>=', ptr);
        fekPercent     : L.AddObject('%', ptr);
        fekUPlus       : L.AddObject('+', ptr);
        fekUMinus      : L.AddObject('-', ptr);
        fekCellInfo    : L.AddObject('CELL', ptr);         // That's the function name!
        else
          begin
            s := GetEnumName(TypeInfo(TFEKind), integer(elem.ElementKind));
            Delete(s, 1, 3);
            L.AddObject(s, ptr);
          end;
      end;
    end;

    i := L.Count-1;
    while (L.Count > 0) and (i >= 0) do begin
      fek := TFEKind(PtrInt(L.Objects[i]));
      case fek of
        fekAdd, fekSub, fekMul, fekDiv, fekPower, fekConcat,
        fekEqual, fekNotEqual, fekLess, fekLessEqual, fekGreater, fekGreaterEqual:
          begin
            L.Strings[i] := Format('%s%s%s', [L[i+2], L[i], L[i+1]]);
            L.Objects[i] := pointer(fekString);
            L.Delete(i+2);
            L.Delete(i+1);
          end;
        fekUPlus, fekUMinus:
          begin
            L.Strings[i] := L[i]+L[i+1];
            L.Objects[i] := Pointer(fekString);
            L.Delete(i+1);
          end;
        fekPercent:
          begin
            L.Strings[i] := L[i+1]+L[i];
            L.Objects[i] := Pointer(fekString);
            L.Delete(i+1);
          end;
        fekParen:
          begin
            L.Strings[i] := Format('(%s)', [L[i+1]]);
            L.Objects[i] := pointer(fekString);
            L.Delete(i+1);
          end;
        else
          if fek >= fekAdd then begin
            elem := ACell^.RPNFormulaValue[i];
            s := '';
            for j:= i+elem.ParamsNum downto i+1 do begin
              s := s + ',' + L[j];
              L.Delete(j);
            end;
            Delete(s, 1, 1);
            L.Strings[i] := Format('%s(%s)', [L[i], s]);
            L.Objects[i] := pointer(fekString);
          end;
      end;
      dec(i);
    end;

    Result := '=' + L[0];

  finally
    L.Free;
  end;
end;

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
  Clears the list of Cells and releases their memory.
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
  Writes UTF-8 encoded text to a determined cell.

  On formats that don't support unicode, the text will be converted
  to ISO Latin 1.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AText     The text to be written encoded in utf-8
}
procedure TsWorksheet.WriteUTF8Text(ARow, ACol: Cardinal; AText: ansistring);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.ContentType := cctUTF8String;
  ACell^.UTF8StringValue := AText;
  ChangedCell(ARow, ACol);
end;

{@@
  Writes a floating-point number to a determined cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  ANumber   The number to be written
  @param  AFormat   The format identifier, e.g. nfFixed (optional)
  @param  ADecimals The number of decimals used for formatting (optional)
  @param  ACurrencySymbol The currency symbol in case of currency format (nfCurrency)
}
procedure TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: double;
  AFormat: TsNumberFormat = nfGeneral; ADecimals: Byte = 2;
  ACurrencySymbol: String = '');
begin
  WriteNumber(GetCell(ARow, ACol), ANumber, AFormat, ADecimals, ACurrencySymbol);
end;


procedure TsWorksheet.WriteNumber(ACell: PCell; ANumber: Double;
  AFormat: TsNumberFormat = nfGeneral; ADecimals: Byte = 2;
  ACurrencySymbol: String = '');
var
  fs: TFormatSettings;
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := ANumber;
    ACell^.Decimals := ADecimals;

    if IsDateTimeFormat(AFormat) then
      raise Exception.Create(lpInvalidNumberFormat);

    {
    if AFormat = nfCustom then
      raise Exception.Create(lpIllegalNumberformat);
     }

    if AFormat <> nfGeneral then begin
      Include(ACell^.UsedFormattingFields, uffNumberFormat);
      ACell^.NumberFormat := AFormat;
      ACell^.Decimals := ADecimals;
      ACell^.CurrencySymbol := ACurrencySymbol;
      ACell^.NumberFormatStr := BuildNumberFormatString(ACell^.NumberFormat,
        Workbook.FormatSettings, ADecimals, ACurrencySymbol);
    end;

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a floating point number to the cell and uses a custom number format
  specified by the format string.
  NOTE that fpspreadsheet may not be able to detect the formatting when reading
  the file. }
procedure TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: Double;
  AFormatString: String);
var
  ACell: PCell;
  parser: TsNumFormatParser;
  nf: TsNumberFormat;
begin
  parser := TsNumFormatParser.Create(Workbook, AFormatString, nfCustom, cdToFPSpreadsheet);
  try
    // Format string ok?
    if parser.Status <> psOK then
      raise Exception.Create(lpNoValidNumberFormatString);
    if IsDateTimeFormat(parser.Builtin_NumFormat)
      then raise Exception.Create(lpInvalidNumberFormat);
    // If format string matches a built-in format use its format identifier,
    // All this is considered when calling Builtin_NumFormat of the parser.
    nf := parser.Builtin_NumFormat;
  finally
    parser.Free;
  end;

  ACell := GetCell(ARow, ACol);
  Include(ACell^.UsedFormattingFields, uffNumberFormat);
  ACell^.ContentType := cctNumber;
  ACell^.NumberValue := ANumber;
  ACell^.NumberFormat := nf;
  ACell^.NumberFormatStr := AFormatString;
  ACell^.Decimals := 0;
  ACell^.CurrencySymbol := '';

  ChangedCell(ARow, ACol);
end;

{@@
  Writes as empty cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell

  Note: an empty cell  is required for formatting.
}
procedure TsWorksheet.WriteBlank(ARow, ACol: Cardinal);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.ContentType := cctEmpty;
  ChangedCell(ARow, ACol);
end;

{@@
  Writes as boolean cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The boolean value
}
procedure TsWorksheet.WriteBoolValue(ARow, ACol: Cardinal; AValue: Boolean);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.ContentType := cctBool;
  ACell^.BoolValue := AValue;
  ChangedCell(ARow, ACol);
end;

{@@
  Writes a date/time value to a determined cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The date/time/datetime to be written
  @param  AFormat    The format specifier, e.g. nfShortDate (optional)
  @param  AFormatStr Format string, used only for nfFmtDateTime.
                     Must follow the rules for "FormatDateTime", or use
                     "dm" as abbreviation for "d/mmm", "my" for "mmm/yy",
                     "ms" for "nn:ss", "msz" for "nn:ss.z" (optional)
                     or use any other free format (at your own risk...)

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) Number, and the cell is formatted
  as a date (either built-in or a custom format).
}
procedure TsWorksheet.WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
  AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = '');
begin
  WriteDateTime(GetCell(ARow, ACol), AValue, AFormat, AFormatStr);
end;

procedure TsWorksheet.WriteDateTime(ACell: PCell; AValue: TDateTime;
  AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = '');
begin
  if ACell <> nil then begin
    if (AFormat in [nfFmtDateTime, nfTimeInterval]) then
      AFormatStr := BuildDateTimeFormatString(AFormat, Workbook.FormatSettings, AFormatStr);

    ACell^.ContentType := cctDateTime;
    ACell^.DateTimeValue := AValue;
    // Date/time is actually a number field in Excel.
    // To make sure it gets saved correctly, set a date format (instead of General).
    // The user can choose another date format if he wants to
    Include(ACell^.UsedFormattingFields, uffNumberFormat);
    ACell^.NumberFormat := AFormat;
    ACell^.NumberFormatStr := AFormatStr;

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

procedure TsWorksheet.WriteDecimals(ARow, ACol: Cardinal; ADecimals: Byte);
begin
  WriteDecimals(FindCell(ARow, ACol), ADecimals);
end;

procedure TsWorksheet.WriteDecimals(ACell: PCell; ADecimals: Byte);
begin
  if (ACell <> nil) and (ACell^.ContentType = cctNumber) and (ACell^.NumberFormat <> nfCustom)
  then begin
    ACell^.Decimals := ADecimals;
    ACell^.NumberFormatStr := BuildNumberFormatString(ACell^.NumberFormat,
      FWorkbook.FormatSettings, ADecimals, ACell^.CurrencySymbol);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a cell with an error.

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The error code value
}
procedure TsWorksheet.WriteErrorValue(ARow, ACol: Cardinal; AValue: TErrorValue);
begin
  WriteErrorValue(GetCell(ARow, ACol), AValue);
end;

procedure TsWorksheet.WriteErrorValue(ACell: PCell; AValue: TErrorValue);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctError;
    ACell^.StatusValue := (ACell^.StatusValue and $F0) or ord(AValue);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@
  Writes a formula to a determined cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AFormula  The formula to be written
}
procedure TsWorksheet.WriteFormula(ARow, ACol: Cardinal; AFormula: TsFormula);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.ContentType := cctFormula;
  ACell^.FormulaValue := AFormula;
  ChangedCell(ARow, ACol);
end;

{@@
  Adds number format to the formatting of a cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  TsNumberFormat What format to apply
  @param  string    Formatstring

  @see    TsNumberFormat
}
procedure TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumberFormat: TsNumberFormat; const AFormatString: String = '');
var
  ACell: PCell;
  oldNumFmt: TsNumberFormat;
begin
  ACell := GetCell(ARow, ACol);
  Include(ACell^.UsedFormattingFields, uffNumberFormat);
  ACell^.NumberFormat := ANumberFormat;
  if (AFormatString = '') then
    ACell^.NumberFormatStr := BuildNumberFormatString(ANumberFormat,
      Workbook.FormatSettings, ACell^.Decimals, ACell^.CurrencySymbol)
  else
    ACell^.NumberFormatStr := AFormatString;
  ChangedCell(ARow, ACol);
end;

procedure TsWorksheet.WriteRPNFormula(ARow, ACol: Cardinal;
  AFormula: TsRPNFormula);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.ContentType := cctRPNFormula;
  ACell^.RPNFormulaValue := AFormula;
  ChangedCell(ARow, ACol);
end;

{@@
  Adds font specification to the formatting of a cell

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontName   Name of the font
  @param  AFontSize   Size of the font, in points
  @param  AFontStyle  Set with font style attributes
                      (don't use those of unit "graphics" !)

  @result             Index of font in font list
}
function TsWorksheet.WriteFont(ARow, ACol: Cardinal; const AFontName: String;
  AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer;
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  Include(lCell^.UsedFormattingFields, uffFont);
  Result := FWorkbook.FindFont(AFontName, AFontSize, AFontStyle, AFontColor);
  if Result = -1 then
    result := FWorkbook.AddFont(AFontName, AFontSize, AFontStyle, AFontColor);
  lCell^.FontIndex := Result;
  ChangedFont(ARow, ACol);
end;

procedure TsWorksheet.WriteFont(ARow, ACol: Cardinal; AFontIndex: Integer);
var
  lCell: PCell;
begin
  if (AFontIndex >= 0) and (AFontIndex < Workbook.GetFontCount) and (AFontIndex <> 4)
    // note: Font index 4 is not defined in BIFF
  then begin
    lCell := GetCell(ARow, ACol);
    Include(lCell^.UsedFormattingFields, uffFont);
    lCell^.FontIndex := AFontIndex;
    ChangedFont(ARow, ACol);
  end else
    raise Exception.Create(lpInvalidFontIndex);
end;

function TsWorksheet.WriteFontColor(ARow, ACol: Cardinal; AFontColor: TsColor): Integer;
var
  lCell: PCell;
  fnt: TsFont;
begin
  lCell := GetCell(ARow, ACol);
  fnt := Workbook.GetFont(lCell^.FontIndex);
  Result := WriteFont(ARow, ACol, fnt.FontName, fnt.Size, fnt.Style, AFontColor);
end;

function TsWorksheet.WriteFontName(ARow, ACol: Cardinal; AFontName: String): Integer;
var
  lCell: PCell;
  fnt: TsFont;
begin
  lCell := GetCell(ARow, ACol);
  fnt := Workbook.GetFont(lCell^.FontIndex);
  result := WriteFont(ARow, ACol, AFontName, fnt.Size, fnt.Style, fnt.Color);
end;

function TsWorksheet.WriteFontSize(ARow, ACol: Cardinal; ASize: Single): Integer;
var
  lCell: PCell;
  fnt: TsFont;
begin
  lCell := GetCell(ARow, ACol);
  fnt := Workbook.GetFont(lCell^.FontIndex);
  Result := WriteFont(ARow, ACol, fnt.FontName, ASize, fnt.Style, fnt.Color);
end;

function TsWorksheet.WriteFontStyle(ARow, ACol: Cardinal;
  AStyle: TsFontStyles): Integer;
var
  lCell: PCell;
  fnt: TsFont;
begin
  lCell := GetCell(ARow, ACol);
  fnt := Workbook.GetFont(lCell^.FontIndex);
  Result := WriteFont(ARow, ACol, fnt.FontName, fnt.Size, AStyle, fnt.Color);
end;

{@@
  Adds text rotation to the formatting of a cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  ARotation How to rotate the text

  @see    TsTextRotation
}
procedure TsWorksheet.WriteTextRotation(ARow, ACol: Cardinal;
  ARotation: TsTextRotation);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  Include(ACell^.UsedFormattingFields, uffTextRotation);
  ACell^.TextRotation := ARotation;
  ChangedFont(ARow, ACol);
end;

procedure TsWorksheet.WriteUsedFormatting(ARow, ACol: Cardinal;
  AUsedFormatting: TsUsedFormattingFields);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.UsedFormattingFields := AUsedFormatting;
  ChangedCell(ARow, ACol);
end;

procedure TsWorksheet.WriteBackgroundColor(ARow, ACol: Cardinal;
  AColor: TsColor);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);
  ACell^.UsedFormattingFields := ACell^.UsedFormattingFields + [uffBackgroundColor];
  ACell^.BackgroundColor := AColor;
  ChangedCell(ARow, ACol);
end;

{ Sets the color of a cell border line.
  Note: the border must be included in Borders set in order to be shown! }
procedure TsWorksheet.WriteBorderColor(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AColor: TsColor);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.BorderStyles[ABorder].Color := AColor;
  ChangedCell(ARow, ACol);
end;

{ Sets the linestyle of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown! }
procedure TsWorksheet.WriteBorderLineStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.BorderStyles[ABorder].LineStyle := ALineStyle;
  ChangedCell(ARow, ACol);
end;

{ Shows the cell borders included in the set ABorders. The borders are drawn
  using the "BorderStyles" assigned to the cell. }
procedure TsWorksheet.WriteBorders(ARow, ACol: Cardinal; ABorders: TsCellBorders);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  Include(lCell^.UsedFormattingFields, uffBorder);
  lCell^.Border := ABorders;
  ChangedCell(ARow, ACol);
end;

{ Sets the style of a cell border, i.e. line style and line color.
  Note: the border must be included in the "Borders" set in order to be shown! }
procedure TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AStyle: TsCellBorderStyle);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.BorderStyles[ABorder] := AStyle;
  ChangedCell(ARow, ACol);
end;

{ Sets line style and line color of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown! }
procedure TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLinestyle; AColor: TsColor);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.BorderStyles[ABorder].LineStyle := ALineStyle;
  lCell^.BorderStyles[ABorder].Color := AColor;
  ChangedCell(ARow, ACol);
end;

{ Sets the style of all cell border of a cell, i.e. line style and line color.
  Note: Only those borders included in the "Borders" set are shown! }
procedure TsWorksheet.WriteBorderStyles(ARow, ACol: Cardinal;
  const AStyles: TsCellBorderStyles);
var
  b: TsCellBorder;
  cell: PCell;
begin
  cell := GetCell(ARow, ACol);
  for b in TsCellBorder do cell^.BorderStyles[b] := AStyles[b];
  ChangedCell(ARow, ACol);
end;

procedure TsWorksheet.WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.UsedFormattingFields := lCell^.UsedFormattingFields + [uffHorAlign];
  lCell^.HorAlignment := AValue;
  ChangedCell(ARow, ACol);
end;

procedure TsWorksheet.WriteVertAlignment(ARow, ACol: Cardinal; AValue: TsVertAlignment);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.UsedFormattingFields := lCell^.UsedFormattingFields + [uffVertAlign];
  lCell^.VertAlignment := AValue;
  ChangedCell(ARow, ACol);
end;

procedure TsWorksheet.WriteWordWrap(ARow, ACol: Cardinal; AValue: Boolean);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  if AValue then
    Include(lCell^.UsedFormattingFields, uffWordwrap)
  else
    Exclude(lCell^.UsedFormattingFields, uffWordwrap);
  ChangedCell(ARow, ACol);
end;

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

function TsWorksheet.GetRow(ARow: Cardinal): PRow;
begin
  Result := FindRow(ARow);
  if (Result = nil) then begin
    Result := GetMem(SizeOf(TRow));
    FillChar(Result^, SizeOf(TRow), #0);
    Result^.Row := ARow;
    FRows.Add(Result);
  end;
end;

function TsWorksheet.GetCol(ACol: Cardinal): PCol;
begin
  Result := FindCol(ACol);
  if (Result = nil) then begin
    Result := GetMem(SizeOf(TCol));
    FillChar(Result^, SizeOf(TCol), #0);
    Result^.Col := ACol;
    FCols.Add(Result);
  end;
end;

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

procedure TsWorksheet.WriteRowInfo(ARow: Cardinal; AData: TRow);
var
  AElement: PRow;
begin
  AElement := GetRow(ARow);
  AElement^.Height := AData.Height;
end;

procedure TsWorksheet.WriteRowHeight(ARow: Cardinal; AHeight: Single);
var
  AElement: PRow;
begin
  AElement := GetRow(ARow);
  AElement^.Height := AHeight;
end;

procedure TsWorksheet.WriteColInfo(ACol: Cardinal; AData: TCol);
var
  AElement: PCol;
begin
  AElement := GetCol(ACol);
  AElement^.Width := AData.Width;
end;

procedure TsWorksheet.WriteColWidth(ACol: Cardinal; AWidth: Single);
var
  AElement: PCol;
begin
  AElement := GetCol(ACol);
  AElement^.Width := AWidth;
end;


{ TsWorkbook }

{@@
  Helper method for clearing the spreadsheet list.
}
procedure TsWorkbook.RemoveWorksheetsCallback(data, arg: pointer);
begin
  TsWorksheet(data).Free;
end;

{@@
  Constructor.
}
constructor TsWorkbook.Create;
begin
  inherited Create;
  FWorksheets := TFPList.Create;
  FormatSettings := DefaultFormatSettings;
  FFontList := TFPList.Create;
  SetDefaultFont('Arial', 10.0);
  InitFonts;
end;

{@@
  Destructor.
}
destructor TsWorkbook.Destroy;
begin
  RemoveAllWorksheets;
  RemoveAllFonts;

  FWorksheets.Free;
  FFontList.Free;

  inherited Destroy;
end;


{@@
  Helper method for determining the spreadsheet type from the file type extension

  Returns: True if the file matches any of the known formats, false otherwise
}
class function TsWorkbook.GetFormatFromFileName(const AFileName: TFileName; var SheetType: TsSpreadsheetFormat): Boolean;
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
  Convenience method which creates the correct
  reader object for a given spreadsheet format.
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
  Convenience method which creates the correct
  writer object for a given spreadsheet format.
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
  Reads the document from a file.
}
procedure TsWorkbook.ReadFromFile(AFileName: string;
  AFormat: TsSpreadsheetFormat);
var
  AReader: TsCustomSpreadReader;
begin
  AReader := CreateSpreadReader(AFormat);
  try
    AReader.ReadFromFile(AFileName, Self);
    FFormat := AFormat;
  finally
    AReader.Free;
  end;
end;

{@@
  Reads the document from a file. This method will try to guess the format from
  the extension. In the case of the ambiguous xls extension, it will simply
  assume that it is BIFF8. Note that it could be BIFF2, 3, 4 or 5 too.
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

procedure TsWorkbook.ReadFromFileIgnoringExtension(AFileName: string);
var
  SheetType: TsSpreadsheetFormat;
  lException: Exception;
begin
  while (SheetType in [sfExcel2..sfExcel8]) and (lException <> nil) do
  begin
    try
      Dec(SheetType);
      ReadFromFile(AFileName, SheetType);
      lException := nil;
    except
      on E: Exception do
           { do nothing } ;
    end;
    if lException = nil then Break;
  end;
end;

{@@
  Reads the document from a seekable stream.
}
procedure TsWorkbook.ReadFromStream(AStream: TStream;
  AFormat: TsSpreadsheetFormat);
var
  AReader: TsCustomSpreadReader;
begin
  AReader := CreateSpreadReader(AFormat);

  try
    AReader.ReadFromStream(AStream, Self);
  finally
    AReader.Free;
  end;
end;

{@@
  Writes the document to a file.

  If the file doesn't exist, it will be created.
}
procedure TsWorkbook.WriteToFile(const AFileName: string;
 const AFormat: TsSpreadsheetFormat; const AOverwriteExisting: Boolean = False);
var
  AWriter: TsCustomSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);
  try
    AWriter.WriteToFile(AFileName, AOverwriteExisting);
  finally
    AWriter.Free;
  end;
end;

{@@
  Writes the document to file based on the extension. If this was an earlier sfExcel type file, it will be upgraded to sfExcel8, 
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
}
procedure TsWorkbook.WriteToStream(AStream: TStream; AFormat: TsSpreadsheetFormat);
var
  AWriter: TsCustomSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);

  try
    AWriter.WriteToStream(AStream);
  finally
    AWriter.Free;
  end;
end;

{@@
  Adds a new worksheet to the workbook

  It is added to the end of the list of worksheets

  @param  AName     The name of the new worksheet
  @return The instace of the newly created worksheet
  @see    TsWorkbook
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


{ Font handling }

{@@
  Adds a font to the font list. Returns the index in the font list.
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

function TsWorkbook.AddFont(const AFont: TsFont): Integer;
begin
  // Font index 4 does not exist in BIFF. Avoid that a real font gets this index.
  if FFontList.Count = 4 then
    FFontList.Add(nil);
  result := FFontList.Add(AFont);
end;

{@@
  Copies the font list "ASource" to the workbook's font list
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
  Checks whether the font with the given specification is already contained in
  the font list. Returns the index, or -1, if not found.
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
      (AColor = fnt.Color)
    then
      exit;
  end;
  Result := -1;
end;

{@@
  Initialized the font list. In case of BIFF format, adds 5 fonts
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

  AddFont(fntName, fntSize, [fssItalic], scBlack);       // FONT2 for uffItalic
  AddFont(fntName, fntSize, [fssUnderline], scBlack);    // FONT3 for uffUnderline
  // FONT4 which does not exist in BIFF is added automatically with nil as place-holder
  AddFont(fntName, fntSize, [fssBold, fssItalic], scBlack); // FONT5 for uffBoldItalic


  FBuiltinFontCount := FFontList.Count;
end;

{@@
  Clears the list of fonts and releases their memory.
}
procedure TsWorkbook.RemoveAllFonts;
var
  i, n: Integer;
  fnt: TsFont;
begin
  for i:=FFontList.Count-1 downto 0 do begin
    fnt := TsFont(FFontList.Items[i]);
    fnt.Free;
    FFontList.Delete(i);
  end;
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
  Returns the font with the given index.
}
function TsWorkbook.GetFont(AIndex: Integer): TsFont;
begin
  if (AIndex >= 0) and (AIndex < FFontList.Count) then
    Result := FFontList.Items[AIndex]
  else
    Result := nil;
end;

{@@
  Returns the count of registered fonts
}
function TsWorkbook.GetFontCount: Integer;
begin
  Result := FFontList.Count;
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
  Result := Format('%x%x%x', [r, g, b]);
end;

{@@
  Returns the name of the color pointed to by the given color index.
  If the name is not known the hex string is returned as RRGGBB.
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
}
function TsWorkbook.GetPaletteColor(AColorIndex: TsColor): TsColorValue;
begin
  if (AColorIndex >= 0) and (AColorIndex < GetPaletteSize) then begin
    if ((FPalette = nil) or (Length(FPalette) = 0)) then
      Result := DEFAULT_PALETTE[AColorIndex]
    else
      Result := FPalette[AColorIndex];
  end else
    Result := $000000;  // "black" as default
end;

{@@
  Replaces a color value of the current palette by a new value. The color must
  be given as ABGR (little-endian), with A=0}
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
  Instructs the Workbook to take colors from the palette pointed to by the parameter
  This palette is only used for writing. When reading the palette found in the
  file is used.
}
procedure TsWorkbook.UsePalette(APalette: PsPalette; APaletteCount: Word;
  ABigEndian: Boolean);
var
  i: Integer;
begin
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


{ TsCustomNumFormatList }

constructor TsCustomNumFormatList.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  AddBuiltinFormats;
end;

destructor TsCustomNumFormatList.Destroy;
begin
  Clear;
  inherited Destroy;
end;

{ Adds a new number format data to the list and returns the list index of the
  new (or present) item. }
function TsCustomNumFormatList.AddFormat(AFormatIndex: Integer;
  AFormatString: String; ANumFormat: TsNumberFormat; ADecimals: byte = 0;
  ACurrencySymbol: String = ''): integer;
var
  item: TsNumFormatData;
begin
  item := TsNumFormatData.Create;
  item.Index := AFormatIndex;
  item.NumFormat := ANumFormat;
  if AFormatString = '' then begin
    if IsDateTimeFormat(ANumFormat) then
      AFormatString := BuildDateTimeFormatString(ANumFormat, Workbook.FormatSettings,
        AFormatString)
    else
    if item.NumFormat <> nfCustom then
      AFormatString := BuildNumberFormatString(ANumFormat, Workbook.FormatSettings,
        ADecimals, ACurrencySymbol);
  end;
  item.FormatString := AFormatString;
  item.Decimals := ADecimals;
  item.CurrencySymbol := ACurrencySymbol;
  Result := inherited Add(item);
end;

function TsCustomNumFormatList.AddFormat(AFormatString: String;
  ANumFormat: TsNumberFormat; ADecimals: Byte = 0;
  ACurrencySymbol: String = ''): Integer;
begin
  if AFormatString = '' then begin
    Result := 0;
    exit;
  end;
  Result := AddFormat(FNextFormatIndex, AFormatString, ANumFormat, ADecimals,
    ACurrencySymbol);
  inc(FNextFormatIndex);
end;

function TsCustomNumFormatList.AddFormat(AFormatCell: PCell): Integer;
var
  item: TsNumFormatData;
begin
  if AFormatCell = nil then
    raise Exception.Create('TsCustomNumFormat.Add: No nil pointers please');

  if Count = 0 then
    raise Exception.Create('TsCustomNumFormatList: Error in program logics: You must provide built-in formats first.');

  Result := AddFormat(FNextFormatIndex,
    AFormatCell^.NumberFormatStr,
    AFormatCell^.NumberFormat,
    AFormatCell^.Decimals,
    AFormatCell^.CurrencySymbol
  );

  inc(FNextFormatIndex);
end;

{ Adds the builtin format items to the list. The formats must be specified in
  a way that is compatible with the destination file format. Conversion of the
  formatstrings can be done by calling "ConvertAfterReadung" bzw. "ConvertBeforeWriting".
  "AddBuiltInFormats" must be called before user items are added.
  Must specify FFirstFormatIndexInFile (BIFF5-8, e.g. doesn't save formats <164)
  and must initialize the index of the first user format (FNextFormatIndex)
  which is automatically incremented when adding user formats. }
procedure TsCustomNumFormatList.AddBuiltinFormats;
begin
  // must be overridden - see xlscommon as an example.
end;

{ Takes the format string (AFormatString) as it is read from the file and
  extracts the number format type and the number of decimals out of it for use by
  fpc. The method also converts the format string to a form that can be used
  by fpc's FormatDateTime and FormatFloat. This conversion should be done in an
  overridden method which known more about the details of the spreadsheet file
  format. }
procedure TsCustomNumFormatList.ConvertAfterReading(AFormatIndex: Integer;
  var AFormatString: String; var ANumFormat: TsNumberFormat;
  var ADecimals: Byte; var ACurrencySymbol: String);
var
  parser: TsNumFormatParser;
  fmt: String;
  lFormatData: TsNumFormatData;
  i: Integer;
  nf: TsNumberFormat;
begin
  i := Find(AFormatIndex);
  if i > 0 then begin
    lFormatData := Items[i];
    fmt := lFormatData.FormatString;
  end else
    fmt := AFormatString;
  nf := nfGeneral;  // not used here.

  // Analyzes the format string and tries to convert it to fpSpreadsheet format.
  parser := TsNumFormatParser.Create(Workbook, fmt, nf, cdToFPSpreadsheet);
  try
    if parser.Status = psOK then begin
      ANumFormat := parser.Builtin_NumFormat;
      AFormatString := parser.FormatString;  // This is the converted string.
      if ANumFormat <> nfCustom then begin
        ADecimals := parser.ParsedSections[0].Decimals;
        ACurrencySymbol := parser.ParsedSections[0].CurrencySymbol;
      end else begin
        ADecimals := 0;
        ACurrencySymbol := '';
      end;
    end;
  finally
    parser.Free;
  end;
end;

{ Is called before collection all number formats of the spreadsheet and before
  writing to file. Its purpose is to convert the format string as used by fpc
  to a format compatible with the spreadsheet file format. }
procedure TsCustomNumFormatList.ConvertBeforeWriting(var AFormatString: String;
  var ANumFormat: TsNumberFormat; var ADecimals: Byte; var ACurrencySymbol: String);
var
  parser: TsNumFormatParser;
  fmt: String;
begin
  parser := TsNumFormatParser.Create(Workbook, AFormatString, ANumFormat, cdFromFPSpreadsheet);
  try
    if parser.Status = psOK then begin
      AFormatString := parser.FormatString;
      ANumFormat := parser.Builtin_NumFormat;
      ADecimals := parser.ParsedSections[0].Decimals;
      ACurrencySymbol := parser.ParsedSections[0].CurrencySymbol;
    end;
  finally
    parser.Free;
  end;
end;

{ Called from the reader when a format item has been read from the file.
  Determines the numFormat type, format string etc and stores the format in the
  list. If necessary, the format string has to be made compatible with fpc
  afterwards - it is used directly for getting the cell text. }
procedure TsCustomNumFormatList.AnalyzeAndAdd(AFormatIndex: Integer;
  AFormatString: String);
var
  nf: TsNumberFormat;
  decs: Byte;
  currsym: String;
begin
  if Find(AFormatIndex) > -1 then
    exit;

  // Analyze & convert the format string, extract infos for internal formatting
  ConvertAfterReading(AFormatIndex, AFormatString, nf, decs, currsym);

  // Add the new item
  AddFormat(AFormatIndex, AFormatString, nf, decs, currSym);
end;

{ Clears the list and frees memory occupied by the format items. }
procedure TsCustomNumFormatList.Clear;
var
  i: Integer;
begin
  for i:=0 to Count-1 do RemoveFormat(i);
  inherited Clear;
end;

{ Deletes a format item from the list, and makes sure that its memory is
  released. }
procedure TsCustomNumFormatList.Delete(AIndex: Integer);
begin
  RemoveFormat(AIndex);
  Delete(AIndex);
end;

{ Seeks a format item with the given properties and returns its list index,
  or -1 if not found. }
function TsCustomNumFormatList.Find(ANumFormat: TsNumberFormat;
  AFormatString: String; ADecimals: Byte; ACurrencySymbol: String): Integer;
var
  item: TsNumFormatData;
  fmt: String;
  itemfmt: String;
begin
  if (ANumFormat = nfFmtDateTime) then begin
    fmt := lowercase(AFormatString);
    for Result := Count-1 downto 0 do begin
      item := Items[Result];
      if (item <> nil) and (item.NumFormat = nfFmtDateTime) then begin
        itemfmt := lowercase(item.FormatString);
        if ((itemfmt = 'dm') or (itemfmt = 'd-mmm') or (itemfmt = 'd mmm') or (itemfmt = 'd. mmm') or (itemfmt ='d/mmm'))
          and ((fmt = 'dm') or (fmt = 'd-mmm') or (fmt = 'd mmm') or (fmt = 'd. mmm') or (fmt = 'd/mmm'))
        then
          exit;
        if ((itemfmt = 'my') or (itemfmt = 'mmm-yy') or (itemfmt = 'mmm yyy') or (itemfmt = 'mmm/yy'))
          and ((fmt = 'my') or (fmt = 'mmm-yy') or (fmt = 'mmm yy') or (fmt = 'mmm/yy'))
        then
          exit;
        if ((itemfmt = 'ms') or (itemfmt = 'nn:ss') or (itemfmt = 'mm:ss'))
          and ((fmt = 'ms') or (fmt = 'nn:ss') or (fmt = 'mm:ss'))
        then
          exit;
        if ((itemfmt = 'msz') or (itemfmt = 'mm:ss.z') or (itemfmt = 'mm:ss.0'))
          and ((fmt = 'msz') or (fmt = 'mm:ss.z') or (fmt = 'mm:ss.0'))
        then
          exit;
      end;
    end;
    for Result := 0 to Count-1 do begin
      item := Items[Result];
      if fmt = lowercase(item.FormatString) then
        exit;
    end;
  end;

  // Check only the format string for nfCustom.
  if (ANumFormat = nfCustom) then
    for Result := Count-1 downto 0 do begin
      item := Items[Result];
      if (item <> nil)
         and (item.NumFormat = ANumFormat)
         and (item.FormatString = AFormatString)
      then
        exit;
    end;

  // The other formats can carry additional information
  for Result := Count-1 downto 0 do begin
    item := Items[Result];
    if (item <> nil)
      and (item.NumFormat = ANumFormat)
      and (item.FormatString = AFormatString)
      and (item.Decimals = ADecimals)
      and (not (item.NumFormat in [nfCurrency, nfCurrencyRed, nfAccounting, nfAccountingRed])
            or (item.CurrencySymbol = ACurrencySymbol))
    then
      exit;
  end;
  Result := -1;
end;

{ Finds the item with the given format index and returns its index in
  the format list. }
function TsCustomNumFormatList.Find(AFormatIndex: Integer): integer;
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

{ Finds the item with the given format string and returns its index in the
  format list. }
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

{ Determines whether the format attributed to the given cell is already
  contained in the list and returns its list index. }
function TsCustomNumFormatList.FindFormatOf(AFormatCell: PCell): integer;
begin
  if AFormatCell = nil then
    Result := -1
  else
    Result := Find(AFormatCell^.NumberFormat, AFormatCell^.NumberFormatStr,
      AFormatCell^.Decimals, AFormatCell^.CurrencySymbol);
end;

{ Determines the format string to be written into the spreadsheet file.
  Needs to be overridden if the format strings are different from the fpc
  convention. }
function TsCustomNumFormatList.FormatStringForWriting(AIndex: Integer): String;
var
  item: TsNumFormatdata;
begin
  item := Items[AIndex];
  if item <> nil then Result := item.FormatString else Result := '';
end;

function TsCustomNumFormatList.GetItem(AIndex: Integer): TsNumFormatData;
begin
  Result := TsNumFormatData(inherited Items[AIndex]);
end;

{ Deletes the memory occupied by the formatting data, but keeps the item in then
  list to maintain the indexes of followint items. }
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

{ Sorts the format data items in ascending order of the format indexes. }
procedure TsCustomNumFormatList.Sort;
begin
  inherited Sort(@CompareNumFormatData);
end;

{ TsCustomSpreadReader }

constructor TsCustomSpreadReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  CreateNumFormatList;
  FNumFormatList.FWorkbook := AWorkbook;
end;

destructor TsCustomSpreadReader.Destroy;
begin
  FNumFormatList.Free;
  inherited Destroy;
end;

procedure TsCustomSpreadReader.CreateNumFormatList;
begin
  { The format list needs to be created by descendants who know about the
    special requirements of the file format. }
end;

{@@
  Default file reading method.

  Opens the file and calls ReadFromStream

  @param  AFileName The input file name.
  @param  AData     The Workbook to be filled with information from the file.

  @see    TsWorkbook
}
procedure TsCustomSpreadReader.ReadFromFile(AFileName: string; AData: TsWorkbook);
var
  InputFile: TFileStream;
begin
  InputFile := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadFromStream(InputFile, AData);
  finally
    InputFile.Free;
  end;
end;

{@@
  This routine should be overriden in descendent classes.
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

procedure TsCustomSpreadReader.ReadFromStrings(AStrings: TStrings;
  AData: TsWorkbook);
begin
  raise Exception.Create(lpUnsupportedReadFormat);
end;

{ TsCustomSpreadWriter }

constructor TsCustomSpreadWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  CreateNumFormatList;
  FNumFormatList.FWorkbook := AWorkbook;
end;

destructor TsCustomSpreadWriter.Destroy;
begin
  FNumFormatList.Free;
  inherited Destroy;
end;

{@@
  Checks if the style of a cell is in the list of manually added FFormattingStyles
  and returns the index or -1 if it isn't
}
function TsCustomSpreadWriter.FindFormattingInList(AFormat: PCell): Integer;
var
  i: Integer;
  b: TsCellBorder;
  equ: Boolean;
begin
  Result := -1;

  for i := Length(FFormattingStyles) - 1 downto 0 do
  begin
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
        if FFormattingStyles[i].BorderStyles[b].Color <> AFormat^.BorderStyles[b].Color
        then begin
          equ := false;
          Break;
        end;
      end;
      if not equ then Continue;
    end;

    if uffBackgroundColor in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].BackgroundColor <> AFormat^.BackgroundColor) then Continue;

    if uffNumberFormat in AFormat^.UsedFormattingFields then begin
      if (FFormattingStyles[i].NumberFormat <> AFormat^.NumberFormat) then Continue;
      case AFormat^.NumberFormat of
        nfFixed, nfFixedTh, nfPercentage, nfExp, nfSci:
          if (FFormattingStyles[i].Decimals <> AFormat^.Decimals) then Continue;
        nfCurrency, nfCurrencyRed, nfAccounting, nfAccountingRed:
          begin
            if (FFormattingStyles[i].Decimals <> AFormat^.Decimals) then Continue;
            if (FFormattingStyles[i].CurrencySymbol <> AFormat^.CurrencySymbol) then Continue;
          end;
        nfShortDateTime, nfShortDate, nfLongDate, nfShortTime, nfLongTime,
        nfShortTimeAM, nfLongTimeAM, nfFmtDateTime, nfTimeInterval, nfCustom:
          if (FFormattingstyles[i].NumberFormatStr <> AFormat^.NumberFormatStr) then Continue;
      end;
    end;

    if uffFont in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].FontIndex <> AFormat^.FontIndex) then Continue;

    // If we arrived here it means that the styles match
    Exit(i);
  end;
end;

{ If formatting features of a cell are not supported by the destination file
  format of the writer, here is the place to apply replacements.
  Must be overridden by descendants. See BIFF2 }
procedure TsCustomSpreadWriter.FixFormat(ACell: PCell);
begin
  // to be overridden
end;

{ Each descendent should define its own default formats, if any.
  Always add the normal, unformatted style first to speed things up. }
procedure TsCustomSpreadWriter.AddDefaultFormats();
begin
  SetLength(FFormattingStyles, 0);
  NextXFIndex := 0;
end;

procedure TsCustomSpreadWriter.CreateNumFormatList;
begin
  { The format list needs to be created by descendants who know about the
    special requirements of the file format. }
end;

procedure TsCustomSpreadWriter.ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
var
  Len: Integer;
begin
  FixFormat(ACell);

  if ACell^.UsedFormattingFields = [] then Exit;
  if FindFormattingInList(ACell) <> -1 then Exit;

  Len := Length(FFormattingStyles);
  SetLength(FFormattingStyles, Len+1);
  FFormattingStyles[Len] := ACell^;

  // Some built-in number formats do not write the format string to the cell
  // But the FormattingStyles need it for comparison later. --> Add the format string.
  if IsDateTimeFormat(FFormattingStyles[Len].NumberFormat) then
    FFormattingStyles[Len].NumberFormatStr := BuildDateTimeFormatString(
      FFormattingStyles[Len].NumberFormat,
      Workbook.FormatSettings,
      FFormattingStyles[Len].NumberFormatStr
    )
  else
  if FFormattingStyles[Len].NumberFormat <> nfCustom then
    FFormattingstyles[Len].NumberFormatStr := BuildNumberFormatString(
      FFormattingStyles[Len].NumberFormat,
      Workbook.FormatSettings,
      FFormattingStyles[Len].Decimals,
      FFormattingStyles[Len].CurrencySymbol
    );

  // We store the index of the XF record that will be assigned to this style in
  // the "row" of the style. Will be needed when writing the XF record.
  FFormattingStyles[Len].Row := NextXFIndex;
  Inc(NextXFIndex);
end;

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

  // Convert the numberformats of the collected styles to be compatible with the destination file
  for i:=0 to High(FFormattingStyles) do
    if (FFormattingStyles[i].NumberFormatStr <> '') and
       (FFormattingStyles[i].NumberFormat <> nfCustom)   // don't touch custom formatstrings!
    then
      FNumFormatList.ConvertBeforeWriting(
        FFormattingStyles[i].NumberFormatStr,
        FFormattingStyles[i].NumberFormat,
        FFormattingStyles[i].Decimals,
        FFormattingStyles[i].CurrencySymbol
      );
end;

{@@
  Adds the number format of the given cell to the NumFormatList, but only if
  it does not yet exist in the list.
}
procedure TsCustomSpreadWriter.ListAllNumFormatsCallback(ACell: PCell; AStream: TStream);
var
  fmt: string;
  nf: TsNumberFormat;
  decs: Byte;
  cs: String;
begin
  if ACell^.NumberFormat = nfGeneral then
    exit;

  // The builtin format list is in "file syntax", but the format string of the
  // cells are in "fpc syntax". Therefore, before seeking, we have to convert
  // the format string of the cell to "file syntax".
  fmt := ACell^.NumberFormatStr;
  nf := ACell^.NumberFormat;
  decs := ACell^.Decimals;
  cs := ACell^.CurrencySymbol;
  if (nf <> nfCustom) then begin
    if IsDateTimeFormat(nf) then
      fmt := BuildDateTimeFormatString(nf, Workbook.FormatSettings, fmt)
    else
      fmt := BuildNumberFormatString(nf, Workbook.FormatSettings, decs, cs);
    FNumFormatList.ConvertBeforeWriting(fmt, nf, decs, cs);
  end;

  // Seek the format string in the current number format list.
  // If not found add the format to the list.
  if FNumFormatList.Find(fmt) = -1 then
    FNumFormatList.AddFormat(fmt, nf, decs, cs);
end;

{@@
  Iterats through all cells and collects the number formats in
  FNumFormatList (without duplicates).
  The index of the list item is needed for the field FormatIndex of the XF record. }
procedure TsCustomSpreadWriter.ListAllNumFormats;
var
  i: Integer;
begin
  for i:=0 to Workbook.GetWorksheetCount-1 do
    IterateThroughCells(nil, Workbook.GetWorksheetByIndex(i).Cells, ListAllNumFormatsCallback);
  NumFormatList.Sort;
end;

{@@
  Expands a formula, separating it in it's constituent parts,
  so that it is already partially parsed and it is easier to
  convert it into the format supported by the writer module
}
function TsCustomSpreadWriter.ExpandFormula(AFormula: TsFormula): TsExpandedFormula;
var
  StrPos: Integer;
  ResPos: Integer;
begin
  ResPos := -1;
  SetLength(Result, 0);

  // The formula needs to start with a "=" character.
  if AFormula.FormulaStr[1] <> '=' then raise Exception.Create('Formula doesn''t start with =');

  StrPos := 2;

  while Length(AFormula.FormulaStr) <= StrPos do
  begin
    // Checks for cell with the format [Letter][Number]
{    if (AFormula.FormulaStr[StrPos] in [a..zA..Z]) and
       (AFormula.FormulaStr[StrPos + 1] in [0..9]) then
    begin
      Inc(ResPos);
      SetLength(Result, ResPos + 1);
      Result[ResPos].ElementKind := fekCell;
//      Result[ResPos].Col1 := fekCell;
      Result[ResPos].Row1 := AFormula.FormulaStr[StrPos + 1];

      Inc(StrPos);
    end
    // Checks for arithmetical operations
    else} if AFormula.FormulaStr[StrPos] = '+' then
    begin
      Inc(ResPos);
      SetLength(Result, ResPos + 1);
      Result[ResPos].ElementKind := fekAdd;
    end;

    Inc(StrPos);
  end;
end;

{@@
  Helper function for the spreadsheet writers.

  @see    TsCustomSpreadWriter.WriteCellsToStream
}
procedure TsCustomSpreadWriter.WriteCellCallback(ACell: PCell; AStream: TStream);
begin
  case ACell.ContentType of
    cctEmpty:      WriteBlank(AStream, ACell^.Row, ACell^.Col, ACell);
    cctDateTime:   WriteDateTime(AStream, ACell^.Row, ACell^.Col, ACell^.DateTimeValue, ACell);
    cctNumber:     WriteNumber(AStream, ACell^.Row, ACell^.Col, ACell^.NumberValue, ACell);
    cctUTF8String: WriteLabel(AStream, ACell^.Row, ACell^.Col, ACell^.UTF8StringValue, ACell);
    cctFormula:    WriteFormula(AStream, ACell^.Row, ACell^.Col, ACell^.FormulaValue, ACell);
    cctRPNFormula: WriteRPNFormula(AStream, ACell^.Row, ACell^.Col, ACell^.RPNFormulaValue, ACell);
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

  @param  AStream   The output stream, passed to the callback routine.
  @param  ACells    List of cells to be iterated
  @param  ACallback The callback routine
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

  @param  AFileName The output file name.
                    If the file already exists it will be replaced.

  @see    TsWorkbook
}
procedure TsCustomSpreadWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean = False);
var
  OutputFile: TFileStream;
  lMode: Word;
begin
  if AOverwriteExisting then lMode := fmCreate or fmOpenWrite
  else lMode := fmCreate;

  OutputFile := TFileStream.Create(AFileName, lMode);
  try
    WriteToStream(OutputFile);
  finally
    OutputFile.Free;
  end;
end;

{@@
  This routine should be overriden in descendent classes.
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

procedure TsCustomSpreadWriter.WriteToStrings(AStrings: TStrings);
begin
  raise Exception.Create(lpUnsupportedWriteFormat);
end;

procedure TsCustomSpreadWriter.WriteFormula(AStream: TStream; const ARow,
  ACol: Cardinal; const AFormula: TsFormula; ACell: PCell);
begin
  // Silently dump the formula; child classes should implement their own support
end;

procedure TsCustomSpreadWriter.WriteRPNFormula(AStream: TStream; const ARow,
  ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell);
begin
  // Silently dump the formula; child classes should implement their own support
end;


{ Simplified creation of RPN formulas }

function NewRPNItem: PRPNItem;
begin
  Result := GetMem(SizeOf(TRPNItem));
  FillChar(Result^.FE, SizeOf(Result^.FE), 0);
  Result^.FE.StringValue := '';
end;

procedure DisposeRPNItem(AItem: PRPNItem);
begin
  if AItem <> nil then
    FreeMem(AItem, SizeOf(TRPNItem));
end;

{@@
  Creates a boolean value entry in the RPN array.
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
}
function RPNCellValue(ACellAddress: String; ANext: PRPNItem): PRPNItem;
var
  r,c: Integer;
  flags: TsRelFlags;
begin
  if not ParseCellString(ACellAddress, r, c, flags) then
    raise Exception.CreateFmt('"%s" is not a valid cell address.', [ACellAddress]);
  Result := RPNCellValue(r,c, flags, ANext);
end;

{@@
  Creates an entry in the RPN array for a cell value, specifed by its
  row and column index and a flag containing information on relative addresses.
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
}
function RPNCellRef(ACellAddress: String; ANext: PRPNItem): PRPNItem;
var
  r,c: Integer;
  flags: TsRelFlags;
begin
  if not ParseCellString(ACellAddress, r, c, flags) then
    raise Exception.CreateFmt('"%s" is not a valid cell address.', [ACellAddress]);
  Result := RPNCellRef(r,c, flags, ANext);
end;

{@@
  Creates an entry in the RPN array for a cell reference, specifed by its
  row and column index and flags containing information on relative addresses.
  "Cell reference" means that all properties of the cell can be handled.
  Note that most Excel formulas with cells require the cell value only
  (--> RPNCellValue)
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
}
function RPNCellRange(ACellRangeAddress: String; ANext: PRPNItem): PRPNItem;
var
  r1,c1, r2,c2: Integer;
  flags: TsRelFlags;
begin
  if not ParseCellRangeString(ACellRangeAddress, r1,c1, r2,c2, flags) then
    raise Exception.CreateFmt('"%s" is not a valid cell range address.', [ACellRangeAddress]);
  Result := RPNCellRange(r1,c1, r2,c2, flags, ANext);
end;

{@@
  Creates an entry in the RPN array for a range of cells, specified by the
  row/column indexes of the top/left and bottom/right corners of the block.
  The flags indicate relative indexes.
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
  Creates an entry in the RPN array with an error value.
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
}
function RPNMissingArg(ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekMissingArg;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a number. Integers and floating-point
  values can be used likewise.
}
function RPNNumber(AValue: Double; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekNum;
  Result^.FE.DoubleValue := AValue;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array which put the curren operator in parenthesis.
  For display purposes only, does not affect calculation.
}
function RPNParenthesis(ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekParen;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for a string.
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
}
function RPNFunc(AToken: TFEKind; ANext: PRPNItem): PRPNItem;
begin
  if FEProps[AToken].MinParams <> FEProps[AToken].MaxParams then
    raise Exception.CreateFmt(lpSpecifyNumberOfParams, [FEProps[AToken].Symbol]);

  Result := RPNFunc(AToken, FEProps[AToken].MinParams, ANext);
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for an Excel function or operation
  specified by its TokenID (--> TFEKind). Specify the number of parameters used.
  They must have been created before.
}
function RPNFunc(AToken: TFEKind; ANumParams: Byte; ANext: PRPNItem): PRPNItem;
begin
  if ord(AToken) < ord(fekAdd) then
    raise Exception.Create('No basic tokens allowed here.');

  if (ANumParams < FEProps[AToken].MinParams) or (ANumParams > FEProps[AToken].MaxParams) then
    raise Exception.CreateFmt(lpIncorrectParamCount, [
      FEProps[AToken].Symbol, FEProps[AToken].MinParams, FEProps[AToken].MaxParams
    ]);

  Result := NewRPNItem;
  Result^.FE.ElementKind := AToken;
  Result^.FE.ParamsNum := ANumParams;
  Result^.Next := ANext;
end;

{@@
  Returns if the function defined by the token requires a fixed number of parameter.
}
function FixedParamCount(AElementKind: TFEKind): Boolean;
begin
  Result := (FEProps[AElementKind].MinParams = FEProps[AElementKind].MaxParams)
        and (FEProps[AElementKind].MinParams >= 0);
end;

{@@
  Creates an RPN formula by a single call using nested RPN items.
}
function CreateRPNFormula(AItem: PRPNItem): TsRPNFormula;
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
  n := 0;
  while item <> nil do begin
    nextitem := item^.Next;
    Result[n] := item^.FE;
    inc(n);
    DisposeRPNItem(item);
    item := nextitem;
  end;
end;

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
  The exception is for the format nfCustom for which the format strings are
  left untouched.

- For writing, the writer creates a TsNumFormatList which stores all formats
  in file syntax.
  - The built-in formats of the file types are coded in the file syntax.
  - The method "ConvertBeforeWriting" converts the cell formats from the
    fpspreadsheet to the file syntax.

- For reading, the reader creates another TsNumFormatList.
  - The built-in formats of the file types are coded again in file syntax.
  - The formats read from the file are added in file syntax.
  - After reading, the formats are converted to fpspreadsheet syntax
    ("ConvertAfterReading").
}

