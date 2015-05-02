unit xlscommon;

{ Comments often have links to sections in the
OpenOffice Microsoft Excel File Format document }

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

interface

uses
  Classes, SysUtils, DateUtils, lconvencoding,
  fpsTypes, fpSpreadsheet, fpsUtils, fpsNumFormatParser,
  fpsReaderWriter;

const
  { RECORD IDs which didn't change across versions 2-8 }
  INT_EXCEL_ID_EOF         = $000A;
  INT_EXCEL_ID_HEADER      = $0014;
  INT_EXCEL_ID_FOOTER      = $0015;
  INT_EXCEL_ID_NOTE        = $001C;
  INT_EXCEL_ID_SELECTION   = $001D;
  INT_EXCEL_ID_DATEMODE    = $0022;
  INT_EXCEL_ID_LEFTMARGIN  = $0026;
  INT_EXCEL_ID_RIGHTMARGIN = $0027;
  INT_EXCEL_ID_TOPMARGIN   = $0028;
  INT_EXCEL_ID_BOTTOMMARGIN= $0029;
  INT_EXCEL_ID_PRINTHEADERS= $002A;
  INT_EXCEL_ID_PRINTGRID   = $002B;
  INT_EXCEL_ID_CONTINUE    = $003C;
  INT_EXCEL_ID_WINDOW1     = $003D;
  INT_EXCEL_ID_PANE        = $0041;
  INT_EXCEL_ID_CODEPAGE    = $0042;
  INT_EXCEL_ID_DEFCOLWIDTH = $0055;

  { RECORD IDs which did not change across versions 2, 5, 8}
  INT_EXCEL_ID_FORMULA     = $0006;    // BIFF3: $0206, BIFF4: $0406
  INT_EXCEL_ID_FONT        = $0031;    // BIFF3-4: $0231

  { RECORD IDs which did not change across version 3-8}
  INT_EXCEL_ID_COLINFO     = $007D;    // does not exist in BIFF2
  INT_EXCEL_ID_SHEETPR     = $0081;    // does not exist in BIFF2
  INT_EXCEL_ID_HCENTER     = $0083;    // does not exist in BIFF2
  INT_EXCEL_ID_VCENTER     = $0084;    // does not exist in BIFF2
  INT_EXCEL_ID_COUNTRY     = $008C;    // does not exist in BIFF2
  INT_EXCEL_ID_PALETTE     = $0092;    // does not exist in BIFF2
  INT_EXCEL_ID_DIMENSIONS  = $0200;    // BIFF2: $0000
  INT_EXCEL_ID_BLANK       = $0201;    // BIFF2: $0001
  INT_EXCEL_ID_NUMBER      = $0203;    // BIFF2: $0003
  INT_EXCEL_ID_LABEL       = $0204;    // BIFF2: $0004
  INT_EXCEL_ID_BOOLERROR   = $0205;    // BIFF2: $0005
  INT_EXCEL_ID_STRING      = $0207;    // BIFF2: $0007
  INT_EXCEL_ID_ROW         = $0208;    // BIFF2: $0008
  INT_EXCEL_ID_INDEX       = $020B;    // BIFF2: $000B
  INT_EXCEL_ID_DEFROWHEIGHT= $0225;    // BIFF2: $0025
  INT_EXCEL_ID_WINDOW2     = $023E;    // BIFF2: $003E
  INT_EXCEL_ID_RK          = $027E;    // does not exist in BIFF2
  INT_EXCEL_ID_STYLE       = $0293;    // does not exist in BIFF2

  { RECORD IDs which did not change across version 4-8 }
  INT_EXCEL_ID_PAGESETUP   = $00A1;    // does not exist before BIFF4
  INT_EXCEL_ID_FORMAT      = $041E;    // BIFF2-3: $001E

  { RECORD IDs which did not change across versions 5-8 }
  INT_EXCEL_ID_OBJ         = $005D;    // does not exist before BIFF5
  INT_EXCEL_ID_BOUNDSHEET  = $0085;    // Renamed to SHEET in the latest OpenOffice docs, does not exist before 5
  INT_EXCEL_ID_MULRK       = $00BD;    // does not exist before BIFF5
  INT_EXCEL_ID_MULBLANK    = $00BE;    // does not exist before BIFF5
  INT_EXCEL_ID_XF          = $00E0;    // BIFF2:$0043, BIFF3:$0243, BIFF4:$0443
  INT_EXCEL_ID_RSTRING     = $00D6;    // does not exist before BIFF5
  INT_EXCEL_ID_SHAREDFMLA  = $04BC;    // does not exist before BIFF5
  INT_EXCEL_ID_BOF         = $0809;    // BIFF2:$0009, BIFF3:$0209; BIFF4:$0409

  { FONT record constants }
  INT_FONT_WEIGHT_NORMAL   = $0190;
  INT_FONT_WEIGHT_BOLD     = $02BC;

  { CODEPAGE record constants }
  WORD_ASCII               = 367;
  WORD_CP_437_DOS_US       = 437;
  WORD_CP_850_DOS_Latin1   = 850;
  WORD_CP_852_DOS_Latin2   = 852;
  WORD_CP_866_DOS_Cyrillic = 866;
  WORD_CP_874_Thai         = 874;
  WORD_UTF_16              = 1200; // BIFF 8
  WORD_CP_1250_Latin2      = 1250;
  WORD_CP_1251_Cyrillic    = 1251;
  WORD_CP_1252_Latin1      = 1252; // BIFF4-BIFF5
  WORD_CP_1253_Greek       = 1253;
  WORD_CP_1254_Turkish     = 1254;
  WORD_CP_1255_Hebrew      = 1255;
  WORD_CP_1256_Arabic      = 1256;
  WORD_CP_1257_Baltic      = 1257;
  WORD_CP_1258_Vietnamese  = 1258;
  WORD_CP_1258_Latin1_BIFF2_3 = 32769; // BIFF2-BIFF3

  { DATEMODE record, 5.28 }
  DATEMODE_1900_BASE       = 1; //1/1/1900 minus 1 day in FPC TDateTime
  DATEMODE_1904_BASE       = 1462; //1/1/1904 in FPC TDateTime

  { WINDOW1 record constants - BIFF5-BIFF8 }
  MASK_WINDOW1_OPTION_WINDOW_HIDDEN             = $0001;
  MASK_WINDOW1_OPTION_WINDOW_MINIMISED          = $0002;
  MASK_WINDOW1_OPTION_HORZ_SCROLL_VISIBLE       = $0008;
  MASK_WINDOW1_OPTION_VERT_SCROLL_VISIBLE       = $0010;
  MASK_WINDOW1_OPTION_WORKSHEET_TAB_VISIBLE     = $0020;

  { WINDOW2 record constants - BIFF3-BIFF8 }
  MASK_WINDOW2_OPTION_SHOW_FORMULAS             = $0001;
  MASK_WINDOW2_OPTION_SHOW_GRID_LINES           = $0002;
  MASK_WINDOW2_OPTION_SHOW_SHEET_HEADERS        = $0004;
  MASK_WINDOW2_OPTION_PANES_ARE_FROZEN          = $0008;
  MASK_WINDOW2_OPTION_SHOW_ZERO_VALUES          = $0010;
  MASK_WINDOW2_OPTION_AUTO_GRIDLINE_COLOR       = $0020;
  MASK_WINDOW2_OPTION_COLUMNS_RIGHT_TO_LEFT     = $0040;
  MASK_WINDOW2_OPTION_SHOW_OUTLINE_SYMBOLS      = $0080;
  MASK_WINDOW2_OPTION_REMOVE_SPLITS_ON_UNFREEZE = $0100;  //BIFF5-BIFF8
  MASK_WINDOW2_OPTION_SHEET_SELECTED            = $0200;  //BIFF5-BIFF8
  MASK_WINDOW2_OPTION_SHEET_ACTIVE              = $0400;  //BIFF5-BIFF8

  { XF substructures }

  { XF_TYPE_PROT - XF Type and Cell protection (3 Bits) - BIFF3-BIFF8 }
  MASK_XF_TYPE_PROT_LOCKED               = $1;
  MASK_XF_TYPE_PROT_FORMULA_HIDDEN       = $2;
  MASK_XF_TYPE_PROT_STYLE_XF             = $4; // 0 = CELL XF

  { XF_USED_ATTRIB - Attributes from parent Style XF (6 Bits) - BIFF3-BIFF8

    - In a CELL XF a cleared bit means that the parent attribute is used,
      while a set bit indicates that the data in this XF is used
    - In a STYLE XF a cleared bit means that the data in this XF is used,
      while a set bit indicates that the attribute should be ignored }

  MASK_XF_USED_ATTRIB_NUMBER_FORMAT      = $01;
  MASK_XF_USED_ATTRIB_FONT               = $02;
  MASK_XF_USED_ATTRIB_TEXT               = $04;
  MASK_XF_USED_ATTRIB_BORDER_LINES       = $08;
  MASK_XF_USED_ATTRIB_BACKGROUND         = $10;
  MASK_XF_USED_ATTRIB_CELL_PROTECTION    = $20;
  { the following values do not agree with the documentation !!!
  MASK_XF_USED_ATTRIB_NUMBER_FORMAT      = $04;
  MASK_XF_USED_ATTRIB_FONT               = $08;
  MASK_XF_USED_ATTRIB_TEXT               = $10;
  MASK_XF_USED_ATTRIB_BORDER_LINES       = $20;
  MASK_XF_USED_ATTRIB_BACKGROUND         = $40;
  MASK_XF_USED_ATTRIB_CELL_PROTECTION    = $80;         }

  { XF record constants }
  MASK_XF_TYPE_PROT                      = $0007;
  MASK_XF_TYPE_PROT_PARENT               = $FFF0;

  MASK_XF_HOR_ALIGN                      = $07;
  MASK_XF_VERT_ALIGN                     = $70;
  MASK_XF_TEXTWRAP                       = $08;

  { XF HORIZONTAL ALIGN }
  MASK_XF_HOR_ALIGN_LEFT                 = $01;
  MASK_XF_HOR_ALIGN_CENTER               = $02;
  MASK_XF_HOR_ALIGN_RIGHT                = $03;
  MASK_XF_HOR_ALIGN_FILLED               = $04;
  MASK_XF_HOR_ALIGN_JUSTIFIED            = $05;  // BIFF4-BIFF8
  MASK_XF_HOR_ALIGN_CENTERED_SELECTION   = $06;  // BIFF4-BIFF8
  MASK_XF_HOR_ALIGN_DISTRIBUTED          = $07;  // BIFF8

  { XF_VERT_ALIGN }
  MASK_XF_VERT_ALIGN_TOP                 = $00;
  MASK_XF_VERT_ALIGN_CENTER              = $10;
  MASK_XF_VERT_ALIGN_BOTTOM              = $20;
  MASK_XF_VERT_ALIGN_JUSTIFIED           = $30;

  { XF FILL PATTERNS }
  MASK_XF_FILL_PATT_EMPTY                = $00;
  MASK_XF_FILL_PATT_SOLID                = $01;

  MASK_XF_FILL_PATT: array[TsFillStyle] of Byte = (
    $00, // fsNoFill
    $01, // fsSolidFill
    $03, // fsGray75
    $02, // fsGray50
    $04, // fsGray25
    $11, // fsGray12
    $12, // fsGray6,
    $05, // fsStripeHor
    $06, // fsStripeVert
    $08, // fsStripeDiagUp
    $07, // fsStripeDiagDown
    $0B, // fsThinStripeHor
    $0C, // fsThinStripeVert
    $0E, // fsThinStripeDiagUp
    $0D, // fsThinStripeDiagDown
    $09, // fsHatchDiag
    $10, // fsThinHatchDiag
    $0A, // fsThickHatchDiag
    $0F  // fsThinHatchHor
  );

  { Cell Addresses constants, valid for BIFF2-BIFF5 }
  MASK_EXCEL_ROW                         = $3FFF;
  MASK_EXCEL_RELATIVE_COL                = $4000;
  MASK_EXCEL_RELATIVE_ROW                = $8000;
  { Note: The assignment of the RELATIVE_COL and _ROW masks is according to
    Microsoft's documentation, but opposite to the OpenOffice documentation. }

  { FORMULA record constants }
  MASK_FORMULA_RECALCULATE_ALWAYS        = $0001;
  MASK_FORMULA_RECALCULATE_ON_OPEN       = $0002;
  MASK_FORMULA_SHARED_FORMULA            = $0008;

  { System colors, for BIFF5-BIFF8 }
  SYS_DEFAULT_FOREGROUND_COLOR           = $0040;
  SYS_DEFAULT_BACKGROUND_COLOR           = $0041;


  { Error codes }
  ERR_INTERSECTION_EMPTY                 = $00;  // #NULL!
  ERR_DIVIDE_BY_ZERO                     = $07;  // #DIV/0!
  ERR_WRONG_TYPE_OF_OPERAND              = $0F;  // #VALUE!
  ERR_ILLEGAL_REFERENCE                  = $17;  // #REF!
  ERR_WRONG_NAME                         = $1D;  // #NAME?
  ERR_OVERFLOW                           = $24;  // #NUM!
  ERR_ARG_ERROR                          = $2A;  // #N/A (not enough, or too many, arguments)

  { Index of last built-in XF format record }
  LAST_BUILTIN_XF                        = 15;

  PAPER_SIZES: array[0..90] of array[0..1] of Double = (  // Dimensions in mm
    (        0.0  ,    0.0       ),  // 0 - undefined
    (2.54*   8.5  ,   11.0  *2.54),  // 1 - Letter
    (2.54*   8.5  ,   11.0  *2.54),  // 2 - Letter small
    (2.54*  11.0  ,   17.0  *2.54),  // 3 - Tabloid
    (2.54*  17.0  ,   11.0  *2.54),  // 4 - Ledger
    (2.54*   8.5  ,   14.0  *2.54),  // 5 - Legal
    (2.54*   5.5  ,    8.5  *2.54),  // 6 - Statement
    (2.54*   7.25 ,   10.5  *2.54),  // 7 - Executive
    (      297.0  ,  420.0       ),  // 8 - A3
    (      210.0  ,  297.0       ),  // 9 - A4
    (      210.0  ,  297.0       ),  // 10 - A4 small
    (      148.0  ,  210.0       ),  // 11 - A5
    (      257.0  ,  364.0       ),  // 12 - B4 (JIS)
    (      182.0  ,  257.0       ),  // 13 - B5 (JIS)
    (2.54*   8.5  ,   13.0  *2.54),  // 14 - Folie
    (      215.0  ,  275.0       ),  // 15 - Quarto
    (2.54*  10.0  ,   14.0  *2.54),  // 16 - 10x14
    (2.54*  11.0  ,   17.0  *2.54),  // 17 - 11x17
    (2.54*   8.5  ,   11.0  *2.54),  // 18 - Note
    (2.54*   3.875,    8.875*2.54),  // 19 - Envelope #9
    (2.54*   4.125,    9.5  *2.54),  // 20 - Envelope #10
    (2.54*   4.5  ,   10.375*2.54),  // 21 - Envelope #11
    (2.54*   4.75 ,   11.0  *2.54),  // 22 - Envelope #12
    (2.54*   5.0  ,   11.5  *2.54),  // 23 - Envelope #14
    (2.54*  17.0  ,   22.0  *2.54),  // 24 - C
    (2.54*  22.0  ,   34.0  *2.54),  // 25 - D
    (2.54*  34.0  ,   44.0  *2.54),  // 26 - E
    (      110.0  ,  220.0       ),  // 27 - Envelope DL
    (      162.0  ,  229.0       ),  // 28 - Envelope C5
    (      324.0  ,  458.0       ),  // 29 - Envelope C3
    (      229.0  ,  324.0       ),  // 30 - Envelope C4
    (      114.0  ,  162.0       ),  // 31 - Envelope C6
    (      114.0  ,  229.0       ),  // 32 - Envelope C6/C5
    (      250.0  ,  353.0       ),  // 33 - B4 (ISO)
    (      176.0  ,  250.0       ),  // 34 - B5 (ISO)
    (      125.0  ,  176.0       ),  // 35 - B6 (ISO)
    (      110.0  ,  230.0       ),  // 36 - Envelope Italy
    (2.54*   3.875,    7.5  *2.54),  // 37 - Envelope Monarch
    (2.54*   3.625,    6.5  *2.54),  // 38 - 6 3/4 Envelope
    (2.54*  14.875,   11.0  *2.54),  // 39 - US Standard Fanfold
    (2.54*   8.5  ,   12.0  *2.54),  // 40 - German Std Fanfold
    (2.54*   8.5  ,   13.0  *2.54),  // 41 - German Legal Fanfold
    (      250.0  ,  353.0       ),  // 42 - B4 (ISO)
    (      100.0  ,  148.0       ),  // 43 - Japanese Postcard
    (2.54*   9.0  ,   11.0  *2.54),  // 44 - 9x11
    (2.54*  10.0  ,   11.0  *2.54),  // 45 - 10x11
    (2.54*  15.0  ,   11.0  *2.54),  // 46 - 15x11
    (      220.0  ,  220.0       ),  // 47 - Envelope Invite
    (        0.0  ,    0.0       ),  // 48 - undefined
    (        0.0  ,    0.0       ),  // 49 - undefined
    (2.54*   9.5  ,   11.0  *2.54),  // 50 - Letter Extra
    (2.54*   9.5  ,   15.0  *2.54),  // 51 - Legal Extra
    (2.54* 11.6875,   18.0  *2.54),  // 52 - Tabloid Extra
    (      235.0  ,  322.0       ),  // 53 - A4 Extra
    (2.54*   8.5  ,   11.0  *2.54),  // 54 - Letter Transverse
    (      210.0  ,  297.0       ),  // 55 - A4 Transverse
    (2.54*   9.5  ,   11.0  *2.54),  // 56 - Letter Extra Transverse
    (      227.0  ,  356.0       ),  // 57 - Super A/A4
    (      305.0  ,  487.0       ),  // 58 - Super B/B4
    (2.54*   8.5  ,  12.6875*2.54),  // 59 - Letter plus
    (      210.0  ,  330.0       ),  // 60 - A4 plus
    (      148.0  ,  210.0       ),  // 61 - A5 transverse
    (      182.0  ,  257.0       ),  // 62 - B5 (JIS) transverse
    (      322.0  ,  445.0       ),  // 63 - A3 Extra
    (      174.0  ,  235.0       ),  // 64 - A5 Extra
    (      201.0  ,  276.0       ),  // 65 - B5 (ISO) Extra
    (      420.0  ,  594.0       ),  // 66 - A2
    (      297.0  ,  420.0       ),  // 67 - A3 Transverse
    (      322.0  ,  445.0       ),  // 68 - A3 Extra Transverse
    (      200.0  ,  148.0       ),  // 69 - Double Japanese Postcard
    (      105.0  ,  148.0       ),  // 70 - A6
    (        0.0  ,    0.0       ),  // 71 - undefined
    (        0.0  ,    0.0       ),  // 72 - undefined
    (        0.0  ,    0.0       ),  // 73 - undefined
    (        0.0  ,    0.0       ),  // 74 - undefined
    (2.54*  11.0  ,    8.5  *2.54),  // 75 - Letter rotated
    (      420.0  ,  297.0       ),  // 76 - A3 rotated
    (      297.0  ,  210.0       ),  // 77 - A4 rotated
    (      210.0  ,  148.0       ),  // 78 - A5 rotated
    (      364.0  ,  257.0       ),  // 79 - B4 (JIS) rotated
    (      257.0  ,  182.0       ),  // 80 - B5 (JIS) rotated
    (      148.0  ,  100.0       ),  // 81 - Japanese Postcard rotated
    (      148.0  ,  200.0       ),  // 82 - Double Japanese Postcard rotated
    (      148.0  ,  105.0       ),  // 83 - A6 rotated
    (        0.0  ,    0.0       ),  // 84 - undefined
    (        0.0  ,    0.0       ),  // 85 - undefined
    (        0.0  ,    0.0       ),  // 86 - undefined
    (        0.0  ,    0.0       ),  // 87 - undefined
    (      128.0  ,  182.0       ),  // 88 - B6 (JIS)
    (      182.0  ,  128.0       ),  // 89 - B6 (JIS) rotated
    (2.54*  12.0  ,   11.0  *2.54)   // 90 - 12x11
  );

type
  TDateMode=(dm1900,dm1904); //DATEMODE values, 5.28

  // Adjusts Excel float (date, date/time, time) with the file's base date to get a TDateTime
  function ConvertExcelDateTimeToDateTime
    (const AExcelDateNum: Double; ADateMode: TDateMode): TDateTime;
  // Adjusts TDateTime with the file's base date to get
  // an Excel float value representing a time/date/datetime
  function ConvertDateTimeToExcelDateTime
    (const ADateTime: TDateTime; ADateMode: TDateMode): Double;

  // Converts the error byte read from cells or formulas to fps error value
  function ConvertFromExcelError(AValue: Byte): TsErrorValue;
  // Converts an fps error value to the byte code needed in xls files
  function ConvertToExcelError(AValue: TsErrorValue): byte;

type
  { TsBIFFHeader }
  TsBIFFHeader = packed record
    RecordID: Word;
    RecordSize: Word;
  end;

  { TsSpreadBIFFReader }
  TsSpreadBIFFReader = class(TsCustomSpreadReader)
  protected
    RecordSize: Word;
    FCodepage: string; // in a format prepared for lconvencoding.ConvertEncoding
    FDateMode: TDateMode;
    FPaletteFound: Boolean;
    FIncompleteCell: PCell;
    FIncompleteNote: String;
    FIncompleteNoteLength: Word;
    FFirstNumFormatIndexInFile: Integer;
    procedure AddBuiltinNumFormats; override;
    procedure ApplyCellFormatting(ACell: PCell; XFIndex: Word); virtual; //overload;
    // Extracts a number out of an RK value
    function DecodeRKValue(const ARK: DWORD): Double;
    // Returns the numberformat for a given XF record
    procedure ExtractNumberFormat(AXFIndex: WORD;
      out ANumberFormat: TsNumberFormat; out ANumberFormatStr: String); virtual;
    // Tries to find if a number cell is actually a date/datetime/time cell and retrieves the value
    function IsDateTime(Number: Double; ANumberFormat: TsNumberFormat;
      ANumberFormatStr: String; out ADateTime: TDateTime): Boolean;
    // Here we can add reading of records which didn't change across BIFF5-8 versions
    // Read a blank cell
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadBool(AStream: TStream); override;
    procedure ReadCodePage(AStream: TStream);
    // Read column info
    procedure ReadColInfo(const AStream: TStream);
    // Read attached comment
    procedure ReadComment(const AStream: TStream);
    // Figures out what the base year for dates is for this file
    procedure ReadDateMode(AStream: TStream);
    // Reads the default column width
    procedure ReadDefColWidth(AStream: TStream);
    // Reas the default row height
    procedure ReadDefRowHeight(AStream: TStream);
    // Read FORMAT record (cell formatting)
    procedure ReadFormat(AStream: TStream); virtual;
    // Read FORMULA record
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadHCENTER(AStream: TStream);
    procedure ReadHeaderFooter(AStream: TStream; AIsHeader: Boolean); virtual;
    procedure ReadMargin(AStream: TStream; AMargin: Integer);
    // Read multiple blank cells
    procedure ReadMulBlank(AStream: TStream);
    // Read multiple RK cells
    procedure ReadMulRKValues(const AStream: TStream);
    // Read floating point number
    procedure ReadNumber(AStream: TStream); override;
    // Read palette
    procedure ReadPalette(AStream: TStream);
    // Read page setup
    procedure ReadPageSetup(AStream: TStream);
    // Read PANE record
    procedure ReadPane(AStream: TStream);
    procedure ReadPrintGridLines(AStream: TStream);
    procedure ReadPrintHeaders(AStream: TStream);
    // Read an RK value cell
    procedure ReadRKValue(AStream: TStream);
    // Read the row, column, and XF index at the current stream position
    procedure ReadRowColXF(AStream: TStream; out ARow, ACol: Cardinal; out AXF: Word); virtual;
    // Read row info
    procedure ReadRowInfo(AStream: TStream); virtual;
    // Read the array of RPN tokens of a formula
    procedure ReadRPNCellAddress(AStream: TStream; out ARow, ACol: Cardinal;
      out AFlags: TsRelFlags); virtual;
    procedure ReadRPNCellAddressOffset(AStream: TStream;
      out ARowOffset, AColOffset: Integer; out AFlags: TsRelFlags); virtual;
    procedure ReadRPNCellRangeAddress(AStream: TStream;
      out ARow1, ACol1, ARow2, ACol2: Cardinal; out AFlags: TsRelFlags); virtual;
    procedure ReadRPNCellRangeOffset(AStream: TStream;
      out ARow1Offset, ACol1Offset, ARow2Offset, ACol2Offset: Integer;
      out AFlags: TsRelFlags); virtual;
    function ReadRPNFunc(AStream: TStream): Word; virtual;
    procedure ReadRPNSharedFormulaBase(AStream: TStream; out ARow, ACol: Cardinal); virtual;
    function ReadRPNTokenArray(AStream: TStream; ACell: PCell;
      ASharedFormulaBase: PCell = nil): Boolean;
    function ReadRPNTokenArraySize(AStream: TStream): word; virtual;
    procedure ReadSharedFormula(AStream: TStream);

    // Helper function for reading a string with 8-bit length
    function ReadString_8bitLen(AStream: TStream): String; virtual;
    // Read STRING record (result of string formula)
    procedure ReadStringRecord(AStream: TStream); virtual;
    procedure ReadVCENTER(AStream: TStream);
    // Read WINDOW2 record (gridlines, sheet headers)
    procedure ReadWindow2(AStream: TStream); virtual;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
  end;


  { TsSpreadBIFFWriter }

  TsSpreadBIFFWriter = class(TsCustomSpreadWriter)
  protected
    FDateMode: TDateMode;
    FCodePage: String;  // in a format prepared for lconvencoding.ConvertEncoding
    FFirstNumFormatIndexInFile: Integer;
    procedure AddBuiltinNumFormats; override;
    function FindXFIndex(ACell: PCell): Integer; virtual;
    function FixColor(AColor: TsColor): TsColor; override;
    function GetLastRowIndex(AWorksheet: TsWorksheet): Integer;
    function GetLastColIndex(AWorksheet: TsWorksheet): Word;
    function GetPrintOptions: Word; virtual;

    // Helper function for writing the BIFF header
    procedure WriteBIFFHeader(AStream: TStream; ARecID, ARecSize: Word);
    // Helper function for writing a string with 8-bit length }
    function WriteString_8BitLen(AStream: TStream; AString: String): Integer; virtual;

    // Write out BLANK cell record
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    // Write out BOOLEAN cell record
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
    // Writes out used codepage for character encoding
    procedure WriteCodePage(AStream: TStream; ACodePage: String); virtual;
    // Writes out column info(s)
    procedure WriteColInfo(AStream: TStream; ACol: PCol);
    procedure WriteColInfos(AStream: TStream; ASheet: TsWorksheet);
    // Writes out NOTE record(s)
    procedure WriteComment(AStream: TStream; ACell: PCell); override;
    // Writes out DATEMODE record depending on FDateMode
    procedure WriteDateMode(AStream: TStream);
    // Writes out a TIME/DATE/TIMETIME
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); override;
    // Writes out ERROR cell record
    procedure WriteError(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TsErrorValue; ACell: PCell); override;
    // Writes out a FORMULA record; formula is stored in cell already
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteHCenter(AStream: TStream);
    procedure WriteHeaderFooter(AStream: TStream; AIsHeader: Boolean); virtual;
    // Writes out page margin for printing
    procedure WriteMARGIN(AStream: TStream; AMargin: Integer);
    // Writes out a FORMAT record
    procedure WriteNumFormat(AStream: TStream; ANumFormatStr: String;
      ANumFormatIndex: Integer); virtual;
    // Writes out all FORMAT records
    procedure WriteNumFormats(AStream: TStream);
    // Writes out a floating point NUMBER record
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Double; ACell: PCell); override;
    procedure WritePageSetup(AStream: TStream);
    // Writes out a PALETTE record containing all colors defined in the workbook
    procedure WritePalette(AStream: TStream);
    // Writes out a PANE record
    procedure WritePane(AStream: TStream; ASheet: TsWorksheet; IsBiff58: Boolean;
      out ActivePane: Byte);
    // Writes out whether grid lines are printed
    procedure WritePrintGridLines(AStream: TStream);
    procedure WritePrintHeaders(AStream: TStream);
    // Writes out a ROW record
    procedure WriteRow(AStream: TStream; ASheet: TsWorksheet;
      ARowIndex, AFirstColIndex, ALastColIndex: Cardinal; ARow: PRow); virtual;
    // Write all ROW records for a sheet
    procedure WriteRows(AStream: TStream; ASheet: TsWorksheet);

    function WriteRPNCellAddress(AStream: TStream; ARow, ACol: Cardinal;
      AFlags: TsRelFlags): Word; virtual;
    function WriteRPNCellOffset(AStream: TStream; ARowOffset, AColOffset: Integer;
      AFlags: TsRelFlags): Word; virtual;
    function WriteRPNCellRangeAddress(AStream: TStream; ARow1, ACol1, ARow2, ACol2: Cardinal;
      AFlags: TsRelFlags): Word; virtual;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal;
      const AFormula: TsRPNFormula; ACell: PCell); virtual;
    function WriteRPNFunc(AStream: TStream; AIdentifier: Word): Word; virtual;
    procedure WriteRPNResult(AStream: TStream; ACell: PCell);
    {
    procedure WriteRPNSharedFormulaLink(AStream: TStream; ACell: PCell;
      var RPNLength: Word); virtual;
      }
    procedure WriteRPNTokenArray(AStream: TStream; ACell: PCell;
      const AFormula: TsRPNFormula; UseRelAddr: Boolean; var RPNLength: Word);
    procedure WriteRPNTokenArraySize(AStream: TStream; ASize: Word); virtual;

    // Writes out a SELECTION record
    procedure WriteSelection(AStream: TStream; ASheet: TsWorksheet; APane: Byte);
    procedure WriteSelections(AStream: TStream; ASheet: TsWorksheet);
    (*
    // Writes out a shared formula
    procedure WriteSharedFormula(AStream: TStream; ACell: PCell); virtual;
    procedure WriteSharedFormulaRange(AStream: TStream;
      AFirstRow, AFirstCol, ALastRow, ALastCol: Cardinal); virtual;
      *)
    procedure WriteSheetPR(AStream: TStream);
    procedure WriteStringRecord(AStream: TStream; AString: String); virtual;
    procedure WriteVCenter(AStream: TStream);
    // Writes cell content received by workbook in OnNeedCellData event
    procedure WriteVirtualCells(AStream: TStream);
    // Writes out a WINDOW1 record
    procedure WriteWindow1(AStream: TStream); virtual;
    // Writes an XF record
    procedure WriteXF(AStream: TStream; ACellFormat: PsCellFormat;
      XFType_Prot: Byte = 0); virtual;
    // Writes all XF records
    procedure WriteXFRecords(AStream: TStream);

  public
    constructor Create(AWorkbook: TsWorkbook); override;
  end;

procedure AddBuiltinBiffFormats(AList: TStringList;
  AFormatSettings: TFormatSettings; ALastIndex: Integer);


implementation

uses
  AVL_Tree, Math, Variants,
  {%H-}fpspatches, fpsStrings, xlsConst, fpsrpn, fpsExprParser;

const
  { Helper table for rpn formulas:
    Assignment of FormulaElementKinds (fekXXXX) to EXCEL_TOKEN IDs. }
  TokenIDs: array[TFEKind] of Word = (
    // Basic operands
    INT_EXCEL_TOKEN_TREFV,          {fekCell}
    INT_EXCEL_TOKEN_TREFR,          {fekCellRef}
    INT_EXCEL_TOKEN_TAREA_R,        {fekCellRange}
    INT_EXCEL_TOKEN_TREFN_V,        {fekCellOffset}
    INT_EXCEL_TOKEN_TNUM,           {fekNum}
    INT_EXCEL_TOKEN_TINT,           {fekInteger}
    INT_EXCEL_TOKEN_TSTR,           {fekString}
    INT_EXCEL_TOKEN_TBOOL,          {fekBool}
    INT_EXCEL_TOKEN_TERR,           {fekErr}
    INT_EXCEL_TOKEN_TMISSARG,       {fekMissArg, missing argument}

    // Basic operations
    INT_EXCEL_TOKEN_TADD,           {fekAdd, +}
    INT_EXCEL_TOKEN_TSUB,           {fekSub, -}
    INT_EXCEL_TOKEN_TMUL,           {fekMul, *}
    INT_EXCEL_TOKEN_TDIV,           {fekDiv, /}
    INT_EXCEL_TOKEN_TPERCENT,       {fekPercent, %}
    INT_EXCEL_TOKEN_TPOWER,         {fekPower, ^}
    INT_EXCEL_TOKEN_TUMINUS,        {fekUMinus, -}
    INT_EXCEL_TOKEN_TUPLUS,         {fekUPlus, +}
    INT_EXCEL_TOKEN_TCONCAT,        {fekConcat, &, for strings}
    INT_EXCEL_TOKEN_TEQ,            {fekEqual, =}
    INT_EXCEL_TOKEN_TGT,            {fekGreater, >}
    INT_EXCEL_TOKEN_TGE,            {fekGreaterEqual, >=}
    INT_EXCEL_TOKEN_TLT,            {fekLess <}
    INT_EXCEL_TOKEN_TLE,            {fekLessEqual, <=}
    INT_EXCEL_TOKEN_TNE,            {fekNotEqual, <>}
    INT_EXCEL_TOKEN_TPAREN,         {Operator in parenthesis}
    Word(-1)                        {fekFunc}
  );

type
  TBIFF58BlankRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
  end;

  TBIFF38BoolErrRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    BoolErrValue: Byte;
    ValueType: Byte;
  end;

  TBIFF58NumberRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    Value: Double;
  end;

  TBIFF25NoteRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    TextLen: Word;
  end;


function ConvertExcelDateTimeToDateTime(const AExcelDateNum: Double;
  ADateMode: TDateMode): TDateTime;
begin
  // Time only:
  if (AExcelDateNum<1) and (AExcelDateNum>=0)  then
  begin
    Result:=AExcelDateNum;
  end
  else
  begin
    case ADateMode of
      dm1900:
        begin
          // Check for Lotus 1-2-3 bug with 1900 leap year
          if AExcelDateNum=61.0 then
            // 29 feb does not exist, change to 28
            // Spell out that we remove a day for ehm "clarity".
            result := 61.0 - 1.0 + DATEMODE_1900_BASE - 1.0
          else
            result := AExcelDateNum + DATEMODE_1900_BASE - 1.0;
        end;
      dm1904:
        result := AExcelDateNum + DATEMODE_1904_BASE;
      else
        raise Exception.CreateFmt('[ConvertExcelDateTimeToDateTime] Unknown datemode %d. Please correct fpspreadsheet source code. ', [ADateMode]);
    end;
  end;
end;

function ConvertDateTimeToExcelDateTime(const ADateTime: TDateTime;
  ADateMode: TDateMode): Double;
begin
  // Time only:
  if (ADateTime<1) and (ADateTime>=0) then
  begin
    Result:=ADateTime;
  end
  else
  begin
    case ADateMode of
    dm1900:
      result:=ADateTime-DATEMODE_1900_BASE+1.0;
    dm1904:
      result:=ADateTime-DATEMODE_1904_BASE;
    else
      raise Exception.CreateFmt('ConvertDateTimeToExcelDateTime: unknown datemode %d. Please correct fpspreadsheet source code. ', [ADateMode]);
    end;
  end;
end;

function ConvertFromExcelError(AValue: Byte): TsErrorValue;
begin
  case AValue of
    ERR_INTERSECTION_EMPTY    : Result := errEmptyIntersection;  // #NULL!
    ERR_DIVIDE_BY_ZERO        : Result := errDivideByZero;       // #DIV/0!
    ERR_WRONG_TYPE_OF_OPERAND : Result := errWrongType;          // #VALUE!
    ERR_ILLEGAL_REFERENCE     : Result := errIllegalRef;         // #REF!
    ERR_WRONG_NAME            : Result := errWrongName;          // #NAME?
    ERR_OVERFLOW              : Result := errOverflow;           // #NUM!
    ERR_ARG_ERROR             : Result := errArgError;           // #N/A!
  end;
end;

function ConvertToExcelError(AValue: TsErrorValue): byte;
begin
  case AValue of
    errEmptyIntersection : Result := ERR_INTERSECTION_EMPTY;     // #NULL!
    errDivideByZero      : Result := ERR_DIVIDE_BY_ZERO;         // #DIV/0!
    errWrongType         : Result := ERR_WRONG_TYPE_OF_OPERAND;  // #VALUE!
    errIllegalRef        : Result := ERR_ILLEGAL_REFERENCE;      // #REF!
    errWrongName         : Result := ERR_WRONG_NAME;             // #NAME?
    errOverflow          : Result := ERR_OVERFLOW;               // #NUM!
    errArgError          : Result := ERR_ARG_ERROR;              // #N/A;
  end;
end;


{@@ ----------------------------------------------------------------------------
  These are the built-in number formats as expected in the biff spreadsheet file.
  In BIFF5+ they are not written to file but they are used for lookup of the
  number format that Excel used.
-------------------------------------------------------------------------------}
procedure AddBuiltinBiffFormats(AList: TStringList;
  AFormatSettings: TFormatSettings; ALastIndex: Integer);
var
  fs: TFormatSettings absolute AFormatSettings;
  cs: String;
  i: Integer;
begin
  cs := fs.CurrencyString;
  AList.Clear;
  AList.Add('');          // 0
  AList.Add('0');         // 1
  AList.Add('0.00');      // 2
  AList.Add('#,##0');     // 3
  AList.Add('#,##0.00');  // 4
  AList.Add(BuildCurrencyFormatString(nfCurrency, fs, 0, fs.CurrencyFormat, fs.NegCurrFormat, cs));     // 5
  AList.Add(BuildCurrencyFormatString(nfCurrencyRed, fs, 0, fs.CurrencyFormat, fs.NegCurrFormat, cs));  // 6
  AList.Add(BuildCurrencyFormatString(nfCurrency, fs, 2, fs.CurrencyFormat, fs.NegCurrFormat, cs));     // 7
  AList.Add(BuildCurrencyFormatString(nfCurrencyRed, fs, 2, fs.CurrencyFormat, fs.NegCurrFormat, cs));  // 8
  AList.Add('0%');                // 9
  AList.Add('0.00%');             // 10
  AList.Add('0.00E+00');          // 11
  AList.Add('# ?/?');             // 12
  AList.Add('# ??/??');           // 13
  AList.Add(BuildDateTimeFormatString(nfShortDate, fs));     // 14
  AList.Add(BuildDateTimeFormatString(nfLongdate, fs));      // 15
  AList.Add(BuildDateTimeFormatString(nfDayMonth, fs));      // 16: 'd/mmm'
  AList.Add(BuildDateTimeFormatString(nfMonthYear, fs));     // 17: 'mmm/yy'
  AList.Add(BuildDateTimeFormatString(nfShortTimeAM, fs));   // 18
  AList.Add(BuildDateTimeFormatString(nfLongTimeAM, fs));    // 19
  AList.Add(BuildDateTimeFormatString(nfShortTime, fs));     // 20
  AList.Add(BuildDateTimeFormatString(nfLongTime, fs));      // 21
  AList.Add(BuildDateTimeFormatString(nfShortDateTime, fs)); // 22
  for i:=23 to 36 do
    AList.Add('');  // not supported
  AList.Add('_(#,##0_);(#,##0)');              // 37
  AList.Add('_(#,##0_);[Red](#,##0)');         // 38
  AList.Add('_(#,##0.00_);(#,##0.00)');        // 39
  AList.Add('_(#,##0.00_);[Red](#,##0.00)');   // 40
  AList.Add('_("'+cs+'"* #,##0_);_("'+cs+'"* (#,##0);_("'+cs+'"* "-"_);_(@_)');  // 41
  AList.Add('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)');          // 42
  AList.Add('_("'+cs+'"* #,##0.00_);_("'+cs+'"* (#,##0.00);_("'+cs+'"* "-"??_);_(@_)'); // 43
  AList.Add('_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)');  // 44
  AList.Add('nn:ss');       // 45
  AList.Add('[h]:nn:ss');   // 46
  AList.Add('nn:ss.z');     // 47
  AList.Add('##0.0E+00');   // 48
  AList.Add('');            // 49: @ ("Text") not supported
  for i:=50 to ALastIndex do AList.Add('');  // not supported/used
end;


{------------------------------------------------------------------------------}
{                           TsSpreadBIFFReader                                 }
{------------------------------------------------------------------------------}

constructor TsSpreadBIFFReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FCellFormatList := TsCellFormatList.Create(true);
  // Allow duplicates! XF indexes get out of sync if not all format records are in list
  // Initial base date in case it won't be read from file
  FDateMode := dm1900;
  // Limitations of BIFF5 and BIFF8 file format
  FLimitations.MaxColCount := 256;
  FLimitations.MaxRowCount := 65536;
  FLimitations.MaxPaletteSize := 64;
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in number formats to the NumFormatList.
  Valid for BIFF5...BIFF8. Needs to be overridden for BIFF2.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.AddBuiltinNumFormats;
begin
  FFirstNumFormatIndexInFile := 164;
  AddBuiltInBiffFormats(
    FNumFormatList, Workbook.FormatSettings, FFirstNumFormatIndexInFile-1
  );
end;

{@@ ----------------------------------------------------------------------------
  Applies the XF formatting referred to by XFIndex to the specified cell
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ApplyCellFormatting(ACell: PCell; XFIndex: Word);
var
  fmt: PsCellFormat;
  i: Integer;
begin
  if Assigned(ACell) then begin
    i := FCellFormatList.FindIndexOfID(XFIndex);
    if i > -1 then
    begin
      fmt := FCellFormatList.Items[i];
      ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt^);  // Adds a copy of fmt to workbook
    end else
      ACell^.FormatIndex := 0;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Extracts a number out of an RK value.
  Valid since BIFF3.
-------------------------------------------------------------------------------}
function TsSpreadBIFFReader.DecodeRKValue(const ARK: DWORD): Double;
var
  Number: Double;
  Tmp: LongInt;
begin
  if ARK and 2 = 2 then begin
    // Signed integer value
    if LongInt(ARK) < 0 then begin
      //Simulates a sar
      Tmp := LongInt(ARK) * (-1);
      Tmp := Tmp shr 2;
      Tmp := Tmp * (-1);
      Number := Tmp - 1;
    end else begin
      Number := ARK shr 2;
    end;
  end else begin
    // Floating point value
    // NOTE: This is endian dependent and IEEE dependent (Not checked) (working win-i386)
    (PDWORD(@Number))^ := $00000000;
    (PDWORD(@Number)+1)^ := ARK and $FFFFFFFC;
  end;
  if ARK and 1 = 1 then begin
    // Encoded value is multiplied by 100
    Number := Number / 100;
  end;
  Result := Number;
end;

{@@ ----------------------------------------------------------------------------
  Extracts number format data from an XF record index by AXFIndex.
  Valid for BIFF5-BIFF8. Needs to be overridden for BIFF2
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ExtractNumberFormat(AXFIndex: WORD;
  out ANumberFormat: TsNumberFormat; out ANumberFormatStr: String);
var
  fmt: PsCellFormat;
  i: Integer;
begin
  i := FCellFormatList.FindIndexOfID(AXFIndex);
  if i > -1 then
  begin
    fmt := FCellFormatList.Items[i];
    ANumberFormat := fmt^.NumberFormat;
    ANumberFormatStr := fmt^.NumberFormatStr;
  end else
  begin
    ANumberFormat := nfGeneral;
    ANumberFormatStr := '';
  end;
end;

{@@ ----------------------------------------------------------------------------
  Converts the number to a date/time and return that if it is
-------------------------------------------------------------------------------}
function TsSpreadBIFFReader.IsDateTime(Number: Double;
  ANumberFormat: TsNumberFormat; ANumberFormatStr: String;
  out ADateTime: TDateTime): boolean;
var
  parser: TsNumFormatParser;
begin
  Result := true;
  if ANumberFormat in [
    nfShortDateTime, nfShortDate, nfLongDate,
    nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM]
  then
    ADateTime := ConvertExcelDateTimeToDateTime(Number, FDateMode)
  else
  if ANumberFormat = nfTimeInterval then
    ADateTime := Number
  else begin
    parser := TsNumFormatParser.Create(Workbook, ANumberFormatStr);
    try
      if (parser.Status = psOK) and parser.IsDateTimeFormat then
        ADateTime := ConvertExcelDateTimeToDateTime(Number, FDateMode)
      else
        Result := false;
    finally
      parser.Free;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads a blank cell
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadBlank(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  rec: TBIFF58BlankRecord;
  cell: PCell;
begin
  rec.Row := 0;  // to silence the compiler...

  // Read entire record into a buffer
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF58BlankRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);

  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  FWorksheet.WriteBlank(cell);

  { Add attributes to cell}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  The name of this method is misleading - it reads a BOOLEAN cell value,
  but also an ERROR value; BIFF stores them in the same record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadBool(AStream: TStream);
var
  rec: TBIFF38BoolErrRecord;
  r, c: Cardinal;
  XF: Word;
  cell: PCell;
begin
  rec.Row := 0;  // to silence the compiler

  { Read entire record into a buffer }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF38BoolErrRecord) - 2*SizeOf(Word));

  r := WordLEToN(rec.Row);
  c := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);

  if FIsVirtualMode then begin
    InitCell(r, c, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(r, c);

  { Retrieve boolean or error value depending on the "ValueType" }
  case rec.ValueType of
    0: FWorksheet.WriteBoolValue(cell, boolean(rec.BoolErrValue));
    1: FWorksheet.WriteErrorValue(cell, ConvertFromExcelError(rec.BoolErrValue));
  end;

  { Add attributes to cell}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, r, c, cell);
end;

{@@ ----------------------------------------------------------------------------
  Reads the code page used in the xls file
  In BIFF8 it seams to always use the UTF-16 codepage
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadCodePage(AStream: TStream);
var
  lCodePage: Word;
begin
  { Codepage }
  lCodePage := WordLEToN(AStream.ReadWord());

  case lCodePage of
    // 016FH = 367 = ASCII
    WORD_CP_437_DOS_US: FCodePage := 'cp437';      // IBM PC CP-437 (US)
    //02D0H = 720 = IBM PC CP-720 (OEM Arabic)
    //02E1H = 737 = IBM PC CP-737 (Greek)
    //0307H = 775 = IBM PC CP-775 (Baltic)
    WORD_CP_850_DOS_Latin1: FCodepage := 'cp850';  // IBM PC CP-850 (Latin I)
    WORD_CP_852_DOS_Latin2: FCodepage := 'cp852';  // IBM PC CP-852 (Latin II (Central European))
    //035AH = 858 = IBM PC CP-858 (Multilingual Latin I with Euro)
    //035CH = 860 = IBM PC CP-860 (Portuguese)
    //035DH = 861 = IBM PC CP-861 (Icelandic)
    //035EH = 862 = IBM PC CP-862 (Hebrew)
    //035FH = 863 = IBM PC CP-863 (Canadian (French))
    //0360H = 864 = IBM PC CP-864 (Arabic)
    //0361H = 865 = IBM PC CP-865 (Nordic)
    WORD_CP_866_DOS_Cyrillic: FCodePage := 'cp866';  // IBM PC CP-866 (Cyrillic Russian)
    //0365H = 869 = IBM PC CP-869 (Greek (Modern))
    WORD_CP_874_Thai: FCodePage := 'cp874';  // 874 = Windows CP-874 (Thai)
    //03A4H = 932 = Windows CP-932 (Japanese Shift-JIS)
    //03A8H = 936 = Windows CP-936 (Chinese Simplified GBK)
    //03B5H = 949 = Windows CP-949 (Korean (Wansung))
    //03B6H = 950 = Windows CP-950 (Chinese Traditional BIG5)
    WORD_UTF_16 : FCodePage := 'ucs2le';            // UTF-16 (BIFF8)
    WORD_CP_1250_Latin2: FCodepage := 'cp1250';     // Windows CP-1250 (Latin II) (Central European)
    WORD_CP_1251_Cyrillic: FCodePage := 'cp1251';   // Windows CP-1251 (Cyrillic)
    WORD_CP_1252_Latin1: FCodePage := 'cp1252';     // Windows CP-1252 (Latin I) (BIFF4-BIFF5)
    WORD_CP_1253_Greek: FCodePage := 'cp1253';      // Windows CP-1253 (Greek)
    WORD_CP_1254_Turkish: FCodepage := 'cp1254';    // Windows CP-1254 (Turkish)
    WORD_CP_1255_Hebrew: FCodePage := 'cp1255';     // Windows CP-1255 (Hebrew)
    WORD_CP_1256_Arabic: FCodePage := 'cp1256';     // Windows CP-1256 (Arabic)
    WORD_CP_1257_Baltic: FCodePage := 'cp1257';     // Windows CP-1257 (Baltic)
    WORD_CP_1258_Vietnamese: FCodePage := 'cp1258'; // Windows CP-1258 (Vietnamese)
    //0551H = 1361 = Windows CP-1361 (Korean (Johab))
    //2710H = 10000 = Apple Roman
    //8000H = 32768 = Apple Roman
    WORD_CP_1258_Latin1_BIFF2_3: FCodePage := 'cp1252';   // Windows CP-1252 (Latin I) (BIFF2-BIFF3)
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads column info (column width) from the stream.
  Valid for BIFF3-BIFF8.
  For BIFF2 use the records COLWIDTH and COLUMNDEFAULT.
-------------------------------------------------------------------------------}
procedure TsSpreadBiffReader.ReadColInfo(const AStream: TStream);
const
  EPS = 1E-2;  // allow for large epsilon because col width calculation is not very well-defined...
var
  c, c1, c2: Cardinal;
  w: Word;
  col: TCol;
begin
  // read column start and end index of column range
  c1 := WordLEToN(AStream.ReadWord);
  c2 := WordLEToN(AStream.ReadWord);
  // read col width in 1/256 of the width of "0" character
  w := WordLEToN(AStream.ReadWord);
  // calculate width in units of "characters"
  col.Width := w / 256;
  // assign width to columns, but only if different from default column width
  if not SameValue(col.Width, FWorksheet.DefaultColWidth, EPS) then
    for c := c1 to c2 do
      FWorksheet.WriteColInfo(c, col);
end;

{@@ ----------------------------------------------------------------------------
  Reads a NOTE record which describes an attached comment
  Valid for BIFF2-BIFF5
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadComment(const AStream: TStream);
var
  rec: TBIFF25NoteRecord;
  r, c: Cardinal;
  n: Word;
  s: ansiString;
  List: TStringList;
begin
  rec.Row := 0; // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF25NoteRecord) - 2*SizeOf(Word));
  r := WordLEToN(rec.Row);
  c := WordLEToN(rec.Col);
  n := WordLEToN(rec.TextLen);
  // First NOTE record
  if r <> $FFFF then
  begin
    // entire note is in this record
    if n <= self.RecordSize - 3*SizeOf(word) then
    begin
      SetLength(s, n);
      AStream.ReadBuffer(s[1], n);
      FIncompleteNote := '';
      FIncompleteNoteLength := 0;
      List := TStringList.Create;
      try
        List.Text := s;  // Fix line endings which are #10 in file
        s := Copy(List.Text, 1, Length(List.Text) - Length(LineEnding));
        s := ConvertEncoding(s, FCodePage, encodingUTF8);
        FWorksheet.WriteComment(r, c, s);
      finally
        List.Free;
      end;
    end else
    // note will be continued in following record(s): Store partial string
    begin
      FIncompleteNoteLength := n;
      n := self.RecordSize - 3*SizeOf(Word);
      SetLength(s, n);
      AStream.ReadBuffer(s[1], n);
      FIncompleteNote := s;
      FIncompleteCell := FWorksheet.GetCell(r, c);  // no AddCell here!
    end;
  end else
  // One of the continuation records
  begin
    SetLength(s, n);
    AStream.ReadBuffer(s[1], n);
    FIncompleteNote := FIncompleteNote + s;
    // last continuation record
    if Length(FIncompleteNote) = FIncompleteNoteLength then
    begin
      List := TStringList.Create;
      try
        List.Text := FIncompleteNote;    // Fix line endings which are #10 in file
        s := Copy(List.Text, 1, Length(List.Text) - Length(LineEnding));
        FWorksheet.WriteComment(FIncompleteCell, s);
      finally
        List.Free;
      end;
      FIncompleteNote := '';
      FIncompleteCell := nil;
      FIncompleteNoteLength := 0;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This record specifies the base date for displaying date values.
  All dates are stored as count of days past this base date.
  In BIFF2-BIFF4 this record is part of the Calculation Settings Block.
  In BIFF5-BIFF8 it is stored in the Workbookk Globals Substream.

  Record DATEMODE, BIFF2-BIFF8:
   Offset Size Contents
     0     2   0 = Base date is 1899-Dec-31 (the cell value 1 represents 1900-Jan-01)
               1 = Base date is 1904-Jan-01 (the cell value 1 represents 1904-Jan-02)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadDateMode(AStream: TStream);
var
  lBaseMode: Word;
begin
  lBaseMode := WordLEtoN(AStream.ReadWord);
  case lBaseMode of
    0: FDateMode := dm1900;
    1: FDateMode := dm1904;
    else raise Exception.CreateFmt('Error reading file. Got unknown date mode number %d.',[lBaseMode]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the default column width
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadDefColWidth(AStream: TStream);
begin
  // The file contains the column width in characters
  FWorksheet.DefaultColWidth := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads the default row height
  Valid for BIFF3 - BIFF8 (override for BIFF2)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadDefRowHeight(AStream: TStream);
var
  hw: Word;
  h: Single;
begin
  // Options
  AStream.ReadWord;

  // Height, in Twips (1/20 pt).
  hw := WordLEToN(AStream.ReadWord);
  h := TwipsToPts(hw) / FWorkbook.GetDefaultFontSize;
  if h > ROW_HEIGHT_CORRECTION then
    FWorksheet.DefaultRowHeight := h - ROW_HEIGHT_CORRECTION;
end;

{@@ ----------------------------------------------------------------------------
  Reads the (number) FORMAT record for formatting numerical data
  To be overridden by descendants.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadFormat(AStream: TStream);
begin
  Unused(AStream);
  // to be overridden
end;

{@@ ----------------------------------------------------------------------------
  Reads a FORMULA record, retrieves the RPN formula and puts the result in the
  corresponding field. The formula is not recalculated here!
  Valid for BIFF5 and BIFF8.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadFormula(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: WORD;
  ResultFormula: Double = 0.0;
  Data: array [0..7] of byte;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  err: TsErrorValue;
  ok: Boolean;
  cell: PCell;





  fmt: TsCellFormat;
begin
  { Index to XF Record }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Result of the formula result in IEEE 754 floating-point value }
  Data[0] := 0;  // to silence the compiler...
  AStream.ReadBuffer(Data, Sizeof(Data));

  { Options flags }
  WordLEtoN(AStream.ReadWord);

  { Not used }
  AStream.ReadDWord;

  { Create cell }
  if FIsVirtualMode then                       // "Virtual" cell
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);    // "Real" cell
    // Don't call "AddCell" because, if the cell belongs to a shared formula, it
    // already has been created before, and then would exist in the tree twice.

  // Now determine the type of the formula result
  if (Data[6] = $FF) and (Data[7] = $FF) then
    case Data[0] of
      0: // String -> Value is found in next record (STRING)
         FIncompleteCell := cell;

      1: // Boolean value
         FWorksheet.WriteBoolValue(cell, Data[2] = 1);

      2: begin  // Error value
           err := ConvertFromExcelError(Data[2]);
           FWorksheet.WriteErrorValue(cell, err);
         end;

      3: FWorksheet.WriteBlank(cell);
    end
  else
  begin
    // Result is a number or a date/time
    Move(Data[0], ResultFormula, SizeOf(Data));

    {Find out what cell type, set content type and value}
    ExtractNumberFormat(XF, nf, nfs);
    if IsDateTime(ResultFormula, nf, nfs, dt) then
      FWorksheet.WriteDateTime(cell, dt) //, nf, nfs)
    else
      FWorksheet.WriteNumber(cell, ResultFormula); //, nf, nfs);
  end;

  { Formula token array }
  if boReadFormulas in FWorkbook.Options then
  begin
    ok := ReadRPNTokenArray(AStream, cell);
    if not ok then
      FWorksheet.WriteErrorValue(cell, errFormulaNotSupported);
  end;

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  Reads whether the page is to be centered horizontally for printing
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadHCENTER(AStream: TStream);
var
  w: word;
begin
  w := WordLEToN(AStream.ReadWord);
  if w = 1 then Include(FWorksheet.PageLayout.Options, poHorCentered);
end;

{@@ ----------------------------------------------------------------------------
  Reads the header/footer to be used for printing.
  Valid for BIFF2-BIFF5, override for BIFF8
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadHeaderFooter(AStream: TStream;
  AIsHeader: Boolean);
var
  s: ansistring;
  len: Byte;
begin
  if RecordSize = 0 then
    exit;

  Len := AStream.ReadByte;
  SetLength(s, len*SizeOf(ansichar));
  AStream.ReadBuffer(s[1], len*SizeOf(ansichar));
  if AIsHeader then
  begin
    FWorksheet.PageLayout.Headers[1] := ConvertEncoding(s, FCodePage, 'utf8');
    FWorksheet.PageLayout.Headers[2] := '';
  end else
  begin
    FWorksheet.PageLayout.Footers[1] := ConvertEncoding(s, FCodePage, 'utf8');
    FWorksheet.PageLayout.Footers[2] := '';
  end;
  Exclude(FWorksheet.PageLayout.Options, poDifferentOddEven);
end;

{@@ ----------------------------------------------------------------------------
  Reads a page margin of the current worksheet (for printing). The margin is
  identified by the parameter "AMargin" (0=left, 1=right, 2=top, 3=bottom)
  The file value is in inches.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadMargin(AStream: TStream; AMargin: Integer);
var
  dbl: Double;
begin
  AStream.ReadBuffer(dbl, SizeOf(dbl));
  case AMargin of
    0: FWorksheet.PageLayout.LeftMargin := InToMM(dbl);
    1: FWorksheet.PageLayout.RightMargin := InToMM(dbl);
    2: FWorksheet.PageLayout.TopMargin := InToMM(dbl);
    3: FWorksheet.PageLayout.BottomMargin := InToMM(dbl);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads multiple blank cell records
  Valid for BIFF5 and BIFF8 (does not exist before)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadMulBlank(AStream: TStream);
var
  ARow, fc, lc, XF: Word;
  pending: integer;
  cell: PCell;
begin
  ARow := WordLEtoN(AStream.ReadWord);
  fc := WordLEtoN(AStream.ReadWord);
  pending := RecordSize - SizeOf(fc) - SizeOf(ARow);
  if FIsVirtualMode then begin
    InitCell(ARow, 0, FVirtualCell);
    cell := @FVirtualCell;
  end;
  while pending > SizeOf(XF) do begin
    XF := AStream.ReadWord; //XF record (not used)
    if FIsVirtualMode then
      cell^.Col := fc
    else
      cell := FWorksheet.AddCell(ARow, fc);
    FWorksheet.WriteBlank(cell);
    ApplyCellFormatting(cell, XF);
    if FIsVirtualMode then
      Workbook.OnReadCellData(Workbook, ARow, fc, cell);
    inc(fc);
    dec(pending, SizeOf(XF));
  end;
  if pending = 2 then begin
    //Just for completeness
    lc := WordLEtoN(AStream.ReadWord);
    if lc + 1 <> fc then begin
      //Stream error... bypass by now
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads multiple RK records.
  Valid for BIFF5 and BIFF8 (does not exist before)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadMulRKValues(const AStream: TStream);
var
  ARow, fc, lc, XF: Word;
  lNumber: Double;
  lDateTime: TDateTime;
  pending: integer;
  RK: DWORD;
  nf: TsNumberFormat;
  nfs: String;
  cell: PCell;
begin
  ARow := WordLEtoN(AStream.ReadWord);
  fc := WordLEtoN(AStream.ReadWord);
  pending := RecordSize - SizeOf(fc) - SizeOf(ARow);
  if FIsVirtualMode then begin
    InitCell(ARow, fc, FVirtualCell);
    cell := @FVirtualCell;
  end;
  while pending > SizeOf(XF) + SizeOf(RK) do begin
    XF := AStream.ReadWord; //XF record (used for date checking)
    if FIsVirtualMode then
      cell^.Col := fc
    else
      cell := FWorksheet.AddCell(ARow, fc);
    RK := DWordLEtoN(AStream.ReadDWord);
    lNumber := DecodeRKValue(RK);
    {Find out what cell type, set contenttype and value}
    ExtractNumberFormat(XF, nf, nfs);
    if IsDateTime(lNumber, nf, nfs, lDateTime) then
      FWorksheet.WriteDateTime(cell, lDateTime, nf, nfs)
    else
      FWorksheet.WriteNumber(cell, lNumber, nf, nfs);
    ApplyCellFormatting(cell, XF);
    if FIsVirtualMode then
      Workbook.OnReadCellData(Workbook, ARow, fc, cell);
    inc(fc);
    dec(pending, SizeOf(XF) + SizeOf(RK));
  end;
  if pending = 2 then begin
    //Just for completeness
    lc := WordLEtoN(AStream.ReadWord);
    if lc + 1 <> fc then begin
      //Stream error... bypass by now
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads a floating point number and seeks the number format
  Valid after BIFF3.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadNumber(AStream: TStream);
var
  rec: TBIFF58NumberRecord;
  ARow, ACol: Cardinal;
  XF: WORD;
  value: Double = 0.0;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  cell: PCell;
begin
  { Read entire record, starting at Row }
  rec.Row := 0;   // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF58NumberRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);
  value := rec.Value;

  {Find out what cell type, set content type and value}
  ExtractNumberFormat(XF, nf, nfs);

  { Create cell }
  if FIsVirtualMode then begin                // "virtual" cell
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);  // "real" cell

  if IsDateTime(value, nf, nfs, dt) then
    FWorksheet.WriteDateTime(cell, dt, nf, nfs)
  else
    FWorksheet.WriteNumber(cell, value, nf, nfs);

  { Add attributes to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  Reads the color palette
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadPalette(AStream: TStream);
var
  i, n: Word;
  pal: Array of TsColorValue;
begin
  n := WordLEToN(AStream.ReadWord) + 8;
  SetLength(pal, n);
  for i:=0 to 7 do
    pal[i] := Workbook.GetPaletteColor(i);
  for i:=8 to n-1 do
    pal[i] := DWordLEToN(AStream.ReadDWord);
  Workbook.UsePalette(@pal[0], n, false);
  FPaletteFound := true;
end;

{@@ ----------------------------------------------------------------------------
  Reads the page setup record containing some parameters for printing
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadPageSetup(AStream: TStream);
var
  w: Word;
  dbl: Double;
begin
  // Paper size
  w := WordLEToN(AStream.ReadWord);
  if (w >= 0) and (w <= High(PAPER_SIZES)) then
  begin
    FWorksheet.PageLayout.PageWidth := PAPER_SIZES[w, 0];
    FWorksheet.PageLayout.PageHeight := PAPER_SIZES[w, 1];
  end;

  // Scaling factor in percent
  FWorksheet.PageLayout.ScalingFactor := WordLEToN(AStream.ReadWord);

  // Start page number
  FWorksheet.PageLayout.StartPageNumber := WordLEToN(AStream.ReadWord);

  // Fit worksheet width to this number of pages (0 = use as many as neede)
  FWorksheet.PageLayout.FitWidthToPages := WordLEToN(AStream.ReadWord);

  // Fit worksheet height to this number of pages (0 = use as many as needed)
  FWorksheet.PageLayout.FitHeightToPages := WordLEToN(AStream.ReadWord);

  // Option flags
  w := WordLEToN(AStream.ReadWord);
  if w and $0001 <> 0 then
    Include(FWorksheet.pageLayout.Options, poPrintPagesByRows);
  if w and $0002 <> 0 then
    FWorksheet.PageLayout.Orientation := spoPortrait else
    FWorksheet.PageLayout.Orientation := spoLandscape;
  if w and $0008 <> 0 then
    Include(FWorksheet.PageLayout.Options, poMonochrome);
  if w and $0010 <> 0 then
    Include(FWorksheet.PageLayout.Options, poDraftQuality);
  if w and $0020 <> 0 then
    Include(FWorksheet.PageLayout.Options, poPrintCellComments);
  if w and $0040 <> 0 then
    Include(FWorksheet.PageLayout.Options, poDefaultOrientation);
  if w and $0080 <> 0 then
    Include(FWorksheet.PageLayout.Options, poUseStartPageNumber);
  if w and $0200 <> 0 then
    Include(FWorksheet.Pagelayout.Options, poCommentsAtEnd);

  // Print resolution in dpi  -- ignoried
  w := WordLEToN(AStream.ReadWord);

  // Vertical print resolution in dpi  -- ignored
  w := WordLEToN(AStream.ReadWord);

  // Header margin
  AStream.ReadBuffer(dbl, SizeOf(dbl));
  FWorksheet.PageLayout.HeaderMargin := InToMM(dbl);

  // Footer margin
  AStream.ReadBuffer(dbl, SizeOf(dbl));
  FWorksheet.PageLayout.FooterMargin := InToMM(dbl);

  // Number of copies
  FWorksheet.PageLayout.Copies := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads pane sizes
  Valid for all BIFF versions
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadPane(AStream: TStream);
begin
  { Position of horizontal split:
    - Unfrozen pane: Width of the left pane(s) (in twips = 1/20 of a point)
    - Frozen pane: Number of visible columns in left pane(s) }
  FWorksheet.LeftPaneWidth := WordLEToN(AStream.ReadWord);

  { Position of vertical split:
    - Unfrozen pane: Height of the top pane(s) (in twips = 1/20 of a point)
    - Frozen pane: Number of visible rows in top pane(s) }
  FWorksheet.TopPaneHeight := WordLEToN(AStream.ReadWord);

  if (FWorksheet.LeftPaneWidth = 0) and (FWorksheet.TopPaneHeight = 0) then
    FWorksheet.Options := FWorksheet.Options - [soHasFrozenPanes];

  { There's more information which is not supported here:
    Offset Size Description
      4     2   Index to first visible row in bottom pane(s)
      6     2   Index to first visible column in right pane(s)
      8     1   Identifier of pane with active cell cursor (see below)
     [9]    1   Not used (BIFF5-BIFF8 only, not written in BIFF2-BIFF4) }
end;

{@@ ----------------------------------------------------------------------------
  Reads whether the gridlines are printed or not
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadPrintGridLines(AStream: TStream);
var
  w: Word;
begin
  w := WordLEToN(AStream.ReadWord);
  if w = 1 then
    Include(FWorksheet.PageLayout.Options, poPrintGridLines);
end;

{@@ ----------------------------------------------------------------------------
  Reads whether the spreadsheet row/column headers are printed or not
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadPrintHeaders(AStream: TStream);
var
  w: Word;
begin
  w := WordLEToN(AStream.ReadWord);
  if w = 1 then
    Include(FWorksheet.PageLayout.Options , poPrintHeaders);
end;

{@@ ----------------------------------------------------------------------------
  Reads the row, column and xf index
  NOT VALID for BIFF2
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRowColXF(AStream: TStream;
  out ARow, ACol: Cardinal; out AXF: WORD);
begin
  { BIFF Record data for row and column}
  ARow := WordLEToN(AStream.ReadWord);
  ACol := WordLEToN(AStream.ReadWord);

  { Index to XF record }
  AXF := WordLEtoN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads an RK value cell from the stream
  Valid since BIFF3.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRKValue(AStream: TStream);
var
  RK: DWord;
  ARow, ACol: Cardinal;
  XF: Word;
  lDateTime: TDateTime;
  Number: Double;
  cell: PCell;
  nf: TsNumberFormat;    // Number format
  nfs: String;           // Number format string
begin
  {Retrieve XF record, row and column}
  ReadRowColXF(AStream, ARow, ACol, XF);

  {Encoded RK value}
  RK := DWordLEtoN(AStream.ReadDWord);

  {Check RK codes}
  Number := DecodeRKValue(RK);

  {Create cell}
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  {Find out what cell type, set contenttype and value}
  ExtractNumberFormat(XF, nf, nfs);
  if IsDateTime(Number, nf, nfs, lDateTime) then
    FWorksheet.WriteDateTime(cell, lDateTime, nf, nfs)
  else
    FWorksheet.WriteNumber(cell, Number, nf, nfs);

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  Reads the part of the ROW record that is common to BIFF3-8 versions
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRowInfo(AStream: TStream);
type
  TRowRecord = packed record
    RowIndex: Word;
    Col1: Word;
    Col2: Word;
    Height: Word;
    NotUsed1: Word;
    NotUsed2: Word;  // not used in BIFF5-BIFF8
    Flags: DWord;
  end;
var
  rowrec: TRowRecord;
  lRow: PRow;
  h: word;
begin
  rowrec.RowIndex := 0;   // to silence the compiler...
  AStream.ReadBuffer(rowrec, SizeOf(TRowRecord));

  // If bit 6 is set in the flags row height does not match the font size.
  // Only for this case we create a row record for fpspreadsheet
  if rowrec.Flags and $00000040 <> 0 then begin
    lRow := FWorksheet.GetRow(WordLEToN(rowrec.RowIndex));
    // row height is encoded into the 15 lower bits in units "twips" (1/20 pt)
    // we need it in "lines", i.e. we divide the points by the point size of the default font
    h := WordLEToN(rowrec.Height) and $7FFF;
    lRow^.Height := TwipsToPts(h) / FWorkbook.GetDefaultFontSize;
    if lRow^.Height > ROW_HEIGHT_CORRECTION then
      lRow^.Height := lRow^.Height - ROW_HEIGHT_CORRECTION
    else
      lRow^.Height := 0;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell address used in an RPN formula element.
  Evaluates the corresponding bits to distinguish between absolute and
  relative addresses.
  Implemented here for BIFF2-BIFF5. BIFF8 must be overridden.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRPNCellAddress(AStream: TStream;
  out ARow, ACol: Cardinal; out AFlags: TsRelFlags);
var
  r: word;
begin
  // 2 bytes for row (including absolute/relative info)
  r := WordLEToN(AStream.ReadWord);
  // 1 byte for column index
  ACol := AStream.ReadByte;
  // Extract row index
  ARow := r and MASK_EXCEL_ROW;
  // Extract absolute/relative flags
  AFlags := [];
  if (r and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (r and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
end;

{@@ ----------------------------------------------------------------------------
  Read the difference between cell row and column indexes of a cell and a
  reference cell.
  Implemented here for BIFF5. BIFF8 must be overridden. Not used by BIFF2.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRPNCellAddressOffset(AStream: TStream;
  out ARowOffset, AColOffset: Integer; out AFlags: TsRelFlags);
var
  r: Word;
  dr: SmallInt;
  dc: ShortInt;
begin
  // 2 bytes for row
  r := WordLEToN(AStream.ReadWord);

  // Check sign bit
  if r and $2000 = 0 then
    dr := SmallInt(r and $3FFF)
  else
    dr := SmallInt($C000 or (r and $3FFF));
//  dr := SmallInt(r and $3FFF);
  ARowOffset := dr;

  // 1 byte for column
  dc := ShortInt(AStream.ReadByte);
  AColOffset := dc;

  // Extract absolute/relative flags
  AFlags := [];
  if (r and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (r and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell range address used in an RPN formula element.
  Evaluates the corresponding bits to distinguish between absolute and
  relative addresses.
  Implemented here for BIFF2-BIFF5. BIFF8 must be overridden.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRPNCellRangeAddress(AStream: TStream;
  out ARow1, ACol1, ARow2, ACol2: Cardinal; out AFlags: TsRelFlags);
var
  r1, r2: word;
begin
  // 2 bytes, each, for first and last row (including absolute/relative info)
  r1 := WordLEToN(AStream.ReadWord);
  r2 := WordLEToN(AStream.ReadWord);
  // 1 byte each for fist and last column index
  ACol1 := AStream.ReadByte;
  ACol2 := AStream.ReadByte;
  // Extract row index of first and last row
  ARow1 := r1 and MASK_EXCEL_ROW;
  ARow2 := r2 and MASK_EXCEL_ROW;
  // Extract absolute/relative flags
  AFlags := [];
  if (r1 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (r2 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol2);
  if (r1 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
  if (r2 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow2);
end;

{@@ ----------------------------------------------------------------------------
  Reads the difference between row and column corner indexes of a cell range
  and a reference cell.
  Implemented here for BIFF5. BIFF8 must be overridden. Not used by BIFF2.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRPNCellRangeOffset(AStream: TStream;
  out ARow1Offset, ACol1Offset, ARow2Offset, ACol2Offset: Integer;
  out AFlags: TsRelFlags);
var
  r1, r2: Word;
  dr1, dr2: SmallInt;
  dc1, dc2: ShortInt;
begin
  // 2 bytes for offset to first row
  r1 := WordLEToN(AStream.ReadWord);
  dr1 := SmallInt(r1 and $3FFF);
  ARow1Offset := dr1;

  // 2 bytes for offset to last row
  r2 := WordLEToN(AStream.ReadWord);
  dr2 := SmallInt(r2 and $3FFF);
  ARow2Offset := dr2;

  // 1 byte for offset to first column
  dc1 := ShortInt(AStream.ReadByte);
  ACol1Offset := dc1;

  // 1 byte for offset to last column
  dc2 := ShortInt(AStream.ReadByte);
  ACol2Offset := dc2;

  // Extract absolute/relative flags
  AFlags := [];
  if (r1 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (r1 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
  if (r2 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol2);
  if (r2 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow2);
end;

{@@ ----------------------------------------------------------------------------
  Reads the identifier for an RPN function with fixed argument count.
  Valid for BIFF4-BIFF8. Override in BIFF2-BIFF3 which read 1 byte only.
-------------------------------------------------------------------------------}
function TsSpreadBIFFReader.ReadRPNFunc(AStream: TStream): Word;
begin
  Result := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell coordinates of the top/left cell of a range using a
  shared formula.
  This cell contains the rpn token sequence of the formula.
  Valid for BIFF3-BIFF8. BIFF2 needs to be overridden (has 1 byte for column).
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadRPNSharedFormulaBase(AStream: TStream;
  out ARow, ACol: Cardinal);
begin
  // 2 bytes for row of first cell in shared formula
  ARow := WordLEToN(AStream.ReadWord);
  // 2 bytes for column of first cell in shared formula
  ACol := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads the array of rpn tokens from the current stream position, creates an
  rpn formula, converts it to a string formula and stores it in the cell.
-------------------------------------------------------------------------------}
function TsSpreadBIFFReader.ReadRPNTokenArray(AStream: TStream;
  ACell: PCell; ASharedFormulaBase: PCell = nil): Boolean;
var
  n: Word;
  p0: Int64;
  token: Byte;
  rpnItem: PRPNItem;
  supported: boolean;
  dblVal: Double = 0.0;   // IEEE 8 byte floating point number
  flags: TsRelFlags;
  r, c, r2, c2: Cardinal;
  dr, dc, dr2, dc2: Integer;
  fek: TFEKind;
  exprDef: TsBuiltInExprIdentifierDef;
  funcCode: Word;
  b: Byte;
  found: Boolean;
  rpnFormula: TsRPNformula;
  strFormula: String;
begin
  rpnItem := nil;
  n := ReadRPNTokenArraySize(AStream);
  p0 := AStream.Position;
  supported := true;
  while (AStream.Position < p0 + n) and supported do begin
    token := AStream.ReadByte;
    case token of
      INT_EXCEL_TOKEN_TREFV:
        begin
          ReadRPNCellAddress(AStream, r, c, flags);
          rpnItem := RPNCellValue(r, c, flags, rpnItem);
        end;
      INT_EXCEL_TOKEN_TREFR:
        begin
          ReadRPNCellAddress(AStream, r, c, flags);
          rpnItem := RPNCellRef(r, c, flags, rpnItem);
        end;
      INT_EXCEL_TOKEN_TAREA_R, INT_EXCEL_TOKEN_TAREA_V:
        begin
          ReadRPNCellRangeAddress(AStream, r, c, r2, c2, flags);
          rpnItem := RPNCellRange(r, c, r2, c2, flags, rpnItem);
        end;
      INT_EXCEL_TOKEN_TREFN_R, INT_EXCEL_TOKEN_TREFN_V:
        begin
          ReadRPNCellAddressOffset(AStream, dr, dc, flags);
          // For compatibility with other formats, convert offsets back to regular indexes.
          if (rfRelRow in flags)
            then r := LongInt(ACell^.Row) + dr
            else r := dr;
          if (rfRelCol in flags)
            then c := LongInt(ACell^.Col) + dc
            else c := dc;
          case token of
            INT_EXCEL_TOKEN_TREFN_V: rpnItem := RPNCellValue(r, c, flags, rpnItem);
            INT_EXCEL_TOKEN_TREFN_R: rpnItem := RPNCellRef(r, c, flags, rpnItem);
          end;
        end;
      INT_EXCEL_TOKEN_TREFN_A:
        begin
          ReadRPNCellRangeOffset(AStream, dr, dc, dr2, dc2, flags);
          // For compatibility with other formats, convert offsets back to regular indexes.
          if (rfRelRow in flags)
            then r := LongInt(ACell^.Row) + dr
            else r := LongInt(ASharedFormulaBase^.Row) + dr;
          if (rfRelRow2 in flags)
            then r2 := LongInt(ACell^.Row) + dr2
            else r2 := LongInt(ASharedFormulaBase^.Row) + dr2;
          if (rfRelCol in flags)
            then c := LongInt(ACell^.Col) + dc
            else c := LongInt(ASharedFormulaBase^.Col) + dc;
          if (rfRelCol2 in flags)
            then c2 := LongInt(ACell^.Col) + dc2
            else c2 := LongInt(ASharedFormulaBase^.Col) + dc2;
          rpnItem := RPNCellRange(r, c, r2, c2, flags, rpnItem);
        end;
      INT_EXCEL_TOKEN_TMISSARG:
        rpnItem := RPNMissingArg(rpnItem);
      INT_EXCEL_TOKEN_TSTR:
        rpnItem := RPNString(ReadString_8BitLen(AStream), rpnItem);
      INT_EXCEL_TOKEN_TERR:
        rpnItem := RPNErr(ConvertFromExcelError(AStream.ReadByte), rpnItem);
      INT_EXCEL_TOKEN_TBOOL:
        rpnItem := RPNBool(AStream.ReadByte=1, rpnItem);
      INT_EXCEL_TOKEN_TINT:
        rpnItem := RPNInteger(WordLEToN(AStream.ReadWord), rpnItem);
      INT_EXCEL_TOKEN_TNUM:
        begin
          AStream.ReadBuffer(dblVal, 8);
          rpnItem := RPNNumber(dblVal, rpnItem);
        end;
      INT_EXCEL_TOKEN_TPAREN:
        rpnItem := RPNParenthesis(rpnItem);

      INT_EXCEL_TOKEN_FUNC_R,
      INT_EXCEL_TOKEN_FUNC_V,
      INT_EXCEL_TOKEN_FUNC_A:
        // functions with fixed argument count
        begin
          funcCode := ReadRPNFunc(AStream);
          exprDef := BuiltInIdentifiers.IdentifierByExcelCode(funcCode);
          if exprDef <> nil then
            rpnItem := RPNFunc(exprDef.Name, rpnItem)
          else
            supported := false;
        end;

      INT_EXCEL_TOKEN_FUNCVAR_R,
      INT_EXCEL_TOKEN_FUNCVAR_V,
      INT_EXCEL_TOKEN_FUNCVAR_A:
        // functions with variable argument count
        begin
          b := AStream.ReadByte;
          funcCode := ReadRPNFunc(AStream);
          exprDef := BuiltinIdentifiers.IdentifierByExcelCode(funcCode);
          if exprDef <> nil then
            rpnItem := RPNFunc(exprDef.Name, b, rpnItem)
          else
            supported := false;
        end;

      INT_EXCEL_TOKEN_TEXP:
        // Indicates that cell belongs to a shared or array formula.
        // This information is not needed any more.
        ReadRPNSharedFormulaBase(AStream, r, c);

      else
        found := false;
        for fek in TBasicOperationTokens do
          if (TokenIDs[fek] = token) then begin
            rpnItem := RPNFunc(fek, rpnItem);
            found := true;
            break;
          end;
        if not found then
          supported := false;
    end;
  end;
  if not supported then begin
    DestroyRPNFormula(rpnItem);
    Result := false;
  end
  else begin
    rpnFormula := CreateRPNFormula(rpnItem, true); // true --> we have to flip the order of items!
    strFormula := FWorksheet.ConvertRPNFormulaToStringFormula(rpnFormula);
    if strFormula <> '' then
      ACell^.FormulaValue := strFormula;
    {
    if (ACell^.SharedFormulaBase = nil) or (ACell = ACell^.SharedFormulaBase) then
      ACell^.FormulaValue := FWorksheet.ConvertRPNFormulaToStringFormula(formula)
    else
      ACell^.FormulaValue := '';
      }
    Result := true;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Helper function for reading of the size of the token array of an RPN formula.
  Is implemented here for BIFF3-BIFF8 where the size is a 2-byte value.
  Needs to be rewritten for BIFF2 using a 1-byte size.
-------------------------------------------------------------------------------}
function TsSpreadBIFFReader.ReadRPNTokenArraySize(AStream: TStream): Word;
begin
  Result := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads a SHAREDFMLA record, i.e. reads cell range coordinates and a rpn
  formula. The formula is applied to all cells in the range. The formula is
  stored only in the top/left cell of the range.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadSharedFormula(AStream: TStream);
var
  r, r1, r2, c, c1, c2: Cardinal;
  cell: PCell;
begin
  // Cell range in which the formula is valid
  r1 := WordLEToN(AStream.ReadWord);
  r2 := WordLEToN(AStream.ReadWord);
  c1 := AStream.ReadByte;         // 8 bit, even for BIFF8
  c2 := AStream.ReadByte;

  { Create cell - this is the "base" of the shared formula }
  if FIsVirtualMode then begin                 // "Virtual" cell
    InitCell(r1, c1, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(r1, c1);        // "Real" cell
    // Don't use "AddCell" here because this cell already exists in files written
    // by Excel, and this would destroy its formatting.

  // Unused
  AStream.ReadByte;

  // Number of existing FORMULA records for this shared formula
  AStream.ReadByte;

  // RPN formula tokens
  ReadRPNTokenArray(AStream, cell, cell); //base);

  // Copy shared formula to individual cells in the specified range
  for r := r1 to r2 do
    for c := c1 to c2 do
      FWorksheet.CopyFormula(cell, r, c);
end;

{@@ ----------------------------------------------------------------------------
  Helper function for reading a string with 8-bit length. Here, we implement
  the version for ansistrings since it is valid for all BIFF versions except
  for BIFF8 where it has to be overridden.
-------------------------------------------------------------------------------}
function TsSpreadBIFFReader.ReadString_8bitLen(AStream: TStream): String;
var
  len: Byte;
  s: ansistring;
begin
  len := AStream.ReadByte;
  SetLength(s, len);
  AStream.ReadBuffer(s[1], len);
  Result := ConvertEncoding(s, FCodePage, encodingUTF8);
end;

{@@ ----------------------------------------------------------------------------
  Reads a STRING record. It immediately precedes a FORMULA record which has a
  string result. The read value is applied to the FIncompleteCell.
  Must be overridden because the implementation depends on BIFF version.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadStringRecord(AStream: TStream);
begin
  Unused(AStream);
end;

{@@ ----------------------------------------------------------------------------
  Reads whether the page is to be centered vertically for printing
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadVCENTER(AStream: TStream);
var
  w: word;
begin
  w := WordLEToN(AStream.ReadWord);
  if w = 1 then Include(FWorksheet.PageLayout.Options, poVertCentered);
end;

{@@ ----------------------------------------------------------------------------
  Reads the WINDOW2 record containing information like "show grid lines",
  "show sheet headers", "panes are frozen", etc.
  The record structure is slightly different for BIFF5 and BIFF8, but we use
  here only the common part.
  BIFF2 has a different structure and has to be re-written.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFReader.ReadWindow2(AStream: TStream);
var
  flags: Word;
begin
  flags := WordLEToN(AStream.ReadWord);

  if (flags and MASK_WINDOW2_OPTION_SHOW_GRID_LINES <> 0) then
    FWorksheet.Options := FWorksheet.Options + [soShowGridLines]
  else
    FWorksheet.Options := FWorksheet.Options - [soShowGridLines];

  if (flags and MASK_WINDOW2_OPTION_SHOW_SHEET_HEADERS <> 0) then
    FWorksheet.Options := FWorksheet.Options + [soShowHeaders]
  else
    FWorksheet.Options := FWorksheet.Options - [soShowHeaders];

  if (flags and MASK_WINDOW2_OPTION_PANES_ARE_FROZEN <> 0) then
    FWorksheet.Options := FWorksheet.Options + [soHasFrozenPanes]
  else
    FWorksheet.Options := FWorksheet.Options - [soHasFrozenPanes];
end;


{------------------------------------------------------------------------------}
{                            TsSpreadBIFFWriter                                }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the general BIFF writer.
  Initializes the date mode and the limitations of the format.
-------------------------------------------------------------------------------}
constructor TsSpreadBIFFWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);

  // Initial base date in case it won't be set otherwise.
  // Use 1900 to get a bit more range between 1900..1904.
  FDateMode := dm1900;

  // Limitations of BIFF5 and BIFF8 file formats
  FLimitations.MaxColCount := 256;
  FLimitations.MaxRowCount := 65536;
  FLimitations.MaxPaletteSize := 64;
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in number formats to the NumFormatList.
  Valid for BIFF5...BIFF8. Needs to be overridden for BIFF2.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.AddBuiltinNumFormats;
begin
  FFirstNumFormatIndexInFile := 164;
  AddBuiltInBiffFormats(
    FNumFormatList, Workbook.FormatSettings, FFirstNumFormatIndexInFile-1
  );
end;

{@@ ----------------------------------------------------------------------------
  Determines the index of the XF record, according to formatting of the
  given cell
-------------------------------------------------------------------------------}
function TsSpreadBIFFWriter.FindXFIndex(ACell: PCell): Integer;
begin
  Result := LAST_BUILTIN_XF + ACell^.FormatIndex;
end;

function TsSpreadBIFFWriter.FixColor(AColor: TsColor): TsColor;
var
  rgb: TsColorValue;
begin
  if AColor >= Limitations.MaxPaletteSize then begin
    rgb := Workbook.GetPaletteColor(AColor);
    Result := Workbook.FindClosestColor(rgb, FLimitations.MaxPaletteSize);
  end else
    Result := AColor;
end;

function TsSpreadBIFFWriter.GetLastRowIndex(AWorksheet: TsWorksheet): Integer;
begin
  Result := AWorksheet.GetLastRowIndex;
end;

function TsSpreadBIFFWriter.GetLastColIndex(AWorksheet: TsWorksheet): Word;
begin
  Result := AWorksheet.GetLastColIndex;
end;

{@@ ----------------------------------------------------------------------------
  Converts the Options of the worksheet's PageLayout to the bitmap required
  by the PageSetup record
  Is overridden by BIFF8 which uses more bits. Not used by BIFF2.
-------------------------------------------------------------------------------}
function TsSpreadBIFFWriter.GetPrintOptions: Word;
begin
  { Options:
     Bit 0: 0 = Print pages in columns; 1 = Print pages in rows
     Bit 1: 0 = Landscape; 1 = Portrait
     Bit 2: 1 = Paper size, scaling factor, paper orientation (portrait/landscape),
                print resolution and number of copies are not initialised
     Bit 3: 0 = Print coloured; 1 = Print black and white
     Bit 4: 0 = Default print quality; 1 = Draft quality
     Bit 5: 0 = Do not print cell notes; 1 = Print cell notes
     Bit 6: 0 = Use paper orientation (portrait/landscape) flag above
            1 = Use default paper orientation (landscape for chart sheets, portrait otherwise)
     Bit 7: 0 = Automatic page numbers; 1 = Use start page number above

     The following flags are valid for BIFF8 only:
     Bit 9: 0 = Print notes as displayed; 1 = Print notes at end of sheet
     Bit 11-10:  00 = Print errors as displayed; 1 = Do not print errors
                 2 = Print errors as “--”; 3 = Print errors as “#N/A” }
  Result := 0;
  if poPrintPagesByRows in FWorksheet.PageLayout.Options then
    Result := Result or $0001;
  if FWorksheet.PageLayout.Orientation = spoPortrait then
    Result := Result or $0002;
  if poMonochrome in FWorksheet.PageLayout.Options then
    Result := Result or $0008;
  if poDraftQuality in FWorksheet.PageLayout.Options then
    Result := Result or $0010;
  if poPrintCellComments in FWorksheet.PageLayout.Options then
    Result := Result or $0020;
  if poDefaultOrientation in FWorksheet.PageLayout.Options then
    Result := Result or $0040;
  if poUseStartPageNumber in FWorksheet.PageLayout.Options then
    Result := Result or $0080;
end;

{@@ ----------------------------------------------------------------------------
  Writes the BIFF record header consisting of the record ID and the size of
  data to be written immediately afterwards.

  @param  ARecID    ID of the record - see the INT_EXCEL_ID_XXXX constants
  @param  ARedSize  Size (in bytes) of the data which follow immediately
                    afterwards
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteBIFFHeader(AStream: TStream;
  ARecID, ARecSize: Word);
var
  rec: TsBIFFHeader;
begin
  rec.RecordID := WordToLE(ARecID);
  rec.RecordSize := WordToLE(ARecSize);
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an empty ("blank") cell. Needed for formatting empty cells.
  Valid for BIFF5 and BIFF8. Needs to be overridden for BIFF2 which has a
  different record structure.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  rec: TBIFF58BlankRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BLANK);
  rec.RecordSize := WordToLE(6);

  { Row and column index }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { Index to XF record, according to formatting }
  rec.XFIndex := WordToLE(FindXFIndex(ACell));

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes a BOOLEAN cell record.
  Valid for BIFF3-BIFF8. Override for BIFF2.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: Boolean; ACell: PCell);
var
  rec: TBIFF38BoolErrRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BOOLERROR);
  rec.RecordSize := WordToLE(8);

  { Row and column index }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { Index to XF record, according to formatting }
  rec.XFIndex := WordToLE(FindXFIndex(ACell));

  { Cell value }
  rec.BoolErrValue := ord(AValue);
  rec.ValueType := 0;  // 0 = boolean value, 1 = error value

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes the code page identifier defined by the workbook to the stream.
  BIFF2 has to be overridden because is uses cp1252, but has a different
  number code.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteCodepage(AStream: TStream; ACodePage: String);
//  AEncoding: TsEncoding);
var
  cp: Word;
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_CODEPAGE, 2);

  { Codepage }
  FCodePage := lowercase(ACodePage);
  case FCodePage of
    'ucs2le': cp := WORD_UTF_16;  // Biff 7
    'cp437' : cp := WORD_CP_437_DOS_US;
    'cp850' : cp := WORD_CP_850_DOS_Latin1;
    'cp852' : cp := WORD_CP_852_DOS_Latin2;
    'cp866' : cp := WORD_CP_866_DOS_Cyrillic;
    'cp874' : cp := WORD_CP_874_Thai;
    'cp1250': cp := WORD_CP_1250_Latin2;
    'cp1251': cp := WORD_CP_1251_Cyrillic;
    'cp1252': cp := WORD_CP_1252_Latin1;
    'cp1253': cp := WORD_CP_1253_Greek;
    'cp1254': cp := WORD_CP_1254_Turkish;
    'cp1255': cp := WORD_CP_1255_Hebrew;
    'cp1256': cp := WORD_CP_1256_Arabic;
    'cp1257': cp := WORD_CP_1257_Baltic;
    'cp1258': cp := WORD_CP_1258_Vietnamese;
  else
    Workbook.AddErrorMsg(rsCodePageNotSupported, [FCodePage]);
    FCodePage := 'cp1252';
    cp := WORD_CP_1252_Latin1;
  end;
  AStream.WriteWord(WordToLE(cp));
end;

{@@ ----------------------------------------------------------------------------
  Writes column info for the given column.
  Currently only the colum width is used.
  Valid for BIFF5 and BIFF8 (BIFF2 uses a different record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteColInfo(AStream: TStream; ACol: PCol);
type
  TColRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    StartCol: Word;
    EndCol: Word;
    ColWidth: Word;
    XFIndex: Word;
    OptionFlags: Word;
    NotUsed: Word;
  end;
var
  rec: TColRecord;
  w: Integer;
begin
  if Assigned(ACol) then
  begin
    if (ACol^.Col >= FLimitations.MaxColCount) then
      exit;

    { BIFF record header }
    rec.RecordID := WordToLE(INT_EXCEL_ID_COLINFO);
    rec.RecordSize := WordToLE(12);

    { Start and end column }
    rec.StartCol := WordToLE(ACol^.Col);
    rec.EndCol := WordToLE(ACol^.Col);

    { calculate width to be in units of 1/256 of pixel width of character "0" }
    w := round(ACol^.Width * 256);

    rec.ColWidth := WordToLE(w);
    rec.XFIndex := WordToLE(15);    // Index of XF record, not used
    rec.OptionFlags := 0;           // not used
    rec.NotUsed := 0;

    { Write out }
    AStream.WriteBuffer(rec, SizeOf(rec));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the column info records for all used columns.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteColInfos(AStream: TStream;
  ASheet: TsWorksheet);
var
  j: Integer;
  col: PCol;
begin
  for j := 0 to ASheet.Cols.Count-1 do begin
    col := PCol(ASheet.Cols[j]);
    WriteColInfo(AStream, col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a NOTE record which describes a comment attached to a cell
  Valid für Biff2 and BIFF5.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteComment(AStream: TStream; ACell: PCell);
const
  CHUNK_SIZE = 2048;
var
  rec: TBIFF25NoteRecord;
  L: Integer;
  base_size: Word;
  p: Integer;
  cmnt: ansistring;
  List: TStringList;
  comment: PsComment;
begin
  Unused(ACell);

  comment := FWorksheet.FindComment(ACell);
  if (comment = nil) or (comment^.Text = '') then
    exit;

  List := TStringList.Create;
  try
    List.Text := ConvertEncoding(comment^.Text, encodingUTF8, FCodePage);
    cmnt := List[0];
    for p := 1 to List.Count-1 do
      cmnt := cmnt + #$0A + List[p];
  finally
    List.Free;
  end;

  L := Length(cmnt);
  base_size := SizeOf(rec) - 2*SizeOf(word);

  // First NOTE record
  rec.RecordID := WordToLE(INT_EXCEL_ID_NOTE);
  rec.Row := WordToLE(ACell^.Row);
  rec.Col := WordToLE(ACell^.Col);
  rec.TextLen := L;
  rec.RecordSize := base_size + Min(L, CHUNK_SIZE);
  AStream.WriteBuffer(rec, SizeOf(rec));
  AStream.WriteBuffer(cmnt[1], Min(L, CHUNK_SIZE));  // Write text

  // If the comment text does not fit into 2048 bytes continuation records
  // have to be written.
  rec.Row := $FFFF;  // indicator that this will be a continuation record
  rec.Col := 0;
  p := CHUNK_SIZE + 1;
  dec(L, CHUNK_SIZE);
  while L > 0 do begin
    rec.TextLen := Min(L, CHUNK_SIZE);
    rec.RecordSize := base_size + rec.TextLen;
    AStream.WriteBuffer(rec, SizeOf(rec));
    AStream.WriteBuffer(cmnt[p], rec.TextLen);
    dec(L, CHUNK_SIZE);
    inc(p, CHUNK_SIZE);
  end;
end;

procedure TsSpreadBIFFWriter.WriteDateMode(AStream: TStream);
begin
  { BIFF Record header }
  // todo: check whether this is in the right place. should end up in workbook globals stream
  WriteBIFFHeader(AStream, INT_EXCEL_ID_DATEMODE, 2);

  case FDateMode of
    dm1900: AStream.WriteWord(WordToLE(0));
    dm1904: AStream.WriteWord(WordToLE(1));
    else raise Exception.CreateFmt('Unknown datemode number %d. Please correct fpspreadsheet code.', [FDateMode]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time/datetime to a BIFF NUMBER record, with a date/time format
  (There is no separate date record type in xls)
  Valid for all BIFF versions.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteDateTime(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
var
  ExcelDateSerial: double;
begin
  ExcelDateSerial := ConvertDateTimeToExcelDateTime(AValue, FDateMode);
  // fpspreadsheet must already have set formatting to a date/datetime format, so
  // this will get written out as a pointer to the relevant XF record.
  // In the end, dates in xls are just numbers with a format. Pass it on to WriteNumber:
  WriteNumber(AStream, ARow, ACol, ExcelDateSerial, ACell);
end;

{@@ ----------------------------------------------------------------------------
  Writes an ERROR cell record.
  Valid for BIFF3-BIFF8. Override for BIFF2.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  rec: TBIFF38BoolErrRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BOOLERROR);
  rec.RecordSize := WordToLE(8);

  { Row and column index }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { Index to XF record, according to formatting }
  rec.XFIndex := WordToLE(FindXFIndex(ACell));

  { Cell value }
  rec.BoolErrValue := ConvertToExcelError(AValue);
  rec.ValueType := 1;  // 0 = boolean value, 1 = error value

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes the a margin record for printing (margin is in inches).
  The margin is identified by the parameter AMargin:
    0=left, 1=right, 2=top, 3=bottom
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteMargin(AStream: TStream; AMargin: Integer);
var
  dbl: double;
begin
  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_LEFTMARGIN + AMargin, SizeOf(Double));
  // the MARGIN IDs are consecutive beginning with the one for left margin

  { Page margin value, written in inches }
  case AMargin of
    0: dbl := FWorksheet.PageLayout.LeftMargin;
    1: dbl := FWorksheet.PageLayout.RightMargin;
    2: dbl := FWorksheet.PageLayout.TopMargin;
    3: dbl := FWorksheet.PageLayout.Bottommargin;
  end;
  dbl := mmToIn(dbl);
  AStream.WriteBuffer(dbl, SizeOf(dbl));
end;

{@@ ----------------------------------------------------------------------------
  Writes a BIFF number format record defined in the specified format string
  (in Excel dialect).
  AFormatIndex is equal to the format index used in the Excel file.
  Needs to be overridden by descendants.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteNumFormat(AStream: TStream;
  ANumFormatStr: String; ANumFormatIndex: Integer);
begin
  Unused(AStream, ANumFormatStr, ANumFormatIndex);
  // needs to be overridden
end;

{@@ ----------------------------------------------------------------------------
  Writes all number formats to the stream. Saving starts at the item with the
  FirstFormatIndexInFile.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteNumFormats(AStream: TStream);
var
  i: Integer;
  parser: TsNumFormatParser;
  fmtStr: String;
begin
  ListAllNumFormats;
  for i:= FFirstNumFormatIndexInFile to NumFormatList.Count-1 do
  begin
    fmtStr := NumFormatList[i];
    parser := TsNumFormatParser.Create(Workbook, fmtStr);
    try
      fmtStr := parser.FormatString;
      WriteNumFormat(AStream, fmtStr, i);
    finally
      parser.Free;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel FORMULA record.
  Note: The formula is already stored in the cell.
  Since BIFF files contain RPN formulas the string formula of the cell is
  converted to an RPN formula and the method calls WriteRPNFormula.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteFormula(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  formula: TsRPNFormula;
begin
  formula := FWorksheet.BuildRPNFormula(ACell);
  WriteRPNFormula(AStream, ARow, ACol, formula, ACell);
  SetLength(formula, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel HCENTER record which determines whether the page is to be
  centered horizontally for printing
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteHCenter(AStream: TStream);
var
  w: Word;
begin
  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_HCENTER, SizeOf(w));

  { Data }
  if poHorCentered in FWorksheet.PageLayout.Options then w := 1 else w := 0;
  AStream.WriteWord(WordToLE(w));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel HEADER or FOOTER record, depending on AIsHeader.
  Valid for BIFF2-5. Override for BIFF7 because of WideStrings
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteHeaderFooter(AStream: TStream;
  AIsHeader: Boolean);
var
  s: AnsiString;
  len: Integer;
  id: Word;
begin
  with FWorksheet.PageLayout do
    if AIsHeader then
    begin
      if (Headers[HEADER_FOOTER_INDEX_ALL] = '') then
        exit;
      s := ConvertEncoding(Headers[HEADER_FOOTER_INDEX_ALL], 'utf8', FCodePage);
      id := INT_EXCEL_ID_HEADER;
    end else
    begin
      if (Footers[HEADER_FOOTER_INDEX_ALL] = '') then
        exit;
      s := ConvertEncoding(Footers[HEADER_FOOTER_INDEX_ALL], 'utf8', FCodePage);
      id := INT_EXCEL_ID_FOOTER;
    end;
  len := Length(s);

  { BIFF record header }
  WriteBiffHeader(AStream, id, 1 + len*sizeOf(AnsiChar));

  { 8-bit string length }
  AStream.WriteByte(len);

  { Characters }
  AStream.WriteBuffer(s[1], len * SizeOf(AnsiChar));
end;


{@@ ----------------------------------------------------------------------------
  Writes a 64-bit floating point NUMBER record.
  Valid for BIFF5 and BIFF8 (BIFF2 has a different record structure).
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteNumber(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: double; ACell: PCell);
var
  rec: TBIFF58NumberRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  { BIFF Record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_NUMBER);
  rec.RecordSize := WordToLE(14);

  { BIFF Record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { Index to XF record }
  rec.XFIndex := FindXFIndex(ACell);

  { IEE 754 floating-point value }
  rec.Value := AValue;

  AStream.WriteBuffer(rec, sizeof(Rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes the PALETTE record for the color palette.
  Valid for BIFF3-BIFF8. BIFF2 has no palette in file.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WritePalette(AStream: TStream);
const
  NUM_COLORS = 56;
var
  i, n: Integer;
  rgb: TsColorValue;
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_PALETTE, 2 + 4*NUM_COLORS);

  { Number of colors }
  AStream.WriteWord(WordToLE(NUM_COLORS));

  { Take the colors from the palette of the Worksheet }
  n := Workbook.GetPaletteSize;

  { Skip the first 8 entries - they are hard-coded into Excel }
  for i := 8 to 8 + NUM_COLORS - 1 do
  begin
    rgb := Math.IfThen(i < n, Workbook.GetPaletteColor(i), $FFFFFF);
    AStream.WriteDWord(DWordToLE(rgb))
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a PAGESETUP record containing information on printing
  Valid for BIFF5-8
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WritePageSetup(AStream: TStream);
var
  dbl: Double;
  i: Integer;
  w: Word;
begin
  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_PAGESETUP, 9*2 + 2*8);

  { Paper size }
  w := 0;
  for i:=0 to High(PAPER_SIZES) do
    if (SameValue(PAPER_SIZES[i,0], FWorksheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,1], FWorksheet.PageLayout.PageWidth))
    or (SameValue(PAPER_SIZES[i,1], FWorksheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,0], FWorksheet.PageLayout.PageWidth))
    then begin
      w := i;
      break;
    end;
  AStream.WriteWord(WordToLE(w));

  { Scaling factor in percent }
  w := Round(FWorksheet.PageLayout.ScalingFactor);
  AStream.WriteWord(WordToLE(w));

  { Start page number }
  w := FWorksheet.PageLayout.StartPageNumber;
  AStream.WriteWord(WordToLE(w));

  { Fit worksheet width to this number of pages, 0 = use as many as needed }
  w := FWorksheet.PageLayout.FitWidthToPages;
  AStream.WriteWord(WordToLE(w));

  { Fit worksheet height to this number of pages, 0 = use as many as needed }
  w := FWorksheet.PageLayout.FitHeightToPages;
  AStream.WriteWord(WordToLE(w));

  { Options:
     Bit 0: 0 = Print pages in columns; 1 = Print pages in rows
     Bit 1: 0 = Landscape; 1 = Portrait
     Bit 2: 1 = Paper size, scaling factor, paper orientation (portrait/landscape),
                print resolution and number of copies are not initialised
     Bit 3: 0 = Print coloured; 1 = Print black and white
     Bit 4: 0 = Default print quality; 1 = Draft quality
     Bit 5: 0 = Do not print cell notes; 1 = Print cell notes
     Bit 6: 0 = Use paper orientation (portrait/landscape) flag above
            1 = Use default paper orientation (landscape for chart sheets, portrait otherwise)
     Bit 7: 0 = Automatic page numbers; 1 = Use start page number above

     The following flags are valid for BIFF8 only:
     Bit 9: 0 = Print notes as displayed; 1 = Print notes at end of sheet
     Bit 11-10:  00 = Print errors as displayed; 1 = Do not print errors
                 2 = Print errors as “--”; 3 = Print errors as “#N/A” }
  w := GetPrintOptions;
  AStream.WriteWord(WordToLE(w));

  { Print resolution in dpi }
  AStream.WriteWord(WordToLE(600));

  { Vertical print resolution in dpi }
  AStream.WriteWord(WordToLE(600));

  { Header margin }
  dbl := mmToIn(FWorksheet.PageLayout.HeaderMargin);
  AStream.WriteBuffer(dbl, SizeOf(dbl));

  { Footer margin }
  dbl := mmToIn(FWorksheet.PageLayout.FooterMargin);
  AStream.WriteBuffer(dbl, SizeOf(dbl));

  { Number of copies to print }
  w := FWorksheet.PageLayout.Copies;
  AStream.WriteWord(WordToLE(w));
end;

{@@ ----------------------------------------------------------------------------
  Writes a PANE record to the stream.
  Valid for all BIFF versions. The difference for BIFF5-BIFF8 is a non-used
  byte at the end. Activate IsBiff58 in these cases.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WritePane(AStream: TStream; ASheet: TsWorksheet;
  IsBiff58: Boolean; out ActivePane: Byte);
var
  n: Word;
begin
  ActivePane := 3;

  if not (soHasFrozenPanes in ASheet.Options) then
    exit;
  if (ASheet.LeftPaneWidth = 0) and (ASheet.TopPaneHeight = 0) then
    exit;

  if not (soHasFrozenPanes in ASheet.Options) then
    exit;
  { Non-frozen panes should work in principle, but they are not read without
    error. They possibly require an additional SELECTION record. }

  { BIFF record header }
  if isBIFF58 then n := 10 else n := 9;
  WriteBIFFHeader(AStream, INT_EXCEL_ID_PANE, n);

  { Position of the vertical split (px, 0 = No vertical split):
    - Unfrozen pane: Width of the left pane(s) (in twips = 1/20 of a point)
    - Frozen pane: Number of visible columns in left pane(s) }
  AStream.WriteWord(WordToLE(ASheet.LeftPaneWidth));

  { Position of the horizontal split (py, 0 = No horizontal split):
    - Unfrozen pane: Height of the top pane(s) (in twips = 1/20 of a point)
    - Frozen pane: Number of visible rows in top pane(s) }
  AStream.WriteWord(WordToLE(ASheet.TopPaneHeight));

  { Index to first visible row in bottom pane(s) }
  if (soHasFrozenPanes in ASheet.Options) then
    AStream.WriteWord(WordToLE(ASheet.TopPaneHeight))
  else
    AStream.WriteWord(WordToLE(0));

  { Index to first visible column in right pane(s) }
  if (soHasFrozenPanes in ASheet.Options) then
    AStream.WriteWord(WordToLE(ASheet.LeftPaneWidth))
  else
    AStream.WriteWord(WordToLE(0));

  { Identifier of pane with active cell cursor
      0 = right-bottom
      1 = right-top
      2 = left-bottom
      3 = left-top }
  if (soHasFrozenPanes in ASheet.Options) then begin
    if (ASheet.LeftPaneWidth = 0) and (ASheet.TopPaneHeight = 0) then
      ActivePane := 3
    else
    if (ASheet.LeftPaneWidth = 0) then
      ActivePane := 2
    else
    if (ASheet.TopPaneHeight =0) then
      ActivePane := 1
    else
      ActivePane := 0;
  end else
    ActivePane := 0;
  AStream.WriteByte(ActivePane);

  if IsBIFF58 then
    AStream.WriteByte(0);
    { Not used (BIFF5-BIFF8 only, not written in BIFF2-BIFF4 }
end;

{@@ ----------------------------------------------------------------------------
  Writes out whether grid lines are printed or not
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WritePrintGridLines(AStream: TStream);
var
  w: Word;
begin
  { Biff record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_PRINTGRID, SizeOf(w));

  { Data }
  if poPrintGridLines in FWorksheet.PageLayout.Options then w := 1 else w := 0;
  AStream.WriteWord(WordToLE(w));
end;

{@@ ----------------------------------------------------------------------------
  Writes out whether column and row headers are printed or not
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WritePrintHeaders(AStream: TStream);
var
  w: Word;
begin
  { Biff record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_PRINTHEADERS, SizeOf(w));

  { Data }
  if poPrintHeaders in FWorksheet.PageLayout.Options then w := 1 else w := 0;
  AStream.WriteWord(WordToLE(w));
end;

{@@ ----------------------------------------------------------------------------
  Writes the address of a cell as used in an RPN formula and returns the
  count of bytes written.
  Valid for BIFF2-BIFF5.
-------------------------------------------------------------------------------}
function TsSpreadBIFFWriter.WriteRPNCellAddress(AStream: TStream;
  ARow, ACol: Cardinal; AFlags: TsRelFlags): Word;
var
  r: Cardinal;  // row index containing encoded relative/absolute address info
begin
  // Encoded row address
  r := ARow and MASK_EXCEL_ROW;
  if (rfRelRow in AFlags) then r := r or MASK_EXCEL_RELATIVE_ROW;
  if (rfRelCol in AFlags) then r := r or MASK_EXCEL_RELATIVE_COL;
  AStream.WriteWord(WordToLE(r));
  // Column address
  AStream.WriteByte(ACol);
  // Number of bytes written
  Result := 3;
end;

{@@ ----------------------------------------------------------------------------
  Writes row and column offset (unsigned integers!)
  Valid for BIFF2-BIFF5.
-------------------------------------------------------------------------------}
function TsSpreadBIFFWriter.WriteRPNCellOffset(AStream: TStream;
  ARowOffset, AColOffset: Integer; AFlags: TsRelFlags): Word;
var
  r: Word;
  c: Byte;
begin
  // Encoded row address
  r := ARowOffset and MASK_EXCEL_ROW;
  if (rfRelRow in AFlags) then r := r + MASK_EXCEL_RELATIVE_ROW;
  if (rfRelCol in AFlags) then r := r + MASK_EXCEL_RELATIVE_COL;
  AStream.WriteWord(WordToLE(r));
  // Column address
  c := Lo(word(AColOffset));
  //c := Lo(AColOffset);
  AStream.WriteByte(c);
  // Number of bytes written
  Result := 3;
end;

{@@ ----------------------------------------------------------------------------
  Writes the address of a cell range as used in an RPN formula and returns the
  count of bytes written.
  Valid for BIFF2-BIFF5.
-------------------------------------------------------------------------------}
function TsSpreadBIFFWriter.WriteRPNCellRangeAddress(AStream: TStream;
  ARow1, ACol1, ARow2, ACol2: Cardinal; AFlags: TsRelFlags): Word;
var
  r: Cardinal;
begin
  r := ARow1 and MASK_EXCEL_ROW;
  if (rfRelRow in AFlags) then r := r or MASK_EXCEL_RELATIVE_ROW;
  if (rfRelCol in AFlags) then r := r or MASK_EXCEL_RELATIVE_COL;
  AStream.WriteWord(WordToLE(r));

  r := ARow2 and MASK_EXCEL_ROW;
  if (rfRelRow2 in AFlags) then r := r or MASK_EXCEL_RELATIVE_ROW;
  if (rfRelCol2 in AFlags) then r := r or MASK_EXCEL_RELATIVE_COL;
  AStream.WriteWord(WordToLE(r));

  AStream.WriteByte(ACol1);
  AStream.WriteByte(ACol2);

  Result := 6;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel FORMULA record
  The formula needs to be converted from usual user-readable string
  to an RPN array
  Valid for BIFF5-BIFF8.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteRPNFormula(AStream: TStream;
  const ARow, ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell);
var
  RPNLength: Word = 0;
  RecordSizePos, StartPos, FinalPos: Int64;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  if Length(AFormula) = 0 then
    exit;
  {
  if not ((Length(AFormula) > 0) or (ACell^.SharedFormulaBase <> nil)) then
    exit;
   }

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMULA));
  RecordSizePos := AStream.Position;
  AStream.WriteWord(0);  // This is the record size which is not yet known here
  StartPos := AStream.Position;

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { Index to XF record, according to formatting }
  AStream.WriteWord(FindXFIndex(ACell));

  { Encoded result of RPN formula }
  WriteRPNResult(AStream, ACell);

  { Options flags }
  AStream.WriteWord(WordToLE(MASK_FORMULA_RECALCULATE_ALWAYS));

  { Not used }
  AStream.WriteDWord(0);

  { Formula data (RPN token array) }
  {
  if ACell^.SharedFormulaBase <> nil then
    WriteRPNSharedFormulaLink(AStream, ACell, RPNLength)
  else
  }
  WriteRPNTokenArray(AStream, ACell, AFormula, false, RPNLength);

  { Write sizes in the end, after we known them }
  FinalPos := AStream.Position;
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(FinalPos - StartPos));
  AStream.Position := FinalPos;
  (*
  { If the cell is the first cell of a range with a shared formula write the
    shared formula RECORD here. The shared formula RECORD must follow the
    first FORMULA record referring to the shared formula}
  if (ACell^.SharedFormulaBase <> nil) and
     (ARow = ACell^.SharedFormulaBase^.Row) and
     (ACol = ACell^.SharedFormulaBase^.Col)
  then
    WriteSharedFormula(AStream, ACell^.SharedFormulaBase);
  *)

  { Write following STRING record if formula result is a non-empty string. }
  if (ACell^.ContentType = cctUTF8String) and (ACell^.UTF8StringValue <> '') then
    WriteSTRINGRecord(AStream, ACell^.UTF8StringValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes the identifier for an RPN function with fixed argument count and
  returns the number of bytes written.
  Valid for BIFF4-BIFF8. Override in BIFF2-BIFF3.
-------------------------------------------------------------------------------}
function TsSpreadBIFFWriter.WriteRPNFunc(AStream: TStream; AIdentifier: Word): Word;
begin
  AStream.WriteWord(WordToLE(AIdentifier));
  Result := 2;
end;

{@@ ----------------------------------------------------------------------------
  Writes the result of an RPN formula.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteRPNResult(AStream: TStream; ACell: PCell);
var
  Data: array[0..3] of word;
  FormulaResult: double;
begin
  Data[0] := 0;  // to silence the compiler...

  { Determine encoded result bytes }
  case ACell^.ContentType of
    cctNumber:
      begin
        FormulaResult := ACell^.NumberValue;
        Move(FormulaResult, Data, 8);
      end;
    cctDateTime:
      begin
        FormulaResult := ACell^.DateTimeValue;
        Move(FormulaResult, Data, 8);
      end;
    cctUTF8String:
      begin
        if ACell^.UTF8StringValue = '' then
          Data[0] := 3;
        Data[3] := $FFFF;
      end;
    cctBool:
      begin
        Data[0] := 1;
        Data[1] := ord(ACell^.BoolValue);
        Data[3] := $FFFF;
      end;
    cctError:
      begin
        Data[0] := 2;
        Data[1] := ConvertToExcelError(ACell^.ErrorValue);
        Data[3] := $FFFF;
      end;
  end;

  { Write result of the formula, encoded above }
  AStream.WriteBuffer(Data, 8);
end;
                  (*
{@@ ----------------------------------------------------------------------------
  Is called from WriteRPNFormula in the case that the cell uses a shared
  formula and writes the token "array" pointing to the shared formula base.
  This implementation is valid for BIFF3-BIFF8. BIFF2 is different, but does not
  support shared formulas; the BIFF2 writer must copy the formula found in the
  SharedFormulaBase field of the cell and adjust the relative references.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteRPNSharedFormulaLink(AStream: TStream;
  ACell: PCell; var RPNLength: Word);
type
  TSharedFormulaLinkRecord = packed record
    FormulaSize: Word;   // Size of token array
    Token: Byte;         // 1st (and only) token of the rpn formula array
    Row: Word;           // row of cell containing the shared formula
    Col: Word;           // column of cell containing the shared formula
  end;
var
  rec: TSharedFormulaLinkRecord;
begin
  rec.FormulaSize := WordToLE(5);
  rec.Token := INT_EXCEL_TOKEN_TEXP;  // Marks the cell for using a shared formula
  rec.Row := WordToLE(ACell^.SharedFormulaBase^.Row);
  rec.Col := WordToLE(ACell^.SharedFormulaBase^.Col);
  AStream.WriteBuffer(rec, SizeOf(rec));
  RPNLength := SizeOf(rec);
end;
                   *)

{@@ ----------------------------------------------------------------------------
  Writes the token array of the given RPN formula and returns its size
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteRPNTokenArray(AStream: TStream;
  ACell: PCell; const AFormula: TsRPNFormula; UseRelAddr: boolean;
  var RPNLength: Word);
var
  i: Integer;
  n: Word;
  dr, dc: Integer;
  TokenArraySizePos: Int64;
  FinalPos: Int64;
  exprDef: TsExprIdentifierDef;
  primaryExcelCode, secondaryExcelCode: Word;
begin
  RPNLength := 0;

  { The size of the token array is written later, because it's necessary to
    calculate it first, and this is done at the same time it is written }
  TokenArraySizePos := AStream.Position;
  WriteRPNTokenArraySize(AStream, 0);

  { Formula data (RPN token array) }
  for i := 0 to Length(AFormula) - 1 do begin

    { Token identifier }
    if AFormula[i].ElementKind = fekFunc then begin
      exprDef := BuiltinIdentifiers.IdentifierByName(Aformula[i].FuncName);
      if exprDef.HasFixedArgumentCount then
        primaryExcelCode := INT_EXCEL_TOKEN_FUNC_V
      else
        primaryExcelCode := INT_EXCEL_TOKEN_FUNCVAR_V;
      secondaryExcelCode := exprDef.ExcelCode;
    end else begin
      primaryExcelCode := TokenIDs[AFormula[i].ElementKind];
      secondaryExcelCode := 0;
    end;

    if UseRelAddr then
      case primaryExcelCode of
        INT_EXCEL_TOKEN_TREFR   : primaryExcelCode := INT_EXCEL_TOKEN_TREFN_R;
        INT_EXCEL_TOKEN_TREFV   : primaryExcelCode := INT_EXCEL_TOKEN_TREFN_V;
        INT_EXCEL_TOKEN_TREFA   : primaryExcelCode := INT_EXCEL_TOKEN_TREFN_A;

        INT_EXCEL_TOKEN_TAREA_R : primaryExcelCode := INT_EXCEL_TOKEN_TAREAN_R;
        INT_EXCEL_TOKEN_TAREA_V : primaryExcelCode := INT_EXCEL_TOKEN_TAREAN_V;
        INT_EXCEL_TOKEN_TAREA_A : primaryExcelCode := INT_EXCEL_TOKEN_TAREAN_A;
      end;

    AStream.WriteByte(primaryExcelCode);
    inc(RPNLength);

    { Token data }
    case primaryExcelCode of
      { Operand Tokens }
      INT_EXCEL_TOKEN_TREFR, INT_EXCEL_TOKEN_TREFV, INT_EXCEL_TOKEN_TREFA:  { fekCell }
        begin
          n := WriteRPNCellAddress(
            AStream,
            AFormula[i].Row, AFormula[i].Col,
            AFormula[i].RelFlags
          );
          inc(RPNLength, n);
        end;

      INT_EXCEL_TOKEN_TAREA_R: { fekCellRange }
        begin
          n := WriteRPNCellRangeAddress(
            AStream,
            AFormula[i].Row, AFormula[i].Col,
            AFormula[i].Row2, AFormula[i].Col2,
            AFormula[i].RelFlags
          );
          inc(RPNLength, n);
        end;

      INT_EXCEL_TOKEN_TREFN_R,
      INT_EXCEL_TOKEN_TREFN_V,
      INT_EXCEL_TOKEN_TREFN_A:  { fekCellOffset }
        begin
          if rfRelRow in AFormula[i].RelFlags
            then dr := integer(AFormula[i].Row) - ACell^.Row
            else dr := integer(AFormula[i].Row);
          if rfRelCol in AFormula[i].RelFlags
            then dc := integer(AFormula[i].Col) - ACell^.Col
            else dc := integer(AFormula[i].Col);
          n := WriteRPNCellOffset(AStream, dr, dc, AFormula[i].RelFlags);
          inc(RPNLength, n);
        end;

      INT_EXCEL_TOKEN_TNUM: { fekNum }
        begin
          AStream.WriteBuffer(AFormula[i].DoubleValue, 8);
          inc(RPNLength, 8);
        end;

      INT_EXCEL_TOKEN_TINT:  { fekNum, but integer }
        begin
          AStream.WriteWord(WordToLE(AFormula[i].IntValue));
          inc(RPNLength, 2);
        end;

      INT_EXCEL_TOKEN_TSTR: { fekString }
      { string constant is stored as widestring in BIFF8, otherwise as ansistring
        Writing is done by the virtual method WriteString_8bitLen. }
        begin
          inc(RPNLength, WriteString_8bitLen(AStream, AFormula[i].StringValue));
        end;

      INT_EXCEL_TOKEN_TBOOL:  { fekBool }
        begin
          AStream.WriteByte(ord(AFormula[i].DoubleValue <> 0.0));
          inc(RPNLength, 1);
        end;

      INT_EXCEL_TOKEN_TERR: { fekErr }
        begin
          AStream.WriteByte(ConvertToExcelError(TsErrorValue(AFormula[i].IntValue)));
          inc(RPNLength, 1);
        end;

      // Functions with fixed parameter count
      INT_EXCEL_TOKEN_FUNC_R, INT_EXCEL_TOKEN_FUNC_V, INT_EXCEL_TOKEN_FUNC_A:
        begin
          n := WriteRPNFunc(AStream, secondaryExcelCode);
          inc(RPNLength, n);
        end;

      // Functions with variable parameter count
      INT_EXCEL_TOKEN_FUNCVAR_V:
        begin
          AStream.WriteByte(AFormula[i].ParamsNum);
          n := WriteRPNFunc(AStream, secondaryExcelCode);
          inc(RPNLength, 1 + n);
        end;

      // Other operations
      INT_EXCEL_TOKEN_TATTR: { fekOpSUM }
      { 3.10, page 71: e.g. =SUM(1) is represented by token array tInt(1),tAttrRum }
        begin
          // Unary SUM Operation
          AStream.WriteByte($10); //tAttrSum token (SUM with one parameter)
          AStream.WriteByte(0); // not used
          AStream.WriteByte(0); // not used
          inc(RPNLength, 3);
        end;

    end;  // case
  end; // for

  // Now update the size of the token array.
  finalPos := AStream.Position;
  AStream.Position := TokenArraySizePos;
  WriteRPNTokenArraySize(AStream, RPNLength);
  AStream.Position := finalPos;
end;

{@@ ----------------------------------------------------------------------------
  Writes the size of the RPN token array. Called from WriteRPNFormula.
  Valid for BIFF3-BIFF8. Override in BIFF2.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteRPNTokenArraySize(AStream: TStream;
  ASize: Word);
begin
  AStream.WriteWord(WordToLE(ASize));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 3-8 ROW record
  Valid for BIFF3-BIFF8
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteRow(AStream: TStream; ASheet: TsWorksheet;
  ARowIndex, AFirstColIndex, ALastColIndex: Cardinal; ARow: PRow);
var
  w: Word;
  dw: DWord;
  cell: PCell;
  spaceabove, spacebelow: Boolean;
  colindex: Cardinal;
  rowheight: Word;
  h: Single;
  fmt: PsCellFormat;
begin
  if (ARowIndex >= FLimitations.MaxRowCount) or
     (AFirstColIndex >= FLimitations.MaxColCount) or
     (ALastColIndex >= FLimitations.MaxColCount)
  then
    exit;

  // Check for additional space above/below row
  spaceabove := false;
  spacebelow := false;
  colindex := AFirstColIndex;
  while colindex <= ALastColIndex do
  begin
    cell := ASheet.FindCell(ARowindex, colindex);
    if (cell <> nil) then
    begin
      fmt := Workbook.GetPointerToCellFormat(cell^.FormatIndex);
      if (uffBorder in fmt^.UsedFormattingFields) then
      begin
        if (cbNorth in fmt^.Border) and (fmt^.BorderStyles[cbNorth].LineStyle = lsThick)
          then spaceabove := true;
        if (cbSouth in fmt^.Border) and (fmt^.BorderStyles[cbSouth].LineStyle = lsThick)
          then spacebelow := true;
      end;
    end;
    if spaceabove and spacebelow then break;
    inc(colindex);
  end;

  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_ROW, 16);;

  { Index of row }
  AStream.WriteWord(WordToLE(Word(ARowIndex)));

  { Index to column of the first cell which is described by a cell record }
  AStream.WriteWord(WordToLE(Word(AFirstColIndex)));

  { Index to column of the last cell which is described by a cell record, increased by 1 }
  AStream.WriteWord(WordToLE(Word(ALastColIndex) + 1));

  { Row height (in twips, 1/20 point) and info on custom row height }
  h := Workbook.GetFont(0).Size;  // Point size of default font
  if (ARow = nil) or (ARow^.Height = ASheet.DefaultRowHeight) then
    rowheight := PtsToTwips((ASheet.DefaultRowHeight + ROW_HEIGHT_CORRECTION) * h)
  else
  if (ARow^.Height = 0) then
    rowheight := 0
  else
    rowheight := PtsToTwips((ARow^.Height + ROW_HEIGHT_CORRECTION)*h);
  w := rowheight and $7FFF;
  AStream.WriteWord(WordToLE(w));

  { 2 words not used }
  AStream.WriteDWord(0);

  { Option flags }
  dw := $00000100;  // bit 8 is always 1
  if spaceabove then dw := dw or $10000000;
  if spacebelow then dw := dw or $20000000;
  if (ARow <> nil) then
    dw := dw or $00000040;  // Row height and font height do not match
  AStream.WriteDWord(DWordToLE(dw));
end;

{@@ ----------------------------------------------------------------------------
  Writes all ROW records for the given sheet.
  Note that the OpenOffice documentation says that rows must be written in
  groups of 32, followed by the cells on these rows, etc. THIS IS NOT NECESSARY!
  Valid for BIFF2-BIFF8.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteRows(AStream: TStream; ASheet: TsWorksheet);
var
  row: PRow;
  i: Integer;
  cell1, cell2: PCell;
begin
  for i := 0 to ASheet.Rows.Count-1 do begin
    row := ASheet.Rows[i];
    cell1 := ASheet.Cells.GetFirstCellOfRow(row^.Row);
    if cell1 <> nil then begin
      cell2 := ASheet.Cells.GetLastCellOfRow(row^.Row);
      WriteRow(AStream, ASheet, row^.Row, cell1^.Col, cell2^.Col, row);
    end else
      WriteRow(AStream, ASheet, row^.Row, 0, 0, row);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2-8 SELECTION record
  Writes just reasonable default values
  APane is 0..3 (see below)
  Valid for BIFF2-BIFF8
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteSelection(AStream: TStream;
  ASheet: TsWorksheet; APane: Byte);
var
  activeCellRow, activeCellCol: Word;
begin
  case APane of
    0: begin   // right-bottom
         activeCellRow := ASheet.TopPaneHeight;
         activeCellCol := ASheet.LeftPaneWidth;
       end;
    1: begin   // right-top
         activeCellRow := 0;
         activeCellCol := ASheet.LeftPaneWidth;
       end;
    2: begin   // left-bottom
         activeCellRow := ASheet.TopPaneHeight;
         activeCellCol := 0;
       end;
    3: begin   // left-top
         activeCellRow := 0;
         activeCellCol := 0;
       end;
  end;

  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_SELECTION, 15);

  { Pane identifier }
  AStream.WriteByte(APane);

  { Index to row of the active cell }
  AStream.WriteWord(WordToLE(activeCellRow));

  { Index to column of the active cell }
  AStream.WriteWord(WordToLE(activeCellCol));

  { Index into the following cell range list to the entry that contains the active cell }
  AStream.WriteWord(WordToLE(0));   // there's only 1 item

  { Cell range array }

  { Count of items }
  AStream.WriteWord(WordToLE(1));  // only 1 item

  { Index to first and last row - are the same here }
  AStream.WriteWord(WordTOLE(activeCellRow));
  AStream.WriteWord(WordTOLE(activeCellRow));

  { Index to first and last column - they are the same here again. }
  { Note: BIFF8 writes bytes here! }
  AStream.WriteByte(activeCellCol);
  AStream.WriteByte(activeCellCol);
end;

procedure TsSpreadBIFFWriter.WriteSelections(AStream: TStream;
  ASheet: TsWorksheet);
begin
  WriteSelection(AStream, ASheet, 3);
  if (ASheet.LeftPaneWidth = 0) then begin
    if ASheet.TopPaneHeight > 0 then WriteSelection(AStream, ASheet, 2);
  end else begin
    WriteSelection(AStream, ASheet, 1);
    if ASheet.TopPaneHeight > 0 then begin
      WriteSelection(AStream, ASheet, 2);
      WriteSelection(AStream, ASheet, 0);
    end;
  end;
end;
                                (*
{@@ ----------------------------------------------------------------------------
  Writes the token array of a shared formula stored in ACell.
  Note: Relative cell addresses of a shared formula are defined by
  token fekCellOffset
  Valid for BIFF5-BIFF8. No shared formulas before BIFF2. But since a worksheet
  containing shared formulas can be written the BIFF2 writer needs to duplicate
  the formulas in each cell (with adjusted cell references). In BIFF2
  WriteSharedFormula must not do anything.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteSharedFormula(AStream: TStream; ACell: PCell);
var
  r1, r2, c1, c2: Cardinal;
  RPNLength: word;
  recordSizePos: Int64;
  startPos, finalPos: Int64;
  formula: TsRPNFormula;
begin
  RPNLength := 0;

  // Write BIFF record ID and size
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_SHAREDFMLA));
  recordSizePos := AStream.Position;
  AStream.WriteWord(0); // This is the record size which is not yet known here
  startPos := AStream.Position;

  // Determine (maximum) cell range covered by the shared formula in ACell.
  // Note: it is possible that the range is not contiguous.
  FWorksheet.FindSharedFormulaRange(ACell, r1, c1, r2, c2);

  // Write borders of cell range covered by the formula
  WriteSharedFormulaRange(AStream, r1, c1, r2, c2);

  // Not used
  AStream.WriteByte(0);

  // Number of existing formula records
  AStream.WriteByte((r2-r1+1) * (c2-c1+1));

  // Create an RPN formula from the shared formula base's string formula
  // and adjust relative references
  formula := FWorksheet.BuildRPNFormula(ACell^.SharedFormulaBase);

  // Writes the rpn token array
  WriteRPNTokenArray(AStream, ACell, formula, true, RPNLength);

  { Write record size at the end after we known it }
  finalPos := AStream.Position;
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(finalPos - startPos));
  AStream.Position := finalPos;
end;

{@@ ----------------------------------------------------------------------------
  Writes the borders of the cell range covered by a shared formula.
  Valid for BIFF5 and BIFF8 - BIFF8 writes 8-bit column index as well.
  Not needed in BIFF2 which does not support shared formulas.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteSharedFormulaRange(AStream: TStream;
  AFirstRow, AFirstCol, ALastRow, ALastCol: Cardinal);
begin
  // Index to first row
  AStream.WriteWord(WordToLE(AFirstRow));
  // Index to last row
  AStream.WriteWord(WordToLE(ALastRow));
  // Index to first column
  AStream.WriteByte(AFirstCol);
  // Index to last rcolumn
  AStream.WriteByte(ALastCol);
end;                              *)

{@@ ----------------------------------------------------------------------------
  Writes a SHEETPR Record.
  Valid for BIFF3-BIFF8.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteSheetPR(AStream: TStream);
var
  flags: Word;
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_SHEETPR, 2);

  flags := $04C1;
  AStream.WriteWord(WordToLE(flags));
end;

{@@ ----------------------------------------------------------------------------
  Helper function for writing a string with 8-bit length. Here, we implement the
  version for ansistrings since it is valid for all BIFF versions except BIFF8
  where it has to be overridden. Is called for writing a string rpn token.
  Returns the count of bytes written.
-------------------------------------------------------------------------------}
function TsSpreadBIFFWriter.WriteString_8bitLen(AStream: TStream;
  AString: String): Integer;
var
  len: Byte;
  s: ansistring;
begin
  s := ConvertEncoding(AString, encodingUTF8, FCodePage);
  len := Length(s);
  AStream.WriteByte(len);
  AStream.WriteBuffer(s[1], len);
  Result := 1 + len;
end;

{@@ ----------------------------------------------------------------------------
  Write the STRING record which immediately follows the RPN formula record if
  the formula result is a non-empty string.
  Must be overridden because implementation depends of BIFF version.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteStringRecord(AStream: TStream;
  AString: String);
begin
  Unused(AStream, AString);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel VCENTER record which determines whether the page is to be
  centered vertically for printing
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteVCenter(AStream: TStream);
var
  w: Word;
begin
  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_VCENTER, SizeOf(w));

  { Data }
  if poVertCentered in FWorksheet.PageLayout.Options then w := 1 else w := 0;
  AStream.WriteWord(WordToLE(w));
end;


procedure TsSpreadBIFFWriter.WriteVirtualCells(AStream: TStream);
var
  r,c: Cardinal;
  lCell: TCell;
  value: variant;
  styleCell: PCell;
begin
  for r := 0 to Workbook.VirtualRowCount-1 do
    for c := 0 to Workbook.VirtualColCount-1 do
    begin
      lCell.Row := r; // to silence a compiler hint...
      InitCell(lCell);
      value := varNull;
      styleCell := nil;
      Workbook.OnWriteCellData(Workbook, r, c, value, styleCell);
      if styleCell <> nil then lCell := styleCell^;
      lCell.Row := r;
      lCell.Col := c;
      if VarIsNull(value) then
      begin         // ignore empty cells that don't have a format
        if styleCell <> nil then
          lCell.ContentType := cctEmpty
        else
          Continue;
      end else
      if VarIsNumeric(value) then
      begin
        lCell.ContentType := cctNumber;
        lCell.NumberValue := value;
      end else
      if VarType(value) = varDate then
      begin
        lCell.ContentType := cctDateTime;
        lCell.DateTimeValue := StrToDateTime(VarToStr(value), Workbook.FormatSettings);
      end else
      if VarIsStr(value) then
      begin
        lCell.ContentType := cctUTF8String;
        lCell.UTF8StringValue := VarToStrDef(value, '');
      end else
      if VarIsBool(value) then
      begin
        lCell.ContentType := cctBool;
        lCell.BoolValue := value <> 0;
      end else
        lCell.ContentType := cctEmpty;
      WriteCellToStream(AStream, @lCell);
//      WriteCellCallback(@lCell, AStream);
      value := varNULL;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5/8 WINDOW1 record
  This record contains general settings for the document window and
  global workbook settings.
  The values written here are reasonable defaults which should work for most
  sheets.
  Valid for BIFF5-BIFF8.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteWindow1(AStream: TStream);
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_WINDOW1, 18);

  { Horizontal position of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE(0));

  { Vertical position of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($0069));

  { Width of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($339F));

  { Height of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($1B5D));

  { Option flags }
  AStream.WriteWord(WordToLE(
   MASK_WINDOW1_OPTION_HORZ_SCROLL_VISIBLE or
   MASK_WINDOW1_OPTION_VERT_SCROLL_VISIBLE or
   MASK_WINDOW1_OPTION_WORKSHEET_TAB_VISIBLE));

  { Index to active (displayed) worksheet }
  AStream.WriteWord(WordToLE($00));

  { Index of first visible tab in the worksheet tab bar }
  AStream.WriteWord(WordToLE($00));

  { Number of selected worksheets }
  AStream.WriteWord(WordToLE(1));

  { Width of worksheet tab bar (in 1/1000 of window width).
    The remaining space is used by the horizontal scroll bar }
  AStream.WriteWord(WordToLE(600));
end;

{@@ ----------------------------------------------------------------------------
  Writes an XF record needed for cell formatting.
  Is called from WriteXFRecords.
  MUST be overridden by descendents.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteXF(AStream: TStream; ACellFormat: PsCellFormat;
  XFType_Prot: Byte = 0);
begin
  Unused(AStream, ACellFormat, XFType_Prot);
end;

{@@ ----------------------------------------------------------------------------
  Writes all XF records of a worksheet. XF records define cell formatting.
  The BIFF file format requires 16 pre-defined records for internal use.
  User-defined records can follow.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFFWriter.WriteXFRecords(AStream: TStream);
var
  i: Integer;
begin
  // XF0
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF1
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF2
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF3
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF4
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF5
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF6
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF7
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF8
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF9
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF10
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF11
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF12
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF13
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF14
  WriteXF(AStream, nil, MASK_XF_TYPE_PROT_STYLE_XF);
  // XF15 - Default, no formatting
  WriteXF(AStream, nil, 0);

  // Add all further non-standard format records
  // The first style was already added --> begin loop with 1
  for i:=1 to Workbook.GetNumCellFormats - 1 do
    WriteXF(AStream, Workbook.GetPointerToCellFormat(i), 0);
end;


end.

