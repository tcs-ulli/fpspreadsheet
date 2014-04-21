unit xlscommon;

{ Comments often have links to sections in the
OpenOffice Microsoft Excel File Format document }

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils, DateUtils,
  fpspreadsheet,
  fpsutils, lconvencoding;

const
  { RECORD IDs which didn't change across versions 2-8 }
  INT_EXCEL_ID_CODEPAGE   = $0042;
  INT_EXCEL_ID_DATEMODE   = $0022;

  { Formula constants TokenID values }

  { Binary Operator Tokens 3.6}
  INT_EXCEL_TOKEN_TADD    = $03;
  INT_EXCEL_TOKEN_TSUB    = $04;
  INT_EXCEL_TOKEN_TMUL    = $05;
  INT_EXCEL_TOKEN_TDIV    = $06;
  INT_EXCEL_TOKEN_TPOWER  = $07; // Power Exponentiation ^
  INT_EXCEL_TOKEN_TCONCAT = $08; // Concatenation &
  INT_EXCEL_TOKEN_TLT     = $09; // Less than <
  INT_EXCEL_TOKEN_TLE     = $0A; // Less than or equal <=
  INT_EXCEL_TOKEN_TEQ     = $0B; // Equal =
  INT_EXCEL_TOKEN_TGE     = $0C; // Greater than or equal >=
  INT_EXCEL_TOKEN_TGT     = $0D; // Greater than >
  INT_EXCEL_TOKEN_TNE     = $0E; // Not equal <>
  INT_EXCEL_TOKEN_TISECT  = $0F; // Cell range intersection
  INT_EXCEL_TOKEN_TLIST   = $10; // Cell range list
  INT_EXCEL_TOKEN_TRANGE  = $11; // Cell range
  INT_EXCEL_TOKEN_TUPLUS  = $12; // Unary plus  +
  INT_EXCEL_TOKEN_TUMINUS = $13; // Unary minus +
  INT_EXCEL_TOKEN_TPERCENT= $14; // Percent (%, divides operand by 100)

  { Constant Operand Tokens, 3.8}
  INT_EXCEL_TOKEN_TMISSARG= $16; //missing operand
  INT_EXCEL_TOKEN_TSTR    = $17; //string
  INT_EXCEL_TOKEN_TBOOL   = $1D; //boolean
  INT_EXCEL_TOKEN_TINT    = $1E; //integer
  INT_EXCEL_TOKEN_TNUM    = $1F; //floating-point

  { Operand Tokens }
  // _R: reference; _V: value; _A: array
  INT_EXCEL_TOKEN_TREFR   = $24;
  INT_EXCEL_TOKEN_TREFV   = $44;
  INT_EXCEL_TOKEN_TREFA   = $64;
  INT_EXCEL_TOKEN_TAREA_R = $25;
  INT_EXCEL_TOKEN_TAREA_V = $45;
  INT_EXCEL_TOKEN_TAREA_A = $65;

  { Function Tokens }
  // _R: reference; _V: value; _A: array
  // Offset 0: token; offset 1: index to a built-in sheet function ( ➜ 3.111)
  INT_EXCEL_TOKEN_FUNC_R = $21;
  INT_EXCEL_TOKEN_FUNC_V = $41;
  INT_EXCEL_TOKEN_FUNC_A = $61;

  //VAR: variable number of arguments:
  INT_EXCEL_TOKEN_FUNCVAR_R = $22;
  INT_EXCEL_TOKEN_FUNCVAR_V = $42;
  INT_EXCEL_TOKEN_FUNCVAR_A = $62;

  { Built-in/worksheet functions }
  INT_EXCEL_SHEET_FUNC_COUNT      = 0;
  INT_EXCEL_SHEET_FUNC_IF         = 1;
  INT_EXCEL_SHEET_FUNC_ISNA       = 2;
  INT_EXCEL_SHEET_FUNC_ISERROR    = 3;
  INT_EXCEL_SHEET_FUNC_SUM        = 4;
  INT_EXCEL_SHEET_FUNC_AVERAGE    = 5;
  INT_EXCEL_SHEET_FUNC_MIN        = 6;
  INT_EXCEL_SHEET_FUNC_MAX        = 7;
  INT_EXCEL_SHEET_FUNC_ROW        = 8;
  INT_EXCEL_SHEET_FUNC_COLUMN     = 9;
  INT_EXCEL_SHEET_FUNC_STDEV      = 12;
  INT_EXCEL_SHEET_FUNC_SIN        = 15;
  INT_EXCEL_SHEET_FUNC_COS        = 16;
  INT_EXCEL_SHEET_FUNC_TAN        = 17;
  INT_EXCEL_SHEET_FUNC_ATAN       = 18;
  INT_EXCEL_SHEET_FUNC_PI         = 19;
  INT_EXCEL_SHEET_FUNC_SQRT       = 20;
  INT_EXCEL_SHEET_FUNC_EXP        = 21;
  INT_EXCEL_SHEET_FUNC_LN         = 22;
  INT_EXCEL_SHEET_FUNC_LOG10      = 23;
  INT_EXCEL_SHEET_FUNC_ABS        = 24; // $18
  INT_EXCEL_SHEET_FUNC_INT        = 25;
  INT_EXCEL_SHEET_FUNC_SIGN       = 26;
  INT_EXCEL_SHEET_FUNC_ROUND      = 27; // $1B
  INT_EXCEL_SHEET_FUNC_MID        = 31;
  INT_EXCEL_SHEET_FUNC_VALUE      = 33;
  INT_EXCEL_SHEET_FUNC_TRUE       = 34;
  INT_EXCEL_SHEET_FUNC_FALSE      = 35;
  INT_EXCEL_SHEET_FUNC_AND        = 36;
  INT_EXCEL_SHEET_FUNC_OR         = 37;
  INT_EXCEL_SHEET_FUNC_NOT        = 38;
  INT_EXCEL_SHEET_FUNC_VAR        = 46;
  INT_EXCEL_SHEET_FUNC_PV         = 56;
  INT_EXCEL_SHEET_FUNC_FV         = 57;
  INT_EXCEL_SHEET_FUNC_NPER       = 58;
  INT_EXCEL_SHEET_FUNC_PMT        = 59;
  INT_EXCEL_SHEET_FUNC_RATE       = 60;
  INT_EXCEL_SHEET_FUNC_RAND       = 63;
  INT_EXCEL_SHEET_FUNC_DATE       = 65; // $41
  INT_EXCEL_SHEET_FUNC_TIME       = 66; // $42
  INT_EXCEL_SHEET_FUNC_DAY        = 67;
  INT_EXCEL_SHEET_FUNC_MONTH      = 68;
  INT_EXCEL_SHEET_FUNC_YEAR       = 69;
  INT_EXCEL_SHEET_FUNC_WEEKDAY    = 70;
  INT_EXCEL_SHEET_FUNC_HOUR       = 71;
  INT_EXCEL_SHEET_FUNC_MINUTE     = 72;
  INT_EXCEL_SHEET_FUNC_SECOND     = 73;
  INT_EXCEL_SHEET_FUNC_NOW        = 74;
  INT_EXCEL_SHEET_FUNC_ROWS       = 76;
  INT_EXCEL_SHEET_FUNC_COLUMNS    = 77;
  INT_EXCEL_SHEET_FUNC_ASIN       = 98;
  INT_EXCEL_SHEET_FUNC_ACOS       = 99;
  INT_EXCEL_SHEET_FUNC_ISREF      = 105;
  INT_EXCEL_SHEET_FUNC_LOG        = 109;
  INT_EXCEL_SHEET_FUNC_CHAR       = 111;
  INT_EXCEL_SHEET_FUNC_LOWER      = 112;
  INT_EXCEL_SHEET_FUNC_UPPER      = 113;
  INT_EXCEL_SHEET_FUNC_PROPER     = 114;
  INT_EXCEL_SHEET_FUNC_LEFT       = 115;
  INT_EXCEL_SHEET_FUNC_RIGHT      = 116;
  INT_EXCEL_SHEET_FUNC_TRIM       = 118;
  INT_EXCEL_SHEET_FUNC_REPLACE    = 119;
  INT_EXCEL_SHEET_FUNC_SUBSTITUTE = 120;
  INT_EXCEL_SHEET_FUNC_CODE       = 121;
  INT_EXCEL_SHEET_FUNC_CELL       = 125;
  INT_EXCEL_SHEET_FUNC_ISERR      = 126;
  INT_EXCEL_SHEET_FUNC_ISTEXT     = 127;
  INT_EXCEL_SHEET_FUNC_ISNUMBER   = 128;
  INT_EXCEL_SHEET_FUNC_ISBLANK    = 129;
  INT_EXCEL_SHEET_FUNC_DATEVALUE  = 140;
  INT_EXCEL_SHEET_FUNC_TIMEVALUE  = 141;
  INT_EXCEL_SHEET_FUNC_COUNTA     = 169;
  INT_EXCEL_SHEET_FUNC_PRODUCT    = 183;
  INT_EXCEL_SHEET_FUNC_ISNONTEXT  = 190;
  INT_EXCEL_SHEET_FUNC_STDEVP     = 193;
  INT_EXCEL_SHEET_FUNC_VARP       = 194;
  INT_EXCEL_SHEET_FUNC_ISLOGICAL  = 198;
  INT_EXCEL_SHEET_FUNC_TODAY      = 221;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_MEDIAN     = 227;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_SINH       = 229;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_COSH       = 230;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_TANH       = 231;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_ASINH      = 232;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_ACOSH      = 233;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_ATANH      = 234;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_INFO       = 244;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_AVEDEV     = 269;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_BETADIST   = 270;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_BETAINV    = 272;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_BINOMDIST  = 273;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_CHIDIST    = 274;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_CHIINV     = 275;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_PERMUT     = 299;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_POISSON    = 300;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_SUMSQ      = 321;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_RADIANS    = 342;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_DEGREES    = 343;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_SUMIF      = 345;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_COUNTIF    = 346;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_COUNTBLANK = 347;  // not available in BIFF2
  INT_EXCEL_SHEET_FUNC_DATEDIF    = 351;  // not available in BIFF2

  { Control Tokens, Special Tokens }
//  01H tExp Matrix formula or shared formula
//  02H tTbl Multiple operation table
//  15H tParen Parentheses
//  18H tNlr Natural language reference (BIFF8)
  INT_EXCEL_TOKEN_TATTR = $19; // tAttr Special attribute
//  1AH tSheet Start of external sheet reference (BIFF2-BIFF4)
//  1BH tEndSheet End of external sheet reference (BIFF2-BIFF4)

  { Built In Color Palette Indexes }
  // Proper spelling
  BUILT_IN_COLOR_PALETTE_BLACK     = $08; // 000000H
  BUILT_IN_COLOR_PALETTE_WHITE     = $09; // FFFFFFH
  BUILT_IN_COLOR_PALETTE_RED       = $0A; // FF0000H
  BUILT_IN_COLOR_PALETTE_GREEN     = $0B; // 00FF00H
  BUILT_IN_COLOR_PALETTE_BLUE      = $0C; // 0000FFH
  BUILT_IN_COLOR_PALETTE_YELLOW    = $0D; // FFFF00H
  BUILT_IN_COLOR_PALETTE_MAGENTA   = $0E; // FF00FFH
  BUILT_IN_COLOR_PALETTE_CYAN      = $0F; // 00FFFFH
  BUILT_IN_COLOR_PALETTE_DARK_RED  = $10; // 800000H
  BUILT_IN_COLOR_PALETTE_DARK_GREEN= $11; // 008000H
  BUILT_IN_COLOR_PALETTE_DARK_BLUE = $12; // 000080H
  BUILT_IN_COLOR_PALETTE_OLIVE     = $13; // 808000H
  BUILT_IN_COLOR_PALETTE_PURPLE    = $14; // 800080H
  BUILT_IN_COLOR_PALETTE_TEAL      = $15; // 008080H
  BUILT_IN_COLOR_PALETTE_SILVER    = $16; // C0C0C0H
  BUILT_IN_COLOR_PALETTE_GREY      = $17; // 808080H

  // Spelling mistake; kept for compatibility
  BUILT_IN_COLOR_PALLETE_BLACK     = $08 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_WHITE     = $09 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_RED       = $0A deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_GREEN     = $0B deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_BLUE      = $0C deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_YELLOW    = $0D deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_MAGENTA   = $0E deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_CYAN      = $0F deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_DARK_RED  = $10 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_DARK_GREEN= $11 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_DARK_BLUE = $12 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_OLIVE     = $13 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_PURPLE    = $14 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_TEAL      = $15 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_SILVER    = $16 deprecated 'Please use the *_PALETTE version';
  BUILT_IN_COLOR_PALLETE_GREY      = $17 deprecated 'Please use the *_PALETTE version';

  EXTRA_COLOR_PALETTE_GREY10PCT    = $18; // E6E6E6H //todo: is $18 correct? see 5.74.3 Built-In Default Colour Tables
  EXTRA_COLOR_PALETTE_GREY20PCT    = $19; // CCCCCCH //todo: is $19 correct? see 5.74.3 Built-In Default Colour Tables

  { CODEPAGE record constants }
  WORD_ASCII = 367;
  WORD_UTF_16 = 1200; // BIFF 8
  WORD_CP_1250_Latin2 = 1250;
  WORD_CP_1251_Cyrillic = 1251;
  WORD_CP_1252_Latin1 = 1252; // BIFF4-BIFF5
  WORD_CP_1253_Greek = 1253;
  WORD_CP_1254_Turkish = 1254;
  WORD_CP_1255_Hebrew = 1255;
  WORD_CP_1256_Arabic = 1256;
  WORD_CP_1257_Baltic = 1257;
  WORD_CP_1258_Vietnamese = 1258;
  WORD_CP_1258_Latin1_BIFF2_3 = 32769; // BIFF2-BIFF3

  { DATEMODE record, 5.28 }
  DATEMODE_1900_BASE=1; //1/1/1900 minus 1 day in FPC TDateTime
  DATEMODE_1904_BASE=1462; //1/1/1904 in FPC TDateTime

  { FORMAT record constants }
  // Subset of the built-in formats for US Excel, 
  // including those needed for date/time output
  FORMAT_GENERAL = 0;             //general/default format
  FORMAT_FIXED_0_DECIMALS = 1;    //fixed, 0 decimals
  FORMAT_FIXED_2_DECIMALS = 2;    //fixed, 2 decimals
  FORMAT_FIXED_THOUSANDS_0_DECIMALS = 3;  //fixed, w/ thousand separator, 0 decs
  FORMAT_FIXED_THOUSANDS_2_DECIMALS = 4;  //fixed, w/ thousand separator, 2 decs
  FORMAT_CURRENCY_0_DECIMALS = 5; //currency (with currency symbol), 0 decs
  FORMAT_CURRENCY_2_DECIMALS = 7; //currency (with currency symbol), 2 decs
  FORMAT_PERCENT_0_DECIMALS = 9;  //percent, 0 decimals
  FORMAT_PERCENT_2_DECIMALS = 10; //percent, 2 decimals
  FORMAT_EXP_2_DECIMALS = 11;     //exponent, 2 decimals
  FORMAT_SCI_1_DECIMAL = 48;      //scientific, 1 decimal
  FORMAT_SHORT_DATE = 14;         //short date
  FORMAT_DATE_DM = 16;            //date D-MMM
  FORMAT_DATE_MY = 17;            //date MMM-YYYY
  FORMAT_SHORT_TIME_AM = 18;      //short time H:MM with AM
  FORMAT_LONG_TIME_AM = 19;       //long time H:MM:SS with AM
  FORMAT_SHORT_TIME = 20;         //short time H:MM
  FORMAT_LONG_TIME = 21;          //long time H:MM:SS
  FORMAT_SHORT_DATETIME = 22;     //short date+time
  FORMAT_TIME_MS = 45;            //time MM:SS
  FORMAT_TIME_MSZ = 47;           //time MM:SS.0
  FORMAT_TIME_INTERVAL = 46;      //time [hh]:mm:ss, hh can be >24


type
  TDateMode=(dm1900,dm1904); //DATEMODE values, 5.28

  // Adjusts Excel float (date, date/time, time) with the file's base date to get a TDateTime
  function ConvertExcelDateTimeToDateTime
    (const AExcelDateNum: Double; ADateMode: TDateMode): TDateTime;
  // Adjusts TDateTime with the file's base date to get
  // an Excel float value representing a time/date/datetime
  function ConvertDateTimeToExcelDateTime
    (const ADateTime: TDateTime; ADateMode: TDateMode): Double;

type
  { TsSpreadBIFFReader }
  TsSpreadBIFFReader = class(TsCustomSpreadReader)
  protected
    FCodepage: string; // in a format prepared for lconvencoding.ConvertEncoding
    FDateMode: TDateMode;
    // converts an Excel color index to a color value.
    function ExcelPaletteToFPSColor(AIndex: Word): TsColor;
    // Here we can add reading of records which didn't change across BIFF2-8 versions
    // Workbook Globals records
    procedure ReadCodePage(AStream: TStream);
    // Figures out what the base year for dates is for this file
    procedure ReadDateMode(AStream: TStream);
    // Read row info
    procedure ReadRowInfo(const AStream: TStream); virtual;
  public
    constructor Create; override;
  end;

  { TsSpreadBIFFWriter }

  TsSpreadBIFFWriter = class(TsCustomSpreadWriter)
  protected
    FDateMode: TDateMode;
    FLastRow: Integer;
    FLastCol: Word;
    function FPSColorToExcelPalette(AColor: TsColor): Word;
    procedure GetLastRowCallback(ACell: PCell; AStream: TStream);
    function GetLastRowIndex(AWorksheet: TsWorksheet): Integer;
    procedure GetLastColCallback(ACell: PCell; AStream: TStream);
    function GetLastColIndex(AWorksheet: TsWorksheet): Word;
    function FormulaElementKindToExcelTokenID(AElementKind: TFEKind; out ASecondaryID: Word): Word;
    // Other records which didn't change
    // Workbook Globals records
    // Write out used codepage for character encoding
    procedure WriteCodepage(AStream: TStream; AEncoding: TsEncoding);
    // Writes out DATEMODE record depending on FDateMode
    procedure WriteDateMode(AStream: TStream);
  public
    constructor Create; override;
  end;

function IsExpNumberFormat(s: String; out Decimals: Word): Boolean;
function IsFixedNumberFormat(s: String; out Decimals: Word): Boolean;
function IsPercentNumberFormat(s: String; out Decimals: Word): Boolean;
function IsThousandSepNumberFormat(s: String; out Decimals: Word): Boolean;

function IsDateFormat(s: String): Boolean;
function IsTimeFormat(s: String; out isLong, isAMPM, isMillisec: Boolean): Boolean;


implementation

function ConvertExcelDateTimeToDateTime(
  const AExcelDateNum: Double; ADateMode: TDateMode): TDateTime;
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
        result:=61.0-1.0+DATEMODE_1900_BASE-1.0
      else
        result:=AExcelDateNum+DATEMODE_1900_BASE-1.0;
    end;
    dm1904:
      result:=AExcelDateNum+DATEMODE_1904_BASE;
    else
      raise Exception.CreateFmt('ConvertExcelDateTimeToDateTime: unknown datemode %d. Please correct fpspreadsheet source code. ', [ADateMode]);
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

{ TsSpreadBIFFReader }

constructor TsSpreadBIFFReader.Create;
begin
  inherited Create;
  // Initial base date in case it won't be read from file
  FDateMode := dm1900;
end;

function TsSpreadBIFFReader.ExcelPaletteToFPSColor(AIndex: Word): TsColor;
begin
  case AIndex of
    BUILT_IN_COLOR_PALLETE_BLACK : Result := scBlack;
    BUILT_IN_COLOR_PALLETE_WHITE: Result := scWhite;
    BUILT_IN_COLOR_PALLETE_RED: Result := scRed;
    BUILT_IN_COLOR_PALLETE_GREEN: Result := scGreen;
    BUILT_IN_COLOR_PALLETE_BLUE: Result := scBlue;
    BUILT_IN_COLOR_PALLETE_YELLOW: Result := scYellow;
    BUILT_IN_COLOR_PALLETE_MAGENTA: Result := scMagenta;
    BUILT_IN_COLOR_PALLETE_CYAN: Result := scCyan;
    BUILT_IN_COLOR_PALLETE_DARK_RED: Result := scDarkRed;
    BUILT_IN_COLOR_PALLETE_DARK_GREEN: Result := scDarkGreen;
    BUILT_IN_COLOR_PALLETE_DARK_BLUE: Result := scDarkBlue;
    BUILT_IN_COLOR_PALLETE_OLIVE: Result := scOlive;
    BUILT_IN_COLOR_PALLETE_PURPLE: Result := scPurple;
    BUILT_IN_COLOR_PALLETE_TEAL: Result := scTeal;
    BUILT_IN_COLOR_PALLETE_SILVER: Result := scSilver;
    BUILT_IN_COLOR_PALLETE_GREY: Result := scGrey;
    //
    EXTRA_COLOR_PALETTE_GREY10PCT: Result := scGrey10pct;
    EXTRA_COLOR_PALETTE_GREY20PCT: Result := scGrey20pct;
  end;
end;

// In BIFF 8 it seams to always use the UTF-16 codepage
procedure TsSpreadBIFFReader.ReadCodePage(AStream: TStream);
var
  lCodePage: Word;
begin
  { Codepage }
  lCodePage := WordLEToN(AStream.ReadWord());

  case lCodePage of
  // 016FH = 367 = ASCII
  // 01B5H = 437 = IBM PC CP-437 (US)
  //02D0H = 720 = IBM PC CP-720 (OEM Arabic)
  //02E1H = 737 = IBM PC CP-737 (Greek)
  //0307H = 775 = IBM PC CP-775 (Baltic)
  //0352H = 850 = IBM PC CP-850 (Latin I)
  $0352: FCodepage := 'cp850';
  //0354H = 852 = IBM PC CP-852 (Latin II (Central European))
  $0354: FCodepage := 'cp852';
  //0357H = 855 = IBM PC CP-855 (Cyrillic)
  $0357: FCodepage := 'cp855';
  //0359H = 857 = IBM PC CP-857 (Turkish)
  $0359: FCodepage := 'cp857';
  //035AH = 858 = IBM PC CP-858 (Multilingual Latin I with Euro)
  //035CH = 860 = IBM PC CP-860 (Portuguese)
  //035DH = 861 = IBM PC CP-861 (Icelandic)
  //035EH = 862 = IBM PC CP-862 (Hebrew)
  //035FH = 863 = IBM PC CP-863 (Canadian (French))
  //0360H = 864 = IBM PC CP-864 (Arabic)
  //0361H = 865 = IBM PC CP-865 (Nordic)
  //0362H = 866 = IBM PC CP-866 (Cyrillic (Russian))
  //0365H = 869 = IBM PC CP-869 (Greek (Modern))
  //036AH = 874 = Windows CP-874 (Thai)
  //03A4H = 932 = Windows CP-932 (Japanese Shift-JIS)
  //03A8H = 936 = Windows CP-936 (Chinese Simplified GBK)
  //03B5H = 949 = Windows CP-949 (Korean (Wansung))
  //03B6H = 950 = Windows CP-950 (Chinese Traditional BIG5)
  //04B0H = 1200 = UTF-16 (BIFF8)
  $04B0: FCodepage := 'utf-16';
  //04E2H = 1250 = Windows CP-1250 (Latin II) (Central European)
  //04E3H = 1251 = Windows CP-1251 (Cyrillic)
  //04E4H = 1252 = Windows CP-1252 (Latin I) (BIFF4-BIFF5)
  //04E5H = 1253 = Windows CP-1253 (Greek)
  //04E6H = 1254 = Windows CP-1254 (Turkish)
  $04E6: FCodepage := 'cp1254';
  //04E7H = 1255 = Windows CP-1255 (Hebrew)
  //04E8H = 1256 = Windows CP-1256 (Arabic)
  //04E9H = 1257 = Windows CP-1257 (Baltic)
  //04EAH = 1258 = Windows CP-1258 (Vietnamese)
  //0551H = 1361 = Windows CP-1361 (Korean (Johab))
  //2710H = 10000 = Apple Roman
  //8000H = 32768 = Apple Roman
  //8001H = 32769 = Windows CP-1252 (Latin I) (BIFF2-BIFF3)
  end;
end;

procedure TsSpreadBIFFReader.ReadDateMode(AStream: TStream);
var
  lBaseMode: Word;
begin
  //5.28 DATEMODE
  //BIFF2 BIFF3 BIFF4 BIFF5 BIFF8
  //0022H 0022H 0022H 0022H 0022H
  //This record specifies the base date for displaying date values. All dates are stored as count of days past this base date. In
  //BIFF2-BIFF4 this record is part of the Calculation Settings Block (➜4.3). In BIFF5-BIFF8 it is stored in the Workbook
  //Globals Substream.
  //Record DATEMODE, BIFF2-BIFF8:
  //Offset Size Contents
  //0 2 0 = Base date is 1899-Dec-31 (the cell value 1 represents 1900-Jan-01)
  //    1 = Base date is 1904-Jan-01 (the cell value 1 represents 1904-Jan-02)
  lBaseMode := WordLEtoN(AStream.ReadWord);
  case lBaseMode of
    0: FDateMode := dm1900;
    1: FDateMode := dm1904;
    else raise Exception.CreateFmt('Error reading file. Got unknown date mode number %d.',[lBaseMode]);
  end;
end;

// Read the part of the ROW record that is common to all BIFF versions
procedure TsSpreadBIFFReader.ReadRowInfo(const AStream: TStream);
type
  TRowRecord = packed record
    RowIndex: Word;
    Col1: Word;
    Col2: Word;
    Height: Word;
  end;
var
  rowrec: TRowRecord;
  lRow: PRow;
  h: word;
begin
  AStream.ReadBuffer(rowrec, SizeOf(TRowRecord));
  h := WordLEToN(rowrec.Height);
  if h and $8000 = 0 then begin // if this bit were set, rowheight would be default
    lRow := FWorksheet.GetRow(WordLEToN(rowrec.RowIndex));
    // Row height is encoded into the 15 remaining bits in units "twips" (1/20 pt)
    lRow^.Height := TwipsToMillimeters(h and $7FFF);
  end;
end;

function TsSpreadBIFFWriter.FPSColorToExcelPalette(AColor: TsColor): Word;
begin
  case AColor of
    scBlack: Result := BUILT_IN_COLOR_PALLETE_BLACK;
    scWhite: Result := BUILT_IN_COLOR_PALLETE_WHITE;
    scRed: Result := BUILT_IN_COLOR_PALLETE_RED;
    scGREEN: Result := BUILT_IN_COLOR_PALLETE_GREEN;
    scBLUE: Result := BUILT_IN_COLOR_PALLETE_BLUE;
    scYELLOW: Result := BUILT_IN_COLOR_PALLETE_YELLOW;
    scMAGENTA: Result := BUILT_IN_COLOR_PALLETE_MAGENTA;
    scCYAN: Result := BUILT_IN_COLOR_PALLETE_CYAN;
    scDarkRed: Result := BUILT_IN_COLOR_PALLETE_DARK_RED;
    scDarkGreen: Result := BUILT_IN_COLOR_PALLETE_DARK_GREEN;
    scDarkBlue: Result := BUILT_IN_COLOR_PALLETE_DARK_BLUE;
    scOLIVE: Result := BUILT_IN_COLOR_PALLETE_OLIVE;
    scPURPLE: Result := BUILT_IN_COLOR_PALLETE_PURPLE;
    scTEAL: Result := BUILT_IN_COLOR_PALLETE_TEAL;
    scSilver: Result := BUILT_IN_COLOR_PALLETE_SILVER;
    scGrey: Result := BUILT_IN_COLOR_PALLETE_GREY;
    //
    scGrey10pct: Result := EXTRA_COLOR_PALETTE_GREY10PCT;
    scGrey20pct: Result := EXTRA_COLOR_PALETTE_GREY20PCT;
  end;
end;

procedure TsSpreadBIFFWriter.GetLastRowCallback(ACell: PCell; AStream: TStream);
begin
  if ACell^.Row > FLastRow then FLastRow := ACell^.Row;
end;

function TsSpreadBIFFWriter.GetLastRowIndex(AWorksheet: TsWorksheet): Integer;
begin
  FLastRow := 0;
  IterateThroughCells(nil, AWorksheet.Cells, GetLastRowCallback);
  Result := FLastRow;
end;

procedure TsSpreadBIFFWriter.GetLastColCallback(ACell: PCell; AStream: TStream);
begin
  if ACell^.Col > FLastCol then FLastCol := ACell^.Col;
end;

function TsSpreadBIFFWriter.GetLastColIndex(AWorksheet: TsWorksheet): Word;
begin
  FLastCol := 0;
  IterateThroughCells(nil, AWorksheet.Cells, GetLastColCallback);
  Result := FLastCol;
end;

function TsSpreadBIFFWriter.FormulaElementKindToExcelTokenID(
  AElementKind: TFEKind; out ASecondaryID: Word): Word;
const
  { Explanation of first index:
     0 --> primary token (basic operands and operations)
     1 --> secondary token of a function with a fixed parameter count
     2 --> secondary token of a function with a variable parameter count }
  TokenIDs: array[fekCell..fekOpSum, 0..1] of Word = (
    // Basic operands
    (0, INT_EXCEL_TOKEN_TREFV),          {fekCell}
    (0, INT_EXCEL_TOKEN_TREFR),          {fekCellRef}
    (0, INT_EXCEL_TOKEN_TAREA_R),        {fekCellRange}
    (0, INT_EXCEL_TOKEN_TNUM),           {fekNum}
    (0, INT_EXCEL_TOKEN_TSTR),           {fekString}
    (0, INT_EXCEL_TOKEN_TBOOL),          {fekBool}
    (0, INT_EXCEL_TOKEN_TMISSARG),       {fekMissArg, missing argument}

    // Basic operations
    (0, INT_EXCEL_TOKEN_TADD),           {fekAdd, +}
    (0, INT_EXCEL_TOKEN_TSUB),           {fekSub, -}
    (0, INT_EXCEL_TOKEN_TDIV),           {fekDiv, /}
    (0, INT_EXCEL_TOKEN_TMUL),           {fekMul, *}
    (0, INT_EXCEL_TOKEN_TPERCENT),       {fekPercent, %}
    (0, INT_EXCEL_TOKEN_TPOWER),         {fekPower, ^}
    (0, INT_EXCEL_TOKEN_TUMINUS),        {fekUMinus, -}
    (0, INT_EXCEL_TOKEN_TUPLUS),         {fekUPlus, +}
    (0, INT_EXCEL_TOKEN_TCONCAT),        {fekConcat, &, for strings}
    (0, INT_EXCEL_TOKEN_TEQ),            {fekEqual, =}
    (0, INT_EXCEL_TOKEN_TGT),            {fekGreater, >}
    (0, INT_EXCEL_TOKEN_TGE),            {fekGreaterEqual, >=}
    (0, INT_EXCEL_TOKEN_TLT),            {fekLess <}
    (0, INT_EXCEL_TOKEN_TLE),            {fekLessEqual, <=}
    (0, INT_EXCEL_TOKEN_TNE),            {fekNotEqual, <>}

    // Math functions
    (1, INT_EXCEL_SHEET_FUNC_ABS),       {fekABS}
    (1, INT_EXCEL_SHEET_FUNC_ACOS),      {fekACOS}
    (1, INT_EXCEL_SHEET_FUNC_ACOSH),     {fekACOSH}
    (1, INT_EXCEL_SHEET_FUNC_ASIN),      {fekASIN}
    (1, INT_EXCEL_SHEET_FUNC_ASINH),     {fekASINH}
    (1, INT_EXCEL_SHEET_FUNC_ATAN),      {fekATAN}
    (1, INT_EXCEL_SHEET_FUNC_ATANH),     {fekATANH}
    (1, INT_EXCEL_SHEET_FUNC_COS),       {fekCOS}
    (1, INT_EXCEL_SHEET_FUNC_COSH),      {fekCOSH}
    (1, INT_EXCEL_SHEET_FUNC_DEGREES),   {fekDEGREES}
    (1, INT_EXCEL_SHEET_FUNC_EXP),       {fekEXP}
    (1, INT_EXCEL_SHEET_FUNC_INT),       {fekINT}
    (1, INT_EXCEL_SHEET_FUNC_LN),        {fekLN}
    (1, INT_EXCEL_SHEET_FUNC_LOG),       {fekLOG}
    (1, INT_EXCEL_SHEET_FUNC_LOG10),     {fekLOG10}
    (1, INT_EXCEL_SHEET_FUNC_PI),        {fekPI}
    (1, INT_EXCEL_SHEET_FUNC_RADIANS),   {fekRADIANS}
    (1, INT_EXCEL_SHEET_FUNC_RAND),      {fekRAND}
    (1, INT_EXCEL_SHEET_FUNC_ROUND),     {fekROUND}
    (1, INT_EXCEL_SHEET_FUNC_SIGN),      {fekSIGN}
    (1, INT_EXCEL_SHEET_FUNC_SIN),       {fekSIN}
    (1, INT_EXCEL_SHEET_FUNC_SINH),      {fekSINH}
    (1, INT_EXCEL_SHEET_FUNC_SQRT),      {fekSQRT}
    (1, INT_EXCEL_SHEET_FUNC_TAN),       {fekTAN}
    (1, INT_EXCEL_SHEET_FUNC_TANH),      {fekTANH}

    // Date/time functions
    (1, INT_EXCEL_SHEET_FUNC_DATE),      {fekDATE}
    (1, INT_EXCEL_SHEET_FUNC_DATEDIF),   {fekDATEDIF}
    (1, INT_EXCEL_SHEET_FUNC_DATEVALUE), {fekDATEVALUE}
    (1, INT_EXCEL_SHEET_FUNC_DAY),       {fekDAY}
    (1, INT_EXCEL_SHEET_FUNC_HOUR),      {fekHOUR}
    (1, INT_EXCEL_SHEET_FUNC_MINUTE),    {fekMINUTE}
    (1, INT_EXCEL_SHEET_FUNC_MONTH),     {fekMONTH}
    (1, INT_EXCEL_SHEET_FUNC_NOW),       {fekNOW}
    (1, INT_EXCEL_SHEET_FUNC_SECOND),    {fekSECOND}
    (1, INT_EXCEL_SHEET_FUNC_TIME),      {fekTIME}
    (1, INT_EXCEL_SHEET_FUNC_TIMEVALUE), {fekTIMEVALUE}
    (1, INT_EXCEL_SHEET_FUNC_TODAY),     {fekTODAY}
    (2, INT_EXCEL_SHEET_FUNC_WEEKDAY),   {fekWEEKDAY}
    (1, INT_EXCEL_SHEET_FUNC_YEAR),      {fekYEAR}

    // Statistical functions
    (2, INT_EXCEL_SHEET_FUNC_AVEDEV),    {fekAVEDEV}
    (2, INT_EXCEL_SHEET_FUNC_AVERAGE),   {fekAVERAGE}
    (2, INT_EXCEL_SHEET_FUNC_BETADIST),  {fekBETADIST}
    (2, INT_EXCEL_SHEET_FUNC_BETAINV),   {fekBETAINV}
    (1, INT_EXCEL_SHEET_FUNC_BINOMDIST), {fekBINOMDIST}
    (1, INT_EXCEL_SHEET_FUNC_CHIDIST),   {fekCHIDIST}
    (1, INT_EXCEL_SHEET_FUNC_CHIINV),    {fekCHIINV}
    (2, INT_EXCEL_SHEET_FUNC_COUNT),     {fekCOUNT}
    (2, INT_EXCEL_SHEET_FUNC_COUNTA),    {fekCOUNTA}
    (1, INT_EXCEL_SHEET_FUNC_COUNTBLANK),{fekCOUNTBLANK}
    (2, INT_EXCEL_SHEET_FUNC_COUNTIF),   {fekCOUNTIF}
    (2, INT_EXCEL_SHEET_FUNC_MAX),       {fekMAX}
    (2, INT_EXCEL_SHEET_FUNC_MEDIAN),    {fekMEDIAN}
    (2, INT_EXCEL_SHEET_FUNC_MIN),       {fekMIN}
    (1, INT_EXCEL_SHEET_FUNC_PERMUT),    {fekPERMUT}
    (1, INT_EXCEL_SHEET_FUNC_POISSON),   {fekPOISSON}
    (2, INT_EXCEL_SHEET_FUNC_PRODUCT),   {fekPRODUCT}
    (2, INT_EXCEL_SHEET_FUNC_STDEV),     {fekSTDEV}
    (2, INT_EXCEL_SHEET_FUNC_STDEVP),    {fekSTDEVP}
    (2, INT_EXCEL_SHEET_FUNC_SUM),       {fekSUM}
    (2, INT_EXCEL_SHEET_FUNC_SUMIF),     {fekSUMIF}
    (2, INT_EXCEL_SHEET_FUNC_SUMSQ),     {fekSUMSQ}
    (2, INT_EXCEL_SHEET_FUNC_VAR),       {fekVAR}
    (2, INT_EXCEL_SHEET_FUNC_VARP),      {fekVARP}

    // Financial functions
    (2, INT_EXCEL_SHEET_FUNC_FV),        {fekFV}
    (2, INT_EXCEL_SHEET_FUNC_NPER),      {fekNPER}
    (2, INT_EXCEL_SHEET_FUNC_PV),        {fekPV}
    (2, INT_EXCEL_SHEET_FUNC_PMT),       {fekPMT}
    (2, INT_EXCEL_SHEET_FUNC_RATE),      {fekRATE}

    // Logical functions
    (2, INT_EXCEL_SHEET_FUNC_AND),       {fekAND}
    (1, INT_EXCEL_SHEET_FUNC_FALSE),     {fekFALSE}
    (2, INT_EXCEL_SHEET_FUNC_IF),        {fekIF}
    (1, INT_EXCEL_SHEET_FUNC_NOT),       {fekNOT}
    (2, INT_EXCEL_SHEET_FUNC_OR),        {fekOR}
    (1, INT_EXCEL_SHEET_FUNC_TRUE),      {fekTRUE}

    // String functions
    (1, INT_EXCEL_SHEET_FUNC_CHAR),      {fekCHAR}
    (1, INT_EXCEL_SHEET_FUNC_CODE),      {fekCODE}
    (2, INT_EXCEL_SHEET_FUNC_LEFT),      {fekLEFT}
    (1, INT_EXCEL_SHEET_FUNC_LOWER),     {fekLOWER}
    (1, INT_EXCEL_SHEET_FUNC_MID),       {fekMID}
    (1, INT_EXCEL_SHEET_FUNC_PROPER),    {fekPROPER}
    (1, INT_EXCEL_SHEET_FUNC_REPLACE),   {fekREPLACE}
    (2, INT_EXCEL_SHEET_FUNC_RIGHT),     {fekRIGHT}
    (2, INT_EXCEL_SHEET_FUNC_SUBSTITUTE),{fekSUBSTITUTE}
    (1, INT_EXCEL_SHEET_FUNC_TRIM),      {fekTRIM}
    (1, INT_EXCEL_SHEET_FUNC_UPPER),     {fekUPPER}

    // lookup/reference functions
    (2, INT_EXCEL_SHEET_FUNC_COLUMN),    {fekCOLUMN}
    (1, INT_EXCEL_SHEET_FUNC_COLUMNS),   {fekCOLUMNS}
    (2, INT_EXCEL_SHEET_FUNC_ROW),       {fekROW}
    (1, INT_EXCEL_SHEET_FUNC_ROWS),      {fekROWS}

    // Info functions
    (2, INT_EXCEL_SHEET_FUNC_CELL),      {fekCELLINFO}
    (1, INT_EXCEL_SHEET_FUNC_INFO),      {fekINFO}
    (1, INT_EXCEL_SHEET_FUNC_ISBLANK),   {fekIsBLANK}
    (1, INT_EXCEL_SHEET_FUNC_ISERR),     {fekIsERR}
    (1, INT_EXCEL_SHEET_FUNC_ISERROR),   {fekIsERROR}
    (1, INT_EXCEL_SHEET_FUNC_ISLOGICAL), {fekIsLOGICAL}
    (1, INT_EXCEL_SHEET_FUNC_ISNA),      {fekIsNA}
    (1, INT_EXCEL_SHEET_FUNC_ISNONTEXT), {fekIsNONTEXT}
    (1, INT_EXCEL_SHEET_FUNC_ISNUMBER),  {fekIsNUMBER}
    (1, INT_EXCEL_SHEET_FUNC_ISREF),     {fekIsREF}
    (1, INT_EXCEL_SHEET_FUNC_ISTEXT),    {fekIsTEXT}
    (1, INT_EXCEL_SHEET_FUNC_VALUE),     {fekValue}

    // Other operations
    (0, INT_EXCEL_TOKEN_TATTR)           {fekOpSum}
  );

begin
  case TokenIDs[AElementKind, 0] of
    0: begin
         Result := TokenIDs[AElementKind, 1];
         ASecondaryID := 0;
       end;
    1: begin
         Result := INT_EXCEL_TOKEN_FUNC_V;
         ASecondaryID := TokenIDs[AElementKind, 1]
       end;
    2: begin
         Result := INT_EXCEL_TOKEN_FUNCVAR_V;
         ASecondaryID := TokenIDs[AElementKind, 1]
       end;
  end;
end;

procedure TsSpreadBIFFWriter.WriteCodepage(AStream: TStream;
  AEncoding: TsEncoding);
var
  lCodepage: Word;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_CODEPAGE));
  AStream.WriteWord(WordToLE(2));

  { Codepage }
  case AEncoding of
    seLatin2:   lCodepage := WORD_CP_1250_Latin2;
    seCyrillic: lCodepage := WORD_CP_1251_Cyrillic;
    seGreek:    lCodepage := WORD_CP_1253_Greek;
    seTurkish:  lCodepage := WORD_CP_1254_Turkish;
    seHebrew:   lCodepage := WORD_CP_1255_Hebrew;
    seArabic:   lCodepage := WORD_CP_1256_Arabic;
  else
    // Default is Latin1
    lCodepage := WORD_CP_1252_Latin1;
  end;
  AStream.WriteWord(WordToLE(lCodepage));
end;

procedure TsSpreadBIFFWriter.WriteDateMode(AStream: TStream);
begin
  { BIFF Record header }
  // todo: check whether this is in the right place. should end up in workbook globals stream
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_DATEMODE));
  AStream.WriteWord(WordToLE(2));

  case FDateMode of
    dm1900: AStream.WriteWord(WordToLE(0));
    dm1904: AStream.WriteWord(WordToLE(1));
    else raise Exception.CreateFmt('Unknown datemode number %d. Please correct fpspreadsheet code.', [FDateMode]);
  end;
end;

constructor TsSpreadBIFFWriter.Create;
begin
  inherited Create;
  // Initial base date in case it won't be set otherwise.
  // Use 1900 to get a bit more range between 1900..1904.
  FDateMode := dm1900;
end;


{ Format checking procedures }

{ This simple parsing procedure of the Excel format string checks for a fixed
  float format s, i.e. s can be '0', '0.00', '000', '0,000', and returns the
  number of decimals, i.e. number of zeros behind the decimal point }
function IsFixedNumberFormat(s: String; out Decimals: Word): Boolean;
var
  i: Integer;
  p: Integer;
  decs: String;
begin
  Decimals := 0;

  // Check if s is a valid format mask.
  try
    FormatFloat(s, 1.0);
  except
    on EConvertError do begin
      Result := false;
      exit;
    end;
  end;

  // If it is count the zeros - each one is a decimal.
  if s = '0' then
    Result := true
  else begin
    p := pos('.', s);  // position of decimal point;
    if p = 0 then begin
      Result := false;
    end else begin
      Result := true;
      for i:= p+1 to Length(s) do
        if s[i] = '0' then begin
          inc(Decimals)
        end
        else
          exit;     // ignore characters after the last 0
    end;
  end;
end;

{ This function checks whether the format string corresponds to a thousand
  separator format like "#,##0.000' and returns the number of fixed decimals
  (i.e. zeros after the decimal point) }
function IsThousandSepNumberFormat(s: String; out Decimals: Word): Boolean;
var
  i, p: Integer;
begin
  Decimals := 0;

  // Check if s is a valid format string
  try
    FormatFloat(s, 1.0);
  except
    on EConvertError do begin
      Result := false;
      exit;
    end;
  end;

  // If it is look for the thousand separator. If found count decimals.
  Result := (Pos(',', s) > 0);
  if Result then begin
    p := pos('.', s);
    if p > 0 then
      for i := p+1 to Length(s) do
        if s[i] = '0' then
          inc(Decimals)
        else
          exit;  // ignore format characters after the last 0
  end;
end;


{ This function checks whether the format string corresponds to percent
  formatting and determines the number of decimals }
function IsPercentNumberFormat(s: String; out Decimals: Word): Boolean;
var
  i, p: Integer;
begin
  Decimals := 0;
  // The signature of the percent format is a percent sign at the end of the
  // format string.
  Result := (s <> '') and (s[Length(s)] = '%');
  if Result then begin
    // Check for a valid format string
    Delete(s, Length(s), 1);
    try
      FormatDateTime(s, 1.0);
    except
      on EConvertError do begin
        Result := false;
        exit;
      end;
    end;
    // Count decimals
    p := pos('.', s);
    if p > 0 then
      for i := p+1 to Length(s)-1 do
        if s[i] = '0' then
          inc(Decimals)
        else
          exit;  // ignore characters after last 0
  end;
end;

{ This function checks whether the format string corresponds to exponential
  formatting and determines the number of decimals  }
function IsExpNumberFormat(s: String; out Decimals: Word): Boolean;
var
  i, p, pe: Integer;
begin
  Result := false;
  Decimals := 0;

  if SameText(s, 'General') then
    exit;

  // Check for a valid format string
  try
    FormatDateTime(s, 1.0);
  except
    on EConvertError do begin
      exit;
    end;
  end;

  // Count decimals
  pe := pos('e', lowercase(s));
  result := pe > 0;
  if Result then begin
    p := pos('.', s);
    if (p > 0) then begin
      if p < pe then
        for i:=1 to pe-1 do
          if s[i] = '0' then
            inc(Decimals)
          else
            exit;   // ignore characters after last 0
    end;
  end;
end;

{ IsDateFormat checks if the format string s corresponds to a date format }
function IsDateFormat(s: String): Boolean;
begin
  // Day, month, year are separated by a slash
  Result := (pos('/', s) > 0);
  if Result then
    // Check validity of format string
    try
      FormatDateTime(s, now);
    except on EConvertError do
      Result := false;
    end;
end;

{ IsTimeFormat checks if the format string s is a time format. isLong is
  true if the string contains hours, minutes and seconds (two colons).
  isAMPM is true if the string contains "AM/PM", "A/P" or "AMPM".
  isMilliSec is true if the string ends with a "z". }
function IsTimeFormat(s: String; out isLong, isAMPM, isMillisec: Boolean): Boolean;
var
  p, i, count: Integer;
begin
  // Time parts are separated by a colon
  p := pos(':', s);
  isLong := false;
  isAMPM := false;
  result := p > 0;

  if Result then begin
    count := 1;
    s := Uppercase(s);

    // If there are is a second colon s is a "long" time format
    for i:=p+1 to Length(s) do
      if s[i] = ':' then begin
        isLong := true;
        break;
      end;

    // Seek for "AM/PM" etc to detect that specific format
    isAMPM := (pos('AM/PM', s) > 0) or (pos('A/P', s) > 0) or (pos('AMPM', s) > 0);

    // Look for the "milliseconds" character z
    isMilliSec := (s[Length(s)] = 'Z');

    // Check validity of format string
    try
      FormatDateTime(s, now);
    except on EConvertError do
      Result := false;
    end;
  end;
end;

end.

