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
  Classes, SysUtils, fpimage, AVL_Tree, avglvltree, lconvencoding, fpsutils;

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

   NOTE: When adding or rearranging items make sure to keep the TokenID table
   in TsSpreadBIFFWriter.FormulaElementKindToExcelTokenID, unit xlscommon,
   in sync !!!
  }

  TFEKind = (
    { Basic operands }
    fekCell, fekCellRef, fekCellRange, fekNum, fekString, fekBool, fekMissingArg,
    { Basic operations }
    fekAdd, fekSub, fekDiv, fekMul, fekPercent, fekPower, fekUMinus, fekUPlus,
    fekConcat,  // string concatenation
    fekEqual, fekGreater, fekGreaterEqual, fekLess, fekLessEqual, fekNotEqual,
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
    fekFV, fekNPER, fekPV, fekPMT, fekRATE,
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
    cctUTF8String, cctDateTime);

  {@@ List of possible formatting fields }

  TsUsedFormattingField = (uffTextRotation, uffBold, uffBorder, uffBackgroundColor,
    uffNumberFormat, uffWordWrap, uffHorAlign, uffVertAlign);

  {@@ Describes which formatting fields are active }

  TsUsedFormattingFields = set of TsUsedFormattingField;

  {@@ Number/cell formatting. Only uses a subset of the default formats,
      enough to be able to read/write date values.
  }

  TsNumberFormat = (nfGeneral, nfFixed, nfFixedTh, nfExp, nfSci, nfPercentage,
    nfShortDateTime, nfFmtDateTime, nfShortDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfTimeInterval);

  {@@ Text rotation formatting. The text is rotated relative to the standard
      orientation, which is from left to right horizontal: --->
                                                           ABC

      So 90 degrees clockwise means that the text will be:
       |  A
       |  B
      \|/ C

      And 90 degree counter clockwise will be:

      /|\ C
       |  B
       |  A
  }

  TsTextRotation = (trHorizontal, rt90DegreeClockwiseRotation,
    rt90DegreeCounterClockwiseRotation);

  {@@ Indicates the border for a cell }

  TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth);

  {@@ Indicates the border for a cell }

  TsCellBorders = set of TsCellBorder;

  {@@ Indicates horizontal and vertical text alignment in cells }
  TsHorAlignment = (haDefault, haLeft, haCenter, haRight);
  TsVertAlignment = (vaDefault, vaTop, vaCenter, vaBottom);

  {@@ Colors in FPSpreadsheet as given by a palette to be compatible with Excel.
   However, please note that they are physically written to XLS file as
   ABGR (where A is 0) }

  TsColor = (   // R G B  color value:
    scBlack ,   // 000000H
    scWhite,    // FFFFFFH
    scRed,      // FF0000H
    scGREEN,    // 00FF00H
    scBLUE,     // 0000FFH
    scYELLOW,   // FFFF00H
    scMAGENTA,  // FF00FFH
    scCYAN,     // 00FFFFH
    scDarkRed,  // 800000H
    scDarkGreen,// 008000H
    scDarkBlue, // 000080H
    scOLIVE,    // 808000H
    scPURPLE,   // 800080H
    scTEAL,     // 008080H
    scSilver,   // C0C0C0H
    scGrey,     // 808080H
    //
    scGrey10pct,// E6E6E6H
    scGrey20pct,// CCCCCCH
    scOrange,   // ffa500H
    scDarkBrown,// a0522dH
    scBrown,    // cd853fH
    scBeige,    // f5f5dcH
    scWheat,    // f5deb3H
    //
    scRGBCOLOR  // Defined via TFPColor
  );

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
    { Formatting fields }
    UsedFormattingFields: TsUsedFormattingFields;
    TextRotation: TsTextRotation;
    HorAlignment: TsHorAlignment;
    VertAlignment: TsVertAlignment;
    Border: TsCellBorders;
    BackgroundColor: TsColor;
    NumberFormat: TsNumberFormat;
    NumberFormatStr: String;
    NumberDecimals: Word;
    RGBBackgroundColor: TFPColor; // only valid if BackgroundColor=scRGBCOLOR
  end;

  PCell = ^TCell;

  TRow = record
    Row: Cardinal;
    Height: Single; // in millimeters
  end;

  PRow = ^TRow;

  TCol = record
    Col: Byte;
    Width: Single; // in "characters". Excel uses the with of char "0" in 1st font
  end;

  PCol = ^TCol;

type

  TsCustomSpreadReader = class;
  TsCustomSpreadWriter = class;

  { TsWorksheet }

  TsWorksheet = class
  private
    FCells: TAvlTree; // Items are TCell
    FCurrentNode: TAVLTreeNode; // For GetFirstCell and GetNextCell
    FRows, FCols: TIndexedAVLTree; // This lists contain only rows or cols with styles different from the standard
    procedure RemoveCallback(data, arg: pointer);
  public
    Name: string;
    { Base methods }
    constructor Create;
    destructor Destroy; override;
    { Utils }
    class function  CellPosToText(ARow, ACol: Cardinal): string;
    { Data manipulation methods - For Cells }
    procedure CopyCell(AFromRow, AFromCol, AToRow, AToCol: Cardinal; AFromWorksheet: TsWorksheet);
    function  FindCell(ARow, ACol: Cardinal): PCell;
    function  GetCell(ARow, ACol: Cardinal): PCell;
    function  GetCellCount: Cardinal;
    function  GetFirstCell(): PCell;
    function  GetNextCell(): PCell;
    function  GetLastColNumber: Cardinal;
    function  GetLastRowNumber: Cardinal;
    function  ReadAsUTF8Text(ARow, ACol: Cardinal): ansistring;
    function  ReadAsNumber(ARow, ACol: Cardinal): Double;
    function  ReadAsDateTime(ARow, ACol: Cardinal; out AResult: TDateTime): Boolean;
    function  ReadUsedFormatting(ARow, ACol: Cardinal): TsUsedFormattingFields;
    function  ReadBackgroundColor(ARow, ACol: Cardinal): TsColor;
    procedure RemoveAllCells;
    procedure WriteUTF8Text(ARow, ACol: Cardinal; AText: ansistring);
    procedure WriteNumber(ARow, ACol: Cardinal; ANumber: double;
      AFormat: TsNumberFormat = nfGeneral; ADecimals: Word = 2);
    procedure WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
      AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = '');
    procedure WriteFormula(ARow, ACol: Cardinal; AFormula: TsFormula);
    procedure WriteNumberFormat(ARow, ACol: Cardinal; ANumberFormat: TsNumberFormat);
    procedure WriteRPNFormula(ARow, ACol: Cardinal; AFormula: TsRPNFormula);
    procedure WriteTextRotation(ARow, ACol: Cardinal; ARotation: TsTextRotation);
    procedure WriteUsedFormatting(ARow, ACol: Cardinal; AUsedFormatting: TsUsedFormattingFields);
    procedure WriteBackgroundColor(ARow, ACol: Cardinal; AColor: TsColor);
    procedure WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment);
    procedure WriteVertAlignment(ARow, ACol: Cardinal; AValue: TsVertAlignment);
    { Data manipulation methods - For Rows and Cols }
    function  FindRow(ARow: Cardinal): PRow;
    function  FindCol(ACol: Cardinal): PCol;
    function  GetRow(ARow: Cardinal): PRow;
    function  GetCol(ACol: Cardinal): PCol;
    procedure RemoveAllRows;
    procedure RemoveAllCols;
    procedure WriteRowInfo(ARow: Cardinal; AData: TRow);
    procedure WriteColInfo(ACol: Cardinal; AData: TCol);
    { Properties }
    property  Cells: TAVLTree read FCells;
    property  Cols: TIndexedAVLTree read FCols;
    property  Rows: TIndexedAVLTree read FRows;
  end;

  { TsWorkbook }

  TsWorkbook = class
  private
    { Internal data }
    FWorksheets: TFPList;
    FEncoding: TsEncoding;
    { Internal methods }
    procedure RemoveCallback(data, arg: pointer);
  public
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
    {@@ This property is only used for formats which don't support unicode
      and support a single encoding for the whole document, like Excel 2 to 5 }
    property Encoding: TsEncoding read FEncoding write FEncoding;
  end;

  {@@ TsSpreadReader class reference type }

  TsSpreadReaderClass = class of TsCustomSpreadReader;
  
  { TsCustomSpreadReader }

  TsCustomSpreadReader = class
  protected
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
  public
    constructor Create; virtual; // To allow descendents to override it
    { General writing methods }
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); virtual;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); virtual;
    procedure ReadFromStrings(AStrings: TStrings; AData: TsWorkbook); virtual;
    { Record reading methods }
    procedure ReadFormula(AStream: TStream); virtual; abstract;
    procedure ReadLabel(AStream: TStream); virtual; abstract;
    procedure ReadNumber(AStream: TStream); virtual; abstract;
  end;

  {@@ TsSpreadWriter class reference type }

  TsSpreadWriterClass = class of TsCustomSpreadWriter;

  TCellsCallback = procedure (ACell: PCell; AStream: TStream) of object;

  { TsCustomSpreadWriter }

  TsCustomSpreadWriter = class
  public
    {@@
    An array with cells which are models for the used styles
    In this array the Row property holds the Index to the corresponding XF field
    }
    FFormattingStyles: array of TCell;
    NextXFIndex: Integer; // Indicates which should be the next XF (Style) Index when filling the styles list
    constructor Create; virtual; // To allow descendents to override it
    { Helper routines }
    function FindFormattingInList(AFormat: PCell): Integer;
    procedure AddDefaultFormats(); virtual;
    procedure ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
    procedure ListAllFormattingStyles(AData: TsWorkbook);
    function  ExpandFormula(AFormula: TsFormula): TsExpandedFormula;
    function  FPSColorToHexString(AColor: TsColor; ARGBColor: TFPColor): string;
    { General writing methods }
    procedure WriteCellCallback(ACell: PCell; AStream: TStream);
    procedure WriteCellsToStream(AStream: TStream; ACells: TAVLTree);
    procedure IterateThroughCells(AStream: TStream; ACells: TAVLTree; ACallback: TCellsCallback);
    procedure WriteToFile(const AFileName: string; AData: TsWorkbook;
      const AOverwriteExisting: Boolean = False); virtual;
    procedure WriteToStream(AStream: TStream; AData: TsWorkbook); virtual;
    procedure WriteToStrings(AStrings: TStrings; AData: TsWorkbook); virtual;
    { Record writing methods }
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell); virtual; abstract;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal; const AFormula: TsFormula; ACell: PCell); virtual;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell); virtual;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal; const AValue: string; ACell: PCell); virtual; abstract;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double; ACell: PCell); virtual; abstract;
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
  function RPNMissingArg(ANext: PRPNItem): PRPNItem;
  function RPNNumber(AValue: Double; ANext: PRPNItem): PRPNItem;
  function RPNString(AValue: String; ANext: PRPNItem): PRPNItem;
  function RPNFunc(AToken: TFEKind; ANext: PRPNItem): PRPNItem; overload;
  function RPNFunc(AToken: TFEKind; ANumParams: Byte; ANext: PRPNItem): PRPNItem; overload;

var
  GsSpreadFormats: array of TsSpreadFormatData;

procedure RegisterSpreadFormat(
  AReaderClass: TsSpreadReaderClass;
  AWriterClass: TsSpreadWriterClass;
  AFormat: TsSpreadsheetFormat);

function SciFloat(AValue: Double; ADecimals: Word): String;
function TimeIntervalToString(AValue: TDateTime): String;


implementation

uses
  Math, StrUtils;

{ Translatable strings }
resourcestring
  lpUnsupportedReadFormat = 'Tried to read a spreadsheet using an unsupported format';
  lpUnsupportedWriteFormat = 'Tried to write a spreadsheet using an unsupported format';
  lpNoValidSpreadsheetFile = '"%s" is not a valid spreadsheet file.';

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
  Formats the number AValue in "scientific" format with the given number of
  decimals. "Scientific" is the same as "exponential", but with exponents rounded
  to multiples of 3.
}
function SciFloat(AValue: Double; ADecimals: Word): String;
var
  m: Double;
  ex: Integer;
begin
  if AValue = 0 then
    Result := '0.0'
  else begin
    ex := floor(log10(abs(AValue)));  // exponent
    // round exponent to multiples of 3
    ex := (ex div 3) * 3;
    if ex < 0 then dec(ex, 3);
    m := AValue * Power(10, -ex);     // mantisse
    Result := Format('%.*fE%d', [ADecimals, m, ex]);
  end;
end;

{@@
  Formats the number AValue as a time string with hours, minutes and seconds.
  Unlike TimeToStr there can be more than 24 hours.
}
function TimeIntervalToString(AValue: TDateTime): String;
var
  hrs: Integer;
  diff: Double;
  h,m,s,z: Word;
  ts: String;
begin
  ts := DefaultFormatSettings.TimeSeparator;
  DecodeTime(frac(abs(AValue)), h, m, s, z);
  hrs := h + trunc(abs(AValue))*24;
  if z > 499 then inc(s);
  if hrs > 0 then
    Result := Format('%d%s%.2d%s%.2d', [hrs, ts, m, ts, s])
  else
    Result := Format('%d%s%.2d', [m, ts, s]);
  if AValue < 0.0 then Result := '-' + Result;
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

procedure TsWorksheet.CopyCell(AFromRow, AFromCol, AToRow, AToCol: Cardinal;
  AFromWorksheet: TsWorksheet);
var
  lCurStr: String;
  lCurUsedFormatting: TsUsedFormattingFields;
  lCurColor: TsColor;
begin
  lCurStr := AFromWorksheet.ReadAsUTF8Text(AFromRow, AFromCol);
  lCurUsedFormatting := AFromWorksheet.ReadUsedFormatting(AFromRow, AFromCol);
  lCurColor := AFromWorksheet.ReadBackgroundColor(AFromRow, AFromCol);
  WriteUTF8Text(AToRow, AToCol, lCurStr);
  WriteUsedFormatting(AToRow, AToCol, lCurUsedFormatting);
  if uffBackgroundColor in lCurUsedFormatting then
  begin
    WriteBackgroundColor(AToRow, AToCol, lCurColor);
  end;
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

  function FloatToStrNoNaN(const Value: Double;
    ANumberFormat: TsNumberFormat; ANumberFormatStr: ansistring): ansistring;
  begin
    if IsNan(Value) then
      Result := ''
    else
    if ANumberFormat = nfSci then
      Result := SciFloat(Value, 1)
    else
    if (ANumberFormat = nfGeneral) or (ANumberFormatStr = '') then
      Result := FloatToStr(Value)
    else
    if (ANumberFormat = nfPercentage) then
      Result := FormatFloat(ANumberFormatStr, Value*100) + '%'
    else
      Result := FormatFloat(ANumberFormatStr, Value);
  end;

  function DateTimeToStrNoNaN(const Value: Double;
    ANumberFormat: TsNumberFormat; ANumberFormatStr: String): ansistring;
  begin
    Result := '';
    if not IsNaN(Value) then begin
      if ANumberFormat = nfTimeInterval then
        Result := TimeIntervalToString(Value)
      else
      if ANumberFormatStr = '' then
        Result := FormatDateTime('c', Value)
      else
        Result := FormatDateTime(ANumberFormatStr, Value);
    end;
  end;

var
  ACell: PCell;
begin
  ACell := FindCell(ARow, ACol);

  if ACell = nil then
  begin
    Result := '';
    Exit;
  end;

  case ACell^.ContentType of
  //cctFormula
  cctNumber:     Result := FloatToStrNoNaN(ACell^.NumberValue, ACell^.NumberFormat, ACell^.NumberFormatStr);
  cctUTF8String: Result := ACell^.UTF8StringValue;
  cctDateTime:   Result := DateTimeToStrNoNaN(ACell^.DateTimeValue, ACell^.NumberFormat, ACell^.NumberFormatStr);
  else
    Result := '';
  end;
end;

function TsWorksheet.ReadAsNumber(ARow, ACol: Cardinal): Double;
var
  ACell: PCell;
  Str: string;
begin
  ACell := FindCell(ARow, ACol);

  if ACell = nil then
  begin
    Result := 0.0;
    Exit;
  end;

  case ACell^.ContentType of

  //cctFormula
  cctDateTime : Result := ACell^.DateTimeValue; //this is in FPC TDateTime format, not Excel
  cctNumber   : Result := ACell^.NumberValue;
  cctUTF8String:
  begin
    // The try is necessary to catch errors while converting the string
    // to a number, an operation which may fail
    try
      Str := ACell^.UTF8StringValue;
      Result := StrToFloat(Str);
    except
      Result := 0.0;
    end;
  end;

  else
    Result := 0.0;
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
end;

{@@
  Writes a floating-point number to a determined cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  ANumber   The number to be written
  @param  AFormat   The format identifier, e.g. nfFixed (optional)
  @param  ADecimals The number of decimals used for formatting (optional)
}
procedure TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: double;
  AFormat: TsNumberFormat = nfGeneral; ADecimals: Word = 2);
var
  ACell: PCell;
  decs: String;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.ContentType := cctNumber;
  ACell^.NumberValue := ANumber;
  ACell^.NumberDecimals := ADecimals;
  if AFormat <> nfGeneral then begin
    Include(ACell^.UsedFormattingFields, uffNumberFormat);
    ACell^.NumberFormat := AFormat;
    decs := DupeString('0', ADecimals);
    if ADecimals > 0 then decs := '.' + decs;
    case AFormat of
      nfFixed:
        ACell^.NumberFormatStr := '0' + decs;
      nfFixedTh:
        ACell^.NumberFormatStr := '#,##0' + decs;
      nfExp:
        ACell^.NumberFormatStr := '0' + decs + 'E+00';
      nfSci:
        ACell^.NumberFormatStr := '';
      nfPercentage:
        ACell^.NumberFormatStr := '0' + decs;
    end;
  end;
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

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) Number, and the cell is formatted
  as a date (either built-in or a custom format).

  Note: custom formats are currently not supported by the writer.
}
procedure TsWorksheet.WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
  AFormat: TsNumberFormat = nfShortDateTime; AFormatStr: String = '');
var
  ACell: PCell;
  fmt: String;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.ContentType := cctDateTime;
  ACell^.DateTimeValue := AValue;
  // Date/time is actually a number field in Excel.
  // To make sure it gets saved correctly, set a date format (instead of General).
  // The user can choose another date format if he wants to
  Include(ACell^.UsedFormattingFields, uffNumberFormat);
  ACell^.NumberFormat := AFormat;
  case AFormat of
    nfShortDateTime:
      ACell^.NumberFormatStr := FormatSettings.ShortDateFormat + ' ' + FormatSettings.ShortTimeFormat;
    nfShortDate:
      ACell^.NumberFormatStr := FormatSettings.ShortDateFormat;
    nfShortTime:
      ACell^.NumberFormatStr := 't';
    nfLongTime:
      ACell^.NumberFormatStr := 'tt';
    nfShortTimeAM:
      ACell^.NumberFormatStr := 't am/pm';
    nfLongTimeAM:
      ACell^.NumberFormatStr := 'tt am/pm';
    nfFmtDateTime:
      begin
        fmt := lowercase(AFormatStr);
        if fmt = 'dm' then ACell^.NumberFormatStr := 'd/mmm'
        else if fmt = 'my' then ACell^.NumberFormatSTr := 'mmm/yy'
        else if fmt = 'ms' then ACell^.NumberFormatStr := 'nn:ss'
        else if fmt = 'msz' then ACell^.NumberFormatStr := 'nn:ss.z'
        else ACell^.NumberFormatStr := AFormatStr;
      end;
    nfTimeInterval:
      ACell^.NumberFormatStr := '';
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
end;

{@@
  Adds number format to the formatting of a cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  TsNumberFormat What format to apply

  @see    TsNumberFormat
}
procedure TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumberFormat: TsNumberFormat);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);

  Include(ACell^.UsedFormattingFields, uffNumberFormat);
  ACell^.NumberFormat := ANumberFormat;
end;

procedure TsWorksheet.WriteRPNFormula(ARow, ACol: Cardinal;
  AFormula: TsRPNFormula);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.ContentType := cctRPNFormula;
  ACell^.RPNFormulaValue := AFormula;
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
end;

procedure TsWorksheet.WriteUsedFormatting(ARow, ACol: Cardinal;
  AUsedFormatting: TsUsedFormattingFields);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.UsedFormattingFields := AUsedFormatting;
end;

procedure TsWorksheet.WriteBackgroundColor(ARow, ACol: Cardinal;
  AColor: TsColor);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.UsedFormattingFields := ACell^.UsedFormattingFields + [uffBackgroundColor];
  ACell^.BackgroundColor := AColor;
end;

procedure TsWorksheet.WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.UsedFormattingFields := lCell^.UsedFormattingFields + [uffHorAlign];
  lCell^.HorAlignment := AValue;
end;

procedure TsWorksheet.WriteVertAlignment(ARow, ACol: Cardinal; AValue: TsVertAlignment);
var
  lCell: PCell;
begin
  lCell := GetCell(ARow, ACol);
  lCell^.UsedFormattingFields := lCell^.UsedFormattingFields + [uffVertAlign];
  lCell^.VertAlignment := AValue;
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

  if (Result = nil) then
  begin
    Result := GetMem(SizeOf(TRow));
    FillChar(Result^, SizeOf(TRow), #0);

    Result^.Row := ARow;

    FRows.Add(Result);
  end;
end;

function TsWorksheet.GetCol(ACol: Cardinal): PCol;
begin
  Result := FindCol(ACol);

  if (Result = nil) then
  begin
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
  for i := FRows.Count-1 downto 0 do
  begin
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
  for i := FCols.Count-1 downto 0 do
  begin
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

procedure TsWorksheet.WriteColInfo(ACol: Cardinal; AData: TCol);
var
  AElement: PCol;
begin
  AElement := GetCol(ACol);
  AElement^.Width := AData.Width;
end;

{ TsWorkbook }

{@@
  Helper method for clearing the spreadsheet list.
}
procedure TsWorkbook.RemoveCallback(data, arg: pointer);
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
end;

{@@
  Destructor.
}
destructor TsWorkbook.Destroy;
begin
  RemoveAllWorksheets;

  FWorksheets.Free;

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
      Result := GsSpreadFormats[i].ReaderClass.Create;

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
      Result := GsSpreadFormats[i].WriterClass.Create;
    
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
    AWriter.WriteToFile(AFileName, Self, AOverwriteExisting);
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
    AWriter.WriteToStream(AStream, Self);
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
  if (integer(AIndex) < FWorksheets.Count) and (integer(AIndex)>=0) then Result := TsWorksheet(FWorksheets.Items[AIndex])
  else Result := nil;
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
  FWorksheets.ForEachCall(RemoveCallback, nil);
end;

{ TsCustomSpreadReader }

constructor TsCustomSpreadReader.Create;
begin
  inherited Create;
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

constructor TsCustomSpreadWriter.Create;
begin
  inherited Create;
end;

{@@
  Checks if the style of a cell is in the list FFormattingStyles and returns the index
  or -1 if it isn't
}
function TsCustomSpreadWriter.FindFormattingInList(AFormat: PCell): Integer;
var
  i: Integer;
begin
  Result := -1;

  for i := 0 to Length(FFormattingStyles) - 1 do
  begin
    if (FFormattingStyles[i].UsedFormattingFields <> AFormat^.UsedFormattingFields) then Continue;

    if uffHorAlign in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].HorAlignment <> AFormat^.HorAlignment) then Continue;

    if uffVertAlign in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].VertAlignment <> AFormat^.VertAlignment) then Continue;

    if uffTextRotation in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].TextRotation <> AFormat^.TextRotation) then Continue;

    if uffBorder in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].Border <> AFormat^.Border) then Continue;

    if uffBackgroundColor in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].BackgroundColor <> AFormat^.BackgroundColor) then Continue;

    if uffNumberFormat in AFormat^.UsedFormattingFields then begin
      if (FFormattingStyles[i].NumberFormat <> AFormat^.NumberFormat) then Continue;
      case AFormat^.NumberFormat of
        nfFixed, nfFixedTh, nfPercentage, nfExp:
          if (FFormattingStyles[i].NumberDecimals <> AFormat^.NumberDecimals) then Continue;
        nfShortDate, nfShortDateTime, nfShortTime, nfLongTime, nfShortTimeAM,
        nfLongTimeAM, nfFmtDateTime, nfTimeInterval:
          if (FFormattingstyles[i].NumberFormatStr <> AFormat^.NumberFormatStr) then Continue;
      end;
    end;

    // If we arrived here it means that the styles match
    Exit(i);
  end;
end;

{ Each descendent should define it's own default formats, if any.
  Always add the normal, unformatted style first to speed up. }
procedure TsCustomSpreadWriter.AddDefaultFormats();
begin
  SetLength(FFormattingStyles, 0);
  NextXFIndex := 0;
end;

procedure TsCustomSpreadWriter.ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
var
  Len: Integer;
begin
  if ACell^.UsedFormattingFields = [] then Exit;

  if FindFormattingInList(ACell) <> -1 then Exit;

  Len := Length(FFormattingStyles);
  SetLength(FFormattingStyles, Len+1);
  FFormattingStyles[Len] := ACell^;
  FFormattingStyles[Len].Row := NextXFIndex;
  Inc(NextXFIndex);
end;

procedure TsCustomSpreadWriter.ListAllFormattingStyles(AData: TsWorkbook);
var
  i: Integer;
begin
  SetLength(FFormattingStyles, 0);

  AddDefaultFormats();

  for i := 0 to AData.GetWorksheetCount - 1 do
  begin
    IterateThroughCells(nil, AData.GetWorksheetByIndex(i).Cells, ListAllFormattingStylesCallback);
  end;
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

  // The formula needs to start with a =
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

function TsCustomSpreadWriter.FPSColorToHexString(AColor: TsColor; ARGBColor: TFPColor): string;
{ We use RGB bytes here, but please note that these are physically written
  to XLS file as ABGR (where A is 0) }
begin
  case AColor of
  scBlack:    Result := '000000';
  scWhite:    Result := 'FFFFFF';
  scRed:      Result := 'FF0000';
  scGREEN:    Result := '00FF00';
  scBLUE:     Result := '0000FF';
  scYELLOW:   Result := 'FFFF00';
  scMAGENTA:  Result := 'FF00FF';
  scCYAN:     Result := '00FFFF';
  scDarkRed:  Result := '800000';
  scDarkGreen:Result := '008000';
  scDarkBlue: Result := '000080';
  scOLIVE:    Result := '808000';
  scPURPLE:   Result := '800080';
  scTEAL:     Result := '008080';
  scSilver:   Result := 'C0C0C0';
  scGrey:     Result := '808080';
  //
  scGrey10pct:Result := 'E6E6E6';
  scGrey20pct:Result := 'CCCCCC';
  scOrange:   Result := 'FFA500';
  scDarkBrown:Result := 'A0522D';
  scBrown:    Result := 'CD853F';
  scBeige:    Result := 'F5F5DC';
  scWheat:    Result := 'F5DEB3';
  //
  scRGBCOLOR: Result := Format('%x%x%x', [ARGBColor.Red div $100, ARGBColor.Green div $100, ARGBColor.Blue div $100]);
  end;
end;

{@@
  Helper function for the spreadsheet writers.

  @see    TsCustomSpreadWriter.WriteCellsToStream
}
procedure TsCustomSpreadWriter.WriteCellCallback(ACell: PCell; AStream: TStream);
begin
  case ACell.ContentType of
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

  @param  AFileName The output file name.
                    If the file already exists it will be replaced.
  @param  AData     The Workbook to be saved.

  @see    TsWorkbook
}
procedure TsCustomSpreadWriter.WriteToFile(const AFileName: string;
  AData: TsWorkbook; const AOverwriteExisting: Boolean = False);
var
  OutputFile: TFileStream;
  lMode: Word;
begin
  if AOverwriteExisting then lMode := fmCreate or fmOpenWrite
  else lMode := fmCreate;

  OutputFile := TFileStream.Create(AFileName, lMode);
  try
    WriteToStream(OutputFile, AData);
  finally
    OutputFile.Free;
  end;
end;

{@@
  This routine should be overriden in descendent classes.
}
procedure TsCustomSpreadWriter.WriteToStream(AStream: TStream; AData: TsWorkbook);
var
  lStringList: TStringList;
begin
  lStringList := TStringList.Create;
  try
    WriteToStrings(lStringList, AData);
    lStringList.SaveToStream(AStream);
  finally
    lStringList.Free;
  end;
end;

procedure TsCustomSpreadWriter.WriteToStrings(AStrings: TStrings;
  AData: TsWorkbook);
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
  if ord(AToken) < ord(fekAdd) then
    raise Exception.Create('No basic tokens allowed here.');
  Result := NewRPNItem;
  Result^.FE.ElementKind := AToken;
  Result^.Next := ANext;
end;

{@@
  Creates an entry in the RPN array for an Excel function or operation
  specified by its TokenID (--> TFEKind). Specify the number of parameters used.
  They must have been created before.
}
function RPNFunc(AToken: TFEKind; ANumParams: Byte; ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(AToken, ANext);
  Result^.FE.ParamsNum := ANumParams;
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


finalization

  SetLength(GsSpreadFormats, 0);

end.

