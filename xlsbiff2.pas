{
xlsbiff2.pas

Writes an Excel 2.x file

Excel 2.x files support only one Worksheet per Workbook, so only the first
will be written.

An Excel file consists of a number of subsequent records.
To ensure a properly formed file, the following order must be respected:

1st record:        BOF
2nd to Nth record: Any record
Last record:       EOF

The row and column numbering in BIFF files is zero-based.

Excel file format specification obtained from:

http://sc.openoffice.org/excelfileformat.pdf

Encoding information: ISO_8859_1 is used, to have support to
other characters, please use a format which support unicode

AUTHORS: Felipe Monteiro de Carvalho
}
unit xlsbiff2;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

uses
  Classes, SysUtils, lconvencoding,
  fpsTypes, fpsNumFormat, fpspreadsheet, fpsUtils, xlscommon;

const
  BIFF2_MAX_PALETTE_SIZE = 8;
  // There are more colors but they do not seem to be controlled by a palette.

type

  { TsBIFF2NumFormatList }
  TsBIFF2NumFormatList = class(TsCustomNumFormatList)
  protected
    procedure AddBuiltinFormats; override;
  public
    constructor Create(AWorkbook: TsWorkbook);
    procedure ConvertBeforeWriting(var AFormatString: String;
      var ANumFormat: TsNumberFormat); override;
    function Find(ANumFormat: TsNumberFormat; ANumFormatStr: String): Integer; override;
  end;

  { TsSpreadBIFF2Reader }

  TsSpreadBIFF2Reader = class(TsSpreadBIFFReader)
  private
//    WorkBookEncoding: TsEncoding;
    FFont: TsFont;
    FPendingXFIndex: Word;
  protected
    procedure CreateNumFormatList; override;
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadBool(AStream: TStream); override;
    procedure ReadColWidth(AStream: TStream);
    procedure ReadDefRowHeight(AStream: TStream);
    procedure ReadFont(AStream: TStream);
    procedure ReadFontColor(AStream: TStream);
    procedure ReadFormat(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadInteger(AStream: TStream);
    procedure ReadIXFE(AStream: TStream);
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
    procedure ReadRowColXF(AStream: TStream; out ARow, ACol: Cardinal; out AXF: Word); override;
    procedure ReadRowInfo(AStream: TStream); override;
    function ReadRPNFunc(AStream: TStream): Word; override;
    procedure ReadRPNSharedFormulaBase(AStream: TStream; out ARow, ACol: Cardinal); override;
    function ReadRPNTokenArraySize(AStream: TStream): Word; override;
    procedure ReadStringRecord(AStream: TStream); override;
    procedure ReadWindow2(AStream: TStream); override;
    procedure ReadXF(AStream: TStream);
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    { General reading methods }
    procedure ReadFromStream(AStream: TStream); override;
  end;

  { TsSpreadBIFF2Writer }

  TsSpreadBIFF2Writer = class(TsSpreadBIFFWriter)
  private
    procedure GetCellAttributes(ACell: PCell; XFIndex: Word;
      out Attrib1, Attrib2, Attrib3: Byte);
    { Record writing methods }
    procedure WriteBOF(AStream: TStream);
    procedure WriteCellFormatting(AStream: TStream; ACell: PCell; XFIndex: Word);
    procedure WriteColWidth(AStream: TStream; ACol: PCol);
    procedure WriteColWidths(AStream: TStream);
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFontIndex: Integer);
    procedure WriteFonts(AStream: TStream);
    procedure WriteFormatCount(AStream: TStream);
    procedure WriteIXFE(AStream: TStream; XFIndex: Word);
  protected
    procedure CreateNumFormatList; override;
    procedure ListAllNumFormats; override;
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
//    procedure WriteCodePage(AStream: TStream; AEncoding: TsEncoding); override;
    procedure WriteCodePage(AStream: TStream; ACodePage: String); override;
    procedure WriteError(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TsErrorValue; ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;
    procedure WriteNumFormat(AStream: TStream; ANumFormatData: TsNumFormatData;
      AListIndex: Integer); override;
    procedure WriteRow(AStream: TStream; ASheet: TsWorksheet;
      ARowIndex, AFirstColIndex, ALastColIndex: Cardinal; ARow: PRow); override;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal;
      const AFormula: TsRPNFormula; ACell: PCell); override;
    function WriteRPNFunc(AStream: TStream; AIdentifier: Word): Word; override;
    procedure WriteRPNSharedFormulaLink(AStream: TStream; ACell: PCell;
      var RPNLength: Word); override;
    procedure WriteRPNTokenArraySize(AStream: TStream; ASize: Word); override;
    procedure WriteSharedFormula(AStream: TStream; ACell: PCell); override;
    procedure WriteStringRecord(AStream: TStream; AString: String); override;
    procedure WriteWindow1(AStream: TStream); override;
    procedure WriteWindow2(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteXF(AStream: TStream; AFormatRecord: PsCellFormat;
      XFType_Prot: Byte = 0); override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    { General writing methods }
    procedure WriteToStream(AStream: TStream); override;
  end;

var
  { the palette of the default BIFF2 colors as "big-endian color" values }
  PALETTE_BIFF2: array[$0..$07] of TsColorValue = (
    $000000,  // $00: black
    $FFFFFF,  // $01: white
    $FF0000,  // $02: red
    $00FF00,  // $03: green
    $0000FF,  // $04: blue
    $FFFF00,  // $05: yellow
    $FF00FF,  // $06: magenta
    $00FFFF   // $07: cyan
  );


implementation

uses
  Math, fpsStrings, fpsReaderWriter, fpsNumFormatParser;

const
  { Excel record IDs }
  INT_EXCEL_ID_DIMENSIONS    = $0000;
  INT_EXCEL_ID_BLANK         = $0001;
  INT_EXCEL_ID_INTEGER       = $0002;
  INT_EXCEL_ID_NUMBER        = $0003;
  INT_EXCEL_ID_LABEL         = $0004;
  INT_EXCEL_ID_BOOLERROR     = $0005;
  INT_EXCEL_ID_ROW           = $0008;
  INT_EXCEL_ID_BOF           = $0009;
  {%H-}INT_EXCEL_ID_INDEX    = $000B;
  INT_EXCEL_ID_FORMAT        = $001E;
  INT_EXCEL_ID_FORMATCOUNT   = $001F;
  INT_EXCEL_ID_COLWIDTH      = $0024;
  INT_EXCEL_ID_DEFROWHEIGHT  = 00025;
  INT_EXCEL_ID_WINDOW2       = $003E;
  INT_EXCEL_ID_XF            = $0043;
  INT_EXCEL_ID_IXFE          = $0044;
  INT_EXCEL_ID_FONTCOLOR     = $0045;

  { BOF record constants }
  INT_EXCEL_SHEET            = $0010;
  {%H-}INT_EXCEL_CHART       = $0020;
  {%H-}INT_EXCEL_MACRO_SHEET = $0040;

type
  TBIFF2_BoolErrRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    BoolErrValue: Byte;
    ValueType: Byte;
  end;

  TBIFF2_DimensionsRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FirstRow: Word;
    LastRowPlus1: Word;
    FirstCol: Word;
    LastColPlus1: Word;
  end;

  TBIFF2_LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    TextLen: Byte;
  end;

  TBIFF2_NumberRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    Value: Double;
  end;

  TBIFF2_IntegerRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    Value: Word;
  end;

  TBIFF2_XFRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FontIndex: Byte;
    NotUsed: Byte;
    NumFormatIndex_Flags: Byte;
    HorAlign_Border_BkGr: Byte;
  end;


{ TsBIFF2NumFormatList }

constructor TsBIFF2NumFormatList.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
end;

{@@ ----------------------------------------------------------------------------
  Prepares the list of built-in number formats. They are created in the default
  dialect for FPC, they have to be converted to Excel syntax before writing.
  Note that Excel2 expects them to be localized. This is something which has to
  be taken account of in ConvertBeforeWriting.
-------------------------------------------------------------------------------}
procedure TsBIFF2NumFormatList.AddBuiltinFormats;
var
  fs: TFormatSettings;
  cs: string;
begin
  fs := FWorkbook.FormatSettings;
  cs := fs.CurrencyString;
  AddFormat( 0, nfGeneral, '');
  AddFormat( 1, nfFixed, '0');
  AddFormat( 2, nfFixed, '0.00');
  AddFormat( 3, nfFixedTh, '#,##0');
  AddFormat( 4, nfFixedTh, '#,##0.00');
  AddFormat( 5, nfCurrency, Format('"%s"#,##0;("%s"#,##0)', [cs, cs]));
  AddFormat( 6, nfCurrencyRed, Format('"%s"#,##0;[Red]("%s"#,##0)', [cs, cs]));
  AddFormat( 7, nfCurrency, Format('"%s"#,##0.00;("%s"#,##0.00)', [cs, cs]));
  AddFormat( 8, nfCurrencyRed, Format('"%s"#,##0.00;[Red]("%s"#,##0.00)', [cs, cs]));
  AddFormat( 9, nfPercentage, '0%');
  AddFormat(10, nfPercentage, '0.00%');
  AddFormat(11, nfExp, '0.00E+00');
  AddFormat(12, nfShortDate, fs.ShortDateFormat);
  AddFormat(13, nfLongDate, fs.LongDateFormat);
  AddFormat(14, nfCustom, 'd/mmm');
  AddFormat(15, nfCustom, 'mmm/yy');
  AddFormat(16, nfShortTimeAM, AddAMPM(fs.ShortTimeFormat, fs));
  AddFormat(17, nfLongTimeAM, AddAMPM(fs.LongTimeFormat, fs));
  AddFormat(18, nfShortTime, fs.ShortTimeFormat);
  AddFormat(19, nfLongTime, fs.LongTimeFormat);
  AddFormat(20, nfShortDateTime, fs.ShortDateFormat + ' ' + fs.ShortTimeFormat);

  FFirstNumFormatIndexInFile := 0; // BIFF2 stores built-in formats to file.
  FNextNumFormatIndex := 21;       // not needed - there are not user-defined formats
end;


procedure TsBIFF2NumFormatList.ConvertBeforeWriting(var AFormatString: String;
  var ANumFormat: TsNumberFormat);
var
  parser: TsNumFormatParser;
begin
  Unused(ANumFormat);

  if AFormatString = '' then
    AFormatString := 'General'
  else begin
    parser := TsNumFormatParser.Create(FWorkbook, AFormatString);
    try
      parser.Localize;
      parser.LimitDecimals;
      AFormatString := parser.FormatString[nfdExcel];
    finally
      parser.Free;
    end;
  end;
end;

function TsBIFF2NumFormatList.Find(ANumFormat: TsNumberFormat;
  ANumFormatStr: String): Integer;
var
  parser: TsNumFormatParser;
  decs: Integer;
  dt: string;
begin
  Result := 0;

  parser := TsNumFormatParser.Create(Workbook, ANumFormatStr);
  try
    decs := parser.Decimals;
    dt := parser.GetDateTimeCode(0);
  finally
    parser.Free;
  end;

  case ANumFormat of
    nfGeneral      : exit;
    nfFixed        : Result := IfThen(decs = 0, 1, 2);
    nfFixedTh      : Result := IfThen(decs = 0, 3, 4);
    nfCurrency     : Result := IfThen(decs = 0, 5, 7);
    nfCurrencyRed  : Result := IfThen(decs = 0, 6, 8);
    nfPercentage   : Result := IfThen(decs = 0, 9, 10);
    nfExp          : Result := 11;
    nfShortDate    : Result := 12;
    nfLongDate     : Result := 13;
    nfShortTimeAM  : Result := 16;
    nfLongTimeAM   : Result := 17;
    nfShortTime    : Result := 18;
    nfLongTime     : Result := 19;
    nfShortDateTime: Result := 20;
    nfCustom       : if dt = 'dm' then Result := 14 else
                     if dt = 'my' then Result := 15;
  end;
end;


{ TsSpreadBIFF2Reader }

constructor TsSpreadBIFF2Reader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FLimitations.MaxPaletteSize := BIFF2_MAX_PALETTE_SIZE;
end;

{@@ ----------------------------------------------------------------------------
  Creates the correct version of the number format list.
  It is for BIFF2 and BIFF3 file formats.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsBIFF2NumFormatList.Create(Workbook);
end;

procedure TsSpreadBIFF2Reader.ReadBlank(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  cell: PCell;
begin
  ReadRowColXF(AStream, ARow, ACol, XF);
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);
  ApplyCellFormatting(cell, XF);
  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  The name of this method is misleading - it reads a BOOLEAN cell value,
  but also an ERROR value; BIFF stores them in the same record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadBool(AStream: TStream);
var
  rec: TBIFF2_BoolErrRecord;
  r, c: Cardinal;
  xf: Word;
  cell: PCell;
begin
  { Read entire record, starting at Row }
  rec.Row := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_BoolErrRecord) - 2*SizeOf(Word));
  r := WordLEToN(rec.Row);
  c := WordLEToN(rec.Col);
  xf := rec.Attrib1 and $3F;
  if xf = 63 then xf := FPendingXFIndex;

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(r, c, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(r, c);

  { Retrieve boolean or error value depending on the "ValueType" }
  case rec.ValueType of
    0: FWorksheet.WriteBoolValue(cell, boolean(rec.BoolErrValue));
    1: FWorksheet.WriteErrorValue(cell, ConvertFromExcelError(rec.BoolErrValue));
  end;

  { Apply formatting }
  ApplyCellFormatting(cell, xf);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, r, c, cell);
end;

procedure TsSpreadBIFF2Reader.ReadColWidth(AStream: TStream);
var
  c, c1, c2: Cardinal;
  w: Word;
  col: TCol;
begin
  // read column start and end index of column range
  c1 := AStream.ReadByte;
  c2 := AStream.ReadByte;
  // read col width in 1/256 of the width of "0" character
  w := WordLEToN(AStream.ReadWord);
  // calculate width in units of "characters"
  col.Width := w / 256;
  // assign width to columns, but only if different from default column width.
  if not SameValue(col.Width, FWorksheet.DefaultColWidth) then
    for c := c1 to c2 do
      FWorksheet.WriteColInfo(c, col);
end;

procedure TsSpreadBIFF2Reader.ReadDefRowHeight(AStream: TStream);
var
  hw: word;
  h : Single;
begin
  hw := WordLEToN(AStream.ReadWord);
  h := TwipsToPts(hw and $8000) / FWorkbook.GetDefaultFontSize;
  if h > ROW_HEIGHT_CORRECTION then
    FWorksheet.DefaultRowHeight := h - ROW_HEIGHT_CORRECTION;
end;

procedure TsSpreadBIFF2Reader.ReadFont(AStream: TStream);
var
  lHeight: Word;
  lOptions: Word;
  Len: Byte;
  lFontName: UTF8String;
begin
  FFont := TsFont.Create;

  { Height of the font in twips = 1/20 of a point }
  lHeight := WordLEToN(AStream.ReadWord);
  FFont.Size := lHeight/20;

  { Option flags }
  lOptions := WordLEToN(AStream.ReadWord);
  FFont.Style := [];
  if lOptions and $0001 <> 0 then Include(FFont.Style, fssBold);
  if lOptions and $0002 <> 0 then Include(FFont.Style, fssItalic);
  if lOptions and $0004 <> 0 then Include(FFont.Style, fssUnderline);
  if lOptions and $0008 <> 0 then Include(FFont.Style, fssStrikeout);

  { Font name: Unicodestring, char count in 1 byte }
  Len := AStream.ReadByte();
  SetLength(lFontName, Len);
  AStream.ReadBuffer(lFontName[1], Len);
  FFont.FontName := lFontName;

  { Add font to workbook's font list }
  FWorkbook.AddFont(FFont);
end;

procedure TsSpreadBIFF2Reader.ReadFontColor(AStream: TStream);
begin
  FFont.Color := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads the FORMAT record required for formatting numerical data
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadFormat(AStream: TStream);
begin
  Unused(AStream);
  // We ignore the formats in the file, everything is known
  // (Using the formats in the file would require de-localizing them).
end;

procedure TsSpreadBIFF2Reader.ReadFromStream(AStream: TStream);
var
  BIFF2EOF: Boolean;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  // Clear existing fonts. They will be replaced by those from the file.
  FWorkbook.RemoveAllFonts;

  { Store some data about the workbook that other routines need }
  //WorkBookEncoding := AData.Encoding;

  BIFF2EOF := False;

  { In BIFF2 files there is only one worksheet, let's create it }
  FWorksheet := FWorkbook.AddWorksheet('Sheet', true);

  { Read all records in a loop }
  while not BIFF2EOF do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);

    CurStreamPos := AStream.Position;

    case RecordType of
      INT_EXCEL_ID_BLANK       : ReadBlank(AStream);
      INT_EXCEL_ID_BOOLERROR   : ReadBool(AStream);
      INT_EXCEL_ID_CODEPAGE    : ReadCodePage(AStream);
      INT_EXCEL_ID_NOTE        : ReadComment(AStream);
      INT_EXCEL_ID_FONT        : ReadFont(AStream);
      INT_EXCEL_ID_FONTCOLOR   : ReadFontColor(AStream);
      INT_EXCEL_ID_FORMAT      : ReadFormat(AStream);
      INT_EXCEL_ID_INTEGER     : ReadInteger(AStream);
      INT_EXCEL_ID_IXFE        : ReadIXFE(AStream);
      INT_EXCEL_ID_NUMBER      : ReadNumber(AStream);
      INT_EXCEL_ID_LABEL       : ReadLabel(AStream);
      INT_EXCEL_ID_FORMULA     : ReadFormula(AStream);
      INT_EXCEL_ID_STRING      : ReadStringRecord(AStream);
      INT_EXCEL_ID_COLWIDTH    : ReadColWidth(AStream);
      INT_EXCEL_ID_DEFCOLWIDTH : ReadDefColWidth(AStream);
      INT_EXCEL_ID_ROW         : ReadRowInfo(AStream);
      INT_EXCEL_ID_DEFROWHEIGHT: ReadDefRowHeight(AStream);
      INT_EXCEL_ID_WINDOW2     : ReadWindow2(AStream);
      INT_EXCEL_ID_PANE        : ReadPane(AStream);
      INT_EXCEL_ID_XF          : ReadXF(AStream);
      INT_EXCEL_ID_BOF         : ;
      INT_EXCEL_ID_EOF         : BIFF2EOF := True;
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    if AStream.Position >= AStream.Size then BIFF2EOF := True;
  end;

  FixCols(FWorksheet);
  FixRows(FWorksheet);
end;

procedure TsSpreadBIFF2Reader.ReadFormula(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  ok: Boolean;
  formulaResult: Double = 0.0;
  Data: array [0..7] of byte;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  err: TsErrorValue;
  cell: PCell;
begin
  { BIFF Record row/column/style }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Result of the formula result in IEEE 754 floating-point value }
  Data[0] := 0;   // to silence the compiler...
  AStream.ReadBuffer(Data, Sizeof(Data));

  { Recalculation byte - currently not used }
  AStream.ReadByte;

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);

  // Now determine the type of the formula result
  if (Data[6] = $FF) and (Data[7] = $FF) then
    case Data[0] of
      0: // String -> Value is found in next record (STRING)
         FIncompleteCell := cell;
      1: // Boolean value
         FWorksheet.WriteBoolValue(cell, Data[2] = 1);
      2: begin  // Error value
           case Data[2] of
             ERR_INTERSECTION_EMPTY   : err := errEmptyIntersection;
             ERR_DIVIDE_BY_ZERO       : err := errDivideByZero;
             ERR_WRONG_TYPE_OF_OPERAND: err := errWrongType;
             ERR_ILLEGAL_REFERENCE    : err := errIllegalRef;
             ERR_WRONG_NAME           : err := errWrongName;
             ERR_OVERFLOW             : err := errOverflow;
             ERR_ARG_ERROR            : err := errArgError;
           end;
           FWorksheet.WriteErrorValue(cell, err);
         end;
      3: // Empty cell
         FWorksheet.WriteBlank(cell);
    end
  else
  begin
    // Result is a number or a date/time
    Move(Data[0], formulaResult, SizeOf(Data));

    {Find out what cell type, set content type and value}
    ExtractNumberFormat(XF, nf, nfs);
    if IsDateTime(formulaResult, nf, nfs, dt) then
      FWorksheet.WriteDateTime(cell, dt, nf, nfs)
    else
      FWorksheet.WriteNumber(cell, formulaResult, nf, nfs);
  end;

  { Formula token array }
  if (boReadFormulas in FWorkbook.Options) then
  begin
    ok := ReadRPNTokenArray(AStream, cell);
    if not ok then FWorksheet.WriteErrorValue(cell, errFormulaNotSupported);
  end;

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadLabel(AStream: TStream);
var
  rec: TBIFF2_LabelRecord;
  L: Byte;
  ARow, ACol: Cardinal;
  XF: Word;
  ansiStr: ansistring;
  valueStr: UTF8String;
  cell: PCell;
begin
  { Read entire record, starting at Row, except for string data }
  rec.Row := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_LabelRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
  if XF = 63 then XF := FPendingXFIndex;

  { String with 8-bit size }
  L := rec.TextLen;
  SetLength(ansiStr, L);
  AStream.ReadBuffer(ansiStr[1], L);

  { Save the data }
  valueStr := ConvertEncoding(ansiStr, FCodePage, encodingUTF8);
  {
  case WorkBookEncoding of
    seLatin2:   AStrValue := CP1250ToUTF8(AValue);
    seCyrillic: AStrValue := CP1251ToUTF8(AValue);
    seGreek:    AStrValue := CP1253ToUTF8(AValue);
    seTurkish:  AStrValue := CP1254ToUTF8(AValue);
    seHebrew:   AStrValue := CP1255ToUTF8(AValue);
    seArabic:   AStrValue := CP1256ToUTF8(AValue);
  else
    // Latin 1 is the default
    AStrValue := CP1252ToUTF8(AValue);
  end;
  }

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);
  FWorksheet.WriteUTF8Text(cell, valueStr);

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadNumber(AStream: TStream);
var
  rec: TBIFF2_NumberRecord;
  ARow, ACol: Cardinal;
  XF: Word;
  value: Double = 0.0;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  cell: PCell;
begin
  { Read entire record, starting at Row }
  rec.Row := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_NumberRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
  if XF = 63 then XF := FPendingXFIndex;
  value := rec.Value;

  {Create cell}
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);

  {Find out what cell type, set content type and value}
  ExtractNumberFormat(XF, nf, nfs);
  if IsDateTime(value, nf, nfs, dt) then
    FWorksheet.WriteDateTime(cell, dt, nf, nfs)
  else
    FWorksheet.WriteNumber(cell, value, nf, nfs);

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadInteger(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  AWord  : Word = 0;
  cell: PCell;
  rec: TBIFF2_IntegerRecord;
begin
  { Read record into buffer }
  rec.Row := 0;   // to silence the comiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_NumberRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
  if XF = 63 then XF := FPendingXFIndex;
  AWord := WordLEToN(rec.Value);

  { Create cell }
  if FIsVirtualMode then
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);

  { Save the data }
  FWorksheet.WriteNumber(cell, AWord);

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  Reads an IXFE record. This record contains the "true" XF index of a cell. It
  is used if there are more than 62 XF records (XF field is only 6-bit). The
  IXFE record is used in front of the cell record using it
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadIXFE(AStream: TStream);
begin
  FPendingXFIndex := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads the row, column and xf index from the stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadRowColXF(AStream: TStream;
  out ARow, ACol: Cardinal; out AXF: WORD);
begin
  { BIFF Record data for row and column}
  ARow := WordLEToN(AStream.ReadWord);
  ACol := WordLEToN(AStream.ReadWord);

  { Index to XF record }
  AXF := AStream.ReadByte and $3F; // to do: if AXF = $3F = 63 then there must be a IXFE record which contains the true XF index!
  if AXF = $3F then
    AXF := FPendingXFIndex;

  { Index to format and font record, cell style - ignored because contained in XF
    Must read to keep the record in sync. }
  AStream.ReadWord;
end;

procedure TsSpreadBIFF2Reader.ReadRowInfo(AStream: TStream);
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
  rowRec.RowIndex := 0;  // to silence the compiler...
  AStream.ReadBuffer(rowrec, SizeOf(TRowRecord));
  h := WordLEToN(rowrec.Height);
  if h and $8000 = 0 then    // if this bit were set, rowheight would be default
  begin
    lRow := FWorksheet.GetRow(WordLEToN(rowrec.RowIndex));
    // Row height is encoded into the 15 remaining bits in units "twips" (1/20 pt)
    // We need it in "lines" units.
    lRow^.Height := TwipsToPts(h and $7FFF) / Workbook.GetFont(0).Size;
    if lRow^.Height > ROW_HEIGHT_CORRECTION then
      lRow^.Height := lRow^.Height - ROW_HEIGHT_CORRECTION
    else
      lRow^.Height := 0;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the identifier for an RPN function with fixed argument count from the
  stream.
  Valid for BIFF2-BIFF3.
-------------------------------------------------------------------------------}
function TsSpreadBIFF2Reader.ReadRPNFunc(AStream: TStream): Word;
var
  b: Byte;
begin
  b := AStream.ReadByte;
  Result := b;
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell coordiantes of the top/left cell of a range using a
  shared formula.
  This cell contains the rpn token sequence of the formula.
  Is overridden because BIFF2 has 1 byte for column.
  Code is not called for shared formulas (which are not supported by BIFF2), but
  maybe for array formulas.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadRPNSharedFormulaBase(AStream: TStream;
  out ARow, ACol: Cardinal);
begin
  // 2 bytes for row of first cell in shared formula
  ARow := WordLEToN(AStream.ReadWord);
  // 1 byte for column of first cell in shared formula
  ACol := AStream.ReadByte;
end;

{@@ ----------------------------------------------------------------------------
  Helper funtion for reading of the size of the token array of an RPN formula.
  Is overridden because BIFF2 uses 1 byte only.
-------------------------------------------------------------------------------}
function TsSpreadBIFF2Reader.ReadRPNTokenArraySize(AStream: TStream): Word;
begin
  Result := AStream.ReadByte;
end;

{@@ ----------------------------------------------------------------------------
  Reads a STRING record which contains the result of string formula.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadStringRecord(AStream: TStream);
var
  len: Byte;
  s: ansistring;
begin
  // The string is a byte-string with 8 bit length
  len := AStream.ReadByte;
  if len > 0 then
  begin
    SetLength(s, Len);
    AStream.ReadBuffer(s[1], len);
    if (FIncompleteCell <> nil) and (s <> '') then
    begin
      // The "IncompleteCell" has been identified in the sheet when reading
      // the FORMULA record which precedes the String record.
//      FIncompleteCell^.UTF8StringValue := AnsiToUTF8(s);
      FIncompleteCell^.UTF8StringValue := ConvertEncoding(s, FCodePage, encodingUTF8);
      FIncompleteCell^.ContentType := cctUTF8String;
      if FIsVirtualMode then
        Workbook.OnReadCellData(Workbook, FIncompleteCell^.Row, FIncompleteCell^.Col, FIncompleteCell);
    end;
  end;
  FIncompleteCell := nil;
end;

{@@ ----------------------------------------------------------------------------
  Reads the WINDOW2 record containing information like "show grid lines",
  "show sheet headers", "panes are frozen", etc.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadWindow2(AStream: TStream);
begin
  // Show formulas, not results
  AStream.ReadByte;

  // Show grid lines
  if AStream.ReadByte <> 0 then
    FWorksheet.Options := FWorksheet.Options + [soShowGridLines]
  else
    FWorksheet.Options := FWorksheet.Options - [soShowGridLines];

  // Show sheet headers
  if AStream.ReadByte <> 0 then
    FWorksheet.Options := FWorksheet.Options + [soShowHeaders]
  else
    FWorksheet.Options := FWorksheet.Options - [soShowHeaders];

  // Panes are frozen
  if AStream.ReadByte <> 0 then
    FWorksheet.Options := FWorksheet.Options + [soHasFrozenPanes]
  else
    FWorksheet.Options := FWorksheet.Options - [soHasFrozenPanes];

  // Show zero values
  AStream.ReadByte;

  // Index to first visible row
  WordLEToN(AStream.ReadWord);

  // Indoex to first visible column
  WordLEToN(AStream.ReadWord);

  // Use automatic grid line color (0= manual)
  AStream.ReadByte;

  // Manual grid line line color (rgb)
  DWordToLE(AStream.ReadDWord);
end;

procedure TsSpreadBIFF2Reader.ReadXF(AStream: TStream);
var
  rec: TBIFF2_XFRecord;
  fmt: TsCellFormat;
  b: Byte;
  nfdata: TsNumFormatData;
  i: Integer;
begin
  // Read entire xf record into buffer
  InitFormatRecord(fmt);
  fmt.ID := FCellFormatList.Count;

  rec.FontIndex := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.FontIndex, SizeOf(rec) - 2*SizeOf(word));

  // Font index
  fmt.FontIndex := rec.FontIndex;
  if fmt.FontIndex = 1 then
    Include(fmt.UsedFormattingFields, uffBold)
  else if fmt.FontIndex > 1 then
    Include(fmt.UsedFormattingFields, uffFont);

  // Number format index
  b := rec.NumFormatIndex_Flags and $3F;
  i := NumFormatList.FindByIndex(b);
  if i > -1 then begin
    nfdata := NumFormatList.Items[i];
    fmt.NumberFormat := nfdata.NumFormat;
    fmt.NumberFormatStr := nfdata.FormatString;
    if nfdata.NumFormat <> nfGeneral then
      Include(fmt.UsedFormattingFields, uffNumberFormat);
  end;

  // Horizontal alignment
  b := rec.HorAlign_Border_BkGr and MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
  begin
    fmt.HorAlignment := TsHorAlignment(b);
    if fmt.HorAlignment <> haDefault then
      Include(fmt.UsedFormattingFields, uffHorAlign);
  end;

  // Vertical alignment - not used in BIFF2
  fmt.VertAlignment := vaDefault;

  // Word wrap - not used in BIFF2
  // -- nothing to do here

  // Text rotation - not used in BIFF2
  // -- nothing to do here

  // Borders
  fmt.Border := [];
  if rec.HorAlign_Border_BkGr and $08 <> 0 then
    Include(fmt.Border, cbWest);
  if rec.HorAlign_Border_BkGr and $10 <> 0 then
    Include(fmt.Border, cbEast);
  if rec.HorAlign_Border_BkGr and $20 <> 0 then
    Include(fmt.Border, cbNorth);
  if rec.HorAlign_Border_BkGr and $40 <> 0 then
    Include(fmt.Border, cbSouth);
  if fmt.Border <> [] then
    Include(fmt.UsedFormattingFields, uffBorder);

  // Background color not supported, only shaded background
  if rec.HorAlign_Border_BkGr and $80 <> 0 then
  begin
    fmt.Background.Style := fsGray50;
    fmt.Background.FgColor := scBlack;
    fmt.Background.BgColor := scTransparent;
    Include(fmt.UsedFormattingFields, uffBackground);
  end;

  // Add the decoded data to the format list
  FCellFormatList.Add(fmt);
end;


{ TsSpreadBIFF2Writer }

constructor TsSpreadBIFF2Writer.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FLimitations.MaxPaletteSize := BIFF2_MAX_PALETTE_SIZE;
end;

{@@ ----------------------------------------------------------------------------
  Creates the correct version of the number format list.
  It is valid for BIFF2 and BIFF3 file formats.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsBIFF2NumFormatList.Create(Workbook);
end;

{@@ ----------------------------------------------------------------------------
  Determines the cell attributes needed for writing a cell content record, such
  as WriteLabel, WriteNumber, etc.
  The cell attributes contain, in bit masks, xf record index, font index,
  borders, etc.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.GetCellAttributes(ACell: PCell; XFIndex: Word;
  out Attrib1, Attrib2, Attrib3: Byte);
var
  fmt: PsCellFormat;
begin
  fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);

  if fmt^.UsedFormattingFields = [] then begin
    Attrib1 := 15;
    Attrib2 := 0;
    Attrib3 := 0;
    exit;
  end;

  // 1st byte:
  //   Mask $3F: Index to XF record
  //   Mask $40: 1 = Cell is locked
  //   Mask $80: 1 = Formula is hidden
  Attrib1 := Min(XFIndex, $3F) and $3F;

  // 2nd byte:
  //   Mask $3F: Index to FORMAT record ("FORMAT" = number format!)
  //   Mask $C0: Index to FONT record
  Attrib2 := fmt^.FontIndex shr 6;

  // 3rd byte
  //   Mask $07: horizontal alignment
  //   Mask $08: Cell has left border
  //   Mask $10: Cell has right border
  //   Mask $20: Cell has top border
  //   Mask $40: Cell has bottom border
  //   Mask $80: Cell has shaded background
  Attrib3 := 0;
  if uffHorAlign in fmt^.UsedFormattingFields then
    Attrib3 := ord (fmt^.HorAlignment);
  if uffBorder in fmt^.UsedFormattingFields then begin
    if cbNorth in fmt^.Border then Attrib3 := Attrib3 or $20;
    if cbWest in fmt^.Border then Attrib3 := Attrib3 or $08;
    if cbEast in fmt^.Border then Attrib3 := Attrib3 or $10;
    if cbSouth in fmt^.Border then Attrib3 := Attrib3 or $40;
  end;
  if (uffBackground in fmt^.UsedFormattingFields) then
    Attrib3 := Attrib3 or $80;
end;

{ Builds up the list of number formats to be written to the biff2 file.
  Unlike biff5+ no formats are added here because biff2 supports only 21
  standard formats; these formats have been added by the NumFormatList's
  AddBuiltInFormats.

  NOT CLEAR IF THIS IS TRUE ????
  }
  // ToDo: check if the BIFF2 format is really restricted to 21 formats.
procedure TsSpreadBIFF2Writer.ListAllNumFormats;
begin
  // Nothing to do here.
end;

{@@ ----------------------------------------------------------------------------
  Attaches cell formatting data for the given cell to the current record.
  Is called from all writing methods of cell contents.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteCellFormatting(AStream: TStream; ACell: PCell;
  XFIndex: Word);
type
  TCellFmtRecord = packed record
    XFIndex_Locked_Hidden: Byte;
    Format_Font: Byte;
    Align_Border_BkGr: Byte;
  end;
var
  rec: TCellFmtRecord;
  fmt: PsCellFormat;
  w: Word;
begin
  fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
  rec.XFIndex_Locked_Hidden := 0;  // to silence the compiler...
  FillChar(rec, SizeOf(rec), 0);

  if fmt^.UsedFormattingFields <> [] then
  begin
    // 1st byte:
    //   Mask $3F: Index to XF record
    //   Mask $40: 1 = Cell is locked
    //   Mask $80: 1 = Formula is hidden
    rec.XFIndex_Locked_Hidden := Min(XFIndex, $3F) and $3F;

    // 2nd byte:
    //   Mask $3F: Index to FORMAT record
    //   Mask $C0: Index to FONT record
    w := fmt^.FontIndex shr 6;   // was shl --> MUST BE shr!   // ??????????????????????
    rec.Format_Font := Lo(w);

    // 3rd byte
    //   Mask $07: horizontal alignment
    //   Mask $08: Cell has left border
    //   Mask $10: Cell has right border
    //   Mask $20: Cell has top border
    //   Mask $40: Cell has bottom border
    //   Mask $80: Cell has shaded background
    if uffHorAlign in fmt^.UsedFormattingFields then
      rec.Align_Border_BkGr := ord(fmt^.HorAlignment);
    if uffBorder in fmt^.UsedFormattingFields then begin
      if cbNorth in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $20;
      if cbWest in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $08;
      if cbEast in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $10;
      if cbSouth in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $40;
    end;
    if uffBackground in fmt^.UsedFormattingFields then
      rec.Align_Border_BkGr := rec.Align_Border_BkGr or $80;
  end;
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 CODEPAGE record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteCodePage(AStream: TStream; ACodePage: String);
//  AEncoding: TsEncoding);
begin
  if ACodePage = 'cp1251' then begin
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_CODEPAGE));
    AStream.WriteWord(WordToLE(2));
    AStream.WriteWord(WordToLE(WORD_CP_1258_Latin1_BIFF2_3));
    FCodePage := ACodePage;
  end else
    inherited;
              (*
  if AEncoding = seLatin1 then begin
    cp := WORD_CP_1258_Latin1_BIFF2_3;
    FCodePage := 'cp1252';

    { BIFF Record header }
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_CODEPAGE));
    AStream.WriteWord(WordToLE(2));
    AStream.WriteWord(WordToLE(cp));
  end else
    inherited; *)
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 COLWIDTH record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteColWidth(AStream: TStream; ACol: PCol);
type
  TColRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    StartCol: Byte;
    EndCol: Byte;
    ColWidth: Word;
  end;
var
  rec: TColRecord;
  w: Integer;
begin
  if Assigned(ACol) then begin
    { BIFF record header }
    rec.RecordID := WordToLE(INT_EXCEL_ID_COLWIDTH);
    rec.RecordSize := WordToLE(4);

    { Start and end column }
    rec.StartCol := ACol^.Col;
    rec.EndCol := ACol^.Col;

    { Column width }
    { calculate width to be in units of 1/256 of pixel width of character "0" }
    w := round(ACol^.Width * 256);
    rec.ColWidth := WordToLE(w);

    { Write out }
    AStream.WriteBuffer(rec, SizeOf(rec));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Write COLWIDTH records for all columns
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteColWidths(AStream: TStream);
var
  j: Integer;
  sheet: TsWorksheet;
  col: PCol;
begin
  sheet := Workbook.GetFirstWorksheet;
  for j := 0 to sheet.Cols.Count-1 do begin
    col := PCol(sheet.Cols[j]);
    WriteColWidth(AStream, col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 DIMENSIONS record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteDimensions(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  firstRow, lastRow, firstCol, lastCol: Cardinal;
  rec: TBIFF2_DimensionsRecord;
begin
  { Determine sheet size }
  GetSheetDimensions(AWorksheet, firstRow, lastRow, firstCol, lastCol);

  { Populate BIFF record }
  rec.RecordID := WordToLE(INT_EXCEL_ID_DIMENSIONS);
  rec.RecordSize := WordToLE(8);
  rec.FirstRow := WordToLE(firstRow);
  if lastRow < $FFFF then             // avoid WORD overflow when adding 1
    rec.LastRowPlus1 := WordToLE(lastRow+1)
  else
    rec.LastRowPlus1 := $FFFF;
  rec.FirstCol := WordToLE(firstCol);
  rec.LastColPlus1 := WordToLE(lastCol+1);

  { Write BIFF record to stream }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{ ------------------------------------------------------------------------------
  Writes an Excel 2 IXFE record
  This record contains the "real" XF index if it is > 62.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteIXFE(AStream: TStream; XFIndex: Word);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_IXFE));
  AStream.WriteWord(WordToLE(2));
  AStream.WriteWord(WordToLE(XFIndex));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 file to a stream

  Excel 2.x files support only one Worksheet per Workbook,
  so only the first one will be written.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteToStream(AStream: TStream);
var
  pane: Byte;
begin
  FWorksheet := Workbook.GetFirstWorksheet;

  WriteBOF(AStream);
    WriteFonts(AStream);
    WriteCodePage(AStream, Workbook.CodePage); //Encoding);
    WriteFormatCount(AStream);
    WriteNumFormats(AStream);
    WriteXFRecords(AStream);
    WriteColWidths(AStream);
    WriteDimensions(AStream, FWorksheet);
    WriteRows(AStream, FWorksheet);

    if (boVirtualMode in Workbook.Options) then
      WriteVirtualCells(AStream)
    else begin
      WriteRows(AStream, FWorksheet);
      WriteCellsToStream(AStream, FWorksheet.Cells);
    end;

    WriteWindow1(AStream);
    //  { -- currently not working
    WriteWindow2(AStream, FWorksheet);
    WritePane(AStream, FWorksheet, false, pane);  // false for "is not BIFF5 or BIFF8"
    WriteSelections(AStream, FWorksheet);
      //}
  WriteEOF(AStream);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 WINDOW1 record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteWindow1(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW1));
  AStream.WriteWord(WordToLE(9));

  { Horizontal position of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE(0));

  { Vertical position of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($0069));

  { Width of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($339F));

  { Height of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($1B5D));

  { Window is visible (1) / hidden (0) }
  AStream.WriteByte(WordToLE(0));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 WINDOW2 record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteWindow2(AStream: TStream;
 ASheet: TsWorksheet);
var
  b: Byte;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW2));
  AStream.WriteWord(WordToLE(14));

  { Show formulas, not results }
  AStream.WriteByte(0);

  { Show grid lines }
  b := IfThen(soShowGridLines in ASheet.Options, 1, 0);
  AStream.WriteByte(b);

  { Show sheet headers }
  b := IfThen(soShowHeaders in ASheet.Options, 1, 0);
  AStream.WriteByte(b);

  { Panes are frozen? }
  b := 0;
  if (soHasFrozenPanes in ASheet.Options) and
     ((ASheet.LeftPaneWidth > 0) or (ASheet.TopPaneHeight > 0))
  then
    b := 1;
  AStream.WriteByte(b);

  { Show zero values as zeros, not empty cells }
  AStream.WriteByte(1);

  { Index to first visible row }
  AStream.WriteWord(0);

  { Index to first visible column }
  AStream.WriteWord(0);

  { Use automatic grid line color }
  AStream.WriteByte(1);

  { RGB of manual grid line color }
  AStream.WriteDWord(0);
end;

procedure TsSpreadBIFF2Writer.WriteXF(AStream: TStream;
 AFormatRecord: PsCellFormat; XFType_Prot: Byte = 0);
var
  rec: TBIFF2_XFRecord;
  b: Byte;
  j: Integer;
begin
  Unused(XFType_Prot);

  { BIFF Record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_XF);
  rec.RecordSize := WordToLE(SizeOf(TBIFF2_XFRecord) - 2*SizeOf(word));

  { Index to FONT record }
  rec.FontIndex := 0;
  if (AFormatRecord <> nil) then
  begin
    if (uffBold in AFormatRecord^.UsedFormattingFields) then
      rec.FontIndex := 1
    else
    if (uffFont in AFormatRecord^.UsedFormattingFields) then
      rec.FontIndex := AFormatRecord^.FontIndex;
  end;

  { Not used byte }
  rec.NotUsed := 0;

  { Number format index and cell flags
      Bit   Mask  Contents
      ----- ----  --------------------------------
      5-0   $3F   Index to (number) FORMAT record
       6    $40   1 = Cell is locked
       7    $80   1 = Formula is hidden }
  rec.NumFormatIndex_Flags := 0;
  if (AFormatRecord <> nil) and (uffNumberFormat in AFormatRecord^.UsedFormattingFields) then
  begin
    // The number formats in the FormatList are still in fpc dialect
    // They will be converted to Excel syntax immediately before writing.
    j := NumFormatList.Find(AFormatRecord^.NumberFormat, AFormatRecord^.NumberFormatStr);
    if j > -1 then
      rec.NumFormatIndex_Flags := NumFormatList[j].Index;

    // Cell flags not used, so far...
  end;

  {Horizontal alignment, border style, and background
  Bit  Mask  Contents
  ---  ----  ------------------------------------------------
  2-0  $07   XF_HOR_ALIGN – Horizontal alignment (0=General, 1=Left, 2=Centred, 3=Right)
   3   $08   1 = Cell has left black border
   4   $10   1 = Cell has right black border
   5   $20   1 = Cell has top black border
   6   $40   1 = Cell has bottom black border
   7   $80   1 = Cell has shaded background }
  b := 0;
  if (AFormatRecord <> nil) then
  begin
    if (uffHorAlign in AFormatRecord^.UsedFormattingFields) then
      b := b + byte(AFormatRecord^.HorAlignment);
    if (uffBorder in AFormatRecord^.UsedFormattingFields) then
    begin
      if cbWest in AFormatRecord^.Border then b := b or $08;
      if cbEast in AFormatRecord^.Border then b := b or $10;
      if cbNorth in AFormatRecord^.Border then b := b or $20;
      if cbSouth in AFormatRecord^.Border then b := b or $40;
    end;
    if (uffBackground in AFormatRecord^.UsedFormattingFields) then
      b := b or $80;
  end;
  rec.HorAlign_Border_BkGr:= b;

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 BOF record
  This must be the first record in an Excel 2 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteBOF(AStream: TStream);
begin
  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_BOF, 4);

  { Unused }
  AStream.WriteWord($0000);

  { Data type }
  AStream.WriteWord(WordToLE(INT_EXCEL_SHEET));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 EOF record
  This must be the last record in an Excel 2 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_EOF, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 font record
  The font data is passed as font index.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteFont(AStream: TStream; AFontIndex: Integer);
var
  Len: Byte;
  lFontName: AnsiString;
  optn: Word;
  font: TsFont;
begin
  font := Workbook.GetFont(AFontIndex);
  if font = nil then  // this happens for FONT4 in case of BIFF
    exit;

  if font.FontName = '' then
    raise Exception.Create('Font name not specified.');
  if font.Size <= 0.0 then
    raise Exception.Create('Font size not specified.');

  lFontName := font.FontName;
  Len := Length(lFontName);

  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_FONT, 4 + 1 + Len * SizeOf(AnsiChar));

  { Height of the font in twips = 1/20 of a point }
  AStream.WriteWord(WordToLE(round(font.Size*20)));

  { Option flags }
  optn := 0;
  if fssBold in font.Style then optn := optn or $0001;
  if fssItalic in font.Style then optn := optn or $0002;
  if fssUnderline in font.Style then optn := optn or $0004;
  if fssStrikeout in font.Style then optn := optn or $0008;
  AStream.WriteWord(WordToLE(optn));

  { Font name: Unicodestring, char count in 1 byte }
  AStream.WriteByte(Len);
  AStream.WriteBuffer(lFontName[1], Len * Sizeof(AnsiChar));

  { Font color: goes into next record! }

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FONTCOLOR));
  AStream.WriteWord(WordToLE(2));

  { Font color index, only first 8 palette entries allowed! }
  AStream.WriteWord(WordToLE(word(FixColor(font.Color))));
end;

{@@ ----------------------------------------------------------------------------
  Writes all font records to the stream
  @see WriteFont
-------------------------------------------------------------------------------}
procedure TsSpreadBiff2Writer.WriteFonts(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to Workbook.GetFontCount-1 do
    WriteFont(AStream, i);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 FORMAT record which describes formatting of numerical data.
-------------------------------------------------------------------------------}
procedure TsSpreadBiff2Writer.WriteNumFormat(AStream: TStream;
  ANumFormatData: TsNumFormatData; AListIndex: Integer);
type
  TNumFormatRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FormatLen: Byte;
  end;
var
  len: Integer;
  s: ansistring;
  rec: TNumFormatRecord;
  buf: array of byte;
begin
  Unused(ANumFormatData);

  s := ConvertEncoding(NumFormatList.FormatStringForWriting(AListIndex), encodingUTF8, FCodePage);
  len := Length(s);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_FORMAT);
  rec.RecordSize := WordToLE(1 + len);

  { Length byte of format string }
  rec.FormatLen := len;

  { Copy the format string characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + SizeOf(ansiChar)*len);
  Move(rec, buf[0], SizeOf(rec));
  Move(s[1], buf[SizeOf(rec)], len*SizeOf(ansiChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(Rec) + SizeOf(ansiChar)*len);

  { Clean up }
  SetLength(buf, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes the number of FORMAT records contained in the file.
  Excel 2 supports only 21 FORMAT records.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteFormatCount(AStream: TStream);
begin
  WriteBiffHeader(AStream, INT_EXCEL_ID_FORMATCOUNT, 2);
  AStream.WriteWord(WordToLE(21)); // there are 21 built-in formats
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 FORMULA record
  The formula is an RPN formula that was converted from usual user-readable
  string to an RPN array by the calling method.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteRPNFormula(AStream: TStream;
  const ARow, ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell);
var
  RPNLength: Word;
  RecordSizePos, FinalPos: Cardinal;
  xf: Word;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  RPNLength := 0;

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMULA));
  RecordSizePos := AStream.Position;
  AStream.WriteWord(0); // We don't know the record size yet. It will be replaced at end.

  { Row and column }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  WriteCellFormatting(AStream, ACell, xf);

  { Encoded result of RPN formula }
  WriteRPNResult(AStream, ACell);

  { 0 = Do not recalculate
    1 = Always recalculate }
  AStream.WriteByte(1);

  { Formula data (RPN token array) }
  if ACell^.SharedFormulaBase <> nil then
    WriteRPNSharedFormulaLink(AStream, ACell, RPNLength)
  else
    WriteRPNTokenArray(AStream, ACell, AFormula, false, RPNLength);

  { Finally write sizes after we know them }
  FinalPos := AStream.Position;
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(17 + RPNLength));
  AStream.Position := FinalPos;

  { Write following STRING record if formula result is a non-empty string }
  if (ACell^.ContentType = cctUTF8String) and (ACell^.UTF8StringValue <> '') then
    WriteStringRecord(AStream, ACell^.UTF8StringValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes the identifier for an RPN function with fixed argument count and
  returns the number of bytes written.
-------------------------------------------------------------------------------}
function TsSpreadBIFF2Writer.WriteRPNFunc(AStream: TStream;
  AIdentifier: Word): Word;
begin
  AStream.WriteByte(Lo(AIdentifier));
  Result := 1;
end;

{@@ ----------------------------------------------------------------------------
  This method is intended to write a link to the cell containing the shared
  formula used by the cell. But since BIFF2 does not support shared formulas
  the writer must copy the shared formula and adapt the relative
  references.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteRPNSharedFormulaLink(AStream: TStream;
  ACell: PCell; var RPNLength: Word);
var
  formula: TsRPNFormula;
begin
  // Create RPN formula from the shared formula base's string formula
  formula := FWorksheet.BuildRPNFormula(ACell);
    // Don't use ACell^.SharedFormulaBase here because this lookup is made
    // by the worksheet automatically.

  // Write adapted copy of shared formula to stream.
  WriteRPNTokenArray(AStream, ACell, formula, false, RPNLength);
  // false --> "do not convert cell addresses to relative offsets", because
  // biff2 does not support shared formulas!

  // Clean up
  SetLength(formula, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes the size of the RPN token array. Called from WriteRPNFormula.
  Overrides xlscommon.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteRPNTokenArraySize(AStream: TStream;
  ASize: Word);
begin
  AStream.WriteByte(ASize);
end;

{@@ ----------------------------------------------------------------------------
  Is intended to write the token array of a shared formula stored in ACell.
  But since BIFF2 does not support shared formulas this method must not do
  anything.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteSharedFormula(AStream: TStream; ACell: PCell);
begin
  Unused(AStream, ACell);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 STRING record which immediately follows a FORMULA record
  when the formula result is a string.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteStringRecord(AStream: TStream;
  AString: String);
var
  s: ansistring;
  len: Integer;
begin
  s := ConvertEncoding(AString, encodingUTF8, FCodePage);
  len := Length(s);

  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_STRING, 1 + len*SizeOf(ansichar));

  { Write string length }
  AStream.WriteByte(len);
  { Write characters }
  AStream.WriteBuffer(s[1], len * SizeOf(ansichar));
end;

{@@ ----------------------------------------------------------------------------
  Writes a Excel 2 BOOLEAN cell record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: Boolean; ACell: PCell);
var
  rec: TBIFF2_BoolErrRecord;
  xf: Integer;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BOOLERROR);
  rec.RecordSize := WordToLE(9);

  { Row and column index }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetCellAttributes(ACell, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Cell value }
  rec.BoolErrValue := ord(AValue);
  rec.ValueType := 0;  // 0 = boolean value, 1 = error value

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 ERROR cell record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  rec: TBIFF2_BoolErrRecord;
  xf: Integer;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BOOLERROR);
  rec.RecordSize := WordToLE(9);

  { Row and column index }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetCellAttributes(ACell, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Cell value }
  rec.BoolErrValue := ConvertToExcelError(AValue);
  rec.ValueType := 1;  // 0 = boolean value, 1 = error value

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 record for an empty cell
  Required if this cell should contain formatting, but no data.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
type
  TBlankRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1, Attrib2, Attrib3: Byte;
  end;
var
  xf: Word;
  rec: TBlankRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BLANK);
  rec.RecordSize := WordToLE(7);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetCellAttributes(ACell, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Write out }
  AStream.WriteBuffer(rec, Sizeof(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 LABEL record
  If the string length exceeds 255 bytes, the string will be truncated and an
  error message will be logged.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  MAXBYTES = 255; //limit for this format
var
  L: Byte;
  AnsiText: ansistring;
  rec: TBIFF2_LabelRecord;
  buf: array of byte;
var
  xf: Word;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  if AValue = '' then Exit; // Writing an empty text doesn't work

  AnsiText := UTF8ToISO_8859_1(AValue);

  if Length(AnsiText) > MAXBYTES then begin
    // BIFF 5 does not support labels/text bigger than 255 chars,
    // so BIFF2 won't either
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    AnsiText := Copy(AnsiText, 1, MAXBYTES);
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
      MAXBYTES, GetCellString(ARow, ACol)
    ]);
  end;
  L := Length(AnsiText);

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_LABEL);
  rec.RecordSize := WordToLE(8 + L);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetCellAttributes(ACell, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Text length: 8 bit }
  rec.TextLen := L;

  { Copy the text characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + SizeOf(ansiChar)*L);
  Move(rec, buf[0], SizeOf(rec));
  Move(AnsiText[1], buf[SizeOf(rec)], L*SizeOf(ansiChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(Rec) + SizeOf(ansiChar)*L);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 NUMBER record
  A "number" is a 64-bit IEE 754 floating point.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
var
  xf: Word;
  rec: TBIFF2_NumberRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_NUMBER);
  rec.RecordSize := WordToLE(15);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetCellAttributes(ACell, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Number value }
  rec.Value := AValue;

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(Rec));
end;

procedure TsSpreadBIFF2Writer.WriteRow(AStream: TStream; ASheet: TsWorksheet;
  ARowIndex, AFirstColIndex, ALastColIndex: Cardinal; ARow: PRow);
var
  containsXF: Boolean;
  rowheight: Word;
  w: Word;
  h: Single;
begin
  if (ARowIndex >= FLimitations.MaxRowCount) or (AFirstColIndex >= FLimitations.MaxColCount)
    or (ALastColIndex >= FLimitations.MaxColCount)
  then
    exit;

  Unused(ASheet);

  containsXF := false;

  { BIFF record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_ROW, IfThen(containsXF, 18, 13));

  { Index of row }
  AStream.WriteWord(WordToLE(Word(ARowIndex)));

  { Index to column of the first cell which is described by a cell record }
  AStream.WriteWord(WordToLE(Word(AFirstColIndex)));

  { Index to column of the last cell which is described by a cell record, increased by 1 }
  AStream.WriteWord(WordToLE(Word(ALastColIndex) + 1));

  { Row height (in twips, 1/20 point) and info on custom row height }
  h := Workbook.GetFont(0).Size;
  if (ARow = nil) or (ARow^.Height = ASheet.DefaultRowHeight) then
    rowheight := PtsToTwips((ASheet.DefaultRowHeight + ROW_HEIGHT_CORRECTION) * h)
  else
  if (ARow^.Height = 0) then
    rowheight := 0
  else
    rowheight := PtsToTwips((ARow^.Height + ROW_HEIGHT_CORRECTION) * h);
  w := rowheight and $7FFF;
  AStream.WriteWord(WordToLE(w));

  { not used }
  AStream.WriteWord(0);

  { Contains row attribute field and XF index }
  AStream.WriteByte(ord(containsXF));

  { Relative offset to calculate stream position of the first cell record for this row }
  AStream.WriteWord(0);

  if containsXF then begin
    { Default row attributes }
    AStream.WriteByte(0);
    AStream.WriteByte(0);
    AStream.WriteByte(0);

    { Index to XF record }
    AStream.WriteWord(WordToLE(15));
  end;
end;

{*******************************************************************
*  Initialization section
*
*  Registers this reader / writer on fpSpreadsheet
*  Converts the palette to litte-endian
*
*******************************************************************}

initialization

  RegisterSpreadFormat(TsSpreadBIFF2Reader, TsSpreadBIFF2Writer, sfExcel2);
  MakeLEPalette(@PALETTE_BIFF2, Length(PALETTE_BIFF2));

end.
