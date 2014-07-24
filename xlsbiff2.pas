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
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  fpspreadsheet, xlscommon, fpsutils, lconvencoding;
  
type

  { TsBIFF2NumFormatList }
  TsBIFF2NumFormatList = class(TsCustomNumFormatList)
  protected
    procedure AddBuiltinFormats; override;
  public
    constructor Create(AWorkbook: TsWorkbook);
    procedure ConvertBeforeWriting(var AFormatString: String;
      var ANumFormat: TsNumberFormat); override;
    function FindFormatOf(AFormatCell: PCell): Integer; override;
  end;

  { TsSpreadBIFF2Reader }

  TsSpreadBIFF2Reader = class(TsSpreadBIFFReader)
  private
    WorkBookEncoding: TsEncoding;
    FWorksheet: TsWorksheet;
    FFont: TsFont;
  protected
    procedure ApplyCellFormatting(ACell: PCell; XFIndex: Word); override;
    procedure CreateNumFormatList; override;
    procedure ExtractNumberFormat(AXFIndex: WORD;
      out ANumberFormat: TsNumberFormat; out ANumberFormatStr: String); override;
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadColWidth(AStream: TStream);
    procedure ReadFont(AStream: TStream);
    procedure ReadFontColor(AStream: TStream);
    procedure ReadFormat(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadInteger(AStream: TStream);
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
    procedure ReadRowColXF(AStream: TStream; out ARow, ACol: Cardinal; out AXF: Word); override;
    procedure ReadRowInfo(AStream: TStream); override;
    function ReadRPNFunc(AStream: TStream): Word; override;
    function ReadRPNTokenArraySize(AStream: TStream): Word; override;
    procedure ReadStringRecord(AStream: TStream); override;
    procedure ReadWindow2(AStream: TStream); override;
    procedure ReadXF(AStream: TStream);
  public
    { General reading methods }
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); override;
  end;

  { TsSpreadBIFF2Writer }

  TsSpreadBIFF2Writer = class(TsSpreadBIFFWriter)
  private
    function FindXFIndex(ACell: PCell): Word;
    procedure GetCellAttributes(ACell: PCell; XFIndex: Word;
      out Attrib1, Attrib2, Attrib3: Byte);
    { Record writing methods }
    procedure WriteBOF(AStream: TStream);
    procedure WriteCellFormatting(AStream: TStream; ACell: PCell; XFIndex: Word);
    procedure WriteColWidth(AStream: TStream; ACol: PCol);
    procedure WriteColWidths(AStream: TStream);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFontIndex: Integer);
    procedure WriteFonts(AStream: TStream);
    procedure WriteFormatCount(AStream: TStream);
    procedure WriteIXFE(AStream: TStream; XFIndex: Word);
    procedure WriteXF(AStream: TStream; AFontIndex, AFormatIndex: byte;
      ABorders: TsCellBorders = []; AHorAlign: TsHorAlignment = haLeft;
      AddBackground: Boolean = false);
    procedure WriteXFFieldsForFormattingStyles(AStream: TStream);
    procedure WriteXFRecords(AStream: TStream);
  protected
    procedure CreateNumFormatList; override;
    procedure ListAllNumFormats; override;
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal; ACell: PCell); override;
    procedure WriteFormat(AStream: TStream; AFormatData: TsNumFormatData;
      AListIndex: Integer); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;
    procedure WriteRow(AStream: TStream; ASheet: TsWorksheet;
      ARowIndex, AFirstColIndex, ALastColIndex: Cardinal; ARow: PRow); override;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal;
      const AFormula: TsRPNFormula; ACell: PCell); override;
    function WriteRPNFunc(AStream: TStream; AIdentifier: Word): Word; override;
    procedure WriteRPNTokenArraySize(AStream: TStream; ASize: Word); override;
    procedure WriteStringRecord(AStream: TStream; AString: String); override;
    procedure WriteWindow1(AStream: TStream); override;
    procedure WriteWindow2(AStream: TStream; ASheet: TsWorksheet);
  public
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
  Math, fpsNumFormatParser;

const
  { Excel record IDs }
  INT_EXCEL_ID_BLANK      = $0001;
  INT_EXCEL_ID_INTEGER    = $0002;
  INT_EXCEL_ID_NUMBER     = $0003;
  INT_EXCEL_ID_LABEL      = $0004;
  INT_EXCEL_ID_ROW        = $0008;
  INT_EXCEL_ID_BOF        = $0009;
  INT_EXCEL_ID_INDEX      = $000B;
  INT_EXCEL_ID_FORMAT     = $001E;
  INT_EXCEL_ID_FORMATCOUNT= $001F;
  INT_EXCEL_ID_COLWIDTH   = $0024;
  INT_EXCEL_ID_WINDOW2    = $003E;
  INT_EXCEL_ID_XF         = $0043;
  INT_EXCEL_ID_IXFE       = $0044;
  INT_EXCEL_ID_FONTCOLOR  = $0045;

  { BOF record constants }
  INT_EXCEL_SHEET         = $0010;
  INT_EXCEL_CHART         = $0020;
  INT_EXCEL_MACRO_SHEET   = $0040;

type
  TBIFF2LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    TextLen: Byte;
  end;

  TBIFF2NumberRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    Value: Double;
  end;

  TBIFF2IntegerRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    Value: Word;
  end;


{ TsBIFF2NumFormatList }

constructor TsBIFF2NumFormatList.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
end;

{ Prepares the list of built-in number formats. They are created in the default
  dialect for FPC, they have to be converted to Excel syntax before writing.
  Note that Excel2 expects them to be localized. This is something which has to
  be taken account of in ConverBeforeWriting.}
procedure TsBIFF2NumFormatList.AddBuiltinFormats;
var
  fs: TFormatSettings;
  cs: string;
begin
  fs := FWorkbook.FormatSettings;
  cs := fs.CurrencyString;
  AddFormat( 0, '', nfGeneral);
  AddFormat( 1, '0', nfFixed);
  AddFormat( 2, '0.00', nfFixed);
  AddFormat( 3, '#,##0', nfFixedTh);
  AddFormat( 4, '#,##0.00', nfFixedTh);
  AddFormat( 5, '"'+cs+'"#,##0_);("'+cs+'"#,##0)', nfCurrency);
  AddFormat( 6, '"'+cs+'"#,##0_);[Red]("'+cs+'"#,##0)', nfCurrencyRed);
  AddFormat( 7, '"'+cs+'"#,##0.00_);("'+cs+'"#,##0.00)', nfCurrency);
  AddFormat( 8, '"'+cs+'"#,##0.00_);[Red]("'+cs+'"#,##0.00)', nfCurrency);
  AddFormat( 9, '0%', nfPercentage);
  AddFormat(10, '0.00%', nfPercentage);
  AddFormat(11, '0.00E+00', nfExp);
  AddFormat(12, fs.ShortDateFormat, nfShortDate);
  AddFormat(13, fs.LongDateFormat, nfLongDate);
  AddFormat(14, 'd/mmm', nfCustom);
  AddFormat(15, 'mmm/yy', nfCustom);
  AddFormat(16, AddAMPM(fs.ShortTimeFormat, fs), nfShortTimeAM);
  AddFormat(17, AddAMPM(fs.LongTimeFormat, fs), nfLongTimeAM);
  AddFormat(18, fs.ShortTimeFormat, nfShortTime);
  AddFormat(19, fs.LongTimeFormat, nfLongTime);
  AddFormat(20, fs.ShortDateFormat + ' ' + fs.ShortTimeFormat, nfShortDateTime);

  FFirstFormatIndexInFile := 0;  // BIFF2 stores built-in formats to file.
  FNextFormatIndex := 21;    // not needed - there are not user-defined formats
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


function TsBIFF2NumFormatList.FindFormatOf(AFormatCell: PCell): Integer;
var
  parser: TsNumFormatParser;
  decs: Integer;
  dt: string;
begin
  Result := 0;

  parser := TsNumFormatParser.Create(Workbook, AFormatCell^.NumberFormatStr);
  try
    decs := parser.Decimals;
    dt := parser.GetDateTimeCode(0);
  finally
    parser.Free;
  end;

  case AFormatCell^.NumberFormat of
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

procedure TsSpreadBIFF2Reader.ApplyCellFormatting(ACell: PCell; XFIndex: Word);
var
  xfData: TXFListData;
begin
  if Assigned(ACell) then begin
    xfData := TXFListData(FXFList.items[XFIndex]);

    // Font index, "bold" attribute
    if xfData.FontIndex = 1 then
      Include(ACell^.UsedFormattingFields, uffBold)
    else
      Include(ACell^.UsedFormattingFields, uffFont);
    ACell^.FontIndex := xfData.FontIndex;

    // Alignment
    ACell^.HorAlignment := xfData.HorAlignment;
    ACell^.VertAlignment := xfData.VertAlignment;

    // Wordwrap not supported by BIFF2
    Exclude(ACell^.UsedFormattingFields, uffWordwrap);
    // Text rotation not supported by BIFF2
    Exclude(ACell^.UsedFormattingFields, uffTextRotation);

    // Border
    if xfData.Borders <> [] then begin
      Include(ACell^.UsedFormattingFields, uffBorder);
      ACell^.Border := xfData.Borders;
    end else
      Exclude(ACell^.UsedFormattingFields, uffBorder);

    // Background, only shaded, color is ignored
    if xfData.BackgroundColor <> 0 then
      Include(ACell^.UsedFormattingFields, uffBackgroundColor)
    else
      Exclude(ACell^.UsedFormattingFields, uffBackgroundColor);
  end;
end;

{ Creates the correct version of the number format list.
  It is for BIFF2 and BIFF3 file formats. }
procedure TsSpreadBIFF2Reader.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsBIFF2NumFormatList.Create(Workbook);
end;

{ Extracts the number format data from an XF record indexed by AXFIndex.
  Note that BIFF2 supports only 21 formats. }
procedure TsSpreadBIFF2Reader.ExtractNumberFormat(AXFIndex: WORD;
  out ANumberFormat: TsNumberFormat; out ANumberFormatStr: String);
var
  lNumFormatData: TsNumFormatData;
begin
  lNumFormatData := FindNumFormatDataForCell(AXFIndex);
  if lNumFormatData <> nil then begin
    ANumberFormat := lNumFormatData.NumFormat;
    ANumberFormatStr := lNumFormatData.FormatString;
  end else begin
    ANumberFormat := nfGeneral;
    ANumberFormatStr := '';
  end;
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

procedure TsSpreadBIFF2Reader.ReadColWidth(AStream: TStream);
var
  c, c1, c2: Cardinal;
  w: Word;
  col: TCol;
  sheet: TsWorksheet;
begin
  sheet := Workbook.GetFirstWorksheet;
  // read column start and end index of column range
  c1 := AStream.ReadByte;
  c2 := AStream.ReadByte;
  // read col width in 1/256 of the width of "0" character
  w := WordLEToN(AStream.ReadWord);
  // calculate width in units of "characters"
  col.Width := w / 256;
  // assign width to columns
  for c := c1 to c2 do
    sheet.WriteColInfo(c, col);
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

// Read the FORMAT record for formatting numerical data
procedure TsSpreadBIFF2Reader.ReadFormat(AStream: TStream);
begin
  Unused(AStream);
  // We ignore the formats in the file, everything is known
  // (Using the formats in the file would require de-localizing them).
end;

procedure TsSpreadBIFF2Reader.ReadFromStream(AStream: TStream; AData: TsWorkbook);
var
  BIFF2EOF: Boolean;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  // Clear existing fonts. They will be replaced by those from the file.
  FWorkbook.RemoveAllFonts;

  { Store some data about the workbook that other routines need }
  WorkBookEncoding := AData.Encoding;

  BIFF2EOF := False;

  { In BIFF2 files there is only one worksheet, let's create it }
  FWorksheet := AData.AddWorksheet('');

  { Read all records in a loop }
  while not BIFF2EOF do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);

    CurStreamPos := AStream.Position;

    case RecordType of
      INT_EXCEL_ID_BLANK     : ReadBlank(AStream);
      INT_EXCEL_ID_FONT      : ReadFont(AStream);
      INT_EXCEL_ID_FONTCOLOR : ReadFontColor(AStream);
      INT_EXCEL_ID_FORMAT    : ReadFormat(AStream);
      INT_EXCEL_ID_INTEGER   : ReadInteger(AStream);
      INT_EXCEL_ID_NUMBER    : ReadNumber(AStream);
      INT_EXCEL_ID_LABEL     : ReadLabel(AStream);
      INT_EXCEL_ID_FORMULA   : ReadFormula(AStream);
      INT_EXCEL_ID_STRING    : ReadStringRecord(AStream);
      INT_EXCEL_ID_COLWIDTH  : ReadColWidth(AStream);
      INT_EXCEL_ID_ROW       : ReadRowInfo(AStream);
      INT_EXCEL_ID_WINDOW2   : ReadWindow2(AStream);
      INT_EXCEL_ID_PANE      : ReadPane(AStream);
      INT_EXCEL_ID_XF        : ReadXF(AStream);
      INT_EXCEL_ID_BOF       : ;
      INT_EXCEL_ID_EOF       : BIFF2EOF := True;
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    if AStream.Position >= AStream.Size then BIFF2EOF := True;
  end;
end;

procedure TsSpreadBIFF2Reader.ReadFormula(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  ok: Boolean;
  formulaResult: Double = 0.0;
//  rpnFormula: TsRPNFormula;
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
  else begin
    if SizeOf(Double) <> 8 then
      raise Exception.Create('Double is not 8 bytes');

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
  if FWorkbook.ReadFormulas then begin
    ok := ReadRPNTokenArray(AStream, cell^.RPNFormulaValue);
    if not ok then FWorksheet.WriteErrorValue(cell, errFormulaNotSupported);
  end;

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadLabel(AStream: TStream);
var
  rec: TBIFF2LabelRecord;
  L: Byte;
  ARow, ACol: Cardinal;
  XF: Word;
  AValue: ansistring;
  AStrValue: UTF8String;
  cell: PCell;
begin
  { Read entire record, starting at Row, except for string data }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2LabelRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;

  { String with 8-bit size }
   L := rec.TextLen;
  SetLength(AValue, L);
  AStream.ReadBuffer(AValue[1], L);

  { Save the data }
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

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);
  FWorksheet.WriteUTF8Text(cell, AStrValue);

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadNumber(AStream: TStream);
var
  rec: TBIFF2NumberRecord;
  ARow, ACol: Cardinal;
  XF: Word;
  value: Double = 0.0;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  cell: PCell;
begin
  { Read entire record, starting at Row }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2NumberRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
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
  rec: TBIFF2IntegerRecord;
begin
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2NumberRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
  AWord := WordLEToN(rec.Value);

  { Create cell }
  if FIsVirtualMode then begin
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

// Read the row, column and xf index
procedure TsSpreadBIFF2Reader.ReadRowColXF(AStream: TStream;
  out ARow, ACol: Cardinal; out AXF: WORD);
begin
  { BIFF Record data for row and column}
  ARow := WordLEToN(AStream.ReadWord);
  ACol := WordLEToN(AStream.ReadWord);

  { Index to XF record }
  AXF := AStream.ReadByte;

  { Index to format and font record, Cell style - ignored because contained in XF
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
  AStream.ReadBuffer(rowrec, SizeOf(TRowRecord));
  h := WordLEToN(rowrec.Height);
  if h and $8000 = 0 then begin // if this bit were set, rowheight would be default
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

{ Reads the identifier for an RPN function with fixed argument count.
  Valid for BIFF2-BIFF3. }
function TsSpreadBIFF2Reader.ReadRPNFunc(AStream: TStream): Word;
var
  b: Byte;
begin
  b := AStream.ReadByte;
  Result := b;
end;

{ Helper funtion for reading of the size of the token array of an RPN formula.
  Is overridden because BIFF2 uses 1 byte only. }
function TsSpreadBIFF2Reader.ReadRPNTokenArraySize(AStream: TStream): Word;
begin
  Result := AStream.ReadByte;
end;

{ Reads a STRING record which contains the result of string formula. }
procedure TsSpreadBIFF2Reader.ReadStringRecord(AStream: TStream);
var
  len: Byte;
  s: ansistring;
begin
  // The string is a byte-string with 8 bit length
  len := AStream.ReadByte;
  if len > 0 then begin
    SetLength(s, Len);
    AStream.ReadBuffer(s[1], len);
    if (FIncompleteCell <> nil) and (s <> '') then begin
      // The "IncompleteCell" has been identified in the sheet when reading
      // the FORMULA record which precedes the String record.
      FIncompleteCell^.UTF8StringValue := AnsiToUTF8(s);
      FIncompleteCell^.ContentType := cctUTF8String;
      if FIsVirtualMode then
        Workbook.OnReadCellData(Workbook, FIncompleteCell^.Row, FIncompleteCell^.Col, FIncompleteCell);
    end;
  end;
  FIncompleteCell := nil;
end;

{ Reads the WINDOW2 record containing information like "show grid lines",
  "show sheet headers", "panes are frozen", etc. }
procedure TsSpreadBIFF2Reader.ReadWindow2(AStream: TStream);
var
  rgb: DWord;
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
  rgb := DWordToLE(AStream.ReadDWord);
end;

procedure TsSpreadBIFF2Reader.ReadXF(AStream: TStream);
{ Offset Size Contents
    0      1   Index to FONT record (➜5.45)
    1      1   Not used
    2      1   Number format and cell flags:
                 Bit  Mask  Contents
                 5-0  3FH   Index to FORMAT record (➜5.49)
                  6   40H   1 = Cell is locked
                  7   80H   1 = Formula is hidden
    3      1   Horizontal alignment, border style, and background:
                 Bit  Mask  Contents
                 2-0  07H   XF_HOR_ALIGN – Horizontal alignment
                              0 General, 1 Left, 2 Center, 3 Right, 4 Filled
                  3   08H   1 = Cell has left black border
                  4   10H   1 = Cell has right black border
                  5   20H   1 = Cell has top black border
                  6   40H   1 = Cell has bottom black border
                  7   80H   1 = Cell has shaded background }
type
  TXFRecord = packed record
    FontIndex: byte;
    NotUsed: byte;
    NumFormat_Flags: byte;
    HorAlign_Border_BackGround: Byte;
  end;
var
  lData: TXFListData;
  xf: TXFRecord;
  b: Byte;
begin
  AStream.ReadBuffer(xf, SizeOf(xf));

  lData := TXFListData.Create;

  // Font index
  lData.FontIndex := xf.FontIndex;

  // Format index
  lData.FormatIndex := xf.NumFormat_Flags and $3F;

  // Horizontal alignment
  b := xf.HorAlign_Border_Background and MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
    lData.HorAlignment := TsHorAlignment(b)
  else
    lData.HorAlignment := haDefault;

  // Vertical alignment - not used in BIFF2
  lData.VertAlignment := vaBottom;

  // Word wrap - not used in BIFF2
  lData.WordWrap := false;

  // Text rotation - not used in BIFF2
  lData.TextRotation := trHorizontal;

  // Borders
  lData.Borders := [];
  if xf.HorAlign_Border_Background and $08 <> 0 then
    Include(lData.Borders, cbWest);
  if xf.HorAlign_Border_Background and $10 <> 0 then
    Include(lData.Borders, cbEast);
  if xf.HorAlign_Border_Background and $20 <> 0 then
    Include(lData.Borders, cbNorth);
  if xf.HorAlign_Border_Background and $40 <> 0 then
    Include(lData.Borders, cbSouth);

  // Background color not supported, only shaded background
  if xf.HorAlign_Border_Background and $80 <> 0 then
    lData.BackgroundColor := 1    // encodes "shaded background = true"
  else
    ldata.BackgroundColor := 0;   // encodes "shaded background = false"

  // Add the decoded data to the list
  FXFList.Add(lData);

  ldata := TXFListData(FXFList.Items[FXFList.Count-1]);
end;


{ TsSpreadBIFF2Writer }

{ Creates the correct version of the number format list.
  It is for BIFF2 and BIFF3 file formats. }
procedure TsSpreadBIFF2Writer.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsBIFF2NumFormatList.Create(Workbook);
end;

function TsSpreadBIFF2Writer.FindXFIndex(ACell: PCell): Word;
var
  lIndex: Integer;
  lCell: TCell;
begin
  // First try the fast methods for default formats
  if ACell^.UsedFormattingFields = [] then
    Result := 15
  else begin
    // If not, then we need to search in the list of dynamic formats
    lCell := ACell^;
    lIndex := FindFormattingInList(@lCell);

    // Carefully check the index
    if (lIndex < 0) or (lIndex > Length(FFormattingStyles)) then
      raise Exception.Create('[TsSpreadBIFF2Writer.WriteXFIndex] Invalid Index, this should not happen!');
    Result := FFormattingStyles[lIndex].Row;
  end;
end;

{ Determines the cell attributes needed for writing a cell content record, such
  as WriteLabel, WriteNumber, etc.
  The cell attributes contain, in bit masks, xf record index, font index, borders, etc.}
procedure TsSpreadBIFF2Writer.GetCellAttributes(ACell: PCell; XFIndex: Word;
  out Attrib1, Attrib2, Attrib3: Byte);
begin
  if ACell^.UsedFormattingFields = [] then begin
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
  //   Mask $3F: Index to FORMAT record
  //   Mask $C0: Index to FONT record
  Attrib2 := ACell^.FontIndex shr 6;

  // 3rd byte
  //   Mask $07: horizontal alignment
  //   Mask $08: Cell has left border
  //   Mask $10: Cell has right border
  //   Mask $20: Cell has top border
  //   Mask $40: Cell has bottom border
  //   Mask $80: Cell has shaded background
  Attrib3 := 0;
  if uffHorAlign in ACell^.UsedFormattingFields then
    Attrib3 := ord (ACell^.HorAlignment);
  if uffBorder in ACell^.UsedFormattingFields then begin
    if cbNorth in ACell^.Border then Attrib3 := Attrib3 or $20;
    if cbWest in ACell^.Border then Attrib3 := Attrib3 or $08;
    if cbEast in ACell^.Border then Attrib3 := Attrib3 or $10;
    if cbSouth in ACell^.Border then Attrib3 := Attrib3 or $40;
  end;
  if uffBackgroundColor in ACell^.UsedFormattingFields then
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

{ Attaches cell formatting data for the given cell to the current record.
  Is called from all writing methods of cell contents. }
procedure TsSpreadBIFF2Writer.WriteCellFormatting(AStream: TStream; ACell: PCell;
  XFIndex: Word);
var
  b: Byte;
  w: Word;
begin
  if ACell^.UsedFormattingFields = [] then
  begin
    AStream.WriteByte($0);
    AStream.WriteByte($0);
    AStream.WriteByte($0);
    Exit;
  end;

  // 1st byte:
  //   Mask $3F: Index to XF record
  //   Mask $40: 1 = Cell is locked
  //   Mask $80: 1 = Formula is hidden
  AStream.WriteByte(Min(XFIndex, $3F) and $3F);

  // 2nd byte:
  //   Mask $3F: Index to FORMAT record
  //   Mask $C0: Index to FONT record
  w := ACell.FontIndex shr 6;   // was shl --> MUST BE shr!
  b := Lo(w);
  //b := ACell.FontIndex shl 6;
  AStream.WriteByte(b);

  // 3rd byte
  //   Mask $07: horizontal alignment
  //   Mask $08: Cell has left border
  //   Mask $10: Cell has right border
  //   Mask $20: Cell has top border
  //   Mask $40: Cell has bottom border
  //   Mask $80: Cell has shaded background
  b := 0;
  if uffHorAlign in ACell^.UsedFormattingFields then
    b := ord (ACell^.HorAlignment);
  if uffBorder in ACell^.UsedFormattingFields then begin
    if cbNorth in ACell^.Border then b := b or $20;
    if cbWest in ACell^.Border then b := b or $08;
    if cbEast in ACell^.Border then b := b or $10;
    if cbSouth in ACell^.Border then b := b or $40;
  end;
  if uffBackgroundColor in ACell^.UsedFormattingFields then
    b := b or $80;
  AStream.WriteByte(b);
end;

{
  Writes an Excel 2 COLWIDTH record
}
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

    (*
    { BIFF Record header }
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_COLWIDTH));  // BIFF record header
    AStream.WriteWord(WordToLE(4));                      // Record size
    AStream.WriteByte(ACol^.Col);                        // start column
    AStream.WriteByte(ACol^.Col);                        // end column
    { calculate width to be in units of 1/256 of pixel width of character "0" }
    w := round(ACol^.Width * 256);
    AStream.WriteWord(WordToLE(w));                     // write width
    *)
  end;
end;

{
  Write COLWIDTH records for all columns
}
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

{
  Writes an Excel 2 IXFE record
  This record contains the "real" XF index if it is > 62.
}
procedure TsSpreadBIFF2Writer.WriteIXFE(AStream: TStream; XFIndex: Word);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_IXFE));
  AStream.WriteWord(WordToLE(2));
  AStream.WriteWord(WordToLE(XFIndex));
end;

{
  Writes an Excel 2 file to a stream

  Excel 2.x files support only one Worksheet per Workbook,
  so only the first will be written.
}
procedure TsSpreadBIFF2Writer.WriteToStream(AStream: TStream);
var
  sheet: TsWorksheet;
  pane: Byte;
begin
  sheet := Workbook.GetFirstWorksheet;

  WriteBOF(AStream);
    WriteFonts(AStream);
    WriteFormatCount(AStream);
    WriteFormats(AStream);
    WriteXFRecords(AStream);
    WriteColWidths(AStream);
    WriteRows(AStream, sheet);

    if (boVirtualMode in Workbook.Options) then
      WriteVirtualCells(AStream)
    else begin
      WriteRows(AStream, sheet);
      WriteCellsToStream(AStream, sheet.Cells);
    end;

    WriteWindow1(AStream);
    //  { -- currently not working
    WriteWindow2(AStream, sheet);
    WritePane(AStream, sheet, false, pane);  // false for "is not BIFF5 or BIFF8"
    WriteSelections(AStream, sheet);
      //}
  WriteEOF(AStream);
end;

{
  Writes an Excel 2 WINDOW1 record
}
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

{
  Writes an Excel 2 WINDOW2 record
}
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
  AFontIndex, AFormatIndex: byte; ABorders: TsCellBorders = [];
  AHorAlign: TsHorAlignment = haLeft; AddBackground: Boolean = false);
var
  b: Byte;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_XF));
  AStream.WriteWord(WordToLE(4));

  { Index to FONT record }
  AStream.WriteByte(AFontIndex);

  { not used }
  AStream.WriteByte(0);

  { Number format index and cell flags }
  b := AFormatIndex and $3F;
  AStream.WriteByte(b);

  { Horizontal alignment, border style, and background }
  b := byte(AHorAlign);
  if cbWest in ABorders then b := b or $08;
  if cbEast in ABorders then b := b or $10;
  if cbNorth in ABorders then b := b or $20;
  if cbSouth in ABorders then b := b or $40;
  if AddBackground then b := b or $80;
  AStream.WriteByte(b);
end;

procedure TsSpreadBIFF2Writer.WriteXFFieldsForFormattingStyles(AStream: TStream);
var
  i, j: Integer;
  lFontIndex: Word;
  lFormatIndex: Word; //number format
  lBorders: TsCellBorders;
  lAddBackground: Boolean;
  lHorAlign: TsHorAlignment;
begin
  // The loop starts with the first style added manually.
  // First style was already added  (see AddDefaultFormats)
  for i := 1 to Length(FFormattingStyles) - 1 do begin
    // Default styles
    lFontIndex := 0;
    lFormatIndex := 0; //General format (one of the built-in number formats)
    lBorders := [];
    lHorAlign := FFormattingStyles[i].HorAlignment;

    // Now apply the modifications.
    if uffNumberFormat in FFormattingStyles[i].UsedFormattingFields then begin
      j := NumFormatList.FindFormatOf(@FFormattingStyles[i]);
      if j > -1 then
        lFormatIndex := NumFormatList[j].Index;
    end;

    if uffBorder in FFormattingStyles[i].UsedFormattingFields then
      lBorders := FFormattingStyles[i].Border;

    if uffBold in FFormattingStyles[i].UsedFormattingFields then
      lFontIndex := 1;   // must be before uffFont which overrides uffBold

    if uffFont in FFormattingStyles[i].UsedFormattingFields then
      lFontIndex := FFormattingStyles[i].FontIndex;

    lAddBackground := (uffBackgroundColor in FFormattingStyles[i].UsedFormattingFields);

    // And finally write the style
    WriteXF(AStream, lFontIndex, lFormatIndex, lBorders, lHorAlign, lAddBackground);
  end;
end;

procedure TsSpreadBIFF2Writer.WriteXFRecords(AStream: TStream);
begin
  WriteXF(AStream, 0, 0);  // XF0
  WriteXF(AStream, 0, 0);  // XF1
  WriteXF(AStream, 0, 0);  // XF2
  WriteXF(AStream, 0, 0);  // XF3
  WriteXF(AStream, 0, 0);  // XF4
  WriteXF(AStream, 0, 0);  // XF5
  WriteXF(AStream, 0, 0);  // XF6
  WriteXF(AStream, 0, 0);  // XF7
  WriteXF(AStream, 0, 0);  // XF8
  WriteXF(AStream, 0, 0);  // XF9
  WriteXF(AStream, 0, 0);  // XF10
  WriteXF(AStream, 0, 0);  // XF11
  WriteXF(AStream, 0, 0);  // XF12
  WriteXF(AStream, 0, 0);  // XF13
  WriteXF(AStream, 0, 0);  // XF14
  WriteXF(AStream, 0, 0);  // XF15 - Default, no formatting

   // Add all further non-standard/built-in formatting styles
  ListAllFormattingStyles;
  WriteXFFieldsForFormattingStyles(AStream);
end;

{
  Writes an Excel 2 BOF record

  This must be the first record on an Excel 2 stream
}
procedure TsSpreadBIFF2Writer.WriteBOF(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BOF));
  AStream.WriteWord(WordToLE($0004));

  { Unused }
  AStream.WriteWord($0000);

  { Data type }
  AStream.WriteWord(WordToLE(INT_EXCEL_SHEET));
end;

{
  Writes an Excel 2 EOF record

  This must be the last record on an Excel 2 stream
}
procedure TsSpreadBIFF2Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_EOF));
  AStream.WriteWord($0000);
end;

{
  Writes an Excel 2 font record
  The font data is passed as font index.
}
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
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FONT));
  AStream.WriteWord(WordToLE(4 + 1 + Len * Sizeof(AnsiChar)));

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
  AStream.WriteWord(WordToLE(word(font.Color)));
end;

procedure TsSpreadBiff2Writer.WriteFonts(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to Workbook.GetFontCount-1 do
    WriteFont(AStream, i);
end;

procedure TsSpreadBiff2Writer.WriteFormat(AStream: TStream;
  AFormatData: TsNumFormatData; AListIndex: Integer);
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
  Unused(AFormatData);

  s := NumFormatList.FormatStringForWriting(AListIndex);
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

                                  (*
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMAT));
  AStream.WriteWord(WordToLE(1 + len));

  { Format string }
  AStream.WriteByte(len);          // AnsiString, char count in 1 byte
  AStream.WriteBuffer(s[1], len);  // String data
  *)
end;

procedure TsSpreadBIFF2Writer.WriteFormatCount(AStream: TStream);
begin
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMATCOUNT));
  AStream.WriteWord(WordToLE(2));
  AStream.WriteWord(WordToLE(21)); // there are 21 built-in formats
end;

{
  Writes an Excel 2 FORMULA record

  The formula needs to be converted from usual user-readable string
  to an RPN array

  // or, in RPN: A1, B1, +
  SetLength(MyFormula, 3);
  MyFormula[0].TokenID := INT_EXCEL_TOKEN_TREFV; A1
  MyFormula[0].Col := 0;
  MyFormula[0].Row := 0;
  MyFormula[1].TokenID := INT_EXCEL_TOKEN_TREFV; B1
  MyFormula[1].Col := 1;
  MyFormula[1].Row := 0;
  MyFormula[2].TokenID := INT_EXCEL_TOKEN_TADD;  +
}
procedure TsSpreadBIFF2Writer.WriteRPNFormula(AStream: TStream;
  const ARow, ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell);
var
  FormulaResult: double;
  i: Integer;
  RPNLength: Word;
  TokenArraySizePos, RecordSizePos, FinalPos: Cardinal;
  FormulaKind, ExtraInfo: Word;
  r: Cardinal;
  len: Integer;
  s: ansistring;
  xf: Word;
begin
  RPNLength := 0;
  FormulaResult := 0.0;

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMULA));
  RecordSizePos := AStream.Position;
  AStream.WriteWord(WordToLE(17 + RPNLength));

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
  WriteRPNTokenArray(AStream, AFormula, RPNLength);

  { Finally write sizes after we know them }
  FinalPos := AStream.Position;
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(17 + RPNLength));
  AStream.Position := FinalPos;

  { Write following STRING record if formula result is a non-empty string }
  if (ACell^.ContentType = cctUTF8String) and (ACell^.UTF8StringValue <> '') then
    WriteStringRecord(AStream, ACell^.UTF8StringValue);

end;

{ Writes the identifier for an RPN function with fixed argument count and
  returns the number of bytes written. }
function TsSpreadBIFF2Writer.WriteRPNFunc(AStream: TStream; AIdentifier: Word): Word;
begin
  AStream.WriteByte(Lo(AIdentifier));
  Result := 1;
end;

{ Writes the size of the RPN token array. Called from WriteRPNFormula.
  Overrides xlscommon. }
procedure TsSpreadBIFF2Writer.WriteRPNTokenArraySize(AStream: TStream;
  ASize: Word);
begin
  AStream.WriteByte(Lo(ASize));
end;

{ Writes an Excel 2 STRING record which immediately follows a FORMULA record
  when the formula result is a string. }
procedure TsSpreadBIFF2Writer.WriteStringRecord(AStream: TStream;
  AString: String);
var
  s: ansistring;
  len: Integer;
begin
//  s := AString;           // Why not call UTF8ToAnsi?
  s := UTF8ToAnsi(AString);
  len := Length(s);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_STRING));
  AStream.WriteWord(WordToLE(1 + len*SizeOf(Char)));

  { Write string length }
  AStream.WriteByte(len);
  { Write characters }
  AStream.WriteBuffer(s[1], len * SizeOf(Char));
end;


{*******************************************************************
*  TsSpreadBIFF2Writer.WriteBlank ()
*
*  DESCRIPTION:    Writes an Excel 2 record for an empty cell
*
*                  Required if this cell should contain formatting
*
*******************************************************************}
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

  (*
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BLANK));
  AStream.WriteWord(WordToLE(7));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  WriteCellFormatting(AStream, ACell, xf);
  *)
end;

{*******************************************************************
*  TsSpreadBIFF2Writer.WriteLabel ()
*
*  DESCRIPTION:    Writes an Excel 2 LABEL record
*
*                  Writes a string to the sheet
*                  If the string length exceeds 255 bytes, the string
*                  will be truncated and an exception will be raised as
*                  a warning.
*
*******************************************************************}
procedure TsSpreadBIFF2Writer.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  MAXBYTES = 255; //limit for this format
var
  L: Byte;
  AnsiText: ansistring;
  TextTooLong: boolean=false;
  rec: TBIFF2LabelRecord;
  buf: array of byte;
var
  xf: Word;
begin
  if AValue = '' then Exit; // Writing an empty text doesn't work

  AnsiText := UTF8ToISO_8859_1(AValue);

  if Length(AnsiText) > MAXBYTES then begin
    // BIFF 5 does not support labels/text bigger than 255 chars,
    // so BIFF2 won't either
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    TextTooLong:=true;
    AnsiText := Copy(AnsiText, 1, MAXBYTES);
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


                  (*
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_LABEL));
  AStream.WriteWord(WordToLE(8 + L));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  WriteCellFormatting(AStream, ACell, xf);

  { String with 8-bit size }
  AStream.WriteByte(L);
  AStream.WriteBuffer(AnsiText[1], L);
                   *)

  {
  //todo: keep a log of errors and show with an exception after writing file or something.
  We can't just do the following
  if TextTooLong then
    Raise Exception.CreateFmt('Text value exceeds %d character limit in cell [%d,%d]. Text has been truncated.',[MaxBytes,ARow,ACol]);
    because the file wouldn't be written.
  }
end;

{*******************************************************************
*  TsSpreadBIFF2Writer.WriteNumber ()
*
*  DESCRIPTION:    Writes an Excel 2 NUMBER record
*
*                  Writes a number (64-bit IEE 754 floating point) to the sheet
*
*******************************************************************}
procedure TsSpreadBIFF2Writer.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
var
  xf: Word;
  rec: TBIFF2NumberRecord;
begin
  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);
                  (*
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_NUMBER));
  AStream.WriteWord(WordToLE(15));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  WriteCellFormatting(AStream, ACell, xf);

  { IEE 754 floating-point value }
  AStream.WriteBuffer(AValue, 8);
  *)

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
  Unused(ASheet);

  containsXF := false;

  { BIFF record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_ROW));
  AStream.WriteWord(WordToLE(IfThen(containsXF, 18, 13)));

  { Index of row }
  AStream.WriteWord(WordToLE(Word(ARowIndex)));

  { Index to column of the first cell which is described by a cell record }
  AStream.WriteWord(WordToLE(Word(AFirstColIndex)));

  { Index to column of the last cell which is described by a cell record, increased by 1 }
  AStream.WriteWord(WordToLE(Word(ALastColIndex) + 1));

  { Row height (in twips, 1/20 point) and info on custom row height }
  h := Workbook.GetFont(0).Size;
  if (ARow = nil) or (ARow^.Height = Workbook.DefaultRowHeight) then
    rowheight := PtsToTwips((Workbook.DefaultRowHeight + ROW_HEIGHT_CORRECTION) * h)
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
