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

  { TsSpreadBIFF2Reader }

  TsSpreadBIFF2Reader = class(TsSpreadBIFFReader)
  private
    WorkBookEncoding: TsEncoding;
    FWorksheet: TsWorksheet;
    FFont: TsFont;
  protected
    procedure ApplyCellFormatting(ARow, ACol: Cardinal; XFIndex: Word); override;
    procedure ExtractNumberFormat(AXFIndex: WORD;
      out ANumberFormat: TsNumberFormat; out ADecimals: Word;
      out ANumberFormatStr: String); override;
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadColWidth(AStream: TStream);
    procedure ReadFont(AStream: TStream);
    procedure ReadFontColor(AStream: TStream);
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadInteger(AStream: TStream);
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
    procedure ReadRowColXF(AStream: TStream; out ARow, ACol: Cardinal; out AXF: Word); override;
    procedure ReadRowInfo(AStream: TStream); override;
    procedure ReadXF(AStream: TStream);
  public
    { General reading methods }
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); override;
  end;

  { TsSpreadBIFF2Writer }

  TsSpreadBIFF2Writer = class(TsSpreadBIFFWriter)
  private
    function  FindXFIndex(ACell: PCell): Word;
    { Record writing methods }
    procedure WriteBOF(AStream: TStream);
    procedure WriteCellFormatting(AStream: TStream; ACell: PCell; XFIndex: Word);
    procedure WriteColWidth(AStream: TStream; ACol: PCol);
    procedure WriteColWidths(AStream: TStream);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFontIndex: Integer);
    procedure WriteFonts(AStream: TStream);
    procedure WriteFormat(AStream: TStream; AFormatCode: String);
    procedure WriteFormatCount(AStream: TStream);
    procedure WriteFormats(AStream: TStream);
    procedure WriteIXFE(AStream: TStream; XFIndex: Word);
    procedure WriteXF(AStream: TStream; AFontIndex, AFormatIndex: byte;
      ABorders: TsCellBorders = []; AHorAlign: TsHorAlignment = haLeft;
      AddBackground: Boolean = false);
    procedure WriteXFFieldsForFormattingStyles(AStream: TStream);
    procedure WriteXFRecords(AStream: TStream);
  protected
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal; ACell: PCell); override;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal; const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double; ACell: PCell); override;
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

  { These are the built-in number formats of BIFF 2. They are not stored in
    the file. Note that, compared to the BUFF5+ built-in formats, two formats
    are missing and the indexes are offset by 2 after #11.
    It seems that BIFF2 can handle only these 21 formats. The other formats
    available in fpspreadsheet are mapped to these 21 formats such that least
    destruction is made. }
  NUMFORMAT_BIFF2: array[0..20] of string = (
    'General',               // 0
    '0',
    '0.00',
    '#,##0',
    '#,##0.00',
    '"$"#,##0_);("$"#,##0)', // 5
    '"$"#,##0_);[Red]("$"#,##0)',
    '"$"#,##0.00_);("$"#,##0.00)',
    '"$"#,##0.00_);[Red]("$"#,##0.00)',
    '0%',
    '0.00%',                 // 10
    '0.00E+00',
    'M/D/YY',
    'D-MMM-YY',
    'D-MMM',
    'MMM-YY',                // 15
    'h:mm AM/PM',
    'h:mm:ss AM/PM',
    'h:mm',
    'h:mm:ss',
    'M/D/YY h:mm'            // 20
  );

implementation

const
  { Excel record IDs }
  INT_EXCEL_ID_BLANK      = $0001;
  INT_EXCEL_ID_INTEGER    = $0002;
  INT_EXCEL_ID_NUMBER     = $0003;
  INT_EXCEL_ID_LABEL      = $0004;
  INT_EXCEL_ID_FORMULA    = $0006;
  INT_EXCEL_ID_ROWINFO    = $0008;
  INT_EXCEL_ID_BOF        = $0009;
  INT_EXCEL_ID_EOF        = $000A;
  INT_EXCEL_ID_FORMAT     = $001E;
  INT_EXCEL_ID_FORMATCOUNT= $001F;
  INT_EXCEL_ID_COLWIDTH   = $0024;
  INT_EXCEL_ID_XF         = $0043;
  INT_EXCEL_ID_IXFE       = $0044;
  INT_EXCEL_ID_FONTCOLOR  = $0045;

  { Cell Addresses constants }
  MASK_EXCEL_ROW          = $3FFF;
  MASK_EXCEL_RELATIVE_COL = $4000;  // This is according to Microsoft documentation,
  MASK_EXCEL_RELATIVE_ROW = $8000;  // but opposite to OpenOffice documentation!

  { BOF record constants }
  INT_EXCEL_SHEET         = $0010;
  INT_EXCEL_CHART         = $0020;
  INT_EXCEL_MACRO_SHEET   = $0040;

  { FORMAT record constants for BIFF2 }
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
  FORMAT_SCI_1_DECIMAL = 11;      //scientific, 1 decimal -- not present in BIFF2 -- mapped to EXP_2_DECIMALS
  FORMAT_SHORT_DATE = 12;         //short date
  FORMAT_DATE_DM = 14;            //date D-MMM
  FORMAT_DATE_MY = 15;            //date MMM-YYYY
  FORMAT_SHORT_TIME_AM = 16;      //short time H:MM with AM
  FORMAT_LONG_TIME_AM = 17;       //long time H:MM:SS with AM
  FORMAT_SHORT_TIME = 18;         //short time H:MM
  FORMAT_LONG_TIME = 19;          //long time H:MM:SS
  FORMAT_SHORT_DATETIME = 20;     //short date+time
  { The next three formats are not available in BIFF2. They should not be used
    when a file is to be saved in BIFF2. If it IS saved as BIFF2 the formats
    are mapped to FORMAT_LONG_TIME}
  FORMAT_TIME_MS = 19;            //time MM:SS
  FORMAT_TIME_MSZ = 19;           //time MM:SS.0
  FORMAT_TIME_INTERVAL = 19;      //time [hh]:mm:ss, hh can be >24

{ TsSpreadBIFF2Reader }

procedure TsSpreadBIFF2Reader.ApplyCellFormatting(ARow, ACol: Cardinal;
  XFIndex: Word);
var
  lCell: PCell;
  xfData: TXFListData;
  style: Byte;
begin
  lCell := FWorksheet.GetCell(ARow, ACol);

  if Assigned(lCell) then begin
    xfData := TXFListData(FXFList.items[XFIndex]);

    // Font index, "bold" attribute
    if xfData.FontIndex = 1 then
      Include(lCell^.UsedFormattingFields, uffBold)
    else
      Include(lCell^.UsedFormattingFields, uffFont);
    lCell^.FontIndex := xfData.FontIndex;

    // Alignment
    lCell^.HorAlignment := xfData.HorAlignment;
    lCell^.VertAlignment := xfData.VertAlignment;

    // Wordwrap not supported by BIFF2
    Exclude(lCell^.UsedFormattingFields, uffWordwrap);
    // Text rotation not supported by BIFF2
    Exclude(lCell^.UsedFormattingFields, uffTextRotation);

    // Border
    if xfData.Borders <> [] then begin
      Include(lCell^.UsedFormattingFields, uffBorder);
      lCell^.Border := xfData.Borders;
    end else
      Exclude(lCell^.UsedFormattingFields, uffBorder);

    // Background, only shaded, color is ignored
    if xfData.BackgroundColor <> 0 then
      Include(lCell^.UsedFormattingFields, uffBackgroundColor)
    else
      Exclude(lCell^.UsedFormattingFields, uffBackgroundColor);
  end;
end;

{ Extracts the number format data from an XF record indexed by AXFIndex.
  Note that BIFF2 supports only 21 formats. }
procedure TsSpreadBIFF2Reader.ExtractNumberFormat(AXFIndex: WORD;
  out ANumberFormat: TsNumberFormat; out ADecimals: Word;
  out ANumberFormatStr: String);
const
  NOT_USED = nfGeneral;
  fmts: array[1..20] of TsNumberFormat = (
    nfFixed, nfFixed, nfFixedTh, nfFixedTh, nfFixedTh,               // 1..5
    nfFixedTh, nfFixedTh, nfFixedTh, nfPercentage, nfPercentage,     // 6..10
    nfExp, nfShortDate, nfShortDate, nfFmtDateTime, nfFmtDateTime,   // 11..15
    nfShortTimeAM, nfLongTimeAM, nfShortTime, nfLongTime, nfShortDateTime// 16..20
  );
  decs: array[1..20] of word = (
    0, 2, 0, 2, 0, 0, 2, 2, 0, 2,     // 1..10
    2, 0, 0, 0, 0, 0, 0, 0, 0, 0      // 11..20
  );
var
  lFormatData: TFormatListData;
  lXFData: TXFListData;
  isAMPM: Boolean;
  isLongTime: Boolean;
  isMilliSec: Boolean;
  t,d: Boolean;
begin
  ANumberFormat := nfGeneral;
  ANumberFormatStr := '';
  ADecimals := 0;

  lFormatData := FindFormatDataForCell(AXFIndex);
  if lFormatData = nil then begin
    // no custom format, so first test for default formats
    lXFData := TXFListData (FXFList.Items[AXFIndex]);
    if (lXFData.FormatIndex > 0) and (lXFData.FormatIndex <= 20) then begin
      ANumberFormat := fmts[lXFData.FormatIndex];
      ADecimals := decs[lXFData.FormatIndex];
    end;
  end else
  // The next is copied from xlscommon - I think it's not necessary here
  if IsPercentNumberFormat(lFormatData.FormatString, ADecimals) then
    ANumberFormat := nfPercentage
  else
  if IsExpNumberFormat(lFormatData.Formatstring, ADecimals) then
    ANumberFormat := nfExp
  else
  if IsThousandSepNumberFormat(lFormatData.FormatString, ADecimals) then
    ANumberFormat := nfFixedTh
  else
  if IsFixedNumberFormat(lFormatData.FormatString, ADecimals) then
    ANumberFormat := nfFixed
  else begin
    t := IsTimeFormat(lFormatData.FormatString, isLongTime, isAMPM, isMilliSec);
    d := IsDateFormat(lFormatData.FormatString);
    if d and t then
      ANumberFormat := nfShortDateTime
    else
    if d then
      ANumberFormat := nfShortDate
    else
    if t then begin
      if isAMPM then begin
        if isLongTime then
          ANumberFormat := nfLongTimeAM
        else
          ANumberFormat := nfShortTimeAM;
      end else begin
        if isLongTime then
          ANumberFormat := nfLongTime
        else
          ANumberFormat := nfShortTime;
      end;
    end;
  end;
end;

procedure TsSpreadBIFF2Reader.ReadBlank(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
begin
  ReadRowColXF(AStream, ARow, ACol, XF);
  ApplyCellFormatting(ARow, ACol, XF);
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
    INT_EXCEL_ID_INTEGER   : ReadInteger(AStream);
    INT_EXCEL_ID_NUMBER    : ReadNumber(AStream);
    INT_EXCEL_ID_LABEL     : ReadLabel(AStream);
    INT_EXCEL_ID_FORMULA   : ReadFormula(AStream);
    INT_EXCEL_ID_COLWIDTH  : ReadColWidth(AStream);
    INT_EXCEL_ID_ROWINFO   : ReadRowInfo(AStream);
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
begin

end;

procedure TsSpreadBIFF2Reader.ReadLabel(AStream: TStream);
var
  L: Byte;
  ARow, ACol: Cardinal;
  XF: Word;
  AValue: array[0..255] of Char;
  AStrValue: UTF8String;
begin
  { BIFF Record row/column/style }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { String with 8-bit size }
  L := AStream.ReadByte();
  AStream.ReadBuffer(AValue, L);
  AValue[L] := #0;

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
  FWorksheet.WriteUTF8Text(ARow, ACol, AStrValue);

  { Apply formatting to cell }
  ApplyCellFormatting(ARow, ACol, XF);
end;

procedure TsSpreadBIFF2Reader.ReadNumber(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  value: Double;
  dt: TDateTime;
  nf: TsNumberFormat;
  nd: Word;
  nfs: String;
begin
  { BIFF Record row/column/style }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { IEE 754 floating-point value }
  AStream.ReadBuffer(value, 8);

  {Find out what cell type, set content type and value}
  ExtractNumberFormat(XF, nf, nd, nfs);
  if IsDateTime(value, nf, dt) then
    FWorksheet.WriteDateTime(ARow, ACol, dt, nf, nfs)
  else
    FWorksheet.WriteNumber(ARow, ACol, value, nf, nd);

  { Apply formatting to cell }
  ApplyCellFormatting(ARow, ACol, XF);
end;

procedure TsSpreadBIFF2Reader.ReadInteger(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  AWord  : Word;
begin
  { BIFF Record row/column/style }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { 16 bit unsigned integer }
  AStream.ReadBuffer(AWord, 2);

  { Save the data }
  FWorksheet.WriteNumber(ARow, ACol, AWord);

  { Apply formatting to cell }
  ApplyCellFormatting(ARow, ACol, XF);
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
    lRow^.Height := TwipsToMillimeters(h and $7FFF);
  end else
    lRow^.AutoHeight := true;
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
    lData.BackgroundColor := 1    // shaded background = "true"
  else
    ldata.BackgroundColor := 0;   // shaded background = "false"

  // Add the decoded data to the list
  FXFList.Add(lData);

  ldata := TXFListData(FXFList.Items[FXFList.Count-1]);
end;


{ TsSpreadBIFF2Writer }

function TsSpreadBIFF2Writer.FindXFIndex(ACell: PCell): Word;
var
  i: Integer;
begin
  if ACell^.UsedFormattingFields = [] then
    Result := 15
  else begin
    // If not, then we need to search in the list of dynamic formats
    i := FindFormattingInList(ACell);
    // Carefully check the index
    if (i < 0) or (i > Length(FFormattingStyles)) then
      raise Exception.Create('[TsSpreadBIFF2Writer.WriteXFIndex] Invalid Index, this should not happen!');
    Result := FFormattingStyles[i].Row;
  end;
end;

{
  Attaches cell formatting data for the given cell to the current record.
  Is called from all writing methods of cell contents.
}
procedure TsSpreadBIFF2Writer.WriteCellFormatting(AStream: TStream; ACell: PCell;
  XFIndex: Word);
var
  b: Byte;
  xf: Word;
  i: Integer;
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
  AStream.WriteByte(XFIndex and $3F);

  // 2nd byte:
  //   Mask $3F: Index to FORMAT record
  //   Mask $C0: Index to FONT record
  b := ACell.FontIndex shl 6;
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
var
  w: Integer;
begin
  if Assigned(ACol) then begin
    { BIFF Record header }
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_COLWIDTH));  // BIFF record header
    AStream.WriteWord(WordToLE(4));                      // Record size
    AStream.WriteByte(ACol^.Col);                        // start column
    AStream.WriteByte(ACol^.Col);                        // end column
    { calculate width to be in units of 1/256 of pixel width of character "0" }
    w := round(ACol^.Width * 256);
    AStream.WriteWord(WordToLE(w));                     // write width
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
begin
  WriteBOF(AStream);

  WriteFonts(AStream);
  WriteFormats(AStream);
  WriteXFRecords(AStream);
  WriteColWidths(AStream);
  WriteCellsToStream(AStream, Workbook.GetFirstWorksheet.Cells);

  WriteEOF(AStream);
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
  i: Integer;
  lFontIndex: Word;
  lFormatIndex: Word; //number format
  lBorders: TsCellBorders;
  lAddBackground: Boolean;
  lHorAlign: TsHorAlignment;
  fmt: String;
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
    if uffNumberFormat in FFormattingStyles[i].UsedFormattingFields then
      case FFormattingStyles[i].NumberFormat of
        nfFixed:
          case FFormattingStyles[i].NumberDecimals of
            0: lFormatIndex := FORMAT_FIXED_0_DECIMALS;
            2: lFormatIndex := FORMAT_FIXED_2_DECIMALS;
          end;
        nfFixedTh:
          case FFormattingStyles[i].NumberDecimals of
            0: lFormatIndex := FORMAT_FIXED_THOUSANDS_0_DECIMALS;
            2: lFormatIndex := FORMAT_FIXED_THOUSANDS_2_DECIMALS;
          end;
        nfExp:
          lFormatIndex := FORMAT_EXP_2_DECIMALS;
        nfSci:
          lFormatIndex := FORMAT_SCI_1_DECIMAL;
        nfPercentage:
          case FFormattingStyles[i].NumberDecimals of
            0: lFormatIndex := FORMAT_PERCENT_0_DECIMALS;
            2: lFormatIndex := FORMAT_PERCENT_2_DECIMALS;
          end;
        {
        nfCurrency:
          case FFormattingStyles[i].NumberDecimals of
            0: lFormatIndex := FORMAT_CURRENCY_0_DECIMALS;
            2: lFormatIndex := FORMAT_CURRENCY_2_DECIMALS;
          end;
        }
        nfShortDate:
          lFormatIndex := FORMAT_SHORT_DATE;
        nfShortTime:
          lFormatIndex := FORMAT_SHORT_TIME;
        nfLongTime:
          lFormatIndex := FORMAT_LONG_TIME;
        nfShortTimeAM:
          lFormatIndex := FORMAT_SHORT_TIME_AM;
        nfLongTimeAM:
          lFormatIndex := FORMAT_LONG_TIME_AM;
        nfShortDateTime:
          lFormatIndex := FORMAT_SHORT_DATETIME;
        nfFmtDateTime:
          begin
            fmt := lowercase(FFormattingStyles[i].NumberFormatStr);
            if (fmt = 'dm') or (fmt = 'd-mmm') or (fmt = 'd mmm') or (fmt = 'd. mmm') or (fmt = 'd/mmm') then
              lFormatIndex := FORMAT_DATE_DM
            else
            if (fmt = 'my') or (fmt = 'mmm-yy') or (fmt = 'mmm yy') or (fmt = 'mmm/yy') then
              lFormatIndex := FORMAT_DATE_MY
            else
            { Because of limitations of BIFF2 the next two formats are mapped
              to the same format index! }
            if (fmt = 'ms') or (fmt = 'nn:ss') or (fmt = 'mm:ss') then
              lFormatIndex := FORMAT_TIME_MS
            else
            if (fmt = 'msz') or (fmt = 'nn:ss.zzz') or (fmt = 'mm:ss.zzz') or (fmt = 'mm:ss.0') or (fmt = 'mm:ss.z') or (fmt = 'nn:ss.z') then
              lFormatIndex := FORMAT_TIME_MSZ;
          end;
        nfTimeInterval:
          lFormatIndex := FORMAT_TIME_INTERVAL;
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

procedure TsSpreadBIFF2Writer.WriteFormat(AStream: TStream; AFormatCode: String);
var
  len: Integer;
  s: AnsiString;
begin
  if AFormatCode = '' then
    exit;

  s := AFormatCode;
  len := Length(s);

  { BIFF record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMAT));
  AStream.WriteWord(WordToLE(len + 1));

  { Write format string }
  AStream.WriteByte(len);
  AStream.WriteBuffer(s[1], len);
end;

procedure TsSpreadBIFF2Writer.WriteFormatCount(AStream: TStream);
begin
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMATCOUNT));
  AStream.WriteWord(WordToLE(2));
  AStream.WriteWord(WordToLE(High(NUMFORMAT_BIFF2)+1));
end;

procedure TsSpreadBIFF2Writer.WriteFormats(AStream: TStream);
var
  i: Integer;
begin
  WriteFormatCount(AStream);
  for i:=0 to High(NUMFORMAT_BIFF2) do
    WriteFormat(AStream, NUMFORMAT_BIFF2[i]);
end;
(*
var
  ds, ts: Char;  //decimal separator, thousand separator
begin
  ds := DefaultFormatSettings.DecimalSeparator;
  ts := DefaultFormatSettings.ThousandSeparator;
  { 0} WriteFormat(AStream, 'General');  // 0
  { 1} WriteFormat(AStream, '0');
  { 2} WriteFormat(AStream, '0'+ds+'00');                           // 0.00
  { 3} WriteFormat(AStream, '#'+ts+'##0');                          // #,##0
  { 4} WriteFormat(AStream, '#'+ts+'##0'+ds+'00');                  // #,##0.00
  { 5} WriteFormat(AStream, '"$"#'+ts+'##0_);("$"#'+ts+'##0)');
  { 6} WriteFormat(AStream, '"$"#'+ts+'##0_);[Red]("$"#'+ts+'##0)');
  { 7} WriteFormat(AStream, '"$"#'+ts+'##0'+ds+'00_);("$"#'+ts+'##0'+ds+'00)');
  { 8} WriteFormat(AStream, '"$"#'+ts+'##0'+ds+'00_);[Red]("$"#'+ts+'##0'+ds+'00)');
  { 9} WriteFormat(AStream, '0%');
  {10} WriteFormat(AStream, '0'+ds+'00%');                          // 0.00%
  {11} WriteFormat(AStream, '0'+ds+'00E+00');                       // 0.00E+00
  {12} WriteFormat(AStream, 'm/d/yy');
  {13} WriteFormat(AStream, 'd-mmm-yy');
  {14} WriteFormat(AStream, 'd-mmm');
  {15} WriteFormat(AStream, 'mmm-yy');
  {16} WriteFormat(AStream, 'h:mm AM/PM');
  {17} WriteFormat(AStream, 'h:mm:ss AM/PM');
  {18} WriteFormat(AStream, 'h:mm');
  {19} WriteFormat(AStream, 'h:mm:ss');
  {20} WriteFormat(AStream, 'm/d/yy h:mm');

  { # TODO: locale support
     0 => 'GENERAL',
     1 => '0',
     2 => '0.00',
     3 => '#,##0',
     4 => '#,##0.00',
     5 => '"$"#,##0_);("$"#,##0)',
     6 => '"$"#,##0_);[Red]("$"#,##0)',
     7 => '"$"#,##0.00_);("$"#,##0.00)',
     8 => '"$"#,##0.00_);[Red]("$"#,##0.00)',
     9 => '0%',
    10 => '0.00%',
    11 => '0.00E+00',
    12 => '# ?/?',
    13 => '# ??/??',
    14 => 'M/D/YY',
    15 => 'D-MMM-YY',
    16 => 'D-MMM',
    17 => 'MMM-YY',
    18 => 'h:mm AM/PM',
    19 => 'h:mm:ss AM/PM',
    20 => 'h:mm',
    21 => 'h:mm:ss',
    22 => 'M/D/YY h:mm',
    37 => '_(#,##0_);(#,##0)',
    38 => '_(#,##0_);[Red](#,##0)',
    39 => '_(#,##0.00_);(#,##0.00)',
    40 => '_(#,##0.00_);[Red](#,##0.00)',
    41 => '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
    42 => '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    43 => '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    44 => '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    45 => 'mm:ss',
    46 => '[h]:mm:ss',
    47 => 'mm:ss.0',
    48 => '##0.0E+0',
    49 => '@',
  }
end;
  *)

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
procedure TsSpreadBIFF2Writer.WriteRPNFormula(AStream: TStream; const ARow,
  ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell);
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

  { Result of the formula in IEEE 754 floating-point value }
  AStream.WriteBuffer(FormulaResult, 8);

  { 0 = Do not recalculate
    1 = Always recalculate }
  AStream.WriteByte($1);

  { Formula }

  { The size of the token array is written later,
    because it's necessary to calculate if first,
    and this is done at the same time it is written }
  TokenArraySizePos := AStream.Position;
  AStream.WriteByte(RPNLength);

  { Formula data (RPN token array) }
  for i := 0 to Length(AFormula) - 1 do
  begin

    { Token identifier }
    FormulaKind := FormulaElementKindToExcelTokenID(AFormula[i].ElementKind, ExtraInfo);
    AStream.WriteByte(FormulaKind);
    Inc(RPNLength);

    { Additional data }
    case FormulaKind of

    { binary operation tokens }

    INT_EXCEL_TOKEN_TADD, INT_EXCEL_TOKEN_TSUB, INT_EXCEL_TOKEN_TMUL,
     INT_EXCEL_TOKEN_TDIV, INT_EXCEL_TOKEN_TPOWER: begin end;

    INT_EXCEL_TOKEN_TNUM:
    begin
      AStream.WriteBuffer(AFormula[i].DoubleValue, 8);
      Inc(RPNLength, 8);
    end;

    INT_EXCEL_TOKEN_TSTR:
    begin
      s := ansistring(AFormula[i].StringValue);
      len := Length(s);
      AStream.WriteByte(len);
      AStream.WriteBuffer(s[1], len);
      Inc(RPNLength, len + 1);
    end;

    INT_EXCEL_TOKEN_TBOOL:
    begin
      AStream.WriteByte(ord(AFormula[i].DoubleValue <> 0.0));
      inc(RPNLength, 1);
    end;

    INT_EXCEL_TOKEN_TREFR, INT_EXCEL_TOKEN_TREFV, INT_EXCEL_TOKEN_TREFA:
    begin
      r := AFormula[i].Row and MASK_EXCEL_ROW;
      if (rfRelRow in AFormula[i].RelFlags) then r := r or MASK_EXCEL_RELATIVE_ROW;
      if (rfRelCol in AFormula[i].RelFlags) then r := r or MASK_EXCEL_RELATIVE_COL;
      AStream.WriteWord(r);
      AStream.WriteByte(AFormula[i].Col);
      Inc(RPNLength, 3);
    end;

    INT_EXCEL_TOKEN_TAREA_R: { fekCellRange }
    begin
      r := AFormula[i].Row and MASK_EXCEL_ROW;
      if (rfRelRow in AFormula[i].RelFlags) then r := r or MASK_EXCEL_RELATIVE_ROW;
      if (rfRelCol in AFormula[i].RelFlags) then r := r or MASK_EXCEL_RELATIVE_COL;
      AStream.WriteWord(WordToLE(r));

      r := AFormula[i].Row2 and MASK_EXCEL_ROW;
      if (rfRelRow2 in AFormula[i].RelFlags) then r := r or MASK_EXCEL_RELATIVE_ROW;
      if (rfRelCol2 in AFormula[i].RelFlags) then r := r or MASK_EXCEL_RELATIVE_COL;
      AStream.WriteWord(WordToLE(r));

      AStream.WriteByte(AFormula[i].Col);
      AStream.WriteByte(AFormula[i].Col2);
      Inc(RPNLength, 6);
    end;

    INT_EXCEL_TOKEN_FUNC_R, INT_EXCEL_TOKEN_FUNC_V, INT_EXCEL_TOKEN_FUNC_A:
    begin
      AStream.WriteByte(Lo(ExtraInfo));
      Inc(RPNLength, 1);
    end;

    INT_EXCEL_TOKEN_FUNCVAR_V:
    begin
      AStream.WriteByte(AFormula[i].ParamsNum);
      AStream.WriteByte(Lo(ExtraInfo));
      // taking only the low-bytes, the high-bytes are needed for compatibility
      // with other BIFF formats...
      Inc(RPNLength, 2);
    end;

    end;
  end;

  { Write sizes in the end, after we known them }
  FinalPos := AStream.Position;
  AStream.position := TokenArraySizePos;
  AStream.WriteByte(RPNLength);
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(17 + RPNLength));
  AStream.position := FinalPos;
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
var
  xf: Word;
begin
  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BLANK));
  AStream.WriteWord(WordToLE(7));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  WriteCellFormatting(AStream, ACell, xf);
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
  MaxBytes=255; //limit for this format
var
  L: Byte;
  AnsiText: ansistring;
  TextTooLong: boolean=false;
var
  xf: Word;
begin
  if AValue = '' then Exit; // Writing an empty text doesn't work

  AnsiText := UTF8ToISO_8859_1(AValue);

  if Length(AnsiText)>MaxBytes then
  begin
    // BIFF 5 does not support labels/text bigger than 255 chars,
    // so BIFF2 won't either
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    TextTooLong:=true;
    AnsiText := Copy(AnsiText,1,MaxBytes);
  end;
  L := Length(AnsiText);

  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

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
begin
  xf := FindXFIndex(ACell);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

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
