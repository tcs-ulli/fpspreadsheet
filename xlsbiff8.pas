{
xlsbiff8.pas

Writes an Excel 8 file

An Excel worksheet stream consists of a number of subsequent records.
To ensure a properly formed file, the following order must be respected:

1st record:        BOF
2nd to Nth record: Any record
Last record:       EOF

Excel 8 files are OLE compound document files, and must be written using the
fpOLE library.

Records Needed to Make a BIFF8 File Microsoft Excel Can Use:

Required Records:

BOF - Set the 6 byte offset to 0x0005 (workbook globals)
Window1
FONT - At least five of these records must be included
XF - At least 15 Style XF records and 1 Cell XF record must be included
STYLE
BOUNDSHEET - Include one BOUNDSHEET record per worksheet
EOF

BOF - Set the 6 byte offset to 0x0010 (worksheet)
INDEX
DIMENSIONS
WINDOW2
EOF

The row and column numbering in BIFF files is zero-based.

Excel file format specification obtained from:
http://sc.openoffice.org/excelfileformat.pdf

see also:
http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP005199291.aspx

AUTHORS:  Felipe Monteiro de Carvalho
          Jose Mejuto
}
unit xlsbiff8;

{$ifdef fpc}
  {$mode delphi}
{$endif}

// The new OLE code is much better, so always use it
{$define USE_NEW_OLE}
{.$define FPSPREADDEBUG} //define to print out debug info to console. Used to be XLSDEBUG;

interface

uses
  Classes, SysUtils, fpcanvas, DateUtils,
  fpspreadsheet, xlscommon,
  {$ifdef USE_NEW_OLE}
  fpolebasic,
  {$else}
  fpolestorage,
  {$endif}
  fpsutils, lazutf8;

type

  { TsSpreadBIFF8Reader }
  TsSpreadBIFF8Reader = class(TsSpreadBIFFReader)
  private
    PendingRecordSize: SizeInt;
    FWorksheetNames: TStringList;
    FCurrentWorksheet: Integer;
    FSharedStringTable: TStringList;
    function ReadWideString(const AStream: TStream; const ALength: WORD): WideString; overload;
    function ReadWideString(const AStream: TStream; const AUse8BitLength: Boolean): WideString; overload;
    procedure ReadWorkbookGlobals(AStream: TStream; AData: TsWorkbook);
    procedure ReadWorksheet(AStream: TStream; AData: TsWorkbook);
    procedure ReadBoundsheet(AStream: TStream);
    function ReadString(const AStream: TStream; const ALength: WORD): UTF8String;
  protected
    procedure ReadFont(const AStream: TStream);
    procedure ReadFormat(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadLabelSST(const AStream: TStream);
    procedure ReadRichString(const AStream: TStream);
    procedure ReadRPNCellAddress(AStream: TStream; out ARow, ACol: Cardinal;
      out AFlags: TsRelFlags); override;
    procedure ReadRPNCellAddressOffset(AStream: TStream;
      out ARowOffset, AColOffset: Integer; out AFlags: TsRelFlags); override;
    procedure ReadRPNCellRangeAddress(AStream: TStream;
      out ARow1, ACol1, ARow2, ACol2: Cardinal; out AFlags: TsRelFlags); override;
    procedure ReadRPNCellRangeOffset(AStream: TStream;
      out ARow1Offset, ACol1Offset, ARow2Offset, ACol2Offset: Integer;
      out AFlags: TsRelFlags); override;
    procedure ReadSST(const AStream: TStream);
    function ReadString_8bitLen(AStream: TStream): String; override;
    procedure ReadStringRecord(AStream: TStream); override;
    procedure ReadXF(const AStream: TStream);
  public
    destructor Destroy; override;
    { General reading methods }
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); override;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); override;
  end;

  { TsSpreadBIFF8Writer }

  TsSpreadBIFF8Writer = class(TsSpreadBIFFWriter)
  private
    // Writes index to XF record according to cell's formatting
    procedure WriteXFFieldsForFormattingStyles(AStream: TStream);
  protected
    { Record writing methods }
    procedure WriteBOF(AStream: TStream; ADataType: Word);
    function  WriteBoundsheet(AStream: TStream; ASheetName: string): Int64;
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFont: TsFont);
    procedure WriteFonts(AStream: TStream);
    procedure WriteFormat(AStream: TStream; AFormatData: TsNumFormatData;
      AListIndex: Integer); override;
    procedure WriteIndex(AStream: TStream);
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    function WriteRPNCellAddress(AStream: TStream; ARow, ACol: Cardinal;
      AFlags: TsRelFlags): word; override;
    function WriteRPNCellOffset(AStream: TStream; ARowOffset, AColOffset: Integer;
      AFlags: TsRelFlags): Word; override;
    function WriteRPNCellRangeAddress(AStream: TStream; ARow1, ACol1, ARow2, ACol2: Cardinal;
      AFlags: TsRelFlags): Word; override;
    function WriteString_8bitLen(AStream: TStream; AString: String): Integer; override;
    procedure WriteStringRecord(AStream: TStream; AString: string); override;
    procedure WriteStyle(AStream: TStream);
    procedure WriteWindow2(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteXF(AStream: TStream; AFontIndex: Word;
      AFormatIndex: Word; AXF_TYPE_PROT, ATextRotation: Byte; ABorders: TsCellBorders;
      const ABorderStyles: TsCellBorderStyles; AHorAlignment: TsHorAlignment = haDefault;
      AVertAlignment: TsVertAlignment = vaDefault; AWordWrap: Boolean = false;
      AddBackground: Boolean = false; ABackgroundColor: TsColor = scSilver);
    procedure WriteXFRecords(AStream: TStream);
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;

var
  // the palette of the 64 default BIFF8 colors as "big-endian color" values
  PALETTE_BIFF8: array[$00..$3F] of TsColorValue = (
    $000000,  // $00: black            // 8 built-in default colors
    $FFFFFF,  // $01: white
    $FF0000,  // $02: red
    $00FF00,  // $03: green
    $0000FF,  // $04: blue
    $FFFF00,  // $05: yellow
    $FF00FF,  // $06: magenta
    $00FFFF,  // $07: cyan

    $000000,  // $08: EGA black
    $FFFFFF,  // $09: EGA white
    $FF0000,  // $0A: EGA red
    $00FF00,  // $0B: EGA green
    $0000FF,  // $0C: EGA blue
    $FFFF00,  // $0D: EGA yellow
    $FF00FF,  // $0E: EGA magenta
    $00FFFF,  // $0F: EGA cyan

    $800000,  // $10: EGA dark red
    $008000,  // $11: EGA dark green
    $000080,  // $12: EGA dark blue
    $808000,  // $13: EGA olive
    $800080,  // $14: EGA purple
    $008080,  // $15: EGA teal
    $C0C0C0,  // $16: EGA silver
    $808080,  // $17: EGA gray
    $9999FF,  // $18:
    $993366,  // $19:
    $FFFFCC,  // $1A:
    $CCFFFF,  // $1B:
    $660066,  // $1C:
    $FF8080,  // $1D:
    $0066CC,  // $1E:
    $CCCCFF,  // $1F:

    $000080,  // $20:
    $FF00FF,  // $21:
    $FFFF00,  // $22:
    $00FFFF,  // $23:
    $800080,  // $24:
    $800000,  // $25:
    $008080,  // $26:
    $0000FF,  // $27:
    $00CCFF,  // $28:
    $CCFFFF,  // $29:
    $CCFFCC,  // $2A:
    $FFFF99,  // $2B:
    $99CCFF,  // $2C:
    $FF99CC,  // $2D:
    $CC99FF,  // $2E:
    $FFCC99,  // $2F:

    $3366FF,  // $30:
    $33CCCC,  // $31:
    $99CC00,  // $32:
    $FFCC00,  // $33:
    $FF9900,  // $34:
    $FF6600,  // $35:
    $666699,  // $36:
    $969696,  // $37:
    $003366,  // $38:
    $339966,  // $39:
    $003300,  // $3A:
    $333300,  // $3B:
    $993300,  // $3C:
    $993366,  // $3D:
    $333399,  // $3E:
    $333333   // $3F:
  );

implementation

uses
  fpsStreams, fpsExprParser;

const
   { Excel record IDs }
     INT_EXCEL_ID_SST        = $00FC; //BIFF8 only
     INT_EXCEL_ID_LABELSST   = $00FD; //BIFF8 only
{%H-}INT_EXCEL_ID_FORCEFULLCALCULATION = $08A3;

   { Cell Addresses constants }
     MASK_EXCEL_COL_BITS_BIFF8     = $00FF;
     MASK_EXCEL_RELATIVE_COL_BIFF8 = $4000;  // This is according to Microsoft documentation,
     MASK_EXCEL_RELATIVE_ROW_BIFF8 = $8000;  // but opposite to OpenOffice documentation!

   { BOF record constants }
     INT_BOF_BIFF8_VER       = $0600;
     INT_BOF_WORKBOOK_GLOBALS= $0005;
{%H-}INT_BOF_VB_MODULE       = $0006;
     INT_BOF_SHEET           = $0010;
{%H-}INT_BOF_CHART           = $0020;
{%H-}INT_BOF_MACRO_SHEET     = $0040;
{%H-}INT_BOF_WORKSPACE       = $0100;
     INT_BOF_BUILD_ID        = $1FD2;
     INT_BOF_BUILD_YEAR      = $07CD;

   { STYLE record constants }
     MASK_STYLE_BUILT_IN     = $8000;

   { XF substructures }

   { XF_ROTATION }
     XF_ROTATION_HORIZONTAL              = 0;
     XF_ROTATION_90DEG_CCW               = 90;
     XF_ROTATION_90DEG_CW                = 180;
     XF_ROTATION_STACKED                 = 255;   // Letters stacked top to bottom, but not rotated

   { XF CELL BORDER LINE STYLES }
     MASK_XF_BORDER_LEFT                 = $0000000F;
     MASK_XF_BORDER_RIGHT                = $000000F0;
     MASK_XF_BORDER_TOP                  = $00000F00;
     MASK_XF_BORDER_BOTTOM               = $0000F000;
     MASK_XF_BORDER_DIAGONAL             = $01E00000;

     MASK_XF_BORDER_SHOW_DIAGONAL_DOWN   = $40000000;
     MASK_XF_BORDER_SHOW_DIAGONAL_UP     = $80000000;

   { XF CELL BORDER COLORS }
     MASK_XF_BORDER_LEFT_COLOR           = $007F0000;
     MASK_XF_BORDER_RIGHT_COLOR          = $3F800000;
     MASK_XF_BORDER_TOP_COLOR            = $0000007F;
     MASK_XF_BORDER_BOTTOM_COLOR         = $00003F80;
     MASK_XF_BORDER_DIAGONAL_COLOR       = $001FC000;

   { XF CELL BACKGROUND PATTERN }
     MASK_XF_BACKGROUND_PATTERN          = $FC000000;

     TEXT_ROTATIONS: Array[TsTextRotation] of Byte = (
       XF_ROTATION_HORIZONTAL,
       XF_ROTATION_90DEG_CW,
       XF_ROTATION_90DEG_CCW,
       XF_ROTATION_STACKED
     );

type
  TBIFF8DimensionsRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FirstRow: DWord;
    LastRowPlus1: DWord;
    FirstCol: Word;
    LastColPlus1: Word;
    NotUsed: Word;
  end;

  TBIFF8LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    TextLen: Word;
    TextFlags: Byte;
  end;

  TBIFF8LabelSSTRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    SSTIndex: DWord;
  end;


{ TsSpreadBIFF8Writer }

constructor TsSpreadBIFF8Writer.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
end;

procedure TsSpreadBIFF8Writer.WriteXFFieldsForFormattingStyles(AStream: TStream);
var
  i, j: Integer;
  lFontIndex: Word;
  lFormatIndex: Word; //number format
  lTextRotation: Byte;
  lBorders: TsCellBorders;
  lBorderStyles: TsCellBorderStyles;
  lAddBackground: Boolean;
  lBackgroundColor: TsColor;
  lHorAlign: TsHorAlignment;
  lVertAlign: TsVertAlignment;
  lWordWrap: Boolean;
begin
  // The first style was already added --> begin loop with 1
  for i := 1 to Length(FFormattingStyles) - 1 do begin
    // Default styles
    lFontIndex := 0;
    lFormatIndex := 0; //General format (one of the built-in number formats)
    lTextRotation := XF_ROTATION_HORIZONTAL;
    lBorders := [];
    lBorderStyles := FFormattingStyles[i].BorderStyles;
    lHorAlign := FFormattingStyles[i].HorAlignment;
    lVertAlign := FFormattingStyles[i].VertAlignment;
    lBackgroundColor := FFormattingStyles[i].BackgroundColor;

    // Now apply the modifications.
    if uffNumberFormat in FFormattingStyles[i].UsedFormattingFields then begin
      // The number formats in the FormattingStyles are still in fpc dialect
      // They will be converted to Excel syntax immediately before writing.
      j := NumFormatList.FindFormatOf(@FFormattingStyles[i]);
      if j > -1 then
        lFormatIndex := NumFormatList[j].Index;
    end;

    if uffBorder in FFormattingStyles[i].UsedFormattingFields then
      lBorders := FFormattingStyles[i].Border;

    if uffTextRotation in FFormattingStyles[i].UsedFormattingFields then
      lTextRotation := TEXT_ROTATIONS[FFormattingStyles[i].TextRotation];

    if uffBold in FFormattingStyles[i].UsedFormattingFields then
      lFontIndex := 1;   // must be before uffFont which overrides uffBold

    if uffFont in FFormattingStyles[i].UsedFormattingFields then
      lFontIndex := FFormattingStyles[i].FontIndex;

    lAddBackground := (uffBackgroundColor in FFormattingStyles[i].UsedFormattingFields);
    lWordwrap := (uffWordwrap in FFormattingStyles[i].UsedFormattingFields);

    // And finally write the style
    WriteXF(AStream, lFontIndex, lFormatIndex, 0, lTextRotation, lBorders,
      lBorderStyles, lHorAlign, lVertAlign, lWordwrap, lAddBackground,
      lBackgroundColor);
  end;
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteToFile ()
*
*  DESCRIPTION:    Writes an Excel BIFF8 file to the disc
*
*                  The BIFF 8 writer overrides this method because
*                  BIFF 8 is written as an OLE document, and our
*                  current OLE document writing method involves:
*
*                  1 - Writing the BIFF data to a memory stream
*
*                  2 - Write the memory stream data to disk using
*                      COM functions
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean);
var
  Stream: TStream;
  OutputStorage: TOLEStorage;
  OLEDocument: TOLEDocument;
begin
  if (boBufStream in Workbook.Options) then begin
    Stream := TBufStream.Create
  end else
    Stream := TMemoryStream.Create;

  OutputStorage := TOLEStorage.Create;
  try
    WriteToStream(Stream);

    // Only one stream is necessary for any number of worksheets
    OLEDocument.Stream := Stream;

    OutputStorage.WriteOLEFile(AFileName, OLEDocument, AOverwriteExisting, 'Workbook');
  finally
    Stream.Free;
    OutputStorage.Free;
  end;
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteToStream ()
*
*  DESCRIPTION:    Writes an Excel BIFF8 record structure
*
*                  Be careful as this method doesn't write the OLE
*                  part of the document, just the BIFF records
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteToStream(AStream: TStream);
const
  isBIFF8 = true;
var
  CurrentPos: Int64;
  Boundsheets: array of Int64;
  i, len: Integer;
  pane: Byte;
begin
  { Write workbook globals }

  WriteBOF(AStream, INT_BOF_WORKBOOK_GLOBALS);

  WriteWindow1(AStream);
  WriteFonts(AStream);
  WriteFormats(AStream);
  WritePalette(AStream);
  WriteXFRecords(AStream);
  WriteStyle(AStream);

  // A BOUNDSHEET for each worksheet
  SetLength(Boundsheets, 0);
  for i := 0 to Workbook.GetWorksheetCount - 1 do
  begin
    len := Length(Boundsheets);
    SetLength(Boundsheets, len + 1);
    Boundsheets[len] := WriteBoundsheet(AStream, Workbook.GetWorksheetByIndex(i).Name);
  end;

  WriteEOF(AStream);

  { Write each worksheet }

  for i := 0 to Workbook.GetWorksheetCount - 1 do
  begin
    FWorksheet := Workbook.GetWorksheetByIndex(i);

    { First goes back and writes the position of the BOF of the
      sheet on the respective BOUNDSHEET record }
    CurrentPos := AStream.Position;
    AStream.Position := Boundsheets[i];
    AStream.WriteDWord(DWordToLE(DWORD(CurrentPos)));
    AStream.Position := CurrentPos;

    WriteBOF(AStream, INT_BOF_SHEET);
      WriteIndex(AStream);
      //WriteSheetPR(AStream);
//      WritePageSetup(AStream);
      WriteColInfos(AStream, FWorksheet);
      WriteDimensions(AStream, FWorksheet);
      //WriteRowAndCellBlock(AStream, sheet);

      if (boVirtualMode in Workbook.Options) then
        WriteVirtualCells(AStream)
      else begin
        WriteRows(AStream, FWorksheet);
        WriteCellsToStream(AStream, FWorksheet.Cells);
      end;

      WriteWindow2(AStream, FWorksheet);
      WritePane(AStream, FWorksheet, isBIFF8, pane);
      WriteSelection(AStream, FWorksheet, pane);
    WriteEOF(AStream);
  end;
  
  { Cleanup }
  SetLength(Boundsheets, 0);
end;


{*******************************************************************
*  TsSpreadBIFF8Writer.WriteBOF ()
*
*  DESCRIPTION:    Writes an Excel 8 BOF record
*
*                  This must be the first record on an Excel 8 stream
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteBOF(AStream: TStream; ADataType: Word);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BOF));
  AStream.WriteWord(WordToLE(16)); //total record size

  { BIFF version. Should only be used if this BOF is for the workbook globals }
  { OpenOffice rejects to correctly read xls files if this field is
    omitted as docs. says, or even if it is being written to zero value,
    Not tested with Excel, but MSExcel reader opens it as expected }
  AStream.WriteWord(WordToLE(INT_BOF_BIFF8_VER));

  { Data type }
  AStream.WriteWord(WordToLE(ADataType));

  { Build identifier, must not be 0 }
  AStream.WriteWord(WordToLE(INT_BOF_BUILD_ID));

  { Build year, must not be 0 }
  AStream.WriteWord(WordToLE(INT_BOF_BUILD_YEAR));

  { File history flags }
  AStream.WriteDWord(DWordToLE(0));

  { Lowest Excel version that can read all records in this file 5?}
  AStream.WriteDWord(DWordToLE(0)); //?????????
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteBoundsheet ()
*
*  DESCRIPTION:    Writes an Excel 8 BOUNDSHEET record
*
*                  Always located on the workbook globals substream.
*
*                  One BOUNDSHEET is written for each worksheet.
*
*  RETURNS:        The stream position where the absolute stream position
*                  of the BOF of this sheet should be written (4 bytes size).
*
*******************************************************************}
function TsSpreadBIFF8Writer.WriteBoundsheet(AStream: TStream; ASheetName: string): Int64;
var
  Len: Byte;
  WideSheetName: WideString;
begin
  WideSheetName:=UTF8Decode(ASheetName);
  Len := Length(WideSheetName);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BOUNDSHEET));
  AStream.WriteWord(WordToLE(6 + 1 + 1 + Len * Sizeof(WideChar)));

  { Absolute stream position of the BOF record of the sheet represented
    by this record }
  Result := AStream.Position;
  AStream.WriteDWord(DWordToLE(0));

  { Visibility }
  AStream.WriteByte(0);

  { Sheet type }
  AStream.WriteByte(0);

  { Sheet name: Unicode string char count 1 byte }
  AStream.WriteByte(Len);
  {String flags}
  AStream.WriteByte(1);
  AStream.WriteBuffer(WideStringToLE(WideSheetName)[1], Len * Sizeof(WideChar));
end;


{
  Writes an Excel 8 DIMENSIONS record

  nm = (rl - rf - 1) / 32 + 1 (using integer division)

  Excel, OpenOffice and FPSpreadsheet ignore the dimensions written in this record,
  but some other applications really use them, so they need to be correct.

  See bug 18886: excel5 files are truncated when imported
}
procedure TsSpreadBIFF8Writer.WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
var
  firstRow, lastRow, firstCol, lastCol: Cardinal;
  rec: TBIFF8DimensionsRecord;
begin
  { Determine sheet size }
  GetSheetDimensions(AWorksheet, firstRow, lastRow, firstCol, lastCol);

  { Populate BIFF record }
  rec.RecordID := WordToLE(INT_EXCEL_ID_DIMENSIONS);
  rec.RecordSize := WordToLE(14);
  rec.FirstRow := DWordToLE(firstRow);
  rec.LastRowPlus1 := DWordToLE(lastRow+1);
  rec.FirstCol := WordToLE(firstCol);
  rec.LastColPlus1 := WordToLE(lastCol+1);
  rec.NotUsed := 0;

  { Write BIFF record to stream }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteEOF ()
*
*  DESCRIPTION:    Writes an Excel 8 EOF record
*
*                  This must be the last record on an Excel 8 stream
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_EOF));
  AStream.WriteWord(WordToLE($0000));
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteFont ()
*
*  DESCRIPTION:    Writes an Excel 8 FONT record
*
*                  The font data is passed in an instance of TsFont
*
*******************************************************************}

procedure TsSpreadBIFF8Writer.WriteFont(AStream: TStream; AFont: TsFont);
var
  Len: Byte;
  WideFontName: WideString;
  optn: Word;
begin
  if AFont = nil then  // this happens for FONT4 in case of BIFF
    exit;

  if AFont.FontName = '' then
    raise Exception.Create('Font name not specified.');
  if AFont.Size <= 0.0 then
    raise Exception.Create('Font size not specified.');

  WideFontName := AFont.FontName;
  Len := Length(WideFontName);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FONT));
  AStream.WriteWord(WordToLE(14 + 1 + 1 + Len * Sizeof(WideChar)));

  { Height of the font in twips = 1/20 of a point }
  AStream.WriteWord(WordToLE(round(AFont.Size*20)));

  { Option flags }
  optn := 0;
  if fssBold in AFont.Style then optn := optn or $0001;
  if fssItalic in AFont.Style then optn := optn or $0002;
  if fssUnderline in AFont.Style then optn := optn or $0004;
  if fssStrikeout in AFont.Style then optn := optn or $0008;
  AStream.WriteWord(WordToLE(optn));

  { Colour index }
  AStream.WriteWord(WordToLE(ord(FixColor(AFont.Color))));

  { Font weight }
  if fssBold in AFont.Style then
    AStream.WriteWord(WordToLE(INT_FONT_WEIGHT_BOLD))
  else
    AStream.WriteWord(WordToLE(INT_FONT_WEIGHT_NORMAL));

  { Escapement type }
  AStream.WriteWord(WordToLE(0));

  { Underline type }
  if fssUnderline in AFont.Style then
    AStream.WriteByte(1)
  else
    AStream.WriteByte(0);

  { Font family }
  AStream.WriteByte(0);

  { Character set }
  AStream.WriteByte(0);

  { Not used }
  AStream.WriteByte(0);

  { Font name: Unicodestring, char count in 1 byte }
  AStream.WriteByte(Len);
  { Widestring flags, 1=regular unicode LE string }
  AStream.WriteByte(1);
  AStream.WriteBuffer(WideStringToLE(WideFontName)[1], Len * Sizeof(WideChar));
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteFonts ()
*
*  DESCRIPTION:    Writes the Excel 8 FONT records needed for the
*                  used fonts in the workbook.
*
*******************************************************************}
procedure TsSpreadBiff8Writer.WriteFonts(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to Workbook.GetFontCount-1 do
    WriteFont(AStream, Workbook.GetFont(i));
end;

procedure TsSpreadBiff8Writer.WriteFormat(AStream: TStream;
  AFormatData: TsNumFormatData; AListIndex: Integer);
type
  TNumFormatRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FormatIndex: Word;
    FormatStringLen: Word;
    FormatStringFlags: Byte;
  end;
var
  len: Integer;
  s: widestring;
  rec: TNumFormatRecord;
  buf: array of byte;
begin
  if (AFormatData = nil) or (AFormatData.FormatString = '') then
    exit;

  s := NumFormatList.FormatStringForWriting(AListIndex);
  len := Length(s);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_FORMAT);
  rec.RecordSize := WordToLE(2 + 2 + 1 + len * SizeOf(WideChar));

  { Format index }
  rec.FormatIndex := WordToLE(AFormatData.Index);

  { Format string }
  { - length of string = 16 bits }
  rec.FormatStringLen := WordToLE(len);
  { - Widestring flags, 1 = regular unicode LE string }
  rec.FormatStringFlags := 1;
  { - Copy the text characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + SizeOf(WideChar)*len);
  Move(rec, buf[0], SizeOf(rec));
  Move(s[1], buf[SizeOf(rec)], len*SizeOf(WideChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(rec) + SizeOf(WideChar)*len);

  { Clean up }
  SetLength(buf, 0);

(*
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMAT));
  AStream.WriteWord(WordToLE(2 + 2 + 1 + len * SizeOf(WideChar)));

  { Format index }
  AStream.WriteWord(WordToLE(AFormatData.Index));

  { Format string }
  { - Unicodestring, char count in 2 bytes  }
  AStream.WriteWord(WordToLE(len));
  { - Widestring flags, 1=regular unicode LE string }
  AStream.WriteByte(1);
  { - String data }
  AStream.WriteBuffer(WideStringToLE(s)[1], len * Sizeof(WideChar));
  *)
end;

{ Writes the address of a cell as used in an RPN formula and returns the
  number of bytes written. }
function TsSpreadBIFF8Writer.WriteRPNCellAddress(AStream: TStream;
  ARow, ACol: Cardinal; AFlags: TsRelFlags): Word;
var
  c: Cardinal;       // column index with encoded relative/absolute address info
begin
  AStream.WriteWord(WordToLE(ARow));
  c := ACol and MASK_EXCEL_COL_BITS_BIFF8;
  if (rfRelRow in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW_BIFF8;
  if (rfRelCol in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL_BIFF8;
  AStream.WriteWord(WordToLE(c));
  Result := 4;
end;

{ Writes row and column offset (unsigned integers!)
  Valid for BIFF2-BIFF5. }
function TsSpreadBIFF8Writer.WriteRPNCellOffset(AStream: TStream;
  ARowOffset, AColOffset: Integer; AFlags: TsRelFlags): Word;
var
  c: Word;
  r: SmallInt;
begin
  // row address
  r := SmallInt(ARowOffset);
  AStream.WriteWord(WordToLE(Word(r)));

  // Encoded column address
  c := word(AColOffset) and MASK_EXCEL_COL_BITS_BIFF8;
  if (rfRelRow in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW_BIFF8;
  if (rfRelCol in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL_BIFF8;
  AStream.WriteWord(WordToLE(c));

  Result := 4;
end;

{ Writes the address of a cell range as used in an RPN formula and returns the
  count of bytes written. }
function TsSpreadBIFF8Writer.WriteRPNCellRangeAddress(AStream: TStream;
  ARow1, ACol1, ARow2, ACol2: Cardinal; AFlags: TsRelFlags): Word;
var
  c: Cardinal;       // column index with encoded relative/absolute address info
begin
  AStream.WriteWord(WordToLE(ARow1));
  AStream.WriteWord(WordToLE(ARow2));

  c := ACol1;
  if (rfRelCol in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL;
  if (rfRelRow in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW;
  AStream.WriteWord(WordToLE(c));

  c := ACol2;
  if (rfRelCol2 in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL;
  if (rfRelRow2 in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW;
  AStream.WriteWord(WordToLE(c));

  Result := 8;
end;
                 (*
{ Writes the borders of the cell range covered by a shared formula.
  Needs to be overridden to write the column data (2 bytes in case of BIFF8). }
procedure TsSpreadBIFF8Writer.WriteSharedFormulaRange(AStream: TStream;
  const ARange: TRect);
begin
  inherited WriteSharedFormulaRange(AStream, ARange);
  {
  // Index to first column
  AStream.WriteWord(WordToLE(ARange.Left));
  // Index to last rcolumn
  AStream.WriteWord(WordToLE(ARange.Right));
  }
end;
*)

{ Helper function for writing a string with 8-bit length. Overridden version
  for BIFF8. Called for writing rpn formula string tokens.
  Returns the count of bytes written}
function TsSpreadBIFF8Writer.WriteString_8BitLen(AStream: TStream;
  AString: String): Integer;
var
  len: Integer;
  wideStr: WideString;
begin
  // string constant is stored as widestring in BIFF8
  wideStr := UTF8Decode(AString);
  len := Length(wideStr);
  AStream.WriteByte(len); // char count in 1 byte
  AStream.WriteByte(1);   // Widestring flags, 1=regular unicode LE string
  AStream.WriteBuffer(WideStringToLE(wideStr)[1], len * Sizeof(WideChar));
  Result := 1 + 1 + len * SizeOf(WideChar);
end;

procedure TsSpreadBIFF8Writer.WriteStringRecord(AStream: TStream;
  AString: String);
var
  wideStr: widestring;
  len: Integer;
begin
  wideStr := UTF8Decode(AString);
  len := Length(wideStr);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_STRING));
  AStream.WriteWord(WordToLE(3 + len*SizeOf(widechar)));

  { Write widestring length }
  AStream.WriteWord(WordToLE(len));
  { Widestring flags, 1=regular unicode LE string }
  AStream.WriteByte(1);
  { Write characters }
  AStream.WriteBuffer(WideStringToLE(wideStr)[1], len * SizeOf(WideChar));
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteIndex ()
*
*  DESCRIPTION:    Writes an Excel 8 INDEX record
*
*                  nm = (rl - rf - 1) / 32 + 1 (using integer division)
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteIndex(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_INDEX));
  AStream.WriteWord(WordToLE(16));

  { Not used }
  AStream.WriteDWord(DWordToLE(0));

  { Index to first used row, rf, 0 based }
  AStream.WriteDWord(DWordToLE(0));

  { Index to first row of unused tail of sheet, rl, last used row + 1, 0 based }
  AStream.WriteDWord(DWordToLE(0));

  { Absolute stream position of the DEFCOLWIDTH record of the current sheet.
    If it doesn't exist, the offset points to where it would occur. }
  AStream.WriteDWord(DWordToLE($00));

  { Array of nm absolute stream positions of the DBCELL record of each Row Block }
  
  { OBS: It seems to be no problem just ignoring this part of the record }
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteLabel ()
*
*  DESCRIPTION:    Writes an Excel 8 LABEL record
*
*                  Writes a string to the sheet
*                  If the string length exceeds 32758 bytes, the string
*                  will be silently truncated.
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  //limit for this format: 32767 bytes - header (see reclen below):
  //37267-8-1=32758
  MAXBYTES = 32758;
var
  L: Word;
  WideValue: WideString;
  rec: TBIFF8LabelRecord;
  buf: array of byte;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  WideValue := UTF8Decode(AValue); //to UTF16
  if WideValue = '' then begin
    // Badly formatted UTF8String (maybe ANSI?)
    if Length(AValue)<>0 then begin
      //Quite sure it was an ANSI string written as UTF8, so raise exception.
      Raise Exception.CreateFmt('Expected UTF8 text but probably ANSI text found in cell [%d,%d]',[ARow,ACol]);
    end;
    Exit;
  end;

  if Length(WideValue) > MAXBYTES then begin
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    SetLength(WideValue, MAXBYTES); //may corrupt the string (e.g. in surrogate pairs), but... too bad.
    Workbook.AddErrorMsg(
      'Text value exceeds %d character limit in cell %s. ' +
      'Text has been truncated.', [
      MAXBYTES, GetCellString(ARow, ACol)
    ]);
  end;
  L := Length(WideValue);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_LABEL);
  rec.RecordSize := 8 + 1 + L * SizeOf(WideChar);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { Index to XF record, according to formatting }
  rec.XFIndex := WordToLE(FindXFIndex(ACell));

  { Byte String with 16-bit length }
  rec.TextLen := WordToLE(L);

  { Byte flags, 1 means regular unicode LE encoding }
  rec.TextFlags := 1;

  { Copy the text characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + L*SizeOf(WideChar));
  Move(rec, buf[0], SizeOf(Rec));
  Move(WideStringToLE(WideValue)[1], buf[SizeOf(Rec)], L*SizeOf(WideChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(rec) + L*SizeOf(WideChar));

  { Clean up }
  SetLength(buf, 0);
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteStyle ()
*
*  DESCRIPTION:    Writes an Excel 8 STYLE record
*
*                  Registers the name of a user-defined style or
*                  specific options for a built-in cell style.
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteStyle(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_STYLE));
  AStream.WriteWord(WordToLE(4));

  { Index to style XF and defines if it's a built-in or used defined style }
  AStream.WriteWord(WordToLE(MASK_STYLE_BUILT_IN));

  { Built-in cell style identifier }
  AStream.WriteByte($00);

  { Level if the identifier for a built-in style is RowLevel or ColLevel, $FF otherwise }
  AStream.WriteByte($FF);
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteWindow2 ()
*
*  DESCRIPTION:    Writes an Excel 8 WINDOW2 record
*
*                  This record contains aditional settings for the
*                  document window (BIFF2-BIFF4) or for a specific
*                  worksheet (BIFF5-BIFF8).
*
*                  The values written here are reasonable defaults,
*                  which should work for most sheets.
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteWindow2(AStream: TStream;
 ASheet: TsWorksheet);
var
  Options: Word;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW2));
  AStream.WriteWord(WordToLE(18));

  { Options flags }
  Options :=
    MASK_WINDOW2_OPTION_SHOW_ZERO_VALUES or
    MASK_WINDOW2_OPTION_AUTO_GRIDLINE_COLOR or
    MASK_WINDOW2_OPTION_SHOW_OUTLINE_SYMBOLS or
    MASK_WINDOW2_OPTION_SHEET_SELECTED or
    MASK_WINDOW2_OPTION_SHEET_ACTIVE;
   { Bug 0026386 -> every sheet must be selected/active, otherwise Excel cannot print }

  if (soShowGridLines in ASheet.Options) then
    Options := Options or MASK_WINDOW2_OPTION_SHOW_GRID_LINES;
  if (soShowHeaders in ASheet.Options) then
    Options := Options or MASK_WINDOW2_OPTION_SHOW_SHEET_HEADERS;
  if (soHasFrozenPanes in ASheet.Options) and ((ASheet.LeftPaneWidth > 0) or (ASheet.TopPaneHeight > 0)) then
    Options := Options or MASK_WINDOW2_OPTION_PANES_ARE_FROZEN;
  AStream.WriteWord(WordToLE(Options));

  { Index to first visible row }
  AStream.WriteWord(WordToLE(0));

  { Index to first visible column }
  AStream.WriteWord(WordToLE(0));

  { Grid line index colour }
  AStream.WriteWord(WordToLE(0));

  { Not used }
  AStream.WriteWord(WordToLE(0));

  { Cached magnification factor in page break preview (in percent); 0 = Default (60%) }
  AStream.WriteWord(WordToLE(0));

  { Cached magnification factor in normal view (in percent); 0 = Default (100%) }
  AStream.WriteWord(WordToLE(0));

  { Not used }
  AStream.WriteDWord(DWordToLE(0));
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteXF ()
*
*  DESCRIPTION:    Writes an Excel 8 XF record
*
*
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteXF(AStream: TStream; AFontIndex: Word;
 AFormatIndex: Word; AXF_TYPE_PROT, ATextRotation: Byte; ABorders: TsCellBorders;
 const ABorderStyles: TsCellBorderStyles; AHorAlignment: TsHorAlignment = haDefault;
 AVertAlignment: TsVertAlignment = vaDefault; AWordWrap: Boolean = false;
 AddBackground: Boolean = false; ABackgroundColor: TsColor = scSilver);
var
  XFOptions: Word;
  XFAlignment, XFOrientationAttrib: Byte;
  XFBorderDWord1, XFBorderDWord2: DWord;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_XF));
  AStream.WriteWord(WordToLE(20));

  { Index to FONT record }
  AStream.WriteWord(WordToLE(AFontIndex));

  { Index to FORMAT record }
  AStream.WriteWord(WordToLE(AFormatIndex));

  { XF type, cell protection and parent style XF }
  XFOptions := AXF_TYPE_PROT and MASK_XF_TYPE_PROT;

  if AXF_TYPE_PROT and MASK_XF_TYPE_PROT_STYLE_XF <> 0 then
   XFOptions := XFOptions or MASK_XF_TYPE_PROT_PARENT;
   
  AStream.WriteWord(WordToLE(XFOptions));

  { Alignment and text break }
  XFAlignment := 0;
  case AHorAlignment of
    haLeft   : XFAlignment := XFAlignment or MASK_XF_HOR_ALIGN_LEFT;
    haCenter : XFAlignment := XFAlignment or MASK_XF_HOR_ALIGN_CENTER;
    haRight  : XFAlignment := XFAlignment or MASK_XF_HOR_ALIGN_RIGHT;
  end;
  case AVertAlignment of
    vaTop    : XFAlignment := XFAlignment or MASK_XF_VERT_ALIGN_TOP;
    vaCenter : XFAlignment := XFAlignment or MASK_XF_VERT_ALIGN_CENTER;
    vaBottom : XFAlignment := XFAlignment or MASK_XF_VERT_ALIGN_BOTTOM;
    else       XFAlignment := XFAlignment or MASK_XF_VERT_ALIGN_BOTTOM;
  end;
  if AWordWrap then
    XFAlignment := XFAlignment or MASK_XF_TEXTWRAP;

  AStream.WriteByte(XFAlignment);

  { Text rotation }
  AStream.WriteByte(ATextRotation); // 0 is horizontal / normal

  { Indentation, shrink and text direction }
  AStream.WriteByte(0);

  { Used attributes }
  XFOrientationAttrib :=
   MASK_XF_USED_ATTRIB_NUMBER_FORMAT or
   MASK_XF_USED_ATTRIB_FONT or
   MASK_XF_USED_ATTRIB_TEXT or
   MASK_XF_USED_ATTRIB_BORDER_LINES or
   MASK_XF_USED_ATTRIB_BACKGROUND or
   MASK_XF_USED_ATTRIB_CELL_PROTECTION;

  AStream.WriteByte(XFOrientationAttrib);

  { Cell border lines and background area }

  // Left and Right line colors
  XFBorderDWord1 := ABorderStyles[cbWest].Color shl 16 +
                    ABorderStyles[cbEast].Color shl 23;

  // Border line styles
  if cbWest in ABorders then
    XFBorderDWord1 := XFBorderDWord1 or (DWord(ABorderStyles[cbWest].LineStyle)+1);
  if cbEast in ABorders then
    XFBorderDWord1 := XFBorderDWord1 or ((DWord(ABorderStyles[cbEast].LineStyle)+1) shl 4);
  if cbNorth in ABorders then
    XFBorderDWord1 := XFBorderDWord1 or ((DWord(ABorderStyles[cbNorth].LineStyle)+1) shl 8);
  if cbSouth in ABorders then
    XFBorderDWord1 := XFBorderDWord1 or ((DWord(ABorderStyles[cbSouth].LineStyle)+1) shl 12);
  if cbDiagDown in ABorders then
    XFBorderDWord1 := XFBorderDWord1 or $40000000;
  if cbDiagUp in ABorders then
    XFBorderDWord1 := XFBorderDWord1 or $80000000;
  AStream.WriteDWord(DWordToLE(XFBorderDWord1));

  // Top, bottom and diagonal line colors
  XFBorderDWord2 := ABorderStyles[cbNorth].Color + ABorderStyles[cbSouth].Color shl 7 +
    ABorderStyles[cbDiagUp].Color shl 14;
    // In BIFF8 both diagonals have the same color - we use the color of the up-diagonal.

  // Diagonal line style
  if (ABorders + [cbDiagUp, cbDiagDown] <> []) then
    XFBorderDWord2 := XFBorderDWord2 or ((DWord(ABorderStyles[cbDiagUp].LineStyle)+1) shl 21);
    // In BIFF8 both diagonals have the same color - we use the color of the up-diagonal.

  // Add a background, if desired
  if AddBackground then XFBorderDWord2 := XFBorderDWord2 or $4000000;
  AStream.WriteDWord(DWordToLE(XFBorderDWord2));

  // Background Pattern Color, always zeroed
  if AddBackground then
    AStream.WriteWord(WordToLE(ABackgroundColor))
  else
    AStream.WriteWord(0);
end;

procedure TsSpreadBIFF8Writer.WriteXFRecords(AStream: TStream);
begin
  // XF0
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF1
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF2
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF3
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF4
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF5
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF6
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF7
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF8
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF9
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF10
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF11
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF12
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF13
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF14
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);
  // XF15 - Default, no formatting
  WriteXF(AStream, 0, 0, 0, XF_ROTATION_HORIZONTAL, [], DEFAULT_BORDERSTYLES);

  // Add all further non-standard/built-in formatting styles
  ListAllFormattingStyles;
  WriteXFFieldsForFormattingStyles(AStream);
end;


{ TsSpreadBIFF8Reader }

destructor TsSpreadBIFF8Reader.Destroy;
begin
  if Assigned(FSharedStringTable) then FSharedStringTable.Free;
  inherited;
end;


var


    counter: Integer = 0;



function TsSpreadBIFF8Reader.ReadWideString(const AStream: TStream;
  const ALength: WORD): WideString;
var
  StringFlags: BYTE;
  DecomprStrValue: WideString;
  AnsiStrValue: ansistring;
  RunsCounter: WORD;
  AsianPhoneticBytes: DWORD;
  i: Integer;
  j: SizeUInt;
  lLen: SizeInt;
  RecordType: WORD;
  RecordSize: WORD;
  C: char;
begin
  StringFlags:=AStream.ReadByte;
  Dec(PendingRecordSize);
  if StringFlags and 4 = 4 then begin
    //Asian phonetics
    //Read Asian phonetics Length (not used)
    AsianPhoneticBytes:=DWordLEtoN(AStream.ReadDWord);
  end;
  if StringFlags and 8 = 8 then begin
    //Rich string
    RunsCounter:=WordLEtoN(AStream.ReadWord);
    dec(PendingRecordSize,2);
  end;
  if StringFlags and 1 = 1 Then begin
    //String is WideStringLE
    if (ALength*SizeOf(WideChar)) > PendingRecordSize then begin
      SetLength(Result,PendingRecordSize div 2);
      AStream.ReadBuffer(Result[1],PendingRecordSize);
      Dec(PendingRecordSize,PendingRecordSize);
    end else begin
      SetLength(Result, ALength);
      AStream.ReadBuffer(Result[1],ALength * SizeOf(WideChar));
      Dec(PendingRecordSize,ALength * SizeOf(WideChar));
    end;
    Result := WideStringLEToN(Result);
  end else begin
    //String is 1 byte per char, this is UTF-16 with the high byte ommited because it is zero
    //so decompress and then convert


    inc(Counter);


    lLen:=ALength;
    SetLength(DecomprStrValue, lLen);
    for i := 1 to lLen do
    begin
      C:=WideChar(AStream.ReadByte());
      DecomprStrValue[i] := C;
      Dec(PendingRecordSize);
      if (PendingRecordSize<=0) and (i<lLen) then begin
        //A CONTINUE may happend here
        RecordType := WordLEToN(AStream.ReadWord);
        RecordSize := WordLEToN(AStream.ReadWord);
        if RecordType<>INT_EXCEL_ID_CONTINUE then begin
          Raise Exception.Create('[TsSpreadBIFF8Reader.ReadWideString] Expected CONTINUE record not found.');
        end else begin
          PendingRecordSize:=RecordSize;
          DecomprStrValue:=copy(DecomprStrValue,1,i)+ReadWideString(AStream,ALength-i);
          break;
        end;
      end;
    end;

    Result := DecomprStrValue;
  end;
  if StringFlags and 8 = 8 then begin
    //Rich string (This only happened in BIFF8)
    for j := 1 to RunsCounter do begin
      if (PendingRecordSize<=0) then begin
        //A CONTINUE may happend here
        RecordType := WordLEToN(AStream.ReadWord);
        RecordSize := WordLEToN(AStream.ReadWord);
        if RecordType<>INT_EXCEL_ID_CONTINUE then begin
          Raise Exception.Create('[TsSpreadBIFF8Reader.ReadWideString] Expected CONTINUE record not found.');
        end else begin
          PendingRecordSize:=RecordSize;
        end;
      end;
      AStream.ReadWord;
      AStream.ReadWord;
      dec(PendingRecordSize,2*2);
    end;
  end;
  if StringFlags and 4 = 4 then begin
    //Asian phonetics
    //Read Asian phonetics, discarded as not used.
    SetLength(AnsiStrValue,AsianPhoneticBytes);
    AStream.ReadBuffer(AnsiStrValue[1],AsianPhoneticBytes);
  end;
end;

function TsSpreadBIFF8Reader.ReadWideString(const AStream: TStream;
  const AUse8BitLength: Boolean): WideString;
var
  Len: Word;
begin
  if AUse8BitLength then
    Len := AStream.ReadByte()
  else
    Len := WordLEtoN(AStream.ReadWord());

  Result := ReadWideString(AStream, Len);
end;

procedure TsSpreadBIFF8Reader.ReadWorkbookGlobals(AStream: TStream;
  AData: TsWorkbook);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  Unused(AData);
  // Clear existing fonts. They will be replaced by those from the file.
  FWorkbook.RemoveAllFonts;
  if Assigned(FSharedStringTable) then FreeAndNil(FSharedStringTable);

  while (not SectionEOF) do begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);
    PendingRecordSize := RecordSize;

    CurStreamPos := AStream.Position;

    if RecordType <> INT_EXCEL_ID_CONTINUE then begin
      case RecordType of
       INT_EXCEL_ID_BOF       : ;
       INT_EXCEL_ID_BOUNDSHEET: ReadBoundSheet(AStream);
       INT_EXCEL_ID_EOF       : SectionEOF := True;
       INT_EXCEL_ID_SST       : ReadSST(AStream);
       INT_EXCEL_ID_CODEPAGE  : ReadCodepage(AStream);
       INT_EXCEL_ID_FONT      : ReadFont(AStream);
       INT_EXCEL_ID_FORMAT    : ReadFormat(AStream);
       INT_EXCEL_ID_XF        : ReadXF(AStream);
       INT_EXCEL_ID_DATEMODE  : ReadDateMode(AStream);
       INT_EXCEL_ID_PALETTE   : ReadPalette(AStream);
      else
        // nothing
      end;
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then SectionEOF := True;
  end;
end;

procedure TsSpreadBIFF8Reader.ReadWorksheet(AStream: TStream; AData: TsWorkbook);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  FWorksheet := AData.AddWorksheet(FWorksheetNames.Strings[FCurrentWorksheet]);

  while (not SectionEOF) do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);
    PendingRecordSize:=RecordSize;

    CurStreamPos := AStream.Position;

    case RecordType of

    INT_EXCEL_ID_BLANK     : ReadBlank(AStream);
    INT_EXCEL_ID_MULBLANK  : ReadMulBlank(AStream);
    INT_EXCEL_ID_NUMBER    : ReadNumber(AStream);
    INT_EXCEL_ID_LABEL     : ReadLabel(AStream);
    INT_EXCEL_ID_FORMULA   : ReadFormula(AStream);
    INT_EXCEL_ID_SHAREDFMLA: ReadSharedFormula(AStream);
    INT_EXCEL_ID_STRING    : ReadStringRecord(AStream);
    //(RSTRING) This record stores a formatted text cell (Rich-Text).
    // In BIFF8 it is usually replaced by the LABELSST record. Excel still
    // uses this record, if it copies formatted text cells to the clipboard.
    INT_EXCEL_ID_RSTRING   : ReadRichString(AStream);
    // (RK) This record represents a cell that contains an RK value
    // (encoded integer or floating-point value). If a floating-point
    // value cannot be encoded to an RK value, a NUMBER record will be written.
    // This record replaces the record INTEGER written in BIFF2.
    INT_EXCEL_ID_RK        : ReadRKValue(AStream);
    INT_EXCEL_ID_MULRK     : ReadMulRKValues(AStream);
    INT_EXCEL_ID_LABELSST  : ReadLabelSST(AStream); //BIFF8 only
    INT_EXCEL_ID_COLINFO   : ReadColInfo(AStream);
    INT_EXCEL_ID_ROW       : ReadRowInfo(AStream);
    INT_EXCEL_ID_WINDOW2   : ReadWindow2(AStream);
    INT_EXCEL_ID_PANE      : ReadPane(AStream);
    INT_EXCEL_ID_BOF       : ;
    INT_EXCEL_ID_EOF       : SectionEOF := True;
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then SectionEOF := True;
  end;

  FixCols(FWorksheet);
  FixRows(FWorksheet);
end;

procedure TsSpreadBIFF8Reader.ReadBoundsheet(AStream: TStream);
var
  Len: Byte;
  WideName: WideString;
begin
  { Absolute stream position of the BOF record of the sheet represented
    by this record }
  // Just assume that they are in order
  AStream.ReadDWord();

  { Visibility }
  AStream.ReadByte();

  { Sheet type }
  AStream.ReadByte();

  { Sheet name: 8-bit length }
  Len := AStream.ReadByte();

  { Read string with flags }
  WideName:=ReadWideString(AStream,Len);

  FWorksheetNames.Add(UTF8Encode(WideName));
end;

function TsSpreadBIFF8Reader.ReadString(const AStream: TStream;
  const ALength: WORD): UTF8String;
begin
  Result := UTF16ToUTF8(ReadWideString(AStream, ALength));
end;

procedure TsSpreadBIFF8Reader.ReadFromFile(AFileName: string; AData: TsWorkbook);
var
  MemStream: TMemoryStream;
  OLEStorage: TOLEStorage;
  OLEDocument: TOLEDocument;
begin
  MemStream := TMemoryStream.Create;
  OLEStorage := TOLEStorage.Create;
  try
    // Only one stream is necessary for any number of worksheets
    OLEDocument.Stream := MemStream;
    OLEStorage.ReadOLEFile(AFileName, OLEDocument, 'Workbook');
      // Can't be shared with BIFF5 because of the parameter "Workbook" !!!)

    // Check if the operation succeded
    if MemStream.Size = 0 then raise Exception.Create('FPSpreadsheet: Reading the OLE document failed');

    // Rewind the stream and read from it
    MemStream.Position := 0;
    ReadFromStream(MemStream, AData);

//    Uncomment to verify if the data was correctly optained from the OLE file
//    MemStream.SaveToFile(SysUtils.ChangeFileExt(AFileName, 'bin.xls'));
  finally
    MemStream.Free;
    OLEStorage.Free;
  end;
end;

procedure TsSpreadBIFF8Reader.ReadFromStream(AStream: TStream; AData: TsWorkbook);
var
  BIFF8EOF: Boolean;
begin
  { Initializations }

  FWorksheetNames := TStringList.Create;
  FWorksheetNames.Clear;
  FCurrentWorksheet := 0;
  BIFF8EOF := False;

  { Read workbook globals }

  ReadWorkbookGlobals(AStream, AData);

  // Check for the end of the file
  if AStream.Position >= AStream.Size then BIFF8EOF := True;

  { Now read all worksheets }

  while (not BIFF8EOF) do
  begin
    //Safe to not read beyond assigned worksheet names.
    if FCurrentWorksheet>FWorksheetNames.Count-1 then break;

    ReadWorksheet(AStream, AData);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then BIFF8EOF := True;

    // Final preparations
    Inc(FCurrentWorksheet);
  end;

  if not FPaletteFound then
    FWorkbook.UsePalette(@PALETTE_BIFF8, Length(PALETTE_BIFF8));

  { Finalizations }

  FWorksheetNames.Free;

end;

procedure TsSpreadBIFF8Reader.ReadLabel(AStream: TStream);
var
  L: Word;
  ARow, ACol: Cardinal;
  XF: Word;
  WideStrValue: WideString;
  cell: PCell;
begin
  { BIFF Record data: Row, Column, XF Index }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Byte String with 16-bit size }
  L := WordLEtoN(AStream.ReadWord());

  { Read string with flags }
  WideStrValue:=ReadWideString(AStream,L);

  { Save the data }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);         // "virtual" cell
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);    // "real" cell

  FWorksheet.WriteUTF8Text(cell, UTF16ToUTF8(WideStrValue));

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF8Reader.ReadRichString(const AStream: TStream);
var
  L: Word;
  B: WORD;
  ARow, ACol: Cardinal;
  XF: Word;
  AStrValue: ansistring;
  cell: PCell;
begin
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Byte String with 16-bit size }
  L := WordLEtoN(AStream.ReadWord());
  AStrValue:=ReadString(AStream,L);        // ???? shouldn't this be a unicode string ????

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);

  { Save the data }
  FWorksheet.WriteUTF8Text(cell, AStrValue);
  //Read formatting runs (not supported)
  B:=WordLEtoN(AStream.ReadWord);
  for L := 0 to B-1 do begin
    AStream.ReadWord; // First formatted character
    AStream.ReadWord; // Index to FONT record
  end;

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{ Reads the cell address used in an RPN formula element. Evaluates the corresponding
  bits to distinguish between absolute and relative addresses.
  Overriding the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellAddress(AStream: TStream;
  out ARow, ACol: Cardinal; out AFlags: TsRelFlags);
var
  c: word;
begin
  // Read row index (2 bytes)
  ARow := WordLEToN(AStream.ReadWord);
  // Read column index; it contains info on absolute/relative address
  c := WordLEToN(AStream.ReadWord);
  // Extract column index
  ACol := c and MASK_EXCEL_COL_BITS_BIFF8;
  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
end;

{ Read the difference between cell row and column indexed of a cell and a reference
  cell.
  Overriding the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellAddressOffset(AStream: TStream;
  out ARowOffset, AColOffset: Integer; out AFlags: TsRelFlags);
var
  dr: SmallInt;
  dc: ShortInt;
  c: Word;
begin
  // 2 bytes for row offset
  dr := ShortInt(WordLEToN(AStream.ReadWord));
  ARowOffset := dr;

  // 2 bytes for column offset
  c := WordLEToN(AStream.ReadWord);
  dc := ShortInt(Lo(c));
  AColOffset := dc;

  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
end;

{ Reads a cell range address used in an RPN formula element.
  Evaluates the corresponding bits to distinguish between absolute and
  relative addresses.
  Overriding the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellRangeAddress(AStream: TStream;
  out ARow1, ACol1, ARow2, ACol2: Cardinal; out AFlags: TsRelFlags);
var
  c1, c2: word;
begin
  // Read row index of first and last rows (2 bytes, each)
  ARow1 := WordLEToN(AStream.ReadWord);
  ARow2 := WordLEToN(AStream.ReadWord);
  // Read column index of first and last columns; they contain info on
  // absolute/relative address
  c1 := WordLEToN(AStream.ReadWord);
  c2 := WordLEToN(AStream.ReadWord);
  // Extract column index of rist and last columns
  ACol1 := c1 and MASK_EXCEL_COL_BITS_BIFF8;
  ACol2 := c2 and MASK_EXCEL_COL_BITS_BIFF8;
  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c1 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c1 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
  if (c2 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol2);
  if (c2 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow2);
end;

{ Reads the difference between row and column corner indexes of a cell range
  and a reference cell.
  Overriding the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellRangeOffset(AStream: TStream;
  out ARow1Offset, ACol1Offset, ARow2Offset, ACol2Offset: Integer;
  out AFlags: TsRelFlags);
var
  c1, c2: Word;
begin
  // 2 bytes for offset of first row
  ARow1Offset := ShortInt(WordLEToN(AStream.ReadWord));

  // 2 bytes for offset to last row
  ARow2Offset := ShortInt(WordLEToN(AStream.ReadWord));

  // 2 bytes for offset of first column
  c1 := WordLEToN(AStream.ReadWord);
  ACol1Offset := Shortint(Lo(c1));

  // 2 bytes for offset of last column
  c2 := WordLEToN(AStream.ReadWord);
  ACol2Offset := ShortInt(Lo(c2));

  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c1 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c1 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
  if (c2 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol2);
  if (c2 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow2);

end;

procedure TsSpreadBIFF8Reader.ReadSST(const AStream: TStream);
var
  Items: DWORD;
  StringLength, CurStrLen: WORD;
  LString: String;
  ContinueIndicator: WORD;
begin
  //Reads the shared string table, only compatible with BIFF8
  if not Assigned(FSharedStringTable) then begin
    //First time SST creation
    FSharedStringTable:=TStringList.Create;

    DWordLEtoN(AStream.ReadDWord); //Apparences not used
    Items:=DWordLEtoN(AStream.ReadDWord);
    Dec(PendingRecordSize,8);
  end else begin
    //A second record must not happend. Garbage so skip.
    Exit;
  end;
  while Items>0 do begin
    StringLength:=0;
    StringLength:=WordLEtoN(AStream.ReadWord);
    Dec(PendingRecordSize,2);
    LString:='';

    // This loop takes care of the string being split between the STT and the CONTINUE, or between CONTINUE records
    while PendingRecordSize>0 do
    begin
      if StringLength>0 then
      begin
        //Read a stream of zero length reads all the stream.
        LString:=LString+ReadString(AStream, StringLength);
      end
      else
      begin
        //String of 0 chars in length, so just read it empty, reading only the mandatory flags
        AStream.ReadByte; //And discard it.
        Dec(PendingRecordSize);
        //LString:=LString+'';
      end;

      // Check if the record finished and we need a CONTINUE record to go on
      if (PendingRecordSize<=0) and (Items>1) then
      begin
        //A Continue will happend, read the
        //tag and continue linking...
        ContinueIndicator:=WordLEtoN(AStream.ReadWord);
        if ContinueIndicator<>INT_EXCEL_ID_CONTINUE then begin
          Raise Exception.Create('[TsSpreadBIFF8Reader.ReadSST] Expected CONTINUE record not found.');
        end;
        PendingRecordSize:=WordLEtoN(AStream.ReadWord);
        CurStrLen := Length(UTF8ToUTF16(LString));
        if StringLength<CurStrLen then Exception.Create('[TsSpreadBIFF8Reader.ReadSST] StringLength<CurStrLen');
        Dec(StringLength, CurStrLen); //Dec the used chars
        if StringLength=0 then break;
      end else begin
        break;
      end;
    end;
    FSharedStringTable.Add(LString);
    {$ifdef FPSPREADDEBUG}
    WriteLn('Adding shared string: ' + LString);
    {$endif}
    dec(Items);
  end;
end;

procedure TsSpreadBIFF8Reader.ReadLabelSST(const AStream: TStream);
var
  ACol,ARow: Cardinal;
  XF: WORD;
  SSTIndex: DWORD;
  rec: TBIFF8LabelSSTRecord;
  cell: PCell;
begin
  { Read entire record, starting at Row }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF8LabelSSTRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);
  SSTIndex := DWordLEToN(rec.SSTIndex);

  if SizeInt(SSTIndex) >= FSharedStringTable.Count then begin
    Raise Exception.CreateFmt('Index %d in SST out of range (0-%d)',[Integer(SSTIndex),FSharedStringTable.Count-1]);
  end;

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);

  FWorksheet.WriteUTF8Text(cell, FSharedStringTable[SSTIndex]);

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{ Helper function for reading a string with 8-bit length. }
function TsSpreadBIFF8Reader.ReadString_8bitLen(AStream: TStream): String;
var
  s: widestring;
begin
  s := ReadWideString(AStream, true);
  Result := UTF8Encode(s);
end;

procedure TsSpreadBIFF8Reader.ReadStringRecord(AStream: TStream);
var
  s: String;
begin
  s := ReadWideString(AStream, false);
  if (FIncompleteCell <> nil) and (s <> '') then begin
    FIncompleteCell^.UTF8StringValue := UTF8Encode(s);
    FIncompleteCell^.ContentType := cctUTF8String;
    if FIsVirtualMode then
      Workbook.OnReadCellData(Workbook, FIncompleteCell^.Row, FIncompleteCell^.Col, FIncompleteCell);
  end;
  FIncompleteCell := nil;
end;

procedure TsSpreadBIFF8Reader.ReadXF(const AStream: TStream);

  function FixLineStyle(dw: DWord): TsLineStyle;
  { Not all line styles defined in BIFF8 are supported by fpspreadsheet. }
  begin
    case dw of
      $01..$07: result := TsLineStyle(dw-1);
//      $07: Result := lsDotted;
      else Result := lsDashed;
    end;
  end;

type
  TXFRecord = packed record                // see p. 224
    FontIndex: Word;                       // Offset 0, Size 2
    FormatIndex: Word;                     // Offset 2, Size 2
    XFType_CellProt_ParentStyleXF: Word;   // Offset 4, Size 2
    Align_TextBreak: Byte;                 // Offset 6, Size 1
    XFRotation: Byte;                      // Offset 7, Size 1
    Indent_Shrink_TextDir: Byte;           // Offset 8, Size 1
    UnusedAttrib: Byte;                    // Offset 9, Size 1
    Border_Background_1: DWord;            // Offset 10, Size 4
    Border_Background_2: DWord;            // Offset 14, Size 4
    Border_Background_3: DWord;            // Offset 18, Size 2
  end;
var
  lData: TXFListData;
  xf: TXFRecord;
  b: Byte;
  dw: DWord;
  fill: Integer;
begin
  AStream.ReadBuffer(xf, SizeOf(xf));

  lData := TXFListData.Create;

  // Font index
  lData.FontIndex := WordLEToN(xf.FontIndex);

  // Format index
  lData.FormatIndex := WordLEToN(xf.FormatIndex);

  // Horizontal text alignment
  b := xf.Align_TextBreak AND MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
    lData.HorAlignment := TsHorAlignment(b)
  else
    lData.HorAlignment := haDefault;

  // Vertical text alignment
  b := (xf.Align_TextBreak AND MASK_XF_VERT_ALIGN) shr 4;
  if (b + 1 <= ord(high(TsVertAlignment))) then
    lData.VertAlignment := tsVertAlignment(b + 1)      // + 1 due to vaDefault
  else
    lData.VertAlignment := vaDefault;

  // Word wrap
  lData.WordWrap := (xf.Align_TextBreak and MASK_XF_TEXTWRAP) <> 0;

  // TextRotation
  case xf.XFRotation of
    XF_ROTATION_HORIZONTAL : lData.TextRotation := trHorizontal;
    XF_ROTATION_90DEG_CCW  : ldata.TextRotation := rt90DegreeCounterClockwiseRotation;
    XF_ROTATION_90DEG_CW   : lData.TextRotation := rt90DegreeClockwiseRotation;
    XF_ROTATION_STACKED    : lData.TextRotation := rtStacked;
  end;

  // Cell borders
  xf.Border_Background_1 := DWordLEToN(xf.Border_Background_1);
  lData.Borders := [];
  lData.BorderStyles := DEFAULT_BORDERSTYLES;

  // the 4 masked bits encode the line style of the border line. 0 = no line
  dw := xf.Border_Background_1 and MASK_XF_BORDER_LEFT;
  if dw <> 0 then begin
    Include(lData.Borders, cbWest);
    lData.BorderStyles[cbWest].LineStyle := FixLineStyle(dw);
  end;
  dw := xf.Border_Background_1 and MASK_XF_BORDER_RIGHT;
  if dw <> 0 then begin
    Include(lData.Borders, cbEast);
    lData.BorderStyles[cbEast].LineStyle := FixLineStyle(dw shr 4);
  end;
  dw := xf.Border_Background_1 and MASK_XF_BORDER_TOP;
  if dw <> 0 then begin
    Include(lData.Borders, cbNorth);
    lData.BorderStyles[cbNorth].LineStyle := FixLineStyle(dw shr 8);
  end;
  dw := xf.Border_Background_1 and MASK_XF_BORDER_BOTTOM;
  if dw <> 0 then begin
    Include(lData.Borders, cbSouth);
    lData.BorderStyles[cbSouth].LineStyle := FixLineStyle(dw shr 12);
  end;
  dw := xf.Border_Background_2 and MASK_XF_BORDER_DIAGONAL;
  if dw <> 0 then begin
    lData.BorderStyles[cbDiagUp].LineStyle := FixLineStyle(dw shr 21);
    lData.BorderStyles[cbDiagDown].LineStyle := lData.BorderStyles[cbDiagUp].LineStyle;
    if xf.Border_Background_1 and MASK_XF_BORDER_SHOW_DIAGONAL_UP <> 0 then
      Include(lData.Borders, cbDiagUp);
    if xf.Border_Background_1 and MASK_XF_BORDER_SHOW_DIAGONAL_DOWN <> 0 then
      Include(lData.Borders, cbDiagDown);
  end;

  // Border line colors
  lData.BorderStyles[cbWest].Color := (xf.Border_Background_1 and MASK_XF_BORDER_LEFT_COLOR) shr 16;
  lData.BorderStyles[cbEast].Color := (xf.Border_Background_1 and MASK_XF_BORDER_RIGHT_COLOR) shr 23;
  lData.BorderStyles[cbNorth].Color := (xf.Border_Background_2 and MASK_XF_BORDER_TOP_COLOR);
  lData.BorderStyles[cbSouth].Color := (xf.Border_Background_2 and MASK_XF_BORDER_BOTTOM_COLOR) shr 7;
  lData.BorderStyles[cbDiagUp].Color := (xf.Border_Background_2 and MASK_XF_BORDER_DIAGONAL_COLOR) shr 14;
  lData.BorderStyles[cbDiagDown].Color := lData.BorderStyles[cbDiagUp].Color;

  // Background fill pattern
  fill := (xf.Border_Background_2 and MASK_XF_BACKGROUND_PATTERN) shr 26;

  // Background color
  xf.Border_Background_3 := DWordLEToN(xf.Border_Background_3);
  if fill <> 0 then
    lData.BackgroundColor := xf.Border_Background_3 and $007F
  else
    lData.BackgroundColor := scTransparent;  // this means "no fill"

  // Add the XF to the list
  FXFList.Add(lData);
end;

procedure TsSpreadBIFF8Reader.ReadFont(const AStream: TStream);
var
  lCodePage: Word;
  lHeight: Word;
  lOptions: Word;
  lColor: Word;
  lWeight: Word;
  Len: Byte;
  font: TsFont;
begin
  font := TsFont.Create;

  { Height of the font in twips = 1/20 of a point }
  lHeight := WordLEToN(AStream.ReadWord); // WordToLE(200)
  font.Size := lHeight/20;

  { Option flags }
  lOptions := WordLEToN(AStream.ReadWord);
  font.Style := [];
  if lOptions and $0001 <> 0 then Include(font.Style, fssBold);
  if lOptions and $0002 <> 0 then Include(font.Style, fssItalic);
  if lOptions and $0004 <> 0 then Include(font.Style, fssUnderline);
  if lOptions and $0008 <> 0 then Include(font.Style, fssStrikeout);

  { Colour index }
  lColor := WordLEToN(AStream.ReadWord);
  //font.Color := TsColor(lColor - 8);  // Palette colors have an offset 8
  font.Color := tsColor(lColor);

  { Font weight }
  lWeight := WordLEToN(AStream.ReadWord);
  if lWeight = 700 then Include(font.Style, fssBold);

  { Escape type }
  AStream.ReadWord();

  { Underline type }
  if AStream.ReadByte > 0 then Include(font.Style, fssUnderline);

  { Font family }
  AStream.ReadByte();

  { Character set }
  lCodepage := AStream.ReadByte();
  {$ifdef FPSPREADDEBUG}
  WriteLn('Reading Font Codepage='+IntToStr(lCodepage));
  {$endif}

  { Not used }
  AStream.ReadByte();

  { Font name: Unicodestring, char count in 1 byte }
  Len := AStream.ReadByte();
  font.FontName := ReadString(AStream, Len);

  { Add font to workbook's font list }
  FWorkbook.AddFont(font);
end;

// Read the FORMAT record for formatting numerical data
procedure TsSpreadBIFF8Reader.ReadFormat(AStream: TStream);
var
  fmtString: String;
  fmtIndex: Integer;
begin
  // Record FORMAT, BIFF 8 (5.49):
  // Offset Size Contents
  // 0      2     Format index used in other records
  // 2      var   Number format string (Unicode string, 16-bit string length)
  // From BIFF5 on: indexes 0..163 are built in
  fmtIndex := WordLEtoN(AStream.ReadWord);

  // 2 var. Number format string (Unicode string, 16-bit string length, ➜2.5.3)
  fmtString := UTF8Encode(ReadWideString(AStream, False));

  // Analyze the format string and add format to the list
  NumFormatList.AnalyzeAndAdd(fmtIndex, fmtString);
end;


{*******************************************************************
*  Initialization section
*
*  Registers this reader / writer on fpSpreadsheet
*  Converts the palette to litte-endian
*
*******************************************************************}

initialization

  RegisterSpreadFormat(TsSpreadBIFF8Reader, TsSpreadBIFF8Writer, sfExcel8);
  MakeLEPalette(@PALETTE_BIFF8, Length(PALETTE_BIFF8));

end.

