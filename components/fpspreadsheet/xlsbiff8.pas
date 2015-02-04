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
  {$mode delphi}{$H+}
{$endif}

// The new OLE code is much better, so always use it
{$define USE_NEW_OLE}
{.$define FPSPREADDEBUG} //define to print out debug info to console. Used to be XLSDEBUG;

interface

uses
  Classes, SysUtils, fpcanvas, DateUtils,
  fpstypes, fpspreadsheet, xlscommon,
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
    procedure ReadWorkbookGlobals(AStream: TStream);
    procedure ReadWorksheet(AStream: TStream);
    procedure ReadBoundsheet(AStream: TStream);
    function ReadString(const AStream: TStream; const ALength: WORD): String;
  protected
    procedure ReadFont(const AStream: TStream);
    procedure ReadFormat(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadLabelSST(const AStream: TStream);
    procedure ReadMergedCells(const AStream: TStream);
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
    procedure ReadFromFile(AFileName: string); override;
    procedure ReadFromStream(AStream: TStream); override;
  end;

  { TsSpreadBIFF8Writer }

  TsSpreadBIFF8Writer = class(TsSpreadBIFFWriter)
  protected
    { Record writing methods }
    procedure WriteBOF(AStream: TStream; ADataType: Word);
    function  WriteBoundsheet(AStream: TStream; ASheetName: string): Int64;
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFont: TsFont);
    procedure WriteFonts(AStream: TStream);
    procedure WriteIndex(AStream: TStream);
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteMergedCells(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteNumFormat(AStream: TStream; AFormatData: TsNumFormatData;
      AListIndex: Integer); override;
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
    procedure WriteXF(AStream: TStream; AFormatRecord: PsCellFormat;
      XFType_Prot: Byte = 0); override;
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
  Math, lconvencoding, fpsStrings, fpsStreams, fpsExprParser;

const
   { Excel record IDs }
     INT_EXCEL_ID_MERGEDCELLS            = $00E5;  // BIFF8 only
     INT_EXCEL_ID_SST                    = $00FC; //BIFF8 only
     INT_EXCEL_ID_LABELSST               = $00FD; //BIFF8 only
{%H-}INT_EXCEL_ID_FORCEFULLCALCULATION   = $08A3;

   { Cell Addresses constants }
     MASK_EXCEL_COL_BITS_BIFF8           = $00FF;
     MASK_EXCEL_RELATIVE_COL_BIFF8       = $4000;  // This is according to Microsoft documentation,
     MASK_EXCEL_RELATIVE_ROW_BIFF8       = $8000;  // but opposite to OpenOffice documentation!

   { BOF record constants }
     INT_BOF_BIFF8_VER                   = $0600;
     INT_BOF_WORKBOOK_GLOBALS            = $0005;
{%H-}INT_BOF_VB_MODULE                   = $0006;
     INT_BOF_SHEET                       = $0010;
{%H-}INT_BOF_CHART                       = $0020;
{%H-}INT_BOF_MACRO_SHEET                 = $0040;
{%H-}INT_BOF_WORKSPACE                   = $0100;
     INT_BOF_BUILD_ID                    = $1FD2;
     INT_BOF_BUILD_YEAR                  = $07CD;

   { STYLE record constants }
     MASK_STYLE_BUILT_IN                 = $8000;

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
  TBIFF8_DimensionsRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FirstRow: DWord;
    LastRowPlus1: DWord;
    FirstCol: Word;
    LastColPlus1: Word;
    NotUsed: Word;
  end;

  TBIFF8_LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    TextLen: Word;
    TextFlags: Byte;
  end;

  TBIFF8_LabelSSTRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    SSTIndex: DWord;
  end;

  TBIFF8_XFRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FontIndex: Word;
    NumFormatIndex: Word;
    XFType_Prot_ParentXF: Word;
    Align_TextBreak: Byte;
    TextRotation: Byte;
    Indent_Shrink_TextDir: Byte;
    UsedAttrib: Byte;
    Border_BkGr1: DWord;
    Border_BkGr2: DWord;
    BkGr3: Word;
  end;


{ TsSpreadBIFF8Reader }

destructor TsSpreadBIFF8Reader.Destroy;
begin
  if Assigned(FSharedStringTable) then FSharedStringTable.Free;
  inherited;
end;

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
  C: WideChar;
begin
  StringFlags:=AStream.ReadByte;
  Dec(PendingRecordSize);
  if StringFlags and 4 = 4 then begin
    //Asian phonetics
    //Read Asian phonetics Length (not used)
    AsianPhoneticBytes := DWordLEtoN(AStream.ReadDWord);
  end;
  if StringFlags and 8 = 8 then begin
    //Rich string
    RunsCounter := WordLEtoN(AStream.ReadWord);
    dec(PendingRecordSize,2);
  end;
  if StringFlags and 1 = 1 Then begin
    //String is WideStringLE
    if (ALength*SizeOf(WideChar)) > PendingRecordSize then begin
      SetLength(Result, PendingRecordSize div 2);
      AStream.ReadBuffer(Result[1], PendingRecordSize);
      Dec(PendingRecordSize, PendingRecordSize);
    end else begin
      SetLength(Result, ALength);
      AStream.ReadBuffer(Result[1], ALength * SizeOf(WideChar));
      Dec(PendingRecordSize, ALength * SizeOf(WideChar));
    end;
    Result := WideStringLEToN(Result);
  end else begin
    // String is 1 byte per char, this is UTF-16 with the high byte ommited
    // because it is zero, so decompress and then convert
    lLen := ALength;
    SetLength(DecomprStrValue, lLen);
    for i := 1 to lLen do
    begin
      C := WideChar(AStream.ReadByte);  // Read 1 byte, but put it into a 2-byte char
      DecomprStrValue[i] := C;
      Dec(PendingRecordSize);
      if (PendingRecordSize <= 0) and (i < lLen) then begin
        //A CONTINUE may have happened here
        RecordType := WordLEToN(AStream.ReadWord);
        RecordSize := WordLEToN(AStream.ReadWord);
        if RecordType <> INT_EXCEL_ID_CONTINUE then begin
          Raise Exception.Create('[TsSpreadBIFF8Reader.ReadWideString] Expected CONTINUE record not found.');
        end else begin
          PendingRecordSize := RecordSize;
          DecomprStrValue := copy(DecomprStrValue,1,i) + ReadWideString(AStream, ALength-i);
          break;
        end;
      end;
    end;
    Result := DecomprStrValue;
  end;
  if StringFlags and 8 = 8 then begin
    //Rich string (This only happened in BIFF8)
    for j := 1 to RunsCounter do begin
      if (PendingRecordSize <= 0) then begin
        //A CONTINUE may happend here
        RecordType := WordLEToN(AStream.ReadWord);
        RecordSize := WordLEToN(AStream.ReadWord);
        if RecordType <> INT_EXCEL_ID_CONTINUE then begin
          Raise Exception.Create('[TsSpreadBIFF8Reader.ReadWideString] Expected CONTINUE record not found.');
        end else begin
          PendingRecordSize := RecordSize;
        end;
      end;
      AStream.ReadWord;
      AStream.ReadWord;
      dec(PendingRecordSize, 2*2);
    end;
  end;
  if StringFlags and 4 = 4 then begin
    //Asian phonetics
    //Read Asian phonetics, discarded as not used.
    SetLength(AnsiStrValue, AsianPhoneticBytes);
    AStream.ReadBuffer(AnsiStrValue[1], AsianPhoneticBytes);
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

procedure TsSpreadBIFF8Reader.ReadWorkbookGlobals(AStream: TStream);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
begin
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

procedure TsSpreadBIFF8Reader.ReadWorksheet(AStream: TStream);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  FWorksheet := FWorkbook.AddWorksheet(FWorksheetNames[FCurrentWorksheet], true);

  while (not SectionEOF) do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);
    PendingRecordSize := RecordSize;

    CurStreamPos := AStream.Position;

    case RecordType of

    INT_EXCEL_ID_BLANK       : ReadBlank(AStream);
    INT_EXCEL_ID_BOOLERROR   : ReadBool(AStream);
    INT_EXCEL_ID_MULBLANK    : ReadMulBlank(AStream);
    INT_EXCEL_ID_NUMBER      : ReadNumber(AStream);
    INT_EXCEL_ID_LABEL       : ReadLabel(AStream);
    INT_EXCEL_ID_FORMULA     : ReadFormula(AStream);
    INT_EXCEL_ID_SHAREDFMLA  : ReadSharedFormula(AStream);
    INT_EXCEL_ID_STRING      : ReadStringRecord(AStream);
    //(RSTRING) This record stores a formatted text cell (Rich-Text).
    // In BIFF8 it is usually replaced by the LABELSST record. Excel still
    // uses this record, if it copies formatted text cells to the clipboard.
    INT_EXCEL_ID_RSTRING     : ReadRichString(AStream);
    // (RK) This record represents a cell that contains an RK value
    // (encoded integer or floating-point value). If a floating-point
    // value cannot be encoded to an RK value, a NUMBER record will be written.
    // This record replaces the record INTEGER written in BIFF2.
    INT_EXCEL_ID_RK          : ReadRKValue(AStream);
    INT_EXCEL_ID_MULRK       : ReadMulRKValues(AStream);
    INT_EXCEL_ID_LABELSST    : ReadLabelSST(AStream); //BIFF8 only
    INT_EXCEL_ID_DEFCOLWIDTH : ReadDefColWidth(AStream);
    INT_EXCEL_ID_COLINFO     : ReadColInfo(AStream);
    INT_EXCEL_ID_MERGEDCELLS : ReadMergedCells(AStream);
    INT_EXCEL_ID_ROW         : ReadRowInfo(AStream);
    INT_EXCEL_ID_WINDOW2     : ReadWindow2(AStream);
    INT_EXCEL_ID_PANE        : ReadPane(AStream);
    INT_EXCEL_ID_BOF         : ;
    INT_EXCEL_ID_EOF         : SectionEOF := True;
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
  const ALength: WORD): String;
begin
  Result := UTF16ToUTF8(ReadWideString(AStream, ALength));
end;

procedure TsSpreadBIFF8Reader.ReadFromFile(AFileName: string);
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
    if MemStream.Size = 0 then raise Exception.Create('[TsSpreadBIFF8Reader.ReadFromFile] Reading of OLE document failed');

    // Rewind the stream and read from it
    MemStream.Position := 0;
    ReadFromStream(MemStream);

//    Uncomment to verify if the data was correctly optained from the OLE file
//    MemStream.SaveToFile(SysUtils.ChangeFileExt(AFileName, 'bin.xls'));
  finally
    MemStream.Free;
    OLEStorage.Free;
  end;
end;

procedure TsSpreadBIFF8Reader.ReadFromStream(AStream: TStream);
var
  BIFF8EOF: Boolean;
begin
  { Initializations }

  FWorksheetNames := TStringList.Create;
  FWorksheetNames.Clear;
  FCurrentWorksheet := 0;
  BIFF8EOF := False;

  { Read workbook globals }
  ReadWorkbookGlobals(AStream);

  // Check for the end of the file
  if AStream.Position >= AStream.Size then BIFF8EOF := True;

  { Now read all worksheets }
  while (not BIFF8EOF) do
  begin
    //Safe to not read beyond assigned worksheet names.
    if FCurrentWorksheet > FWorksheetNames.Count-1 then break;

    ReadWorksheet(AStream);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then BIFF8EOF := True;

    // Final preparations
    Inc(FCurrentWorksheet);
    if FCurrentWorksheet = FWorksheetNames.Count then BIFF8EOF := True;
    // It can happen in files written by Office97 that the OLE directory is
    // at the end of the file.
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
    InitCell(ARow, ACol, FVirtualCell);        // "virtual" cell
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);    // "real" cell

  FWorksheet.WriteUTF8Text(cell, UTF16ToUTF8(WideStrValue));

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF8Reader.ReadMergedCells(const AStream: TStream);
var
  rng: packed record Row1, Row2, Col1, Col2: Word; end;
  i, n: word;
begin
  rng.Row1 := 0;  // to silence the compiler...

  // Count of merged ranges
  n := WordLEToN(AStream.ReadWord);

  for i:=1 to n do begin
    // Read range
    AStream.ReadBuffer(rng, SizeOf(rng));
    // Transfer cell range to worksheet
    FWorksheet.MergeCells(
      WordLEToN(rng.Row1), WordLEToN(rng.Col1),
      WordLEToN(rng.Row2), WordLEToN(rng.Col2)
    );
  end;
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
  B := WordLEtoN(AStream.ReadWord);
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
  rec: TBIFF8_LabelSSTRecord;
  cell: PCell;
begin
  rec.Row := 0;  // to silence the compiler...

  { Read entire record, starting at Row }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF8_LabelSSTRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);
  SSTIndex := DWordLEToN(rec.SSTIndex);

  if SizeInt(SSTIndex) >= FSharedStringTable.Count then begin
    raise Exception.CreateFmt(rsIndexInSSTOutOfRange, [
      Integer(SSTIndex),FSharedStringTable.Count-1
    ]);
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
const
  HAS_8BITLEN = true;
var
  wideStr: widestring;
begin
  wideStr := ReadWideString(AStream, HAS_8BITLEN);
  Result := UTF8Encode(wideStr);
end;

procedure TsSpreadBIFF8Reader.ReadStringRecord(AStream: TStream);
var
  wideStr: WideString;
begin
  wideStr := ReadWideString(AStream, false);
  if (FIncompleteCell <> nil) and (wideStr <> '') then begin
    FIncompleteCell^.UTF8StringValue := UTF8Encode(wideStr);
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
      else Result := lsDashed;
    end;
  end;
var
  rec: TBIFF8_XFRecord;
  fmt: TsCellFormat;
  b: Byte;
  dw: DWord;
  fill: Integer;
  nfidx: Integer;
  nfdata: TsNumFormatData;
  i: Integer;
begin
  InitFormatRecord(fmt);
  fmt.ID := FCellFormatList.Count;

  rec.FontIndex := 0;  // to silence the compiler...
  // Read entire xf record into a buffer
  AStream.ReadBuffer(rec.FontIndex, SizeOf(rec) - 2*SizeOf(word));

  // Font index
  fmt.FontIndex := WordLEToN(rec.FontIndex);
  if fmt.FontIndex = 1 then
    Include(fmt.UsedFormattingFields, uffBold)
  else if fmt.FontIndex > 1 then
    Include(fmt.UsedFormattingFields, uffFont);

  // Number format index
  nfidx := WordLEToN(rec.NumFormatIndex);
  i := NumFormatList.FindByIndex(nfidx);
  if i > -1 then begin
    nfdata := NumFormatList.Items[i];
    fmt.NumberFormat := nfdata.NumFormat;
    fmt.NumberFormatStr := nfdata.FormatString;
    if nfdata.NumFormat <> nfGeneral then
      Include(fmt.UsedFormattingFields, uffNumberFormat);
  end;

  // Horizontal text alignment
  b := rec.Align_TextBreak AND MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
  begin
    fmt.HorAlignment := TsHorAlignment(b);
    if fmt.HorAlignment <> haDefault then
      Include(fmt.UsedFormattingFields, uffHorAlign);
  end;

  // Vertical text alignment
  b := (rec.Align_TextBreak AND MASK_XF_VERT_ALIGN) shr 4;
  if (b + 1 <= ord(high(TsVertAlignment))) then
  begin
    fmt.VertAlignment := TsVertAlignment(b + 1);      // + 1 due to vaDefault
    // Unfortunately BIFF does not provide a "default" vertical alignment code.
    // Without the following correction "non-formatted" cells would always have
    // the uffVertAlign FormattingField set which contradicts the statement of
    // not being formatted.
    if fmt.VertAlignment = vaBottom then
      fmt.VertAlignment := vaDefault;
    if fmt.VertAlignment <> vaDefault then
      Include(fmt.UsedFormattingFields, uffVertAlign);
  end;

  // Word wrap
  if (rec.Align_TextBreak and MASK_XF_TEXTWRAP) <> 0 then
    Include(fmt.UsedFormattingFields, uffWordwrap);

  // TextRotation
  case rec.TextRotation of
    XF_ROTATION_HORIZONTAL : fmt.TextRotation := trHorizontal;
    XF_ROTATION_90DEG_CCW  : fmt.TextRotation := rt90DegreeCounterClockwiseRotation;
    XF_ROTATION_90DEG_CW   : fmt.TextRotation := rt90DegreeClockwiseRotation;
    XF_ROTATION_STACKED    : fmt.TextRotation := rtStacked;
  end;
  if fmt.TextRotation <> trHorizontal then
    Include(fmt.UsedFormattingFields, uffTextRotation);

  // Cell borders
  rec.Border_BkGr1 := DWordLEToN(rec.Border_BkGr1);
  rec.Border_BkGr2 := DWordLEToN(rec.Border_BkGr2);

  // the 4 masked bits encode the line style of the border line. 0 = no line
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_LEFT;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbWest);
    fmt.BorderStyles[cbWest].LineStyle := FixLineStyle(dw);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_RIGHT;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbEast);
    fmt.BorderStyles[cbEast].LineStyle := FixLineStyle(dw shr 4);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_TOP;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbNorth);
    fmt.BorderStyles[cbNorth].LineStyle := FixLineStyle(dw shr 8);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_BOTTOM;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbSouth);
    fmt.BorderStyles[cbSouth].LineStyle := FixLineStyle(dw shr 12);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr2 and MASK_XF_BORDER_DIAGONAL;
  if dw <> 0 then
  begin
    fmt.BorderStyles[cbDiagUp].LineStyle := FixLineStyle(dw shr 21);
    fmt.BorderStyles[cbDiagDown].LineStyle := fmt.BorderStyles[cbDiagUp].LineStyle;
    if rec.Border_BkGr1 and MASK_XF_BORDER_SHOW_DIAGONAL_UP <> 0 then
      Include(fmt.Border, cbDiagUp);
    if rec.Border_BkGr1 and MASK_XF_BORDER_SHOW_DIAGONAL_DOWN <> 0 then
      Include(fmt.Border, cbDiagDown);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;

  // Border line colors
  fmt.BorderStyles[cbWest].Color := (rec.Border_BkGr1 and MASK_XF_BORDER_LEFT_COLOR) shr 16;
  fmt.BorderStyles[cbEast].Color := (rec.Border_BkGr1 and MASK_XF_BORDER_RIGHT_COLOR) shr 23;
  fmt.BorderStyles[cbNorth].Color := (rec.Border_BkGr2 and MASK_XF_BORDER_TOP_COLOR);
  fmt.BorderStyles[cbSouth].Color := (rec.Border_BkGr2 and MASK_XF_BORDER_BOTTOM_COLOR) shr 7;
  fmt.BorderStyles[cbDiagUp].Color := (rec.Border_BkGr2 and MASK_XF_BORDER_DIAGONAL_COLOR) shr 14;
  fmt.BorderStyles[cbDiagDown].Color := fmt.BorderStyles[cbDiagUp].Color;

  // Background fill pattern
  fill := (rec.Border_BkGr2 and MASK_XF_BACKGROUND_PATTERN) shr 26;

  // Background color
  rec.BkGr3 := DWordLEToN(rec.BkGr3);
  if fill <> 0 then begin
    fmt.BackgroundColor := rec.BkGr3 and $007F;
    Include(fmt.UsedFormattingFields, uffBackgroundColor);
  end else
    fmt.BackgroundColor := scTransparent;  // this means "no fill"

  // Add the XF to the list
  FCellFormatList.Add(fmt);
end;

procedure TsSpreadBIFF8Reader.ReadFont(const AStream: TStream);
var
  {%H-}lCodePage: Word;
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

// Read the (number) FORMAT record for formatting numerical data
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


{ TsSpreadBIFF8Writer }

constructor TsSpreadBIFF8Writer.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel BIFF8 file to the disc

  The BIFF 8 writer overrides this method because BIFF 8 is written
  as an OLE document, and our current OLE document writing method involves:

         1 - Writing the BIFF data to a memory stream
         2 - Write the memory stream data to disk using COM functions
-------------------------------------------------------------------------------}
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
  WriteCodePage(AStream, 'ucs2le'); //seUTF16);
  WriteWindow1(AStream);
  WriteFonts(AStream);
  WriteNumFormats(AStream);
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

      // View settings block
      WriteWindow2(AStream, FWorksheet);
      WritePane(AStream, FWorksheet, isBIFF8, pane);
      WriteSelection(AStream, FWorksheet, pane);

      WriteMergedCells(AStream, FWorksheet);

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


{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 DIMENSIONS record

  nm = (rl - rf - 1) / 32 + 1 (using integer division)

  Excel, OpenOffice and FPSpreadsheet ignore the dimensions written in this
  record, but some other applications really use them, so they need to be correct.

  See bug 18886: excel5 files are truncated when imported
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
var
  firstRow, lastRow, firstCol, lastCol: Cardinal;
  rec: TBIFF8_DimensionsRecord;
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

  WideFontName := UTF8Decode(AFont.FontName);
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

procedure TsSpreadBiff8Writer.WriteNumFormat(AStream: TStream;
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
  s: String;
  ws: widestring;
  rec: TNumFormatRecord;
  buf: array of byte;
begin
  if (AFormatData = nil) or (AFormatData.FormatString = '') then
    exit;

  s := NumFormatList.FormatStringForWriting(AListIndex);
  ws := UTF8Decode(s);
  len := Length(ws);

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
  Move(ws[1], buf[SizeOf(rec)], len*SizeOf(WideChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(rec) + SizeOf(WideChar)*len);

  { Clean up }
  SetLength(buf, 0);
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
  rec: TBIFF8_LabelRecord;
  buf: array of byte;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  WideValue := UTF8Decode(AValue); //to UTF16
  if WideValue = '' then begin
    // Badly formatted UTF8String (maybe ANSI?)
    if Length(AValue)<>0 then begin
      //Quite sure it was an ANSI string written as UTF8, so raise exception.
      raise Exception.CreateFmt(rsUTF8TextExpectedButANSIFoundInCell, [GetCellString(ARow,ACol)]);
    end;
    Exit;
  end;

  if Length(WideValue) > MAXBYTES then begin
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    SetLength(WideValue, MAXBYTES); //may corrupt the string (e.g. in surrogate pairs), but... too bad.
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
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

procedure TsSpreadBIFF8Writer.WriteMergedCells(AStream: TStream;
  AWorksheet: TsWorksheet);
const
  MAX_PER_RECORD = 1026;
var
  i, n0, n: Integer;
  rngList: TsCellRangeArray;
begin
  AWorksheet.GetMergedCellRanges(rngList);
  n0 := Length(rngList);
  i := 0;

  while n0 > 0 do begin
    n := Min(n0, MAX_PER_RECORD);
    // at most 1026 merged ranges per BIFF record, the rest goes into a new record

    { BIFF record header }
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_MERGEDCELLS));
    AStream.WriteWord(WordToLE(2 + n*8));

    // Count of cell ranges in this record
    AStream.WriteWord(WordToLE(n));

    // Loop writing the merged cell ranges
    while (n > 0) and (i < Length(rngList)) do begin
      AStream.WriteWord(WordToLE(rngList[i].Row1));
      AStream.WriteWord(WordToLE(rngList[i].Row2));
      AStream.WriteWord(WordToLE(rngList[i].Col1));
      AStream.WriteWord(WordToLE(rngList[i].Col2));
      inc(i);
      dec(n);
    end;

    dec(n0, MAX_PER_RECORD);
  end;
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
  { BIFF record header }
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
procedure TsSpreadBIFF8Writer.WriteXF(AStream: TStream;
 AFormatRecord: PsCellFormat; XFType_Prot: Byte = 0);
var
  rec: TBIFF8_XFRecord;
  j: Integer;
  b: Byte;
  dw1, dw2: DWord;
begin
  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_XF);
  rec.RecordSize := WordToLE(SizeOf(TBIFF8_XFRecord) - 2*SizeOf(Word));

  { Index to font record }
  rec.FontIndex := 0;
  if (AFormatRecord <> nil) then begin
    if (uffBold in AFormatRecord^.UsedFormattingFields) then
      rec.FontIndex := 1
    else
    if (uffFont in AFormatRecord^.UsedFormattingFields) then
      rec.FontIndex := AFormatRecord^.FontIndex;
  end;
  rec.FontIndex := WordToLE(rec.FontIndex);

  { Index to number format }
  rec.NumFormatIndex := 0;
  if (AFormatRecord <> nil) and (uffNumberFormat in AFormatRecord^.UsedFormattingFields)
  then begin
    // The number formats in the FormatList are still in fpc dialect
    // They will be converted to Excel syntax immediately before writing.
    j := NumFormatList.Find(AFormatRecord^.NumberFormat, AFormatRecord^.NumberFormatStr);
    if j > -1 then
      rec.NumFormatIndex := NumFormatList[j].Index;
  end;
  rec.NumFormatIndex := WordToLE(rec.NumFormatIndex);

  { XF type, cell protection and parent style XF }
  rec.XFType_Prot_ParentXF := XFType_Prot and MASK_XF_TYPE_PROT;
  if XFType_Prot and MASK_XF_TYPE_PROT_STYLE_XF <> 0 then
    rec.XFType_Prot_ParentXF := rec.XFType_Prot_ParentXF or MASK_XF_TYPE_PROT_PARENT;

  { Text alignment and text break }
  if AFormatRecord = nil then
    b := MASK_XF_VERT_ALIGN_BOTTOM
  else
  begin
    b := 0;
    if (uffHorAlign in AFormatRecord^.UsedFormattingFields) then
      case AFormatRecord^.HorAlignment of
        haDefault: ;
        haLeft   : b := b or MASK_XF_HOR_ALIGN_LEFT;
        haCenter : b := b or MASK_XF_HOR_ALIGN_CENTER;
        haRight  : b := b or MASK_XF_HOR_ALIGN_RIGHT;
      end;
    // Since the default vertical alignment is vaDefault but "0" corresponds
    // to vaTop, we alwys have to write the vertical alignment.
    case AFormatRecord^.VertAlignment of
      vaTop    : b := b or MASK_XF_VERT_ALIGN_TOP;
      vaCenter : b := b or MASK_XF_VERT_ALIGN_CENTER;
      vaBottom : b := b or MASK_XF_VERT_ALIGN_BOTTOM;
      else       b := b or MASK_XF_VERT_ALIGN_BOTTOM;
    end;
    if (uffWordWrap in AFormatRecord^.UsedFormattingFields) then
      b := b or MASK_XF_TEXTWRAP;
  end;
  rec.Align_TextBreak := b;

  { Text rotation }
  rec.TextRotation := 0;
  if (AFormatRecord <> nil) and (uffTextRotation in AFormatRecord^.UsedFormattingFields)
    then rec.TextRotation := TEXT_ROTATIONS[AFormatRecord^.TextRotation];

  { Indentation, shrink, merge and text direction:
    see "Excel97-2007BinaryFileFormat(xls)Specification.pdf", p281 ff
    Bits 0-3: Indent value
    Bit 4: Shrink to fit
    Bit 5: MergeCell
    Bits 6-7: Reading direction  }
  rec.Indent_Shrink_TextDir := 0;

  { Used attributes }
  rec.UsedAttrib :=
   MASK_XF_USED_ATTRIB_NUMBER_FORMAT or
   MASK_XF_USED_ATTRIB_FONT or
   MASK_XF_USED_ATTRIB_TEXT or
   MASK_XF_USED_ATTRIB_BORDER_LINES or
   MASK_XF_USED_ATTRIB_BACKGROUND or
   MASK_XF_USED_ATTRIB_CELL_PROTECTION;

  { Cell border lines and background area }

  dw1 := 0;
  dw2 := 0;
  rec.BkGr3 := 0;
  if (AFormatRecord <> nil) and (uffBorder in AFormatRecord^.UsedFormattingFields) then
  begin
    // Left and right line colors
    dw1 := AFormatRecord^.BorderStyles[cbWest].Color shl 16 +
           AFormatRecord^.BorderStyles[cbEast].Color shl 23;
    // Border line styles
    if cbWest in AFormatRecord^.Border then
      dw1 := dw1 or (DWord(AFormatRecord^.BorderStyles[cbWest].LineStyle)+1);
    if cbEast in AFormatRecord^.Border then
      dw1 := dw1 or ((DWord(AFormatRecord^.BorderStyles[cbEast].LineStyle)+1) shl 4);
    if cbNorth in AFormatRecord^.Border then
      dw1 := dw1 or ((DWord(AFormatRecord^.BorderStyles[cbNorth].LineStyle)+1) shl 8);
    if cbSouth in AFormatRecord^.Border then
      dw1 := dw1 or ((DWord(AFormatRecord^.BorderStyles[cbSouth].LineStyle)+1) shl 12);
    if cbDiagDown in AFormatRecord^.Border then
      dw1 := dw1 or $40000000;
    if cbDiagUp in AFormatRecord^.Border then
      dw1 := dw1 or $80000000;

    // Top, bottom and diagonal line colors
    dw2 := FixColor(AFormatRecord^.BorderStyles[cbNorth].Color) +
           FixColor(AFormatRecord^.BorderStyles[cbSouth].Color) shl 7 +
           FixColor(AFormatRecord^.BorderStyles[cbDiagUp].Color) shl 14;
    // In BIFF8 both diagonals have the same color - we use the color of the up-diagonal.

    // Diagonal line style
    if (AFormatRecord^.Border * [cbDiagUp, cbDiagDown] <> []) then
      dw2 := dw2 or ((DWord(AFormatRecord^.BorderStyles[cbDiagUp].LineStyle)+1) shl 21);
    // In BIFF8 both diagonals have the same line style - we use the color of the up-diagonal.
  end;

  if (AFormatRecord <> nil) and (uffBackgroundColor in AFormatRecord^.UsedFormattingFields) then
  begin
    dw2 := dw2 or DWORD(MASK_XF_FILL_PATT_SOLID shl 26);
    rec.BkGr3 := FixColor(AFormatRecord^.BackgroundColor);
  end;

  rec.Border_BkGr1 := DWordToLE(dw1);
  rec.Border_BkGr2 := DWordToLE(dw2);
  rec.BkGr3 := WordToLE(rec.BkGr3);

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
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

