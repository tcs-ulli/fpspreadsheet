{
xlsbiff5.pas

Writes an Excel 5 file

An Excel worksheet stream consists of a number of subsequent records.
To ensure a properly formed file, the following order must be respected:

1st record:        BOF
2nd to Nth record: Any record
Last record:       EOF

Excel 5 files are OLE compound document files, and must be written using the
fpOLE library.

Records Needed to Make a BIFF5 File Microsoft Excel Can Use:

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

Records Needed to Make a BIFF5 File Microsoft Excel Can Use obtained from:

http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q147732&ID=KB;EN-US;Q147732&LN=EN-US&rnk=2&SD=msdn&FR=0&qry=BIFF&src=DHCS_MSPSS_msdn_SRCH&SPR=MSALL&

Microsoft BIFF 5 writer example:

http://support.microsoft.com/kb/150447/en-us

Encoding information: ISO_8859_1 is used, to have support to
other characters, please use a format which support unicode

AUTHORS: Felipe Monteiro de Carvalho
}
unit xlsbiff5;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

{$define USE_NEW_OLE}
{.$define FPSPREADDEBUG} //define to print out debug info to console. Used to be XLSDEBUG;

interface

uses
  Classes, SysUtils, fpcanvas,
  fpstypes, fpspreadsheet,
  xlscommon,
  {$ifdef USE_NEW_OLE}
  fpolebasic,
  {$else}
  fpolestorage,
  {$endif}
  fpsutils, lconvencoding;

type

  { TsSpreadBIFF5Reader }

  TsSpreadBIFF5Reader = class(TsSpreadBIFFReader)
  private
    FWorksheetNames: TStringList;
    FCurrentWorksheet: Integer;
  protected
    { Record writing methods }
    procedure ReadBoundsheet(AStream: TStream);
    procedure ReadFont(const AStream: TStream);
    procedure ReadFormat(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadWorkbookGlobals(AStream: TStream; AData: TsWorkbook);
    procedure ReadWorksheet(AStream: TStream; AData: TsWorkbook);
    procedure ReadRichString(AStream: TStream);
    procedure ReadStandardWidth(AStream: TStream; ASheet: TsWorksheet);
    procedure ReadStringRecord(AStream: TStream); override;
    procedure ReadXF(AStream: TStream);
  public
    { General reading methods }
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); override;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); override;
  end;

  { TsSpreadBIFF5Writer }

  TsSpreadBIFF5Writer = class(TsSpreadBIFFWriter)
  private
    WorkBookEncoding: TsEncoding;
  protected
    { Record writing methods }
    procedure WriteBOF(AStream: TStream; ADataType: Word);
    function  WriteBoundsheet(AStream: TStream; ASheetName: string): Int64;
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream;  AFont: TsFont);
    procedure WriteFonts(AStream: TStream);
    procedure WriteIndex(AStream: TStream);
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumFormat(AStream: TStream; ANumFormatData: TsNumFormatData;
      AListIndex: Integer); override;
    procedure WriteStringRecord(AStream: TStream; AString: String); override;
    procedure WriteStyle(AStream: TStream);
    procedure WriteWindow2(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteXF(AStream: TStream; AFormatRecord: PsCellFormat;
      XFType_Prot: Byte = 0); override;
  public
    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;

var
  // the palette of the default BIFF5 colors as "big-endian color" values
  PALETTE_BIFF5: array[$00..$3F] of TsColorValue = (
    $000000,  // $00: black
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

    $8080FF,  // $18:
    $802060,  // $19:
    $FFFFC0,  // $1A:
    $A0E0F0,  // $1B:
    $600080,  // $1C:
    $FF8080,  // $1D:
    $0080C0,  // $1E:
    $C0C0FF,  // $1F:

    $000080,  // $20:
    $FF00FF,  // $21:
    $FFFF00,  // $22:
    $00FFFF,  // $23:
    $800080,  // $24:
    $800000,  // $25:
    $008080,  // $26:
    $0000FF,  // $27:
    $00CFFF,  // $28:
    $69FFFF,  // $29:
    $E0FFE0,  // $2A:
    $FFFF80,  // $2B:
    $A6CAF0,  // $2C:
    $DD9CB3,  // $2D:
    $B38FEE,  // $2E:
    $E3E3E3,  // $2F:

    $2A6FF9,  // $30:
    $3FB8CD,  // $31:
    $488436,  // $32:
    $958C41,  // $33:
    $8E5E42,  // $34:
    $A0627A,  // $35:
    $624FAC,  // $36:
    $969696,  // $37:
    $1D2FBE,  // $38:
    $286676,  // $39:
    $004500,  // $3A:
    $453E01,  // $3B:
    $6A2813,  // $3C:
    $85396A,  // $3D:
    $4A3285,  // $3E:
    $424242   // $3F:
  );


implementation

uses
  fpsStrings, fpsStreams;

const
   { Excel record IDs }
     // see: in xlscommon

   { BOF record constants }
     INT_BOF_BIFF5_VER          = $0500;
     INT_BOF_WORKBOOK_GLOBALS   = $0005;
{%H-}INT_BOF_VB_MODULE          = $0006;
     INT_BOF_SHEET              = $0010;
{%H-}INT_BOF_CHART              = $0020;
{%H-}INT_BOF_MACRO_SHEET        = $0040;
{%H-}INT_BOF_WORKSPACE          = $0100;
     INT_BOF_BUILD_ID           = $1FD2;
     INT_BOF_BUILD_YEAR         = $07CD;

   { Record IDs }
     INT_EXCEL_ID_STANDARDWIDTH = $0099;

   { FONT record constants }
{%H-}BYTE_ANSILatin1            = $00;
{%H-}BYTE_SYSTEM_DEFAULT        = $01;
{%H-}BYTE_SYMBOL                = $02;
{%H-}BYTE_Apple_Roman           = $4D;
{%H-}BYTE_ANSI_Japanese_Shift_JIS = $80;
{%H-}BYTE_ANSI_Korean_Hangul    = $81;
{%H-}BYTE_ANSI_Korean_Johab     = $81;
{%H-}BYTE_ANSI_Chinese_Simplified_GBK = $86;
{%H-}BYTE_ANSI_Chinese_Traditional_BIG5 = $88;
{%H-}BYTE_ANSI_Greek            = $A1;
{%H-}BYTE_ANSI_Turkish          = $A2;
{%H-}BYTE_ANSI_Vietnamese       = $A3;
{%H-}BYTE_ANSI_Hebrew           = $B1;
{%H-}BYTE_ANSI_Arabic           = $B2;
{%H-}BYTE_ANSI_Baltic           = $BA;
{%H-}BYTE_ANSI_Cyrillic         = $CC;
{%H-}BYTE_ANSI_Thai             = $DE;
{%H-}BYTE_ANSI_Latin2           = $EE;
{%H-}BYTE_OEM_Latin1            = $FF;

   { FORMULA record constants }
{%H-}MASK_FORMULA_RECALCULATE_ALWAYS  = $0001;
{%H-}MASK_FORMULA_RECALCULATE_ON_OPEN = $0002;
{%H-}MASK_FORMULA_SHARED_FORMULA      = $0008;

   { STYLE record constants }
     MASK_STYLE_BUILT_IN     = $8000;

   { WINDOW1 record constants }
{%H-}MASK_WINDOW1_OPTION_WINDOW_HIDDEN             = $0001;
{%H-}MASK_WINDOW1_OPTION_WINDOW_MINIMISED          = $0002;
{%H-}MASK_WINDOW1_OPTION_HORZ_SCROLL_VISIBLE       = $0008;
{%H-}MASK_WINDOW1_OPTION_VERT_SCROLL_VISIBLE       = $0010;
{%H-}MASK_WINDOW1_OPTION_WORKSHEET_TAB_VISIBLE     = $0020;

  { XF substructures }

   { XF substructures --- see xlscommon! }
     MASK_XF_ORIENTATION                 = $03;
     XF_ROTATION_HORIZONTAL              = 0;
     XF_ROTATION_STACKED                 = 1;
     XF_ROTATION_90DEG_CCW               = 2;
     XF_ROTATION_90DEG_CW                = 3;

   { XF CELL BORDER }
     MASK_XF_BORDER_LEFT                 = $00000038;
     MASK_XF_BORDER_RIGHT                = $000001C0;
     MASK_XF_BORDER_TOP                  = $00000007;
     MASK_XF_BORDER_BOTTOM               = $01C00000;

   { XF CELL BORDER COLORS }
     MASK_XF_BORDER_LEFT_COLOR           = $007F0000;
     MASK_XF_BORDER_RIGHT_COLOR          = $3F800000;
     MASK_XF_BORDER_TOP_COLOR            = $0000FE00;
     MASK_XF_BORDER_BOTTOM_COLOR         = $FE000000;

   { XF CELL BACKGROUND }
     MASK_XF_BKGR_PATTERN_COLOR          = $0000007F;
{%H-}MASK_XF_BKGR_BACKGROUND_COLOR       = $00003F80;
     MASK_XF_BKGR_FILLPATTERN            = $003F0000;

     TEXT_ROTATIONS: Array[TsTextRotation] of Byte = (
       XF_ROTATION_HORIZONTAL,
       XF_ROTATION_90DEG_CW,
       XF_ROTATION_90DEG_CCW,
       XF_ROTATION_STACKED
     );

type
  TBIFF5_DimensionsRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FirstRow: Word;
    LastRowPlus1: Word;
    FirstCol: Word;
    LastColPlus1: Word;
    NotUsed: Word;
  end;

  TBIFF5_LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    TextLen: Word;
  end;

  TBIFF5_XFRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FontIndex: Word;
    NumFormatIndex: Word;
    XFType_Prot_ParentXF: Word;
    Align_TextBreak: Byte;
    TextOrient_UnusedAttrib: Byte;
    Border_BkGr1: DWord;
    Border_BkGr2: DWord;
  end;


{ TsSpreadBIFF5Reader }

procedure TsSpreadBIFF5Reader.ReadWorkbookGlobals(AStream: TStream;
  AData: TsWorkbook);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  Unused(AData);

  // Clear existing fonts. They will be replaced by those from the file.
  FWorkbook.RemoveAllFonts;

  while (not SectionEOF) do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);

    CurStreamPos := AStream.Position;

    case RecordType of
     INT_EXCEL_ID_BOF        : ;
     INT_EXCEL_ID_BOUNDSHEET : ReadBoundSheet(AStream);
     INT_EXCEL_ID_FONT       : ReadFont(AStream);
     INT_EXCEL_ID_FORMAT     : ReadFormat(AStream);
     INT_EXCEL_ID_XF         : ReadXF(AStream);
     INT_EXCEL_ID_PALETTE    : ReadPalette(AStream);
     INT_EXCEL_ID_EOF        : SectionEOF := True;
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then SectionEOF := True;
  end;
end;

procedure TsSpreadBIFF5Reader.ReadWorksheet(AStream: TStream; AData: TsWorkbook);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  FWorksheet := AData.AddWorksheet(FWorksheetNames[FCurrentWorksheet], true);

  while (not SectionEOF) do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);

    CurStreamPos := AStream.Position;

    case RecordType of
      INT_EXCEL_ID_BLANK         : ReadBlank(AStream);
      INT_EXCEL_ID_BOOLERROR     : ReadBool(AStream);
      INT_EXCEL_ID_MULBLANK      : ReadMulBlank(AStream);
      INT_EXCEL_ID_NOTE          : ReadComment(AStream);
      INT_EXCEL_ID_NUMBER        : ReadNumber(AStream);
      INT_EXCEL_ID_LABEL         : ReadLabel(AStream);
      INT_EXCEL_ID_RSTRING       : ReadRichString(AStream); //(RSTRING) This record stores a formatted text cell (Rich-Text). In BIFF8 it is usually replaced by the LABELSST record. Excel still uses this record, if it copies formatted text cells to the clipboard.
      INT_EXCEL_ID_RK            : ReadRKValue(AStream); //(RK) This record represents a cell that contains an RK value (encoded integer or floating-point value). If a floating-point value cannot be encoded to an RK value, a NUMBER record will be written. This record replaces the record INTEGER written in BIFF2.
      INT_EXCEL_ID_MULRK         : ReadMulRKValues(AStream);
      INT_EXCEL_ID_COLINFO       : ReadColInfo(AStream);
      INT_EXCEL_ID_STANDARDWIDTH : ReadStandardWidth(AStream, FWorksheet);
      INT_EXCEL_ID_DEFCOLWIDTH   : ReadDefColWidth(AStream);
      INT_EXCEL_ID_ROW           : ReadRowInfo(AStream);
      INT_EXCEL_ID_FORMULA       : ReadFormula(AStream);
      INT_EXCEL_ID_SHAREDFMLA    : ReadSharedFormula(AStream);
      INT_EXCEL_ID_STRING        : ReadStringRecord(AStream);
      INT_EXCEL_ID_WINDOW2       : ReadWindow2(AStream);
      INT_EXCEL_ID_PANE          : ReadPane(AStream);
      INT_EXCEL_ID_BOF           : ;
      INT_EXCEL_ID_EOF           : SectionEOF := True;

     {$IFDEF FPSPREADDEBUG} // Only write out if debugging
    else
      // Show unsupported record types to console.
      case RecordType of
        $000C: ; //(CALCCOUNT) This record is part of the Calculation Settings Block. It specifies the maximum number of times the formulas should be iteratively calculated. This is a fail-safe against mutually recursive formulas locking up a spreadsheet application.
        $000D: ; //(CALCMODE) This record is part of the Calculation Settings Block. It specifies whether to calculate formulas manually, automatically or automatically except for multiple table operations.
        $000F: ; //(REFMODE) This record is part of the Calculation Settings Block. It stores which method is used to show cell addresses in formulas.
        $0010: ; //(DELTA) This record is part of the Calculation Settings Block. It stores the maximum change of the result to exit an iteration.
        $0011: ; //(ITERATION) This record is part of the Calculation Settings Block. It stores if iterations are allowed while calculating recursive formulas.
        $0014: ; //(HEADER) This record is part of the Page Settings Block. It specifies the page header string for the current worksheet. If this record is not present or completely empty (record size is 0), the sheet does not contain a page header.
        $0015: ; //(FOOTER) This record is part of the Page Settings Block. It specifies the page footer string for the current worksheet. If this record is not present or completely empty (record size is 0), the sheet does not contain a page footer.
        $001D: ; //(SELECTION) This record contains the addresses of all selected cell ranges and the position of the active cell for a pane in the current sheet.
        $0026: ; //(LEFTMARGIN) This record is part of the Page Settings Block. It contains the left page margin of the current worksheet.
        $0027: ; //(RIGHTMARGIN) This record is part of the Page Settings Block. It contains the right page margin of the current worksheet.
        $0028: ; //(TOPMARGIN) This record is part of the Page Settings Block. It contains the top page margin of the current worksheet.
        $0029: ; //(BOTTOMMARGIN) This record is part of the Page Settings Block. It contains the bottom page margin of the current worksheet.
        $002A: ; //(PRINTHEADERS) This record stores if the row and column headers (the areas with row numbers and column letters) will be printed.
        $002B: ; //(PRINTGRIDLINES) This record stores if sheet grid lines will be printed.
        $0055: ; //(DEFCOLWIDTH) This record specifies the default column width for columns that do not have a specific width set using the records COLWIDTH (BIFF2), COLINFO (BIFF3-BIFF8), or STANDARDWIDTH.
        $005F: ; //(SAVERECALC) This record is part of the Calculation Settings Block. It contains the “Recalculate before save” option in Excel's calculation settings dialogue.
        $007D: ; //(COLINFO) This record specifies the width and default cell formatting for a given range of columns.
        $0080: ; //(GUTS) This record contains information about the layout of outline symbols.
        $0081: ; //(SHEETPR) This record stores a 16-bit value with Boolean options for the current sheet. From BIFF5 on the “Save external linked values” option is moved to the record BOOKBOOL. This record is also used to distinguish standard sheets from dialogue sheets.
        $0082: ; //(GRIDSET) This record specifies if the option to print sheet grid lines (record PRINTGRIDLINES) has ever been changed.
        $0083: ; //(HCENTER) This record is part of the Page Settings Block. It specifies if the sheet is centred horizontally when printed.
        $0084: ; //(VCENTER) This record is part of the Page Settings Block. It specifies if the sheet is centred vertically when printed.
        $008C: ; //(COUNTRY) This record stores two Windows country identifiers. The first represents the user interface language of the Excel version that has saved the file, and the second represents the system regional settings at the time the file was saved.
        $00A1: ; //(PAGESETUP) This record is part of the Page Settings Block. It stores the page format settings of the current sheet. The pages may be scaled in percent or by using an absolute number of pages.
        $00BE: ; //(MULBLANK) This record represents a cell range of empty cells. All cells are located in the same row.
        $0200: ; //(DIMENSION) This record contains the range address of the used area in the current sheet.
        $0201: ; //(BLANK) This record represents an empty cell. It contains the cell address and formatting information.
        $0208: ; //(ROW) This record contains the properties of a single row in a sheet. Rows and cells in a sheet are divided into blocks of 32 rows.
        $0225: ; //(DEFAULTROWHEIGHT) This record specifies the default height and default flags for rows that do not have a corresponding ROW record.
        $023E: ; //(WINDOW2) This record contains the range address of the used area in the current sheet.
      else
        WriteLn(format('Record type: %.4X Record Size: %.4X',[RecordType,RecordSize]));
      end;
      {$ENDIF}
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then SectionEOF := True;
  end;

  FixCols(FWorksheet);
  FixRows(FWorksheet);
end;

procedure TsSpreadBIFF5Reader.ReadBoundsheet(AStream: TStream);
var
  Len: Byte;
  s: AnsiString;
  sheetName: String;
begin
  { Absolute stream position of the BOF record of the sheet represented
    by this record }
  // Just assume that they are in order
  AStream.ReadDWord();

  { Visibility }
  AStream.ReadByte();

  { Sheet type }
  AStream.ReadByte();

  { Sheet name: Byte string, 8-bit length }
  Len := AStream.ReadByte();

  SetLength(s, Len);
  AStream.ReadBuffer(s[1], Len*SizeOf(AnsiChar));
  sheetName := AnsiToUTF8(s);
  FWorksheetNames.Add(sheetName);
end;

procedure TsSpreadBIFF5Reader.ReadRichString(AStream: TStream);
var
  L: Word;
  B: BYTE;
  ARow, ACol: Cardinal;
  XF: Word;
  AStrValue: ansistring;
  cell: PCell;
begin
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Byte String with 16-bit size }
  L := WordLEtoN(AStream.ReadWord());
  SetLength(AStrValue,L);
  AStream.ReadBuffer(AStrValue[1], L);

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);

  { Save the data }
  FWorksheet.WriteUTF8Text(cell, ISO_8859_1ToUTF8(AStrValue));
  //Read formatting runs (not supported)
  B := AStream.ReadByte;
  for L := 0 to B-1 do begin
    AStream.ReadByte; // First formatted character
    AStream.ReadByte; // Index to FONT record
  end;

  { Add attributes to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{ Reads the default column width that is used when a bit in the GCW bit structure
  is set for the corresponding column. The GCW is ignored here. The column
  width read from the STANDARDWIDTH record overrides the one from the
  DEFCOLWIDTH record. }
procedure TsSpreadBIFF5Reader.ReadStandardWidth(AStream: TStream; ASheet: TsWorksheet);
var
  w: Word;
begin
  // read width in 1/256 of the width of "0" character
  w := WordLEToN(AStream.ReadWord);
  // calculate width in units of "characters" and use it as DefaultColWidth
  ASheet.DefaultColWidth := w / 256;
end;

{ Reads a STRING record which contains the result of string formula. }
procedure TsSpreadBIFF5Reader.ReadStringRecord(AStream: TStream);
var
  len: Word;
  s: ansistring;
begin
  // The string is a byte-string with 16 bit length
  len := WordLEToN(AStream.ReadWord);
  if len > 0 then begin
    SetLength(s, Len);
    AStream.ReadBuffer(s[1], len);
    if (FIncompleteCell <> nil) and (s <> '') then begin
      FIncompleteCell^.UTF8StringValue := AnsiToUTF8(s);
      FIncompleteCell^.ContentType := cctUTF8String;
      if FIsVirtualMode then
        Workbook.OnReadCellData(Workbook, FIncompleteCell^.Row, FIncompleteCell^.Col, FIncompleteCell);
    end;
  end;
  FIncompleteCell := nil;
end;

procedure TsSpreadBIFF5Reader.ReadFromFile(AFileName: string; AData: TsWorkbook);
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
    OLEStorage.ReadOLEFile(AFileName, OLEDocument);

    // Check if the operation succeded
    if MemStream.Size = 0 then raise Exception.Create('[TsSpreadBIFF5Reader.ReadFromFile] Reading of OLE document failed');

    // Rewind the stream and read from it
    MemStream.Position := 0;
    ReadFromStream(MemStream, AData);

//    Uncomment to verify if the data was correctly obtained from the OLE file
//    MemStream.SaveToFile(SysUtils.ChangeFileExt(AFileName, 'bin.xls'));
  finally
    MemStream.Free;
    OLEStorage.Free;
  end;
end;

procedure TsSpreadBIFF5Reader.ReadXF(AStream: TStream);
var
  rec: TBIFF5_XFRecord;
  fmt: TsCellFormat;
  nfidx: Integer;
  i: Integer;
  nfdata: TsNumFormatData;
  b: Byte;
  dw: DWord;
  fill: Word;
begin
  InitFormatRecord(fmt);
  fmt.ID := FCellFormatList.Count;

  // Read the complete xf record into a buffer
  rec.FontIndex := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.FontIndex, SizeOf(rec) - 2*SizeOf(Word));

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

  // Text rotation
  case rec.TextOrient_UnusedAttrib and MASK_XF_ORIENTATION of
    XF_ROTATION_HORIZONTAL : fmt.TextRotation := trHorizontal;
    XF_ROTATION_90DEG_CCW  : fmt.TextRotation := rt90DegreeCounterClockwiseRotation;
    XF_ROTATION_90DEG_CW   : fmt.TextRotation := rt90DegreeClockwiseRotation;
    XF_ROTATION_STACKED    : fmt.TextRotation := rtStacked;
  end;
  if fmt.TextRotation <> trHorizontal then
    Include(fmt.UsedFormattingFields, uffTextRotation);

  // Cell borders and background
  rec.Border_BkGr1 := DWordLEToN(rec.Border_BkGr1);
  rec.Border_BkGr2 := DWordLEToN(rec.Border_BkGr2);
  // The 4 masked bits encode the line style of the border line. 0 = no line.
  // The case of "no line" is not included in the TsLineStyle enumeration.
  // --> correct by subtracting 1!
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_BOTTOM;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbSouth);
    fmt.BorderStyles[cbSouth].LineStyle := TsLineStyle(dw shr 22 - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr2 and MASK_XF_BORDER_LEFT;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbWest);
    fmt.BorderStyles[cbWest].LineStyle := TsLineStyle(dw shr 3 - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr2 and MASK_XF_BORDER_RIGHT;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbEast);
    fmt.BorderStyles[cbEast].LineStyle := TsLineStyle(dw shr 6 - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr2 and MASK_XF_BORDER_TOP;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbNorth);
    fmt.BorderStyles[cbNorth].LineStyle := TsLineStyle(dw - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;

  // Border line colors
  fmt.BorderStyles[cbWest].Color := (rec.Border_BkGr2 and MASK_XF_BORDER_LEFT_COLOR) shr 16;
  fmt.BorderStyles[cbEast].Color := (rec.Border_BkGr2 and MASK_XF_BORDER_RIGHT_COLOR) shr 23;
  fmt.BorderStyles[cbNorth].Color := (rec.Border_BkGr2 and MASK_XF_BORDER_TOP_COLOR) shr 9;
  fmt.BorderStyles[cbSouth].Color := (rec.Border_BkGr1 and MASK_XF_BORDER_BOTTOM_COLOR) shr 25;

  // Background fill style
  fill := (rec.Border_BkGr1 and MASK_XF_BKGR_FILLPATTERN) shr 16;

  // Background color
  if fill = 0 then
    fmt.BackgroundColor := scTransparent
  else begin
    fmt.BackgroundColor := rec.Border_BkGr1 and MASK_XF_BKGR_PATTERN_COLOR;
    Include(fmt.UsedFormattingFields, uffBackgroundColor);
  end;

  // Add the XF to the list
  FCellFormatList.Add(fmt);
end;

procedure TsSpreadBIFF5Reader.ReadFromStream(AStream: TStream; AData: TsWorkbook);
var
  BIFF5EOF: Boolean;
  p,s: Int64;
begin
  { Initializations }

  FWorksheetNames := TStringList.Create;
  FWorksheetNames.Clear;
  FCurrentWorksheet := 0;
  BIFF5EOF := False;

  { Read workbook globals }

  ReadWorkbookGlobals(AStream, AData);

  // Check for the end of the file
  if AStream.Position >= AStream.Size then BIFF5EOF := True;

  { Now read all worksheets }

  while (not BIFF5EOF) do
  begin
    ReadWorksheet(AStream, AData);

    // Check for the end of the file
    p := AStream.Position;
    s := AStream.Size;
    if AStream.Position >= AStream.Size then BIFF5EOF := True;

    // Final preparations
    Inc(FCurrentWorksheet);
    if FCurrentWorksheet = FWorksheetNames.Count then BIFF5EOF := True;
    // It can happen in files written by Office97 that the OLE directory is
    // at the end of the file.
  end;

  if not FPaletteFound then
    FWorkbook.UsePalette(@PALETTE_BIFF5, Length(PALETTE_BIFF5));

  { Finalizations }

  FWorksheetNames.Free;
end;

procedure TsSpreadBIFF5Reader.ReadFont(const AStream: TStream);
var
  {%H-}lCodePage: Word;
  lHeight: Word;
  lOptions: Word;
  lColor: Word;
  lWeight: Word;
  Len: Byte;
  fontname: ansistring;
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

  { Escapement type }
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

  { Font name: Ansistring, char count in 1 byte }
  Len := AStream.ReadByte();
  SetLength(fontname, Len);
  AStream.ReadBuffer(fontname[1], Len);
  font.FontName := fontname;

  { Add font to workbook's font list }
  FWorkbook.AddFont(font);
end;

// Read the FORMAT record for formatting numerical data
procedure TsSpreadBIFF5Reader.ReadFormat(AStream: TStream);
var
  len: byte;
  fmtIndex: Integer;
  fmtString: AnsiString;
begin
  // Record FORMAT, BIFF 8 (5.49):
  // Offset Size Contents
  // 0      2     Format index used in other records
  // 2      var   Number format string (byte string, 8-bit string length)
  // From BIFF5 on: indexes 0..163 are built in

  // format index
  fmtIndex := WordLEtoN(AStream.ReadWord);

  // number format string
  len := AStream.ReadByte;
  SetLength(fmtString, len);
  AStream.ReadBuffer(fmtString[1], len);

  // Add to the list
  NumFormatList.AnalyzeAndAdd(fmtIndex, AnsiToUTF8(fmtString));
end;

procedure TsSpreadBIFF5Reader.ReadLabel(AStream: TStream);
var
  rec: TBIFF5_LabelRecord;
  L: Word;
  ARow, ACol: Cardinal;
  XF: WORD;
  cell: PCell;
  AValue: ansistring;
begin
  rec.Row := 0;  // to silence the compiler...

  { Read entire record, starting at Row, except for string data }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF5_LabelRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);

  { Byte String with 16-bit size }
  L := WordLEToN(rec.TextLen);
  SetLength(AValue, L);
  AStream.ReadBuffer(AValue[1], L);

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.GetCell(ARow, ACol);

  { Save the data }
  FWorksheet.WriteUTF8Text(cell, ISO_8859_1ToUTF8(AValue));

  { Add attributes }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;


{ TsSpreadBIFF5Writer }

{@@ ----------------------------------------------------------------------------
  Writes an Excel BIFF5 file to the disc

  The BIFF 5 writer overrides this method because BIFF 5 is written as
  an OLE document, and our current OLE document writing method involves:

     1 - Writing the BIFF data to a memory stream
     2 - Write the memory stream data to disk using COM functions
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteToFile(const AFileName: string;
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

    OutputStorage.WriteOLEFile(AFileName, OLEDocument, AOverwriteExisting);
  finally
    Stream.Free;
    OutputStorage.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel BIFF5 record structure

  Be careful as this method doesn't write the OLE part of the document,
  just the BIFF records
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteToStream(AStream: TStream);
var
  CurrentPos: Int64;
  Boundsheets: array of Int64;
  i, len: Integer;
  pane: Byte;
begin
  { Store some data about the workbook that other routines need }
  WorkBookEncoding := Workbook.Encoding;

  { Write workbook globals }

  WriteBOF(AStream, INT_BOF_WORKBOOK_GLOBALS);

  WriteCodepage(AStream, WorkBookEncoding);
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
    AStream.WriteDWord(CurrentPos);
    AStream.Position := CurrentPos;

    WriteBOF(AStream, INT_BOF_SHEET);

      WriteIndex(AStream);
//      WritePageSetup(AStream);
      WriteColInfos(AStream, FWorksheet);
      WriteDimensions(AStream, FWorksheet);
      WriteWindow2(AStream, FWorksheet);
      WritePane(AStream, FWorksheet, true, pane);  // true for "is BIFF5 or BIFF8"
      WriteSelection(AStream, FWorksheet, pane);
      //WriteRows(AStream, sheet);

      if (boVirtualMode in Workbook.Options) then
        WriteVirtualCells(AStream)
      else begin
        WriteRows(AStream, FWorksheet);
        WriteCellsToStream(AStream, FWorksheet.Cells);
      end;

    WriteEOF(AStream);
  end;

  { Cleanup }

  SetLength(Boundsheets, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 BOF record

  This must be the first record of an Excel 5 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteBOF(AStream: TStream; ADataType: Word);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BOF));
  AStream.WriteWord(WordToLE(8));

  { BIFF version. Should only be used if this BOF is for the workbook globals }
  if ADataType = INT_BOF_WORKBOOK_GLOBALS then
   AStream.WriteWord(WordToLE(INT_BOF_BIFF5_VER))
  else AStream.WriteWord(0);

  { Data type }
  AStream.WriteWord(WordToLE(ADataType));

  { Build identifier, must not be 0 }
  AStream.WriteWord(WordToLE(INT_BOF_BUILD_ID));

  { Build year, must not be 0 }
  AStream.WriteWord(WordToLE(INT_BOF_BUILD_YEAR));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 BOUNDSHEET record

  Always located on the workbook globals substream.
  One BOUNDSHEET is written for each worksheet.

  @return   The stream position where the absolute stream position
            of the BOF of this sheet should be written (4 bytes size).
-------------------------------------------------------------------------------}
function TsSpreadBIFF5Writer.WriteBoundsheet(AStream: TStream; ASheetName: string): Int64;
var
  Len: Byte;
  LatinSheetName: string;
begin
  LatinSheetName := UTF8ToISO_8859_1(ASheetName);
  Len := Length(LatinSheetName);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BOUNDSHEET));
  AStream.WriteWord(WordToLE(6 + 1 + Len));

  { Absolute stream position of the BOF record of the sheet represented
    by this record }
  Result := AStream.Position;
  AStream.WriteDWord(WordToLE(0));

  { Visibility }
  AStream.WriteByte(0);

  { Sheet type }
  AStream.WriteByte(0);

  { Sheet name: Byte string, 8-bit length }
  AStream.WriteByte(Len);
  AStream.WriteBuffer(LatinSheetName[1], Len);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 DIMENSIONS record

  nm = (rl - rf - 1) / 32 + 1 (using integer division)

  Excel, OpenOffice and FPSpreadsheet ignore the dimensions written in this
  record, but some other applications really use them, so they need to be correct.

  See bug 18886: excel5 files are truncated when imported
--------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
var
  rec: TBIFF5_DimensionsRecord;
  firstCol, lastCol, firstRow, lastRow: Cardinal;
begin
  { Determine sheet size }
  GetSheetDimensions(AWorksheet, firstRow, lastRow, firstCol, lastCol);

  { Setup BIFF record }
  rec.RecordID := WordToLE(INT_EXCEL_ID_DIMENSIONS);
  rec.RecordSize := WordToLE(10);
  rec.FirstRow := WordToLE(firstRow);
  if lastRow < $FFFF then   // avoid WORD overflow
    rec.LastRowPlus1 := WordToLE(lastRow + 1)
  else
    rec.LastRowPlus1 := $FFFF;
  rec.FirstCol := WordToLe(firstCol);
  rec.LastColPlus1 := WordToLE(lastCol+1);
  rec.NotUsed := 0;

  { Write BIFF record }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 EOF record
  This must be the last record of an Excel 5 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_EOF));
  AStream.WriteWord($0000);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 FONT record
  The font data is passed as an instance of TsFont.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteFont(AStream: TStream; AFont: TsFont);
var
  Len: Byte;
  optn: Word;
begin
  if AFont = nil then  // this happens for FONT4 in case of BIFF
    exit;

  if AFont.FontName = '' then
    raise Exception.Create('Font name not specified.');
  if AFont.Size <= 0.0 then
    raise Exception.Create('Font size not specified.');

  Len := Length(AFont.FontName);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FONT));
  AStream.WriteWord(WordToLE(14 + 1 + Len));

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
  AStream.WriteWord(0);

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

  { Font name: Byte string, 8-bit length }
  AStream.WriteByte(Len);
  AStream.WriteBuffer(AFont.FontName[1], Len);
end;

{@@ ----------------------------------------------------------------------------
  Writes the Excel 5 FONT records needed for the used fonts in the workbook.
-------------------------------------------------------------------------------}
procedure TsSpreadBiff5Writer.WriteFonts(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to Workbook.GetFontCount-1 do
    WriteFont(AStream, Workbook.GetFont(i));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 FORMAT record which is needed for formatting of numerical
  data.
-------------------------------------------------------------------------------}
procedure TsSpreadBiff5Writer.WriteNumFormat(AStream: TStream;
  ANumFormatData: TsNumFormatData; AListIndex: Integer);
type
  TNumFormatRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FormatIndex: Word;
    FormatStringLen: Byte;
  end;
var
  len: Integer;
  s: ansistring;
  rec: TNumFormatRecord;
  buf: array of byte;
begin
  if (ANumFormatData = nil) or (ANumFormatData.FormatString = '') then
    exit;

  s := UTF8ToAnsi(NumFormatList.FormatStringForWriting(AListIndex));
  len := Length(s);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_FORMAT);
  rec.RecordSize := WordToLE(2 + 1 + len * SizeOf(AnsiChar));

  { Format index }
  rec.FormatIndex := WordToLE(ANumFormatData.Index);

  { Format string }
  { Length in 1 byte }
  rec.FormatStringLen := len;
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
  Writes an Excel 5 INDEX record

  nm = (rl - rf - 1) / 32 + 1 (using integer division)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteIndex(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_INDEX));
  AStream.WriteWord(WordToLE(12));

  { Not used }
  AStream.WriteDWord(0);

  { Index to first used row, rf, 0 based }
  AStream.WriteWord(0);

  { Index to first row of unused tail of sheet, rl, last used row + 1, 0 based }
  AStream.WriteWord(33);

  { Absolute stream position of the DEFCOLWIDTH record of the current sheet.
    If it doesn't exist, the offset points to where it would occur. }
  AStream.WriteDWord($00);

  { Array of nm absolute stream positions of the DBCELL record of each Row Block }

  { OBS: It seams to be no problem just ignoring this part of the record }
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 LABEL record

  If the string length exceeds 255 bytes, the string will be truncated and
  an error message will be logged as a warning.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  MAXBYTES = 255; //limit for this format
var
  L: Word;
  AnsiValue: ansistring;
  rec: TBIFF5_LabelRecord;
  buf: array of byte;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  case WorkBookEncoding of
    seLatin2:   AnsiValue := UTF8ToCP1250(AValue);
    seCyrillic: AnsiValue := UTF8ToCP1251(AValue);
    seGreek:    AnsiValue := UTF8ToCP1253(AValue);
    seTurkish:  AnsiValue := UTF8ToCP1254(AValue);
    seHebrew:   AnsiValue := UTF8ToCP1255(AValue);
    seArabic:   AnsiValue := UTF8ToCP1256(AValue);
  else
    // Latin 1 is the default
    AnsiValue := UTF8ToCP1252(AValue);
  end;

  if AnsiValue = '' then begin
    // Bad formatted UTF8String (maybe ANSI?)
    if Length(AValue) <> 0 then begin
      //It was an ANSI string written as UTF8 quite sure, so raise exception.
      Raise Exception.CreateFmt(rsUTF8TextExpectedButANSIFoundInCell, [
        GetCellString(ARow, ACol)
      ]);
    end;
    Exit;
  end;

  if Length(AnsiValue) > MAXBYTES then begin
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    AnsiValue := Copy(AnsiValue, 1, MAXBYTES);
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
      MAXBYTES, GetCellString(ARow, ACol)
    ]);
  end;
  L := Length(AnsiValue);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_LABEL);
  rec.RecordSize := WordToLE(8 + L);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { Index to XF record }
  rec.XFIndex := WordToLE(FindXFIndex(ACell));

  { String length, 16 bit }
  rec.TextLen := WordToLE(L);

  { Copy the text characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + SizeOf(ansiChar)*L);
  Move(rec, buf[0], SizeOf(rec));
  Move(AnsiValue[1], buf[SizeOf(rec)], L*SizeOf(ansiChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(Rec) + SizeOf(ansiChar)*L);

  { Clean up }
  SetLength(buf, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 STRING record which immediately follows a FORMULA record
  when the formula result is a string.
  BIFF5 writes a byte-string, but uses a 16-bit length here!
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteStringRecord(AStream: TStream;
  AString: String);
var
  s: ansistring;
  len: Integer;
begin
  s := UTF8ToAnsi(AString);
  len := Length(s);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_STRING));
  AStream.WriteWord(WordToLE(2 + len*SizeOf(Char)));

  { Write string length }
  AStream.WriteWord(WordToLE(len));
  { Write characters }
  AStream.WriteBuffer(s[1], len * SizeOf(Char));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 STYLE record

  Registers the name of a user-defined style or specific options for
  a built-in cell style.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteStyle(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_STYLE));
  AStream.WriteWord(WordToLE(4));

  { Index to style XF and defines if it's a built-in or used defined style }
  AStream.WriteWord(WordToLE(MASK_STYLE_BUILT_IN));

  { Built-in cell style identifier }
  AStream.WriteByte($00);

  { Level if the identifier for a built-in style is RowLevel or ColLevel, $FF otherwise }
  AStream.WriteByte(WordToLE($FF));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 WINDOW2 record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteWindow2(AStream: TStream;
  ASheet: TsWorksheet);
var
  Options: Word;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW2));
  AStream.WriteWord(WordToLE(10));

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

  { Grid line RGB colour }
  AStream.WriteDWord(DWordToLE(0));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 5 XF record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF5Writer.WriteXF(AStream: TStream;
  AFormatRecord: PsCellFormat; XFType_Prot: Byte = 0);
var
  rec: TBIFF5_XFRecord;
  j: Integer;
  b: Byte;
  dw1, dw2: DWord;
begin
  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_XF);
  rec.RecordSize := WordToLE(SizeOf(TBIFF5_XFRecord) - 2*SizeOf(Word));

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
        haLeft   : b := b or MASK_XF_HOR_ALIGN_LEFT;
        haCenter : b := b or MASK_XF_HOR_ALIGN_CENTER;
        haRight  : b := b or MASK_XF_HOR_ALIGN_RIGHT;
        haDefault: ;
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
  rec.TextOrient_UnusedAttrib := 0;
  if (AFormatRecord <> nil) and (uffTextRotation in AFormatRecord^.UsedFormattingFields)
    then rec.TextOrient_UnusedAttrib := TEXT_ROTATIONS[AFormatRecord^.TextRotation];

  { Cell border lines and background area }
  dw1 := 0;
  dw2 := 0;
  if (AFormatRecord <> nil) then
  begin
    if (uffBackgroundColor in AFormatRecord^.UsedFormattingFields) then
    begin
      // Background color
      dw1 := dw1 or (FixColor(AFormatRecord^.BackgroundColor) and $0000007F);
      dw1 := dw1 or (MASK_XF_FILL_PATT_SOLID shl 16);
    end;
    // Border lines
    if (uffBorder in AFormatRecord^.UsedFormattingFields) then
    begin
      dw1 := dw1 or (AFormatRecord^.BorderStyles[cbSouth].Color shl 25);      // Bottom line color
      dw2 := (FixColor(AFormatRecord^.BorderStyles[cbNorth].Color) shl 9) or  // Top line color
             (FixColor(AFormatRecord^.BorderStyles[cbWest].Color) shl 16) or  // Left line color
             (FixColor(AFormatRecord^.BorderStyles[cbEast].Color) shl 23);    // Right line color
      if cbSouth in AFormatRecord^.Border then
        dw1 := dw1 or ((DWord(AFormatRecord^.BorderStyles[cbSouth].LineStyle)+1) shl 22);
      if cbNorth in AFormatRecord^.Border then
        dw2 := dw2 or  (DWord(AFormatRecord^.BorderStyles[cbNorth].LineStyle)+1);
      if cbWest in AFormatRecord^.Border then
        dw2 := dw2  or ((DWord(AFormatRecord^.BorderStyles[cbWest].LineStyle)+1) shl 3);
      if cbEast in AFormatRecord^.Border then
        dw2 := dw2  or ((DWord(AFormatRecord^.BorderStyles[cbEast].LineStyle)+1) shl 6);
    end;
  end;
  rec.Border_BkGr1 := dw1;
  rec.Border_BkGr2 := dw2;

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;


initialization

  RegisterSpreadFormat(TsSpreadBIFF5Reader, TsSpreadBIFF5Writer, sfExcel5);
  MakeLEPalette(@PALETTE_BIFF5, Length(PALETTE_BIFF5));

end.

