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

AUTHORS:  Felipe Monteiro de Carvalho
          Jose Mejuto
}
unit xlsbiff8;

{$ifdef fpc}
  {$mode delphi}
{$endif}

// The new OLE code is much better, so always use it
{$define USE_NEW_OLE}
{.$define FPSPREADDEBUG} //used to be XLSDEBUG

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
    function DecodeRKValue(const ARK: DWORD): Double;
    function ReadWideString(const AStream: TStream; const ALength: WORD): WideString; overload;
    function ReadWideString(const AStream: TStream; const AUse8BitLength: Boolean): WideString; overload;
    procedure ReadWorkbookGlobals(AStream: TStream; AData: TsWorkbook);
    procedure ReadWorksheet(AStream: TStream; AData: TsWorkbook);
    procedure ReadBoundsheet(AStream: TStream);
    procedure ReadRKValue(const AStream: TStream);
    procedure ReadMulRKValues(const AStream: TStream);
    function ReadString(const AStream: TStream; const ALength: WORD): UTF8String;
    procedure ReadSST(const AStream: TStream);
    procedure ReadLabelSST(const AStream: TStream);
    // Read XF record
    procedure ReadXF(const AStream: TStream);
    // Workbook Globals records
    // procedure ReadCodepage in xlscommon
    // procedure ReadDateMode in xlscommon
    procedure ReadFont(const AStream: TStream);
    procedure ReadFormat(AStream: TStream); override;
    { Record reading methods }
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadRichString(const AStream: TStream);
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
    procedure WriteXFIndex(AStream: TStream; ACell: PCell);
    procedure WriteXFFieldsForFormattingStyles(AStream: TStream);
  protected
    { Record writing methods }
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBOF(AStream: TStream; ADataType: Word);
    function  WriteBoundsheet(AStream: TStream; ASheetName: string): Int64;
//    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
//      const AValue: TDateTime; ACell: PCell); override;
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFont: TsFont);
    procedure WriteFonts(AStream: TStream);
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      const AFormula: TsFormula; ACell: PCell); override;
    procedure WriteIndex(AStream: TStream);
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal;
      const AFormula: TsRPNFormula; ACell: PCell); override;
    procedure WriteStyle(AStream: TStream);
    procedure WriteWindow1(AStream: TStream);
    procedure WriteWindow2(AStream: TStream; ASheetSelected: Boolean);
    procedure WriteXF(AStream: TStream; AFontIndex: Word;
      AFormatIndex: Word; AXF_TYPE_PROT, ATextRotation: Byte; ABorders: TsCellBorders;
      AHorAlignment: TsHorAlignment = haDefault; AVertAlignment: TsVertAlignment = vaDefault;
      AWordWrap: Boolean = false; AddBackground: Boolean = false;
      ABackgroundColor: TsColor = scSilver);
    procedure WriteXFRecords(AStream: TStream);
  public
    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;

var
  // the palette of the default BIFF8 colors as "big-endian color" values
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

const
  { Excel record IDs }
  INT_EXCEL_ID_BLANK      = $0201;
  INT_EXCEL_ID_BOF        = $0809;
  INT_EXCEL_ID_BOUNDSHEET = $0085; // Renamed to SHEET in the latest OpenOffice docs
  INT_EXCEL_ID_COUNTRY    = $008C;
  INT_EXCEL_ID_EOF        = $000A;
  INT_EXCEL_ID_DIMENSIONS = $0200;
  INT_EXCEL_ID_FORMULA    = $0006;
  INT_EXCEL_ID_INDEX      = $020B;
  INT_EXCEL_ID_LABEL      = $0204;
  INT_EXCEL_ID_NUMBER     = $0203;
  INT_EXCEL_ID_ROWINFO    = $0208;
  INT_EXCEL_ID_STYLE      = $0293;
  INT_EXCEL_ID_WINDOW1    = $003D;
  INT_EXCEL_ID_WINDOW2    = $023E;
  INT_EXCEL_ID_RSTRING    = $00D6;
  INT_EXCEL_ID_RK         = $027E;
  INT_EXCEL_ID_MULRK      = $00BD;
  INT_EXCEL_ID_SST        = $00FC; //BIFF8 only
  INT_EXCEL_ID_CONTINUE   = $003C;
  INT_EXCEL_ID_LABELSST   = $00FD; //BIFF8 only
  INT_EXCEL_ID_FORMAT     = $041E;
  INT_EXCEL_ID_FORCEFULLCALCULATION = $08A3;

  { Cell Addresses constants }
  MASK_EXCEL_ROW          = $3FFF;
  MASK_EXCEL_COL_BITS_BIFF8=$00FF;
  MASK_EXCEL_RELATIVE_COL = $4000;  // This is according to Microsoft documentation,
  MASK_EXCEL_RELATIVE_ROW = $8000;  // but opposite to OpenOffice documentation!

  { BOF record constants }
  INT_BOF_BIFF8_VER       = $0600;
  INT_BOF_WORKBOOK_GLOBALS= $0005;
  INT_BOF_VB_MODULE       = $0006;
  INT_BOF_SHEET           = $0010;
  INT_BOF_CHART           = $0020;
  INT_BOF_MACRO_SHEET     = $0040;
  INT_BOF_WORKSPACE       = $0100;
  INT_BOF_BUILD_ID        = $1FD2;
  INT_BOF_BUILD_YEAR      = $07CD;

  { FORMULA record constants }
  MASK_FORMULA_RECALCULATE_ALWAYS  = $0001;
  MASK_FORMULA_RECALCULATE_ON_OPEN = $0002;
  MASK_FORMULA_SHARED_FORMULA      = $0008;

  { STYLE record constants }
  MASK_STYLE_BUILT_IN     = $8000;

  { WINDOW1 record constants }
  MASK_WINDOW1_OPTION_WINDOW_HIDDEN             = $0001;
  MASK_WINDOW1_OPTION_WINDOW_MINIMISED          = $0002;
  MASK_WINDOW1_OPTION_HORZ_SCROLL_VISIBLE       = $0008;
  MASK_WINDOW1_OPTION_VERT_SCROLL_VISIBLE       = $0010;
  MASK_WINDOW1_OPTION_WORKSHEET_TAB_VISIBLE     = $0020;

  { WINDOW2 record constants }
  MASK_WINDOW2_OPTION_SHOW_FORMULAS             = $0001;
  MASK_WINDOW2_OPTION_SHOW_GRID_LINES           = $0002;
  MASK_WINDOW2_OPTION_SHOW_SHEET_HEADERS        = $0004;
  MASK_WINDOW2_OPTION_PANES_ARE_FROZEN          = $0008;
  MASK_WINDOW2_OPTION_SHOW_ZERO_VALUES          = $0010;
  MASK_WINDOW2_OPTION_AUTO_GRIDLINE_COLOR       = $0020;
  MASK_WINDOW2_OPTION_COLUMNS_RIGHT_TO_LEFT     = $0040;
  MASK_WINDOW2_OPTION_SHOW_OUTLINE_SYMBOLS      = $0080;
  MASK_WINDOW2_OPTION_REMOVE_SPLITS_ON_UNFREEZE = $0100;
  MASK_WINDOW2_OPTION_SHEET_SELECTED            = $0200;
  MASK_WINDOW2_OPTION_SHEET_ACTIVE              = $0400;

  { XF substructures }

  { XF_ROTATION }
  XF_ROTATION_HORIZONTAL              = 0;
  XF_ROTATION_90DEG_CCW               = 90;
  XF_ROTATION_90DEG_CW                = 180;
  XF_ROTATION_STACKED                 = 255;   // Letters stacked top to bottom, but not rotated

  { XF CELL BORDER }
  MASK_XF_BORDER_LEFT                 = $0000000F;
  MASK_XF_BORDER_RIGHT                = $000000F0;
  MASK_XF_BORDER_TOP                  = $00000F00;
  MASK_XF_BORDER_BOTTOM               = $0000F000;


{ TsSpreadBIFF8Writer }

{ Index to XF record, according to formatting }
procedure TsSpreadBIFF8Writer.WriteXFIndex(AStream: TStream; ACell: PCell);
var
  lIndex: Integer;
  lXFIndex: Word;
begin
  // First try the fast methods for default formats
  if ACell^.UsedFormattingFields = [] then
  begin
    AStream.WriteWord(WordToLE(15)); //XF15; see TsSpreadBIFF8Writer.AddDefaultFormats
    Exit;
  end;

{
  if ACell^.UsedFormattingFields = [uffTextRotation] then
  begin
    case ACell^.TextRotation of
      rt90DegreeCounterClockwiseRotation: AStream.WriteWord(WordToLE(16)); //XF_16
      rt90DegreeClockwiseRotation: AStream.WriteWord(WordToLE(17)); //XF_17
    else
      AStream.WriteWord(WordToLE(15)); //XF_15
    end;
    Exit;
  end;
 }

  {
  uffNumberFormat does not seem to have default XF indexes, but perhaps look at XF_21
  if ACell^.UsedFormattingFields = [uffNumberFormat] then
  begin
    case ACell^.NumberFormat of
      nfShortDate:     AStream.WriteWord(WordToLE(???)); //what XF index?
      nfShortDateTime: AStream.WriteWord(WordToLE(???)); //what XF index?
    else
      AStream.WriteWord(WordToLE(15)); //e.g. nfGeneral: XF_15
    end;
    Exit;
  end;
  }

  // If not, then we need to search in the list of dynamic formats
  lIndex := FindFormattingInList(ACell);
  // Carefully check the index
  if (lIndex < 0) or (lIndex > Length(FFormattingStyles)) then
    raise Exception.Create('[TsSpreadBIFF8Writer.WriteXFIndex] Invalid Index, this should not happen!');

  lXFIndex := FFormattingStyles[lIndex].Row;

  AStream.WriteWord(WordToLE(lXFIndex));
end;

procedure TsSpreadBIFF8Writer.WriteXFFieldsForFormattingStyles(AStream: TStream);
var
  i: Integer;
  lFontIndex: Word;
  lFormatIndex: Word; //number format
  lTextRotation: Byte;
  lBorders: TsCellBorders;
  lAddBackground: Boolean;
  lBackgroundColor: TsColor;
  lHorAlign: TsHorAlignment;
  lVertAlign: TsVertAlignment;
  lWordWrap: Boolean;
  fmt: String;
begin
  // The first style was already added
  for i := 1 to Length(FFormattingStyles) - 1 do begin
    // Default styles
    lFontIndex := 0;
    lFormatIndex := 0; //General format (one of the built-in number formats)
    lTextRotation := XF_ROTATION_HORIZONTAL;
    lBorders := [];
    lHorAlign := FFormattingStyles[i].HorAlignment;
    lVertAlign := FFormattingStyles[i].VertAlignment;
    lBackgroundColor := FFormattingStyles[i].BackgroundColor;

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
            if (fmt = 'ms') or (fmt = 'nn:ss') or (fmt = 'mm:ss') then
              lFormatIndex := FORMAT_TIME_MS
            else
            if (fmt = 'msz') or (fmt = 'nn:ss.zzz') or (fmt = 'mm:ss.zzz') or (fmt = 'mm:ss.0') or (fmt = 'mm:ss.z') or (fmt = 'nn:ss.z') then
              lFormatIndex := FORMAT_TIME_MSZ
          end;
        nfTimeInterval:
          lFormatIndex := FORMAT_TIME_INTERVAL;
      end;

    if uffBorder in FFormattingStyles[i].UsedFormattingFields then
      lBorders := FFormattingStyles[i].Border;

    if uffTextRotation in FFormattingStyles[i].UsedFormattingFields then
    begin
      case FFormattingStyles[i].TextRotation of
      trHorizontal:                       lTextRotation := XF_ROTATION_HORIZONTAL;
      rt90DegreeClockwiseRotation:        lTextRotation := XF_ROTATION_90DEG_CW;
      rt90DegreeCounterClockwiseRotation: lTextRotation := XF_ROTATION_90DEG_CCW;
      rtStacked:                          lTextRotation := XF_ROTATION_STACKED;
      end;
    end;

    if uffBold in FFormattingStyles[i].UsedFormattingFields then
      lFontIndex := 1;   // must be before uffFont which overrides uffBold

    if uffFont in FFormattingStyles[i].UsedFormattingFields then
      lFontIndex := FFormattingStyles[i].FontIndex;

    lAddBackground := (uffBackgroundColor in FFormattingStyles[i].UsedFormattingFields);
    lWordwrap := (uffWordwrap in FFormattingStyles[i].UsedFormattingFields);

    // And finally write the style
    WriteXF(AStream, lFontIndex, lFormatIndex, 0, lTextRotation, lBorders,
      lHorAlign, lVertAlign, lWordwrap, lAddBackground, lBackgroundColor);
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
  MemStream: TMemoryStream;
  OutputStorage: TOLEStorage;
  OLEDocument: TOLEDocument;
begin
  MemStream := TMemoryStream.Create;
  OutputStorage := TOLEStorage.Create;
  try
    WriteToStream(MemStream);

    // Only one stream is necessary for any number of worksheets
    OLEDocument.Stream := MemStream;

    OutputStorage.WriteOLEFile(AFileName, OLEDocument, AOverwriteExisting, 'Workbook');
  finally
    MemStream.Free;
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
var
  MyData: TMemoryStream;
  CurrentPos: Int64;
  Boundsheets: array of Int64;
  sheet: TsWorksheet;
  i, j, len: Integer;
  col: PCol;
begin
  { Write workbook globals }

  WriteBOF(AStream, INT_BOF_WORKBOOK_GLOBALS);

  WriteWindow1(AStream);
  WriteFonts(AStream);
  WritePalette(AStream);
  WriteXFRecords(AStream);
  WriteStyle(AStream);

  // A BOUNDSHEET for each worksheet
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
    sheet := Workbook.GetWorksheetByIndex(i);

    { First goes back and writes the position of the BOF of the
      sheet on the respective BOUNDSHEET record }
    CurrentPos := AStream.Position;
    AStream.Position := Boundsheets[i];
    AStream.WriteDWord(DWordToLE(DWORD(CurrentPos)));
    AStream.Position := CurrentPos;

    WriteBOF(AStream, INT_BOF_SHEET);
      WriteIndex(AStream);
      WriteColInfos(AStream, sheet);
      WriteDimensions(AStream, sheet);
      WriteWindow2(AStream, True);
      WriteCellsToStream(AStream, sheet.Cells);
    WriteEOF(AStream);
  end;
  
  { Cleanup }
  
  SetLength(Boundsheets, 0);
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteBlank
*
*  DESCRIPTION:    Writes the record for an empty cell
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BLANK));
  AStream.WriteWord(WordToLE(6));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { Index to XF record, according to formatting }
  WriteXFIndex(AStream, ACell);
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
  lLastCol: Word;
  lLastRow: Integer;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_DIMENSIONS));
  AStream.WriteWord(WordToLE(14));

  { Index to first used row }
  AStream.WriteDWord(DWordToLE(0));

  { Index to last used row, increased by 1 }
  lLastRow := GetLastRowIndex(AWorksheet)+1;
  AStream.WriteDWord(DWordToLE(lLastRow)); // Old dummy value: 33

  { Index to first used column }
  AStream.WriteWord(WordToLE(0));

  { Index to last used column, increased by 1 }
  lLastCol := GetLastColIndex(AWorksheet)+1;
  AStream.WriteWord(WordToLE(lLastCol)); // Old dummy value: 10

  { Not used }
  AStream.WriteWord(WordToLE(0));
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
  AStream.WriteWord(WordToLE(ord(AFont.Color)));

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
*  DESCRIPTION:    Writes the Excel 8 FONT records neede for the
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

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteFormula ()
*
*  DESCRIPTION:    Writes an Excel 5 FORMULA record
*
*                  To input a formula to this method, first convert it
*                  to RPN, and then list all it's members in the
*                  AFormula array
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteFormula(AStream: TStream; const ARow,
  ACol: Cardinal; const AFormula: TsFormula; ACell: PCell);
{var
  FormulaResult: double;
  i: Integer;
  RPNLength: Word;
  TokenArraySizePos, RecordSizePos, FinalPos: Int64;}
begin
(*  RPNLength := 0;
  FormulaResult := 0.0;

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMULA));
  RecordSizePos := AStream.Position;
  AStream.WriteWord(WordToLE(22 + RPNLength));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { Index to XF Record }
  AStream.WriteWord($0000);

  { Result of the formula in IEE 754 floating-point value }
  AStream.WriteBuffer(FormulaResult, 8);

  { Options flags }
  AStream.WriteWord(WordToLE(MASK_FORMULA_RECALCULATE_ALWAYS));

  { Not used }
  AStream.WriteDWord(0);

  { Formula }

  { The size of the token array is written later,
    because it's necessary to calculate if first,
    and this is done at the same time it is written }
  TokenArraySizePos := AStream.Position;
  AStream.WriteWord(RPNLength);

  { Formula data (RPN token array) }
  for i := 0 to Length(AFormula) - 1 do
  begin
    { Token identifier }
    AStream.WriteByte(AFormula[i].TokenID);
    Inc(RPNLength);

    { Additional data }
    case AFormula[i].TokenID of

    { binary operation tokens }

    INT_EXCEL_TOKEN_TADD, INT_EXCEL_TOKEN_TSUB, INT_EXCEL_TOKEN_TMUL,
     INT_EXCEL_TOKEN_TDIV, INT_EXCEL_TOKEN_TPOWER: begin end;

    INT_EXCEL_TOKEN_TNUM:
    begin
      AStream.WriteBuffer(AFormula[i].DoubleValue, 8);
      Inc(RPNLength, 8);
    end;

    INT_EXCEL_TOKEN_TREFR, INT_EXCEL_TOKEN_TREFV, INT_EXCEL_TOKEN_TREFA:
    begin
      AStream.WriteWord(AFormula[i].Row and MASK_EXCEL_ROW);
      AStream.WriteByte(AFormula[i].Col);
      Inc(RPNLength, 3);
    end;

    end;
  end;

  { Write sizes in the end, after we known them }
  FinalPos := AStream.Position;
  AStream.position := TokenArraySizePos;
  AStream.WriteByte(RPNLength);
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(22 + RPNLength));
  AStream.position := FinalPos;*)
end;

procedure TsSpreadBIFF8Writer.WriteRPNFormula(AStream: TStream; const ARow,
  ACol: Cardinal; const AFormula: TsRPNFormula; ACell: PCell);
var
  FormulaResult: double;
  i: Integer;
  len: Integer;
  RPNLength: Word;
  TokenArraySizePos, RecordSizePos, FinalPos: Int64;
  TokenID: Word;
  lSecondaryID: Word;
  c: Cardinal;
  wideStr: WideString;
begin
  RPNLength := 0;
  FormulaResult := 0.0;

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMULA));
  RecordSizePos := AStream.Position;
  AStream.WriteWord(WordToLE(22 + RPNLength));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { Index to XF record, according to formatting }
  //AStream.WriteWord(0);
  WriteXFIndex(AStream, ACell);

  { Result of the formula in IEEE 754 floating-point value }
  AStream.WriteBuffer(FormulaResult, 8);

  { Options flags }
  AStream.WriteWord(WordToLE(MASK_FORMULA_RECALCULATE_ALWAYS));

  { Not used }
  AStream.WriteDWord(0);

  { Formula }

  { The size of the token array is written later,
    because it's necessary to calculate if first,
    and this is done at the same time it is written }
  TokenArraySizePos := AStream.Position;
  AStream.WriteWord(RPNLength);

  { Formula data (RPN token array) }
  for i := 0 to Length(AFormula) - 1 do
  begin
    { Token identifier }
    TokenID := FormulaElementKindToExcelTokenID(AFormula[i].ElementKind, lSecondaryID);
    AStream.WriteByte(TokenID);
    Inc(RPNLength);

    { Additional data }
    case TokenID of
    { Operand Tokens }
    INT_EXCEL_TOKEN_TREFR, INT_EXCEL_TOKEN_TREFV, INT_EXCEL_TOKEN_TREFA: { fekCell }
    begin
      AStream.WriteWord(AFormula[i].Row);
      c := AFormula[i].Col and MASK_EXCEL_COL_BITS_BIFF8;
      if (rfRelRow in AFormula[i].RelFlags) then c := c or MASK_EXCEL_RELATIVE_ROW;
      if (rfRelCol in AFormula[i].RelFlags) then c := c or MASK_EXCEL_RELATIVE_COL;
      AStream.WriteWord(c);
      Inc(RPNLength, 4);
    end;

    INT_EXCEL_TOKEN_TAREA_R: { fekCellRange }
    begin
      {
      Cell range address, BIFF8:
      Offset Size Contents
      0 2 Index to first row (0…65535) or offset of first row (method [B], -32768…32767)
      2 2 Index to last row (0…65535) or offset of last row (method [B], -32768…32767)
      4 2 Index to first column or offset of first column, with relative flags (see table above)
      6 2 Index to last column or offset of last column, with relative flags (see table above)
      }
      AStream.WriteWord(WordToLE(AFormula[i].Row));
      AStream.WriteWord(WordToLE(AFormula[i].Row2));
      c := AFormula[i].Col;
      if (rfRelCol in AFormula[i].RelFlags) then c := c or MASK_EXCEL_RELATIVE_COL;
      if (rfRelRow in AFormula[i].RelFlags) then c := c or MASK_EXCEL_RELATIVE_ROW;
      AStream.WriteWord(WordToLE(c));
      c := AFormula[i].Col2;
      if (rfRelCol2 in AFormula[i].RelFlags) then c := c or MASK_EXCEL_RELATIVE_COL;
      if (rfRelRow2 in AFormula[i].RelFlags) then c := c or MASK_EXCEL_RELATIVE_ROW;
      AStream.WriteWord(WordToLE(c));
      Inc(RPNLength, 8);
    end;

    INT_EXCEL_TOKEN_TNUM: { fekNum }
    begin
      AStream.WriteBuffer(AFormula[i].DoubleValue, 8);
      Inc(RPNLength, 8);
    end;

    INT_EXCEL_TOKEN_TSTR: { fekString }
    begin
      // string constant is stored as widestring in BIFF8
      wideStr := AFormula[i].StringValue;
      len := Length(wideStr);
      AStream.WriteByte(len); // char count in 1 byte
      AStream.WriteByte(1);   // Widestring flags, 1=regular unicode LE string
      AStream.WriteBuffer(WideStringToLE(wideStr)[1], len * Sizeof(WideChar));
      Inc(RPNLength, 1 + 1 + len*SizeOf(WideChar));
    end;

    INT_EXCEL_TOKEN_TBOOL:  { fekBool }
    begin
      AStream.WriteByte(ord(AFormula[i].DoubleValue <> 0.0));
      inc(RPNLength, 1);
    end;

    { binary operation tokens }
    INT_EXCEL_TOKEN_TADD, INT_EXCEL_TOKEN_TSUB, INT_EXCEL_TOKEN_TMUL,
     INT_EXCEL_TOKEN_TDIV, INT_EXCEL_TOKEN_TPOWER: begin end;

    { Other operations }
    INT_EXCEL_TOKEN_TATTR: { fekOpSUM }
    { 3.10, page 71: e.g. =SUM(1) is represented by token array
    tInt(1),tAttrRum
    }
    begin
      // Unary SUM Operation
      AStream.WriteByte($10); //tAttrSum token (SUM with one parameter)
      AStream.WriteByte(0); // not used
      AStream.WriteByte(0); // not used
      Inc(RPNLength, 3);
    end;

    // Functions with fixed parameter count
    INT_EXCEL_TOKEN_FUNC_R, INT_EXCEL_TOKEN_FUNC_V, INT_EXCEL_TOKEN_FUNC_A:
    begin
      AStream.WriteWord(WordToLE(lSecondaryID));
      Inc(RPNLength, 2);
    end;

    // Functions with variable parameter count
    INT_EXCEL_TOKEN_FUNCVAR_V:
    begin
      AStream.WriteByte(AFormula[i].ParamsNum);
      AStream.WriteWord(WordToLE(lSecondaryID));
      Inc(RPNLength, 3);
    end;

    else
    end;
  end;

  { Write sizes in the end, after we known them }
  FinalPos := AStream.Position;
  AStream.position := TokenArraySizePos;
  AStream.WriteByte(RPNLength);
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(22 + RPNLength));
  AStream.position := FinalPos;
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
  MaxBytes=32758;
var
  L, RecLen: Word;
  TextTooLong: boolean=false;
  WideValue: WideString;
begin
  WideValue := UTF8Decode(AValue); //to UTF16
  if WideValue = '' then
  begin
    // Badly formatted UTF8String (maybe ANSI?)
    if Length(AValue)<>0 then begin
      //Quite sure it was an ANSI string written as UTF8, so raise exception.
      Raise Exception.CreateFmt('Expected UTF8 text but probably ANSI text found in cell [%d,%d]',[ARow,ACol]);
    end;
    Exit;
  end;

  if Length(WideValue)>MaxBytes then
  begin
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    TextTooLong := true;
    SetLength(WideValue,MaxBytes); //may corrupt the string (e.g. in surrogate pairs), but... too bad.
  end;
  L := Length(WideValue);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_LABEL));
  RecLen := 8 + 1 + L * SizeOf(WideChar);
  AStream.WriteWord(WordToLE(RecLen));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { Index to XF record, according to formatting }
  WriteXFIndex(AStream, ACell);

  { Byte String with 16-bit size }
  AStream.WriteWord(WordToLE(L));

  { Byte flags. 1 means regular Unicode LE encoding}
  AStream.WriteByte(1);
  AStream.WriteBuffer(WideStringToLE(WideValue)[1], L * Sizeof(WideChar));

  {
  //todo: keep a log of errors and show with an exception after writing file or something.
  We can't just do the following
  if TextTooLong then
    Raise Exception.CreateFmt('Text value exceeds %d character limit in cell [%d,%d]. Text has been truncated.',[MaxBytes,ARow,ACol]);
  because the file wouldn't be written.
  }
end;

{*******************************************************************
*  TsSpreadBIFF8Writer.WriteNumber ()
*
*  DESCRIPTION:    Writes an Excel 8 NUMBER record
*
*                  Writes a number (64-bit floating point) to the sheet
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_NUMBER));
  AStream.WriteWord(WordToLE(14)); //total record size

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { Index to XF record, according to formatting }
  WriteXFIndex(AStream, ACell);

  { IEE 754 floating-point value (is different in BIGENDIAN???) }
  AStream.WriteBuffer(AValue, 8);
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
*  TsSpreadBIFF8Writer.WriteWindow1 ()
*
*  DESCRIPTION:    Writes an Excel 8 WINDOW1 record
*
*                  This record contains general settings for the
*                  document window and global workbook settings.
*
*                  The values written here are reasonable defaults,
*                  which should work for most sheets.
*
*******************************************************************}
procedure TsSpreadBIFF8Writer.WriteWindow1(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW1));
  AStream.WriteWord(WordToLE(18));

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
 ASheetSelected: Boolean);
var
  Options: Word;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW2));
  AStream.WriteWord(WordToLE(18));

  { Options flags }
  Options := MASK_WINDOW2_OPTION_SHOW_GRID_LINES or
   MASK_WINDOW2_OPTION_SHOW_SHEET_HEADERS or
   MASK_WINDOW2_OPTION_SHOW_ZERO_VALUES or
   MASK_WINDOW2_OPTION_AUTO_GRIDLINE_COLOR or
   MASK_WINDOW2_OPTION_SHOW_OUTLINE_SYMBOLS or
   MASK_WINDOW2_OPTION_SHEET_ACTIVE;

  if ASheetSelected then Options := Options or MASK_WINDOW2_OPTION_SHEET_SELECTED;

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
 AHorAlignment: TsHorAlignment = haDefault; AVertAlignment: TsVertAlignment = vaDefault;
 AWordWrap: Boolean = false; AddBackground: Boolean = false;
 ABackgroundColor: TsColor = scSilver);
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

  // Left and Right line colors, use black
  XFBorderDWord1 := 8 * $10000 {left line - black} + 8 * $800000 {right line - black};

  if cbNorth in ABorders then XFBorderDWord1 := XFBorderDWord1 or $100;
  if cbWest in ABorders  then XFBorderDWord1 := XFBorderDWord1 or $1;
  if cbEast in ABorders  then XFBorderDWord1 := XFBorderDWord1 or $10;
  if cbSouth in ABorders then XFBorderDWord1 := XFBorderDWord1 or $1000;

  AStream.WriteDWord(DWordToLE(XFBorderDWord1));

  // Top and Bottom line colors, use black
  XFBorderDWord2 := 8 {top line - black} + 8 * $80 {bottom line - black};
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
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF1
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF2
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF3
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF4
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF5
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF6
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF7
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF8
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF9
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF10
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF11
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF12
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF13
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF14
  WriteXF(AStream, 0, 0, MASK_XF_TYPE_PROT_STYLE_XF, XF_ROTATION_HORIZONTAL, []);
  // XF15 - Default, no formatting
  WriteXF(AStream, 0, 0, 0, XF_ROTATION_HORIZONTAL, []);

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

function TsSpreadBIFF8Reader.DecodeRKValue(const ARK: DWORD): Double;
var
  Number: Double;
  Tmp: LongInt;
begin
  if ARK and 2 = 2 then begin
    // Signed integer value
    if LongInt(ARK)<0 then begin
      //Simulates a sar
      Tmp:=LongInt(ARK)*-1;
      Tmp:=Tmp shr 2;
      Tmp:=Tmp*-1;
      Number:=Tmp-1;
    end else begin
      Number:=ARK shr 2;
    end;
  end else begin
    // Floating point value
    // NOTE: This is endian dependent and IEEE dependent (Not checked) (working win-i386)
    (PDWORD(@Number))^:= $00000000;
    (PDWORD(@Number)+1)^:=(ARK and $FFFFFFFC);
  end;
  if ARK and 1 = 1 then begin
    // Encoded value is multiplied by 100
    Number:=Number / 100;
  end;
  Result:=Number;
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
      SetLength(Result,ALength);
      AStream.ReadBuffer(Result[1],ALength * SizeOf(WideChar));
      Dec(PendingRecordSize,ALength * SizeOf(WideChar));
    end;
    Result:=WideStringLEToN(Result);
  end else begin
    //String is 1 byte per char, this is UTF-16 with the high byte ommited because it is zero
    //so decompress and then convert
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
  WideName: WideString;
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
       INT_EXCEL_ID_BOF:        ;
       INT_EXCEL_ID_BOUNDSHEET: ReadBoundSheet(AStream);
       INT_EXCEL_ID_EOF:        SectionEOF := True;
       INT_EXCEL_ID_SST:        ReadSST(AStream);
       INT_EXCEL_ID_CODEPAGE:   ReadCodepage(AStream);
       INT_EXCEL_ID_FONT:       ReadFont(AStream);
       INT_EXCEL_ID_XF:         ReadXF(AStream);
       INT_EXCEL_ID_FORMAT:     ReadFormat(AStream);
       INT_EXCEL_ID_DATEMODE:   ReadDateMode(AStream);
       INT_EXCEL_ID_PALETTE:    ReadPalette(AStream);
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

    INT_EXCEL_ID_BLANK:   ReadBlank(AStream);
    INT_EXCEL_ID_NUMBER:  ReadNumber(AStream);
    INT_EXCEL_ID_LABEL:   ReadLabel(AStream);
    INT_EXCEL_ID_FORMULA: ReadFormula(AStream);
    //(RSTRING) This record stores a formatted text cell (Rich-Text).
    // In BIFF8 it is usually replaced by the LABELSST record. Excel still
    // uses this record, if it copies formatted text cells to the clipboard.
    INT_EXCEL_ID_RSTRING: ReadRichString(AStream);
    // (RK) This record represents a cell that contains an RK value
    // (encoded integer or floating-point value). If a floating-point
    // value cannot be encoded to an RK value, a NUMBER record will be written.
    // This record replaces the record INTEGER written in BIFF2.
    INT_EXCEL_ID_RK:      ReadRKValue(AStream);
    INT_EXCEL_ID_MULRK:   ReadMulRKValues(AStream);
    INT_EXCEL_ID_LABELSST:ReadLabelSST(AStream); //BIFF8 only
    INT_EXCEL_ID_COLINFO: ReadColInfo(AStream);
    INT_EXCEL_ID_ROWINFO: ReadRowInfo(AStream);
    INT_EXCEL_ID_BOF:     ;
    INT_EXCEL_ID_EOF:     SectionEOF := True;
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then SectionEOF := True;
  end;
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

procedure TsSpreadBIFF8Reader.ReadRKValue(const AStream: TStream);
var
  RK: DWORD;
  ARow, ACol: Cardinal;
  XF: WORD;
  lDateTime: TDateTime;
  Number: Double;
  nf: TsNumberFormat;    // Number format
  nd: word;              // decimals
  nfs: String;           // Number format string
begin
  {Retrieve XF record, row and column}
  ReadRowColXF(AStream, ARow, ACol, XF);

  {Encoded RK value}
  RK:=DWordLEtoN(AStream.ReadDWord);

  {Check RK codes}
  Number:=DecodeRKValue(RK);

  {Find out what cell type, set contenttype and value}
  ExtractNumberFormat(XF, nf, nd, nfs);
  if IsDateTime(Number, nf, lDateTime) then
    FWorksheet.WriteDateTime(ARow, ACol, lDateTime, nf, nfs)
  else
    FWorksheet.WriteNumber(ARow, ACol, Number, nf);

  {Add attributes}
  ApplyCellFormatting(ARow, ACol, XF);
end;

procedure TsSpreadBIFF8Reader.ReadMulRKValues(const AStream: TStream);
var
  ARow, fc,lc,XF: Word;
  lDateTime: TDateTime;
  Pending: integer;
  RK: DWORD;
  Number: Double;
  nf: TsNumberFormat;
  nd: word;
  nfs: String;
begin
  ARow:=WordLEtoN(AStream.ReadWord);
  fc:=WordLEtoN(AStream.ReadWord);
  Pending:=RecordSize-sizeof(fc)-Sizeof(ARow);
  while Pending > (sizeof(XF)+sizeof(RK)) do begin
    XF:=AStream.ReadWord; //XF record (used for date checking)
    RK:=DWordLEtoN(AStream.ReadDWord);
    Number:=DecodeRKValue(RK);
    {Find out what cell type, set contenttype and value}
    ExtractNumberFormat(XF, nf, nd, nfs);
    if IsDateTime(Number, nf, lDateTime) then
      FWorksheet.WriteDateTime(ARow, fc, lDateTime, nf, nfs)
    else
      FWorksheet.WriteNumber(ARow, fc, Number, nf, nd);
    inc(fc);
    dec(Pending,(sizeof(XF)+sizeof(RK)));
  end;
  if Pending=2 then begin
    //Just for completeness
    lc:=WordLEtoN(AStream.ReadWord);
    if lc+1<>fc then begin
      //Stream error... bypass by now
    end;
  end;
end;

function TsSpreadBIFF8Reader.ReadString(const AStream: TStream;
  const ALength: WORD): UTF8String;
begin
  Result:=UTF16ToUTF8(ReadWideString(AStream, ALength));
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
    OLEStorage.ReadOLEFile(AFileName, OLEDocument,'Workbook');

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

procedure TsSpreadBIFF8Reader.ReadBlank(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
begin
  { Read row, column, and XF index from BIFF file }
  ReadRowColXF(AStream, ARow, ACol, XF);
  { Add attributes to cell}
  ApplyCellFormatting(ARow, ACol, XF);
end;

procedure TsSpreadBIFF8Reader.ReadFormula(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: WORD;
  ResultFormula: Double;
  Data: array [0..7] of BYTE;
  Flags: WORD;
  FormulaSize: BYTE;
  i: Integer;
begin
  { BIFF Record header }
  { BIFF Record data }
  { Index to XF Record }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Result of the formula in IEE 754 floating-point value }
  AStream.ReadBuffer(Data, Sizeof(Data));

  { Options flags }
  Flags := WordLEtoN(AStream.ReadWord);

  { Not used }
  AStream.ReadDWord;

  { Formula size }
  FormulaSize := WordLEtoN(AStream.ReadWord);

  { Formula data, output as debug info }
{  Write('Formula Element: ');
  for i := 1 to FormulaSize do
    Write(IntToHex(AStream.ReadByte, 2) + ' ');
  WriteLn('');}

  //RPN data not used by now
  AStream.Position := AStream.Position + FormulaSize;

  if SizeOf(Double) <> 8 then
    raise Exception.Create('Double is not 8 bytes');
  Move(Data[0], ResultFormula, SizeOf(Data));
  FWorksheet.WriteNumber(ARow, ACol, ResultFormula);

  {Add attributes}
  ApplyCellFormatting(ARow, ACol, XF);
end;

procedure TsSpreadBIFF8Reader.ReadLabel(AStream: TStream);
var
  L: Word;
  StringFlags: BYTE;
  ARow, ACol: Cardinal;
  XF: Word;
  WideStrValue: WideString;
  AnsiStrValue: AnsiString;
begin
  { BIFF Record data: Row, Column, XF Index }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Byte String with 16-bit size }
  L := WordLEtoN(AStream.ReadWord());

  { Read string with flags }
  WideStrValue:=ReadWideString(AStream,L);

  { Save the data }
  FWorksheet.WriteUTF8Text(ARow, ACol, UTF16ToUTF8(WideStrValue));

  {Add attributes}
  ApplyCellFormatting(ARow, ACol, XF);
end;

procedure TsSpreadBIFF8Reader.ReadRichString(const AStream: TStream);
var
  L: Word;
  B: WORD;
  ARow, ACol: Cardinal;
  XF: Word;
  AStrValue: ansistring;
begin
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Byte String with 16-bit size }
  L := WordLEtoN(AStream.ReadWord());
  AStrValue:=ReadString(AStream,L);

  { Save the data }
  FWorksheet.WriteUTF8Text(ARow, ACol, AStrValue);
  //Read formatting runs (not supported)
  B:=WordLEtoN(AStream.ReadWord);
  for L := 0 to B-1 do begin
    AStream.ReadWord; // First formatted character
    AStream.ReadWord; // Index to FONT record
  end;

  {Add attributes}
  ApplyCellFormatting(ARow, ACol, XF);
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
begin
  ReadRowColXF(AStream, ARow, ACol, XF);
  SSTIndex := DWordLEtoN(AStream.ReadDWord);
  if SizeInt(SSTIndex) >= FSharedStringTable.Count then begin
    Raise Exception.CreateFmt('Index %d in SST out of range (0-%d)',[Integer(SSTIndex),FSharedStringTable.Count-1]);
  end;
  FWorksheet.WriteUTF8Text(ARow, ACol, FSharedStringTable[SSTIndex]);

  {Add attributes}
  ApplyCellFormatting(ARow, ACol, XF);
end;

procedure TsSpreadBIFF8Reader.ReadXF(const AStream: TStream);
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
  // the 4 masked bits encode the line style of the border line. 0 = no line
  // We ignore the line style here. --> check against "no line"
  if xf.Border_Background_1 and MASK_XF_BORDER_LEFT <> 0 then
    Include(lData.Borders, cbWest);
  if xf.Border_Background_1 and MASK_XF_BORDER_RIGHT <> 0 then
    Include(lData.Borders, cbEast);
  if xf.Border_Background_1 and MASK_XF_BORDER_TOP <> 0 then
    Include(lData.Borders, cbNorth);
  if xf.Border_Background_1 and MASK_XF_BORDER_BOTTOM <> 0 then
    Include(lData.Borders, cbSouth);

  // Background color;
  xf.Border_Background_3 := DWordLEToN(xf.Border_Background_3);
  lData.BackgroundColor := xf.Border_Background_3 AND $007F;

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
  lFontName: UTF8String;
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

  { Font name: Unicodestring, char count in 1 byte }
  Len := AStream.ReadByte();
  font.FontName := ReadString(AStream, Len);

  { Add font to workbook's font list }
  FWorkbook.AddFont(font);
end;

// Read the FORMAT record for formatting numerical data
procedure TsSpreadBIFF8Reader.ReadFormat(AStream: TStream);
var
  lData: TFormatListData;
begin
  lData := TFormatListData.Create;

  // Record FORMAT, BIFF 8 (5.49):
  // Offset Size Contents
  // 0      2     Format index used in other records
  // 2      var   Number format string (Unicode string, 16-bit string length)
  // From BIFF5 on: indexes 0..163 are built in
  lData.Index := WordLEtoN(AStream.ReadWord);

  // 2 var. Number format string (Unicode string, 16-bit string length, ➜2.5.3)
  lData.FormatString := ReadWideString(AStream, False);

  // Add to the list
  FFormatList.Add(lData);
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

