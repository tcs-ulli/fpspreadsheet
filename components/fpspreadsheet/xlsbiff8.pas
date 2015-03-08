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
  Classes, SysUtils, fpcanvas, DateUtils, contnrs, lazutf8,
  fpstypes, fpsnumformat, fpspreadsheet, xlscommon,
  {$ifdef USE_NEW_OLE}
  fpolebasic,
  {$else}
  fpolestorage,
  {$endif}
  fpsutils;

type

  { TsSpreadBIFF8Reader }
  TsSpreadBIFF8Reader = class(TsSpreadBIFFReader)
  private
    PendingRecordSize: SizeInt;
    FWorksheetNames: TStringList;
    FCurrentWorksheet: Integer;
    FSharedStringTable: TStringList;
    FCommentList: TObjectList;
    FCommentPending: Boolean;
    FCommentID: Integer;
    FCommentLen: Integer;
    function ReadWideString(const AStream: TStream; const ALength: WORD): WideString; overload;
    function ReadWideString(const AStream: TStream; const AUse8BitLength: Boolean): WideString; overload;
    procedure ReadWorkbookGlobals(AStream: TStream);
    procedure ReadWorksheet(AStream: TStream);
    procedure ReadBoundsheet(AStream: TStream);
    function ReadString(const AStream: TStream; const ALength: WORD): String;
  protected
    procedure ReadCONTINUE(const AStream: TStream);
    procedure ReadFONT(const AStream: TStream);
    procedure ReadFORMAT(AStream: TStream); override;
    procedure ReadHyperLink(AStream: TStream);
    procedure ReadHyperlinkToolTip(AStream: TStream);
    procedure ReadLABEL(AStream: TStream); override;
    procedure ReadLabelSST(const AStream: TStream);
    procedure ReadMergedCells(const AStream: TStream);
    procedure ReadNOTE(const AStream: TStream);
    procedure ReadOBJ(const AStream: TStream);
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
    procedure ReadTXO(const AStream: TStream);
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
    procedure WriteComment(AStream: TStream; ACell: PCell); override;
    procedure WriteComments(AStream: TStream; AWorksheet: TsWorksheet);
    {
    procedure WriteCommentEscher(AStream: TStream; AComment: PsComment;
      ACommentIndex: Integer);
    procedure WriteCommentNote(AStream: TStream; AComment: PsComment;
      ACommentIndex: Integer);
      }
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFont: TsFont);
    procedure WriteFonts(AStream: TStream);
    procedure WriteIndex(AStream: TStream);
    procedure WriteHyperlink(AStream: TStream; AHyperlink: PsHyperlink;
      AWorksheet: TsWorksheet);
    procedure WriteHyperlinks(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteHyperlinkToolTip(AStream: TStream; const ARow, ACol: Cardinal;
      const ATooltip: String);
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteMergedCells(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteMSODrawing1(AStream: TStream; ANumShapes: Word; AComment: PsComment);
    procedure WriteMSODrawing2(AStream: TStream; AComment: PsComment; AObjID: Word);
    procedure WriteMSODrawing2_Data(AStream: TStream; AComment: PsComment; AShapeID: Word);
    procedure WriteMSODrawing3(AStream: TStream);
    procedure WriteNOTE(AStream: TStream; AComment: PsComment; AObjID: Word);
    procedure WriteNumFormat(AStream: TStream; AFormatData: TsNumFormatData;
      AListIndex: Integer); override;
    procedure WriteOBJ(AStream: TStream; AObjID: Word);
    function WriteRPNCellAddress(AStream: TStream; ARow, ACol: Cardinal;
      AFlags: TsRelFlags): word; override;
    function WriteRPNCellOffset(AStream: TStream; ARowOffset, AColOffset: Integer;
      AFlags: TsRelFlags): Word; override;
    function WriteRPNCellRangeAddress(AStream: TStream; ARow1, ACol1, ARow2, ACol2: Cardinal;
      AFlags: TsRelFlags): Word; override;
    function WriteString_8bitLen(AStream: TStream; AString: String): Integer; override;
    procedure WriteStringRecord(AStream: TStream; AString: string); override;
    procedure WriteSTYLE(AStream: TStream);
    procedure WriteTXO(AStream: TStream; AComment: PsComment);
    procedure WriteWINDOW2(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteXF(AStream: TStream; AFormatRecord: PsCellFormat;
      XFType_Prot: Byte = 0); override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;

  TExcel8Settings = record
    DateMode: TDateMode;
  end;

var
  Excel8Settings: TExcel8Settings = (
    DateMode: dm1900;
  );

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
  Math, lconvencoding, LazFileUtils, URIParser,
  fpsStrings, fpsStreams, fpsReaderWriter, fpsExprParser, xlsEscher;

const
   { Excel record IDs }
     INT_EXCEL_ID_MERGEDCELLS            = $00E5;  // BIFF8 only
     INT_EXCEL_ID_MSODRAWING             = $00EC;  // BIFF8 only
     INT_EXCEL_ID_SST                    = $00FC;  // BIFF8 only
     INT_EXCEL_ID_LABELSST               = $00FD;  // BIFF8 only
     INT_EXCEL_ID_TXO                    = $01B6;  // BIFF8 only
     INT_EXCEL_ID_HYPERLINK              = $01B8;  // BIFF8 only
     INT_EXCEL_ID_HLINKTOOLTIP           = $0800;  // BIFF8 only
{%H-}INT_EXCEL_ID_FORCEFULLCALCULATION   = $08A3;

   { Excel OBJ subrecord IDs }
     INT_EXCEL_OBJID_FTEND               = $0000;
{%H-}INT_EXCEL_OBJID_FTMACRO             = $0004;
{%H-}INT_EXCEL_OBJID_FTBUTTON            = $0005;
{%H-}INT_EXCEL_OBJID_FTGMO               = $0006;  // Group marker
{%H-}INT_EXCEL_OBJID_CF                  = $0007;  // Clipboard format
{%H-}INT_EXCEL_OBJID_PIOGRBIT            = $0008;  // Picture option flags
{%H-}INT_EXCEL_OBJID_PICTFMLA            = $0009;  // Picture fmla-style macro
{%H-}INT_EXCEL_OBJID_FTCBLS              = $000A;  // Checkbox link
{%H-}INT_EXCEL_OBJID_FTRBO               = $000B;  // Radio button
{%H-}INT_EXCEL_OBJID_FTSBS               = $000C;  // Scrollbar
{%H-}INT_EXCEL_OBJID_FTNTS               = $000D;  // Notes structure (= Comment)
{%H-}INT_EXCEL_OBJID_FTSBSFMLA           = $000E;  // Scroll bar fmla-style macro
{%H-}INT_EXCEL_OBJID_FTGBODATA           = $000F;  // Group box data
{%H-}INT_EXCEL_OBJID_FTEDODATA           = $0010;  // Edit control data
{%H-}INT_EXCEL_OBJID_FTRBODATA           = $0011;  // Radio button data
{%H-}INT_EXCEL_OBJID_FTCBLSDATA          = $0012;  // Check box data
{%H-}INT_EXCEL_OBJID_FTLBSDATA           = $0013;  // List box data
{%H-}INT_EXCEL_OBJID_FTCBLSFMLA          = $0014;  // Check box link fmla-style macro
     INT_EXCEL_OBJID_FTCMO               = $0015;  // Common object data

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

     TEXT_ROTATIONS: Array[TsTextRotation] of Byte = (
       XF_ROTATION_HORIZONTAL,
       XF_ROTATION_90DEG_CW,
       XF_ROTATION_90DEG_CCW,
       XF_ROTATION_STACKED
     );

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

   { HLINK FLAGS }
     MASK_HLINK_LINK                     = $00000001;
     MASK_HLINK_ABSOLUTE                 = $00000002;
     MASK_HLINK_DESCRIPTION              = $00000014;
     MASK_HLINK_TEXTMARK                 = $00000008;
{%H-}MASK_HLINK_TARGETFRAME              = $00000080;
{%H-}MASK_HLINK_UNCPATH                  = $00000100;

     SHAPEID_BASE = 1024;


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

  TBIFF8TXORecord = packed record
    RecordID: Word;
    RecordSize: Word;
    OptionFlags: Word;
    TextRot: Word;
    Reserved1: Word;
    Reserved2: Word;
    Reserved3: Word;
    TextLen: Word;
    NumFormattingRuns: Word;
    Reserved4: Word;
    Reserved5: Word;
  end;

  TBIFF8Comment = class
    ID: Integer;
    Text: String;
  end;


{ TsSpreadBIFF8Reader }

destructor TsSpreadBIFF8Reader.Destroy;
begin
  if Assigned(FSharedStringTable) then FSharedStringTable.Free;
  if Assigned(FCommentList) then FCommentList.Free;
  inherited;
end;

{ Reads a CONTINUE record. If the Flag "FCommentPending" is active then this
  record contains the text of a comment assigned to a cell. The length of the
  string is taken from the preceeding TXO record, and the ID of the comment is
  extracted in another preceeding record, an OBJ record. }
procedure TsSpreadBIFF8Reader.ReadCONTINUE(const AStream: TStream);
var
  commentStr: String;
  comment: TBIFF8Comment;
begin
  if FCommentPending then begin
    commentStr := ReadWideString(AStream, FCommentLen);
    if commentStr <> '' then
    begin
      comment := TBIFF8Comment.Create;
      comment.ID := FCommentID;
      comment.Text := commentStr;
      FCommentList.Add(comment);
    end;
    FCommentPending := false;
  end;
end;

{ Reads a NOTE record (comment associated with a cell). All comments have been
  collected in the FCommentList of the reader from preceding OBJ, TXO and
  CONTINUE records. }
procedure TsSpreadBIFF8Reader.ReadNOTE(const AStream: TStream);
var
  r, c: Word;
  commentID: Word;
  commentText: String;
  i: Integer;
begin
  { Row of the comment }
  r := WordLEToN(AStream.ReadWord);
  { Column of the comment }
  c := WordLEToN(AStream.ReadWord);
  { Option flags, not needed }
  WordLEToN(AStream.ReadWord);
  { Comment ID }
  commentID := WordLEToN(AStream.ReadWord);
  { Next would be the author - ignored... }

  { Seek comment with this ID in the comment list of the reader. }
  for i:=0 to FCommentList.Count-1 do
    if TBIFF8Comment(FCommentList[i]).ID = commentID then
    begin
      commentText := TBIFF8Comment(FCommentList[i]).Text;
      FWorksheet.WriteComment(r, c, commentText);
      exit;
    end;
end;

{ Reads an OBJ record. So far, we only evaluate it to get the ID of the comment
  stored in a following TXO and CONTINUE record. }
procedure TsSpreadBIFF8Reader.ReadOBJ(const AStream: TStream);
var
  subrecID, subrecSize: Word;
  streamPos: Int64;
  lastSubRec: Boolean;
  objType: Word;
  objID: Word;
begin
  lastSubRec := false;
  while not lastSubRec do begin
    subrecID := WordLEToN(AStream.ReadWord);
    subrecSize := WordLEToN(AStream.ReadWord);
    streamPos := AStream.Position;
    case subrecID of
      INT_EXCEL_OBJID_FTCMO:  // common object data
        begin
          objType := WordLEToN(AStream.ReadWord);
          objID := WordLEToN(AStream.ReadWord);
          if objType = $19 then begin  // $19 = object is a "comment"
            FCommentPending := true;
            FCommentID := objID;
            exit;
          end else
            FCommentPending := false;
        end;
      INT_EXCEL_OBJID_FTEND:
        lastSubRec := true;
    end;
    AStream.Position := streamPos + subrecSize;
  end;
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
    INT_EXCEL_ID_CONTINUE    : ReadCONTINUE(AStream);
    INT_EXCEL_ID_FORMULA     : ReadFormula(AStream);
    INT_EXCEL_ID_HYPERLINK   : ReadHyperlink(AStream);
    INT_EXCEL_ID_HLINKTOOLTIP: ReadHyperlinkToolTip(AStream);
    INT_EXCEL_ID_LABEL       : ReadLabel(AStream);
    INT_EXCEL_ID_MULBLANK    : ReadMulBlank(AStream);
    INT_EXCEL_ID_NOTE        : ReadNOTE(AStream);
    INT_EXCEL_ID_NUMBER      : ReadNumber(AStream);
    INT_EXCEL_ID_OBJ         : ReadOBJ(AStream);
    INT_EXCEL_ID_SHAREDFMLA  : ReadSharedFormula(AStream);
    INT_EXCEL_ID_STRING      : ReadStringRecord(AStream);
    INT_EXCEL_ID_TXO         : ReadTXO(AStream);
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
  BIFF8EOF := False;

  FWorksheetNames := TStringList.Create;
  FWorksheetNames.Clear;
  FCurrentWorksheet := 0;

  if FCommentList = nil then FCommentList := TObjectList.Create
    else FCommentList.Clear;

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

procedure TsSpreadBIFF8Reader.ReadLABEL(AStream: TStream);
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

{ Reads a TXO record (TEXT OBJECT). Needed to retrieve cell comments.
  We only extract the length of the comment text (in characters). The text itself
  is contained in the following CONTINUE record. }
procedure TsSpreadBIFF8Reader.ReadTXO(const AStream: TStream);
var
  rec: TBIFF8TXORecord;
begin
  rec.OptionFlags := 0;  // to silence the compiler
  AStream.ReadBuffer(rec.OptionFlags, Sizeof(Rec) - 2*SizeOf(Word));
  FCommentLen := WordLEToN(rec.TextLen);
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
  fs: TsFillStyle;
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

  // Background fill pattern and color
  fill := (rec.Border_BkGr2 and MASK_XF_BACKGROUND_PATTERN) shr 26;
  if fill <> MASK_XF_FILL_PATT_EMPTY then
  begin
    for fs in TsFillStyle do
      if fill = MASK_XF_FILL_PATT[fs] then
      begin
        rec.BkGr3 := DWordLEToN(rec.BkGr3);
        // Pattern color
        fmt.Background.FgColor := rec.BkGr3 and $007F;
        if fmt.Background.FgColor = SYS_DEFAULT_FOREGROUND_COLOR then
          fmt.Background.FgColor := scBlack;
        // Background color
        fmt.Background.BgColor := (rec.BkGr3 and $3F80) shr 7;
        if fmt.Background.BgColor = SYS_DEFAULT_BACKGROUND_COLOR then
          fmt.Background.BgColor := scTransparent;
        // Fill style
        fmt.Background.Style := fs;
        Include(fmt.UsedFormattingFields, uffBackground);
        break;
      end;
  end;

  // Add the XF to the list
  FCellFormatList.Add(fmt);
end;

{ Reads a FONT record. The retrieved font is stored in the workbook's FontList. }
procedure TsSpreadBIFF8Reader.ReadFONT(const AStream: TStream);
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
procedure TsSpreadBIFF8Reader.ReadFORMAT(AStream: TStream);
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

{@@ ----------------------------------------------------------------------------
  Reads a HYPERLINK record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.ReadHyperlink(AStream: TStream);
var
  row, col, row1, col1, row2, col2: word;
  guid: TGUID;
  flags: DWord;
  widestr: widestring;
  len: DWord;
  link: String;
  linkDos: String;
  mark: String;
  dirUpCount: Word;
  ansistr: ansistring;
  size: DWord;
begin
  { Row and column index range of cells using the hyperlink }
  row1 := WordLEToN(AStream.ReadWord);
  row2 := WordLEToN(AStream.ReadWord);
  col1 := WordLEToN(AStream.ReadWord);
  col2 := WordLEToN(AStream.ReadWord);

  { GUID of standard link }
  AStream.ReadBuffer(guid{%H-}, SizeOf(guid));

  { unknown DWord }
  AStream.ReadDWord;

  { Flags }
  flags := DWordLEToN(AStream.ReadDWord);

  { Description }
  if flags and MASK_HLINK_DESCRIPTION = MASK_HLINK_DESCRIPTION then
  begin
    // not used because there is always a "normal" cell to which the hyperlink is associated.
    // character count of description incl trailing zero
    len := DWordLEToN(AStream.ReadDWord);
    // Character array (16-bit characters, with trailing zero word)
    SetLength(wideStr, len);
    AStream.ReadBuffer(wideStr[1], len*SizeOf(wideChar));
  end;

  { Target frame: external link (URI or local file) }
  link := '';
  if flags and MASK_HLINK_LINK <> 0 then
  begin
    AStream.ReadBuffer(guid, SizeOf(guid));

    // Check for URL
    if GuidToString(guid) = '{79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}' then
    begin
      // Size of character array incl trailing zero
      size := DWordLEToN(AStream.ReadDWord);
      len := size div 2 -1;
      // Character array of URL (16-bit-characters, with trailing zero word)
      SetLength(wideStr, len);
      AStream.ReadBuffer(wideStr[1], size);
      link := UTF8Encode(widestr);
    end else
    // Check for local file
    if GuidToString(guid) = '{00000303-0000-0000-C000-000000000046}' then
    begin
      dirUpCount := WordLEToN(AStream.ReadWord);
      // Character count of the shortened file path and name, incl trailing zero byte
      len := DWordLEToN(AStream.ReadDWord);
      // Character array of the shortened file path and name in 8.3-DOS-format.
      // This field can be filled with a long file name too.
      // No unicode string header, always 8-bit characters, zero-terminated.
      SetLength(ansiStr, len);
      AStream.ReadBuffer(ansiStr[1], len*SizeOf(ansiChar));
      SetLength(ansistr, len-1);  // Remove trailing zero
      while dirUpCount > 0 do
      begin
        ansistr := '..' + PathDelim + ansistr;
        dec(dirUpCount);
      end;
      linkDos := AnsiToUTF8(ansiStr);
      // 6 unknown DWord values
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      // Size of the following file link field including string length field
      // and additional data field
      size := DWordLEToN(AStream.ReadDWord);
      if size > 0 then
      begin
        // Size of the extended file path and name.
        size := DWordLEToN(AStream.ReadDWord);
        len := size div 2;
        // Unknown
        AStream.ReadWord;
        // Character array of the extended file path and name
        // no Unicode string header, always 16-bit characters, not zero-terminated
        SetLength(wideStr, len);
        AStream.ReadBuffer(wideStr[1], size);
        link := UTF8Encode(wideStr);
      end else
        link := linkDos;

      // An absolute path must be a fully qualified URI to be compatible with fps
      if flags and MASK_HLINK_ABSOLUTE <> 0 then
        FilenameToURI(link);
    end;
  end;

  { Text mark }
  if flags and MASK_HLINK_TEXTMARK = MASK_HLINK_TEXTMARK then
  begin
    // Character count of the text mark, including trailing zero word
    len := DWordLEToN(AStream.ReadDWord);
    // Character array of the text mark without "#" sign
    // no Unicode string header, always 16-bit characters, zero-terminated
    SetLength(wideStr, len);
    AStream.ReadBuffer(wideStr[1], len*SizeOf(wideChar));
    SetLength(wideStr, len-1);  // Remove trailing zero word
    mark := UTF8Encode(wideStr);
  end;

  // Add bookmark to hyperlink target
  if (link <> '') and (mark <> '') then
    link := link + '#' + mark
  else
  if (link = '') then
    link := '#' + mark;

  // Add hyperlink(s) to worksheet
  for row := row1 to row2 do
    for col := col1 to col2 do
      FWorksheet.WriteHyperlink(row, col, link);
end;


{@@ ----------------------------------------------------------------------------
  Reads a HYPERLINK TOOLTIP record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.ReadHyperlinkToolTip(AStream: TStream);
var
  txt: String;
  widestr: widestring;
  //row, col,
  row1, col1, row2, col2: Word;
  hyperlink: PsHyperlink;
  numbytes: Integer;
begin
  { Record type; this matches the BIFF record type }
  AStream.ReadWord;

  { Row and column index range of cells using the hyperlink tooltip }
  row1 := WordLEToN(AStream.ReadWord);
  row2 := WordLEToN(AStream.ReadWord);
  col1 := WordLEToN(AStream.ReadWord);
  col2 := WordLEToN(AStream.ReadWord);

  { Hyperlink tooltip, a null-terminated unicode string }
  numbytes := RecordSize - 5*SizeOf(word);
  SetLength(wideStr, numbytes div 2);
  AStream.ReadBuffer(wideStr[1], numbytes);
  SetLength(wideStr, Length(wideStr)-1);  // Remove trailing zero word
  txt := UTF8Encode(wideStr);

  { Add tooltip to hyperlinks }
  for hyperlink in FWorksheet.Hyperlinks.GetRangeEnumerator(row1, col1, row2, col2) do
    hyperlink^.ToolTip := txt;

  {
  for row := row1 to row2 do
    for col := col1 to col2 do
    begin
      hyperlink := PsHyperlink(FWorksheet.Hyperlinks.Find(row, col));
      if hyperlink <> nil then
        hyperlink^.ToolTip := txt;
    end;
    }
end;


{------------------------------------------------------------------------------}
{                          TsSpreadBIFF8Writer                                 }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the Excel 8 writer
-------------------------------------------------------------------------------}
constructor TsSpreadBIFF8Writer.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FDateMode := Excel8Settings.DateMode;
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
  i: Integer;
  pane: Byte;
begin
  { Write workbook globals }
  WriteBOF(AStream, INT_BOF_WORKBOOK_GLOBALS);
  WriteCodePage(AStream, 'ucs2le'); // = utf8
  WriteWindow1(AStream);
  WriteFonts(AStream);
  WriteNumFormats(AStream);
  WritePalette(AStream);
  WriteXFRecords(AStream);
  WriteStyle(AStream);

  // A BOUNDSHEET for each worksheet
  SetLength(Boundsheets, Workbook.GetWorksheetCount);
  for i := 0 to Workbook.GetWorksheetCount - 1 do
    Boundsheets[i] := WriteBoundsheet(AStream, Workbook.GetWorksheetByIndex(i).Name);

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
        WriteComments(AStream, FWorksheet);
      end;

      // View settings block
      WriteWindow2(AStream, FWorksheet);
      WritePane(AStream, FWorksheet, isBIFF8, pane);
      WriteSelection(AStream, FWorksheet, pane);
      WriteHyperlinks(AStream, FWorksheet);

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
  WriteBIFFHeader(AStream, INT_EXCEL_ID_BOF, 16);

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

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 BOUNDSHEET record
  Always located in the workbook globals substream.
  One BOUNDSHEET is written for each worksheet.

  @return   The stream position where the absolute stream position
            of the BOF of this sheet should be written (4 bytes size).
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.WriteBoundsheet(AStream: TStream; ASheetName: string): Int64;
var
  Len: Byte;
  WideSheetName: WideString;
begin
  WideSheetName:=UTF8Decode(ASheetName);
  Len := Length(WideSheetName);

  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_BOUNDSHEET, 8 + Len * Sizeof(WideChar));

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
  Inherited method for writing a cell comment immediately after cell content.
  A writing method has been implemented by xlscommon. But in BIFF8, this
  must not do anything because comments are collected in a list and
  written en-bloc. See WriteComments.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteComment(AStream: TStream; ACell: PCell);
begin
  // Nothing to do. Reverts the behavior introduced by xlscommon.
  Unused(AStream, ACell);
end;

{@@ ----------------------------------------------------------------------------
  Writes all comments to the worksheet stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteComments(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  index: Integer;
  comment: PsComment;
begin
  exit;      // Remove after comments can be written correctly
  {$warning TODO: Fix writing of cell comments in BIFF8 (file is readable by OpenOffice, but not by Excel)}

  { At first we have to write all Escher-related records for all comments;
      MSODRAWING - OBJ - MSODRAWING - TXO }
  index := 1;
  for comment in AWorksheet.Comments do
  begin
    if index = 1 then
      WriteMSODrawing1(AStream, FWorksheet.Comments.Count, comment)
    else
      WriteMSODrawing2(AStream, comment, index);
    WriteOBJ(AStream, index);
    WriteMSODrawing3(AStream);
    WriteTXO(AStream, comment);
    inc(index);
  end;
//  IterateThroughComments(AStream, AWorksheet.Comments, WriteCommentsEscherCallback);

  { The NOTE records for all comments follow subsequently. }
  index := 1;
  for comment in AWorksheet.Comments do
  begin
    WriteNOTE(AStream, comment, index);
    inc(index);
  end;
//  IterateThroughComments(AStream, AWorksheet.Comments, WriteCommentsNoteCallback);
end;
  (*
{@@ ----------------------------------------------------------------------------
  Helper method which writes all Escher-related records required for a cell
  comment:
    MSODRAWING - OBJ - MSODRAWING - TXT
  The NOTE records are written separately
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteCommentsEscher(AStream: TStream;
  AComment: PsComment; ACommentIndex: Integer);
begin
  if ACommentIndex = 0 then
    WriteMSODrawing1(AStream, FWorksheet.Comments.Count, AComment)
  else
    WriteMSODrawing2(AStream, AComment, ACommentIndex+1);
  WriteOBJ(AStream, ACommentIndex+1);
  WriteMSODrawing3(AStream);
  WriteTXO(AStream, AComment);
end;

{@@ ----------------------------------------------------------------------------
  Helper method for writing all NOTE records of a worksheet.
  The Escher-related records required for each cell comment already have been
  written.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteCommentNote(AStream: TStream;
  AComment: PsComment; ACommentIndex: Integer);
begin
  WriteNOTE(AStream, AComment, ACommentIndex+1);
end;
*)
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

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 EOF record.
  This must be the last record on an Excel 8 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_EOF, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 FONT record.
  The font data is passed as an instance of TsFont
-------------------------------------------------------------------------------}
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
  WriteBIFFHeader(AStream, INT_EXCEL_ID_FONT, 16 + Len * Sizeof(WideChar));

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
    WriteFONT(AStream, Workbook.GetFont(i));
end;

{@@ ----------------------------------------------------------------------------
  Writes the first MSODRAWING record to file. It is needed for a comment
  attached to a cell, but also for embedded shapes (currently not supported).

<pre>
  Structure of this record:
                                                    Type    Ver   Inst
     Dg container                                   $F002          0
       |--- FDG record                              $F008    0     1
       |--- SpGr container                          $F003          0
             |---- Sp container  (group shape)      $F004          0
             |       |---- FSpGr record             $F009    1     0
MSODRAWING1  |       |---- FSp record               $F00A    2     0
................................................................................
MSODRAWING2  |---- Sp container  (child shape)      $F004          0
                     |---- FSp record               $F00A    2    202 (Textbox)
                     |---- FOpt record              $F00B    3    13 (num props)
                     |---- Client anchor record     $F010    0     0
                     |---- Client data record       $F011    0     0
</pre>
-------------------------------------------------------------------------------}
procedure TsSpreadBiff8Writer.WriteMSODrawing1(AStream: TStream; ANumShapes: Word;
  AComment: PsComment);
const
  DRAWING_ID = 1;
var
  len: DWord;
  tmpStream: TMemoryStream;
begin
  tmpStream := TMemoryStream.Create;
  try
    { OfficeArtDgContainer record (container of drawing) }
    len := 224 + 152*(ANumShapes - 1);
    WriteMSODgContainer(tmpStream, len);

    { OfficeArtFdg record (info on shapes: num shapes, drawing ID, last Obj ID ) }
    WriteMSOFdgRecord(tmpStream, ANumShapes + 1, DRAWING_ID, SHAPEID_BASE + ANumShapes);

    { OfficeArtSpGrContainer record (shape group container) }
    len := 200 + 152*(ANumShapes - 1);
    WriteMSOSpGrContainer(tmpStream, len);

    { OfficeArtSpContainer record }
    WriteMSOSpContainer(tmpStream, 40);

    { OfficeArtFSpGr record }
    WriteMSOFSpGrRecord(tmpStream, 0, 0, 0, 0);       // 16 + 8 bytes

    { OfficeArtFSp record }
    WriteMSOFSpRecord(tmpStream, SHAPEID_BASE, MSO_SPT_NOTPRIMITIVE,
      MSO_FSP_BITS_GROUP + MSO_FSP_BITS_PATRIARCH);   // 8 + 8 bytes

    { Data for the 1st comment }
    WriteMSODrawing2_Data(tmpStream, AComment, SHAPEID_BASE + 1);

    { Write the BIFF stream }
    tmpStream.Position := 0;
    len := tmpStream.Size;
    WriteBiffHeader(AStream, INT_EXCEL_ID_MSODRAWING, tmpStream.Size);
    AStream.CopyFrom(tmpStream, tmpStream.Size);
  finally
    tmpStream.Free;
  end;
end;

{ Write the MSODRAWING record which occurs before the OBJ record.
  Do not use for the very first OBJ record where the record must be
  WriteMSODrawing1 + WriteMSODrawing2_Data}
procedure TsSpreadBiff8Writer.WriteMSODrawing2(AStream: TStream;
  AComment: PsComment; AObjID: Word);
var
  tmpStream: TStream;
  len: Word;
begin
  tmpStream := TMemoryStream.Create;
  try
    { Shape data for cell comment }
    WriteMSODrawing2_Data(tmpStream, AComment, SHAPEID_BASE + AObjID);

    { Get size of data stream }
    len := tmpStream.Size;

    { BIFF Header }
    WriteBiffHeader(AStream, INT_EXCEL_ID_MSODRAWING, len);

    { Copy MSO data to BIFF stream }
    tmpStream.Position := 0;
    AStream.CopyFrom(tmpStream, len);
  finally
    tmpStream.Free;
  end;
end;

procedure TsSpreadBiff8Writer.WriteMSODrawing2_Data(AStream: TStream;
  AComment: PsComment; AShapeID: Word);
var
  tmpStream: TStream;
  len: Cardinal;
begin
  // We write all the record data to a temporary stream to get the record
  // size (it depends on the number of properties written to the FOPT record.
  // The record size is needed in the very first SpContainer record...

  tmpStream := TMemoryStream.Create;
  try
    { OfficeArtFSp record }
    WriteMSOFSpRecord(tmpStream, AShapeID, MSO_SPT_TEXTBOX,
      MSO_FSP_BITS_HASANCHOR + MSO_FSP_BITS_HASSHAPETYPE);

    { OfficeArtFOpt record }
    WriteMSOFOptRecord_Comment(tmpStream);

    { OfficeArtClientAnchor record }
    WriteMSOClientAnchorSheetRecord(tmpStream,
      AComment^.Row + 1, AComment^.Col + 1, AComment^.Row + 3, AComment^.Col + 5,
      691, 486, 38, 26,
      true, true
    );

    { OfficeArtClientData record }
    WriteMSOClientDataRecord(tmpStream);

    // Now we know the record size
    len := tmpStream.Size;

    // Write an OfficeArtSpContainer to the stream provided...
    WriteMSOSpContainer(AStream, len+8);    // !!! for some reason, Excel wants here 8 additional bytes !!!

    // ... and append the data from the temporary stream.
    tmpStream.Position := 0;
    AStream.Copyfrom(tmpStream, len);

  finally
    tmpStream.Free;
  end;
end;

{ Writes the MSODRAWING record which must occur immediately before a TXO record }
procedure TsSpreadBiff8Writer.WriteMSODRAWING3(AStream: TStream);
begin
  { BIFF Header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_MSODRAWING, 8);

  { OfficeArtClientTextbox record: Text-related data for a shape }
  WriteMSOClientTextBoxRecord(AStream);
end;

{ Writes a NOTE record for a comment attached to a cell }
procedure TsSpreadBiff8Writer.WriteNOTE(AStream: TStream; AComment: PsComment;
  AObjID: Word);
const
  AUTHOR: ansistring = 'author';
var
  len: Integer;
begin
  len := Length(AUTHOR) * sizeOf(ansichar);

  { BIFF Header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_NOTE));  // ID of NOTE record
  AStream.WriteWord(WordToLE(12+len));             // Size of NOTE record

  { Record data }
  AStream.WriteWord(WordToLE(AComment^.Row));      // Row index of cell
  AStream.WriteWord(WordToLE(AComment^.Col));      // Column index of cell
  AStream.WriteWord(0);                            // Flags
  AStream.WriteWord(WordToLE(AObjID));             // Object identifier (1, ...)
  AStream.WriteWord(len);                          // Char length of author string
  AStream.WriteByte(0);                            // Flag for 8-bit characters
  AStream.WriteBuffer(AUTHOR[1], len);             // Author
  AStream.WriteByte(0);                            // Unused
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

{ Writes an OBJ record - belongs to the record required for cell comments }
procedure TsSpreadBIFF8Writer.WriteOBJ(AStream: TStream; AObjID: Word);
var
  guid: TGuid;
begin
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_OBJ));
  AStream.WriteWord(WordToLE(52));

  AStream.WriteWord(WordToLE($0015));  // Subrecord ftCmo
  AStream.WriteWord(WordToLE(18));     // Subrecord size: 18 bytes
  AStream.WriteWord(WordToLE($0019));  // Object type: Comment
  AStream.WriteWord(WordToLE(AObjID)); // Object ID number (1, ... )
  AStream.WriteWord(WordToLE($4011));  // Option flags automatic line style, locked when sheet is protected
  AStream.WriteDWord(0);               // Unused
  AStream.WriteDWord(0);               // Unused
  AStream.WriteDWord(0);               // Unused

  AStream.WriteWord(WordToLE($000D));  // Subrecord ftNts
  AStream.WriteWord(WordToLE(22));     // Size of subrecord: 22 bytes
//  CreateGUID(guid);
  FillChar(guid{%H-}, SizeOf(guid), 0);
  AStream.WriteBuffer(guid, 16);       // GUID of comment
  AStream.WriteWord(WordToLE(0));      // shared note (0 = false)
  AStream.WriteDWord(0);               // unused

  AStream.WriteWord(WordToLE($0000));  // Subrecord ftEnd
  AStream.WriteWord(0);                // Size of subrecord: 0 bytes
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

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 INDEX record

    nm = (rl - rf - 1) / 32 + 1 (using integer division)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteIndex(AStream: TStream);
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_INDEX, 16);

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

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 HYPERLINK record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteHyperlink(AStream: TStream;
  AHyperlink: PsHyperlink; AWorksheet: TsWorksheet);
var
  temp: TStream;
  guid: TGUID;
  widestr: widestring;
  ansistr: ansistring;
  descr: String;
  fn: String;
  flags: DWord;
  size: Integer;
  cell: PCell;
  target, bookmark: String;
  u: TUri;
  isInternal: Boolean;
  dirUpCounter: Integer;
begin
  cell := AWorksheet.FindCell(AHyperlink^.Row, AHyperlink^.Col);
  if (cell = nil) or (AHyperlink^.Target='') then
    exit;

  descr := AWorksheet.ReadAsUTF8Text(cell);      // Hyperlink description
  SplitHyperlink(AHyperlink^.Target, target, bookmark);
  u := ParseURI(AHyperlink^.Target);
  isInternal := (target = '') and (bookmark <> '');
  fn := '';        // Name of local file
  if target <> '' then
  begin
    if (u.Protocol='') then
      fn := target
    else
      UriToFileName(target, fn);
    ForcePathDelims(fn);
  end;

  // Since the length of the record is not known in the first place we write
  // the data to a temporary stream at first.
  temp := TMemoryStream.Create;
  try
    { Cell range using the same hyperlink - we support only single cells }
    temp.WriteWord(WordToLE(cell^.Row));   // first row
    temp.WriteWord(WordToLE(cell^.Row));   // last row
    temp.WriteWord(WordToLE(cell^.Col));   // first column
    temp.WriteWord(WordToLE(cell^.Col));   // last column

    { GUID of standard link }
    guid := StringToGuid('{79EAC9D0-BAF9-11CE-8C82-00AA004BA90B}');
    temp.WriteBuffer(guid, SizeOf(guid));

    { unknown }
    temp.WriteDWord(DWordToLe($00000002));

    { option flags }
    flags := 0;
    if isInternal then
      flags := MASK_HLINK_TEXTMARK or MASK_HLINK_DESCRIPTION
    else
      flags := MASK_HLINK_LINK;
    if SameText(u.Protocol, 'file') then
      flags := flags or MASK_HLINK_ABSOLUTE;
    if descr <> AHyperlink^.Target then
      flags := flags or MASK_HLINK_DESCRIPTION;  // has description
    if bookmark <> '' then
      flags := flags or MASK_HLINK_TEXTMARK;     // link contains a bookmark
    temp.WriteDWord(DWordToLE(flags));

    { description }
    if flags and MASK_HLINK_DESCRIPTION <> 0 then
    begin
      widestr := UTF8Decode(descr);
      { Character count incl trailing zero }
      temp.WriteDWord(DWordToLE(Length(wideStr) + 1));
      { Character array (16-bit characters), plus trailing zeros }
      temp.WriteBuffer(wideStr[1], (Length(wideStr)+1)*SizeOf(widechar));
    end;

    if target <> '' then
    begin
      if (fn <> '') then  // URI is a local file
      begin
       { GUID of file moniker }
        guid := StringToGuid('{00000303-0000-0000-C000-000000000046}');
        temp.WriteBuffer(guid, SizeOf(guid));
        { Convert to ansi - should be DOS 8.3, but this is not necessary }
        ansistr := UTF8ToAnsi(fn);
        { Directory-up level counter  }
        dirUpCounter := 0;
        if not FileNameIsAbsolute(ansistr) then
          while (pos ('..' + PathDelim, ansistr) = 1) do
          begin
            inc(dirUpCounter);
            Delete(ansistr, 1, Length('..'+PathDelim));
          end;
        temp.WriteWord(WordToLE(dirUpCounter));
        { Character count of file name incl trailing zero }
        temp.WriteDWord(DWordToLe(Length(ansistr)+1));
        { Character array of file name (8-bit characters), plus trailing zero }
        temp.WriteBuffer(ansistr[1], Length(ansistr)+1);
        { Unknown }
        temp.WriteDWord(DWordToLE($DEADFFFF));
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        { Size of following file link fields }
        widestr := UTF8ToUTF16(fn);
        size := 4 + 2 + Length(wideStr)*SizeOf(widechar);
        temp.WriteDWord(DWordToLE(size));
        if size > 0 then
        begin
          { Character count of extended file name }
          temp.WriteDWord(DWordToLE(Length(widestr)*SizeOf(WideChar)));
          { Unknown }
          temp.WriteWord(WordToLE($0003));
          { Character array, 16-bit characters, NOT ZERO-TERMINATED! }
          temp.WriteBuffer(widestr[1], Length(wideStr)*SizeOf(WideChar));
        end;
      end
      else begin  { Hyperlink target is a URL }
        widestr := UTF8Decode(target);
        { GUID of URL Moniker }
        guid := StringToGUID('{79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}');
        temp.WriteBuffer(guid, SizeOf(guid));
        { Character count incl trailing zero }
        temp.WriteDWord(DWordToLE(Length(wideStr)+1)*SizeOf(wideChar));
        { Character array plus trailing zero (16-bit characters), plus trailing zeros }
        temp.WriteBuffer(wideStr[1], (length(wideStr)+1)*SizeOf(wideChar));
      end;
    end; // hkURI

    // Hyperlink contains a text mark (#)
    if bookmark <> '' then
    begin
      // Convert to 16-bit characters
      widestr := UTF8Decode(bookmark);
      { Character count of text mark, incl trailing zero }
      temp.WriteDWord(DWordToLE(Length(wideStr) + 1));
      { Character array (16-bit characters) plus trailing zeros }
      temp.WriteBuffer(wideStr[1], (Length(wideStr)+1) * SizeOf(WideChar));
    end;

    { BIFF record header }
    WriteBIFFHeader(AStream, INT_EXCEL_ID_HYPERLINK, temp.Size);

    { Record data }
    temp.Position := 0;
    AStream.CopyFrom(temp, temp.Size);

  finally
    temp.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes all hyperlinks
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteHyperlinks(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  hyperlink: PsHyperlink;
begin
  for hyperlink in AWorksheet.Hyperlinks do begin
    { Write HYPERLINK record }
    WriteHyperlink(AStream, hyperlink, AWorksheet);
    { Write HYPERLINK TOOLTIP record }
    if hyperlink^.Tooltip <> '' then
      WriteHyperlinkTooltip(AStream, hyperlink^.Row, hyperlink^.Col, hyperlink^.Tooltip);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a HYPERLINK TOOLTIP record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteHyperlinkTooltip(AStream: TStream;
  const ARow, ACol: Cardinal; const ATooltip: String);
var
  widestr: widestring;
begin
  widestr := UTF8Decode(ATooltip);

  { BIFF record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_HLINKTOOLTIP,
    10 + (Length(wideStr)+1) * SizeOf(widechar));

  { Repeated record ID }
  AStream.WriteWord(WordToLe(INT_EXCEL_ID_HLINKTOOLTIP));

  { Cell range using the same hyperlink tooltip - we support only single cells }
  AStream.WriteWord(WordToLE(ARow));   // first row
  AStream.WriteWord(WordToLE(ARow));   // last row
  AStream.WriteWord(WordToLE(ACol));   // first column
  AStream.WriteWord(WordToLE(ACol));   // last column

  { Tooltop characters, no length, but trailing zero }
  AStream.WriteBuffer(wideStr[1], (Length(widestr)+1)*SizeOf(wideChar));
end;


{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 LABEL record (string cell value)

  If the string length exceeds 32758 bytes, the string will be truncated,
  a note will be left in the workbooks log.
-------------------------------------------------------------------------------}
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
  n0, n: Integer;
  rng: PsCellRange;
  newRecord: Boolean;
begin
  n0 := AWorksheet.MergedCells.Count;
  n := Min(n0, MAX_PER_RECORD);
  newRecord := true;
  for rng in AWorksheet.MergedCells do
  begin
    if newRecord then
    begin
      newRecord := false;
      { BIFF record header }
      WriteBIFFHeader(AStream, INT_EXCEL_ID_MERGEDCELLS, 2 + n*8);
      { Number of cell ranges in this record }
      AStream.WriteWord(WordToLE(n));
    end;
    { Write range data }
    AStream.WriteWord(WordToLE(rng^.Row1));
    AStream.WriteWord(WordToLE(rng^.Row2));
    AStream.WriteWord(WordToLE(rng^.Col1));
    AStream.WriteWord(WordToLE(rng^.Col2));

    dec(n);
    if n = 0 then begin
      newRecord := true;
      dec(n0, MAX_PER_RECORD);
      n := Min(n0, MAX_PER_RECORD);
    end;
  end;
end;

{@@-----------------------------------------------------------------------------
  Writes an Excel 8 STYLE record

  Registers the name of a user-defined style or specific options
  for a built-in cell style.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteStyle(AStream: TStream);
begin
  { BIFF record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_STYLE, 4);

  { Index to style XF and defines if it's a built-in or used defined style }
  AStream.WriteWord(WordToLE(MASK_STYLE_BUILT_IN));

  { Built-in cell style identifier }
  AStream.WriteByte($00);

  { Level if the identifier for a built-in style is RowLevel or ColLevel, $FF otherwise }
  AStream.WriteByte($FF);
end;

{@@ ----------------------------------------------------------------------------
  Writes a TXO and two CONTINUE records as needed for cell comments.
  It can safely be assumed that the cell exists and contains a comment.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteTXO(AStream: TStream; AComment: PsComment);
var
  recTXO: TBIFF8TXORecord;
  comment: widestring;
  compressed: ansistring;
  len: Integer;
  wchar: widechar;
  i: Integer;
  bytesFmtRuns: Integer;
begin
  { Prepare comment string. It is stored as a string with 8-bit characters }
  comment := UTF8Decode(AComment^.Text);
  SetLength(compressed, length(comment));
  for i:= 1 to Length(comment) do
  begin
    wchar := comment[i];
    compressed[i] := wchar;
  end;
  len := Length(compressed);

  { (1)  TXO record ---------------------------------------------------------- }
  { BIFF record header }
  FillChar(recTXO{%H-}, SizeOf(recTXO), 0);
  recTXO.RecordID := WordToLE(INT_EXCEL_ID_TXO);
  recTXO.RecordSize := SizeOf(recTXO) - 2*SizeOf(word);
  { Record data }
  recTXO.OptionFlags := WordToLE($0212);  // Left & top aligned, lock option on
  recTXO.TextRot := 0;                    // Comment text not rotated
  recTXO.TextLen := WordToLE(len);
  bytesFmtRuns := 8*SizeOf(Word);         // see (3) below
  recTXO.NumFormattingRuns := WordToLE(bytesFmtRuns);
  { Write out to file }
  AStream.WriteBuffer(recTXO, SizeOf(recTXO));

  { (2)  1st CONTINUE record containing the comment text --------------------- }
  { BIFF record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_CONTINUE));
  AStream.WriteWord(len+1);
  { Record data }
  AStream.WriteByte(0);
  AStream.WriteBuffer(compressed[1], len);

  { (3)  2nd CONTINUE record containing the formatting runs ------------------ }
  { BIFF record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_CONTINUE));
  AStream.WriteWord(bytesFmtRuns);
  { Record data }
  AStream.WriteWord(0);             // start index of 1st formatting run (we only use 1 run)
  AStream.WriteWord(WordToLE(1));   // Font index to be used (default font)
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(WordToLE(len)); // lastRun: number of characters
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(0);             // Not used
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 WINDOW2 record

  This record contains additional settings for the document window (BIFF2-BIFF4)
  or for a specific worksheet (BIFF5-BIFF8).

  The values written here are reasonable defaults, which should work for most
  sheets.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteWINDOW2(AStream: TStream;
  ASheet: TsWorksheet);
var
  Options: Word;
begin
  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_WINDOW2, 18);

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

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 XF record (cell format)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteXF(AStream: TStream;
 AFormatRecord: PsCellFormat; XFType_Prot: Byte = 0);
var
  rec: TBIFF8_XFRecord;
  j: Integer;
  b: Byte;
  dw1, dw2: DWord;
  w3: Word;
begin
  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_XF);
  rec.RecordSize := WordToLE(SizeOf(TBIFF8_XFRecord) - SizeOf(TsBIFFHeader));

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
  w3 := 0;
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

  { Background fill }
  if (AFormatRecord <> nil) and (uffBackground in AFormatRecord^.UsedFormattingFields) then
  begin
    // Fill pattern style
    dw2 := dw2 or DWORD(MASK_XF_FILL_PATT[AFormatRecord^.Background.Style] shl 26);
    // Pattern color
    if AFormatRecord^.Background.FgColor = scTransparent
      then w3 := w3 or SYS_DEFAULT_FOREGROUND_COLOR
      else w3 := w3 or FixColor(AFormatRecord^.Background.FgColor);
    // Background color
    if AFormatRecord^.Background.BgColor = scTransparent
      then w3 := w3 or SYS_DEFAULT_BACKGROUND_COLOR shl 7
      else w3 := w3 or (FixColor(AFormatRecord^.Background.BgColor) shl 7);
  end;

  rec.Border_BkGr1 := DWordToLE(dw1);
  rec.Border_BkGr2 := DWordToLE(dw2);
  rec.BkGr3 := WordToLE(w3);

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;


{@@ ----------------------------------------------------------------------------
  Initialization section

  Registers this reader / writer on fpSpreadsheet
  Converts the palette to litte-endian
-------------------------------------------------------------------------------}

initialization

  RegisterSpreadFormat(TsSpreadBIFF8Reader, TsSpreadBIFF8Writer, sfExcel8);
  MakeLEPalette(@PALETTE_BIFF8, Length(PALETTE_BIFF8));

end.

