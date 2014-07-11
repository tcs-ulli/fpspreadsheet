{
xlsxooxml.pas

Writes an OOXML (Office Open XML) document

An OOXML document is a compressed ZIP file with the following files inside:

[Content_Types].xml         -
_rels/.rels                 -
xl/_rels\workbook.xml.rels  -
xl/workbook.xml             - Global workbook data and list of worksheets
xl/styles.xml               -
xl/sharedStrings.xml        -
xl/worksheets\sheet1.xml    - Contents of each worksheet
...
xl/worksheets\sheetN.xml

Specifications obtained from:

http://openxmldeveloper.org/default.aspx

AUTHORS: Felipe Monteiro de Carvalho
}
unit xlsxooxml;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  {$IF FPC_FULLVERSION >= 20701}
  zipper,
  {$ELSE}
  fpszipper,
  {$ENDIF}
  {xmlread, DOM,} AVL_Tree,
  fpspreadsheet, fpsutils;
  
type

  { TsOOXMLFormatList }
  TsOOXMLNumFormatList = class(TsCustomNumFormatList)
  protected
    {
    procedure AddBuiltinFormats; override;
    procedure Analyze(AFormatIndex: Integer; var AFormatString: String;
      var ANumFormat: TsNumberFormat; var ADecimals: Word); override;
      }
  public
    {
    function FormatStringForWriting(AIndex: Integer): String; override;
    }
  end;

  { TsSpreadOOXMLWriter }

  TsSpreadOOXMLWriter = class(TsCustomSpreadWriter)
  private
  protected
    FPointSeparatorSettings: TFormatSettings;
    FSharedStringsCount: Integer;
  protected
    { Helper routines }
    procedure CreateNumFormatList; override;
    procedure CreateStreams;
    procedure DestroyStreams;
    procedure ResetStreams;
    function GetStyleIndex(ACell: PCell): Cardinal;
  protected
    { Streams with the contents of files }
    FSContentTypes: TStream;
    FSRelsRels: TStream;
    FSWorkbook: TStream;
    FSWorkbookRels: TStream;
    FSStyles: TStream;
    FSSharedStrings: TStream;
    FSSharedStrings_complete: TStream;
    FSSheets: array of TStream;
    FCurSheetNum: Integer;
  protected
    { Routines to write the files }
    procedure WriteGlobalFiles;
    procedure WriteContent;
    procedure WriteWorksheet(CurSheet: TsWorksheet);
  protected
    { Record writing methods }
    //todo: add WriteDate
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal; ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal; const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double; ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell); override;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    { General writing methods }
    procedure WriteStringToFile(AFileName, AString: string);
    procedure WriteToFile(const AFileName: string; const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;

implementation

uses
  variants;

const
  { OOXML general XML constants }
  XML_HEADER           = '<?xml version="1.0" encoding="utf-8" ?>';

  { OOXML Directory structure constants }
  // Note: directory separators are always / because the .xlsx is a zip file which
  // requires / instead of \, even on Windows; see 
  // http://www.pkware.com/documents/casestudies/APPNOTE.TXT
  // 4.4.17.1 All slashes MUST be forward slashes '/' as opposed to backwards slashes '\'
  OOXML_PATH_TYPES     = '[Content_Types].xml';
  OOXML_PATH_RELS      = '_rels/';
  OOXML_PATH_RELS_RELS = '_rels/.rels';
  OOXML_PATH_XL        = 'xl/';
  OOXML_PATH_XL_RELS   = 'xl/_rels/';
  OOXML_PATH_XL_RELS_RELS = 'xl/_rels/workbook.xml.rels';
  OOXML_PATH_XL_WORKBOOK = 'xl/workbook.xml';
  OOXML_PATH_XL_STYLES   = 'xl/styles.xml';
  OOXML_PATH_XL_STRINGS  = 'xl/sharedStrings.xml';
  OOXML_PATH_XL_WORKSHEETS = 'xl/worksheets/';

  { OOXML schemas constants }
  SCHEMAS_TYPES        = 'http://schemas.openxmlformats.org/package/2006/content-types';
  SCHEMAS_RELS         = 'http://schemas.openxmlformats.org/package/2006/relationships';
  SCHEMAS_DOC_RELS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  SCHEMAS_DOCUMENT     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  SCHEMAS_WORKSHEET    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  SCHEMAS_STYLES       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  SCHEMAS_STRINGS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
  SCHEMAS_SPREADML     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  { OOXML mime types constants }
  MIME_XML             = 'application/xml';
  MIME_RELS            = 'application/vnd.openxmlformats-package.relationships+xml';
  MIME_SPREADML        = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
  MIME_SHEET           = MIME_SPREADML + '.sheet.main+xml';
  MIME_WORKSHEET       = MIME_SPREADML + '.worksheet+xml';
  MIME_STYLES          = MIME_SPREADML + '.styles+xml';
  MIME_STRINGS         = MIME_SPREADML + '.sharedStrings+xml';


{ TsSpreadOOXMLWriter }

procedure TsSpreadOOXMLWriter.WriteGlobalFiles;
var
  i: Integer;
begin
  { --- Content Types --- }
  AppendToStream(FSContentTypes,
    XML_HEADER);
  AppendToStream(FSContentTypes,
    '<Types xmlns="' + SCHEMAS_TYPES + '">');
  AppendToStream(FSContentTypes,
      '<Override PartName="/_rels/.rels" ContentType="' + MIME_RELS + '" />');
  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />');
  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/workbook.xml" ContentType="' + MIME_SHEET + '" />');

  for i:=1 to Workbook.GetWorksheetCount do
    AppendToStream(FSContentTypes, Format(
      '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%s" />',
        [i, MIME_WORKSHEET]));

  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/styles.xml" ContentType="' + MIME_STYLES + '" />');
  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/sharedStrings.xml" ContentType="' + MIME_STRINGS + '" />');
  AppendToStream(FSContentTypes,
    '</Types>');

  { --- RelsRels --- }
  AppendToStream(FSRelsRels,
    XML_HEADER);
  AppendToStream(FSRelsRels, Format(
    '<Relationships xmlns="%s">', [SCHEMAS_RELS]));
  AppendToStream(FSRelsRels, Format(
    '<Relationship Type="%s" Target="xl/workbook.xml" Id="rId1" />', [SCHEMAS_DOCUMENT]));
  AppendToStream(FSRelsRels,
    '</Relationships>');

  { --- Styles --- }
  AppendToStream(FSStyles,
    XML_Header);
  AppendToStream(FSStyles, Format(
    '<styleSheet xmlns="%s">', [SCHEMAS_SPREADML]));
  AppendToStream(FSStyles,
      '<fonts count="2">');
  AppendToStream(FSStyles,
        '<font><sz val="10" /><name val="Arial" /></font>',
        '<font><sz val="10" /><name val="Arial" /><b val="true"/></font>');
  AppendToStream(FSStyles,
      '</fonts>');
  AppendToStream(FSStyles,
      '<fills count="2">');
  AppendToStream(FSStyles,
        '<fill>',
          '<patternFill patternType="none" />',
        '</fill>');
  AppendToStream(FSStyles,
        '<fill>',
          '<patternFill patternType="gray125" />',
        '</fill>');
  AppendToStream(FSStyles,
      '</fills>');
  AppendToStream(FSStyles,
      '<borders count="1">');
  AppendToStream(FSStyles,
        '<border>',
          '<left /><right /><top /><bottom /><diagonal />',
        '</border>');
  AppendToStream(FSStyles,
      '</borders>');
  AppendToStream(FSStyles,
      '<cellStyleXfs count="2">');
  AppendToStream(FSStyles,
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />',
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" />');
  AppendToStream(FSStyles,
      '</cellStyleXfs>');
  AppendToStream(FSStyles,
      '<cellXfs count="2">');
  AppendToStream(FSStyles,
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />',
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" />');
  AppendToStream(FSStyles,
      '</cellXfs>');
  AppendToStream(FSStyles,
      '<cellStyles count="1">',
        '<cellStyle name="Normal" xfId="0" builtinId="0" />',
      '</cellStyles>');
  AppendToStream(FSStyles,
      '<dxfs count="0" />');
  AppendToStream(FSStyles,
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />');
  AppendToStream(FSStyles,
    '</styleSheet>');
end;

procedure TsSpreadOOXMLWriter.WriteContent;
var
  i: Integer;
begin
  { --- WorkbookRels ---
  { Workbook relations - Mark relation to all sheets }
  AppendToStream(FSWorkbookRels,
    XML_HEADER);
  AppendToStream(FSWorkbookRels,
    '<Relationships xmlns="' + SCHEMAS_RELS + '">');
  AppendToStream(FSWorkbookRels,
      '<Relationship Id="rId1" Type="' + SCHEMAS_STYLES + '" Target="styles.xml" />');
  AppendToStream(FSWorkbookRels,
      '<Relationship Id="rId2" Type="' + SCHEMAS_STRINGS + '" Target="sharedStrings.xml" />');

  for i:=1 to Workbook.GetWorksheetCount do
    AppendToStream(FSWorkbookRels, Format(
      '<Relationship Type="%s" Target="worksheets/sheet%d.xml" Id="rId%d" />',
        [SCHEMAS_WORKSHEET, i, i+2]));

  AppendToStream(FSWOrkbookRels,
    '</Relationships>');

  { --- Workbook --- }
  { Global workbook data - Mark all sheets }
  AppendToStream(FSWorkbook,
    XML_HEADER);
  AppendToStream(FSWorkbook, Format(
    '<workbook xmlns="%s" xmlns:r="%s">', [SCHEMAS_SPREADML, SCHEMAS_DOC_RELS]));
  AppendToStream(FSWorkbook,
      '<fileVersion appName="fpspreadsheet" />');
  AppendToStream(FSWorkbook,
      '<workbookPr defaultThemeVersion="124226" />');
  AppendToStream(FSWorkbook,
      '<bookViews>',
        '<workbookView xWindow="480" yWindow="90" windowWidth="15195" windowHeight="12525" />',
      '</bookViews>');
  AppendToStream(FSWorkbook,
      '<sheets>');
  for i:=1 to Workbook.GetWorksheetCount do
    AppendToStream(FSWorkbook, Format(
        '<sheet name="Sheet%d" sheetId="%d" r:id="rId%d" />', [i, i, i+2]));
  AppendToStream(FSWorkbook,
      '</sheets>');
  AppendToStream(FSWorkbook,
      '<calcPr calcId="114210" />');
  AppendToStream(FSWorkbook,
    '</workbook>');

  // Preparation for shared strings
  FSharedStringsCount := 0;

  // Write all worksheets which fills also the shared strings
  for i := 0 to Workbook.GetWorksheetCount - 1 do
    WriteWorksheet(Workbook.GetWorksheetByIndex(i));

  // Finalization of the shared strings document
  AppendToStream(FSSharedStrings_complete,
    XML_HEADER, Format(
    '<sst xmlns="%s" count="%d" uniqueCount="%d">', [SCHEMAS_SPREADML, FSharedStringsCount, FSharedStringsCount]
  ));
  FSSharedStrings.Position := 0;
  FSSharedStrings_complete.CopyFrom(FSSharedStrings, FSSharedStrings.Size);
  AppendToStream(FSSharedStrings_complete,
    '</sst>');
  FSSharedStrings_complete.Position := 0;
end;

{
FSheets[CurStr] :=
 XML_HEADER + LineEnding +
 '<worksheet xmlns="' + SCHEMAS_SPREADML + '" xmlns:r="' + SCHEMAS_DOC_RELS + '">' + LineEnding +
 '  <sheetViews>' + LineEnding +
 '    <sheetView workbookViewId="0" />' + LineEnding +
 '  </sheetViews>' + LineEnding +
 '  <sheetData>' + LineEnding +
 '  <row r="1" spans="1:4">' + LineEnding +
 '    <c r="A1">' + LineEnding +
 '      <v>1</v>' + LineEnding +
 '    </c>' + LineEnding +
 '    <c r="B1">' + LineEnding +
 '      <v>2</v>' + LineEnding +
 '    </c>' + LineEnding +
 '    <c r="C1">' + LineEnding +
 '      <v>3</v>' + LineEnding +
 '    </c>' + LineEnding +
 '    <c r="D1">' + LineEnding +
 '      <v>4</v>' + LineEnding +
 '    </c>' + LineEnding +
 '  </row>' + LineEnding +
 '  <row r="2" spans="1:4">' + LineEnding +
 '    <c r="A2" t="s">' + LineEnding +
 '      <v>0</v>' + LineEnding +
 '    </c>' + LineEnding +
 '    <c r="B2" t="s">' + LineEnding +
 '      <v>1</v>' + LineEnding +
 '    </c>' + LineEnding +
 '    <c r="C2" t="s">' + LineEnding +
 '      <v>2</v>' + LineEnding +
 '    </c>' + LineEnding +
 '    <c r="D2" t="s">' + LineEnding +
 '      <v>3</v>' + LineEnding +
 '    </c>' + LineEnding +
 '  </row>' + LineEnding +
 '  </sheetData>' + LineEnding +
 '</worksheet>';
}
procedure TsSpreadOOXMLWriter.WriteWorksheet(CurSheet: TsWorksheet);
var
  r, c: Cardinal;
  LastColIndex: Cardinal;
  lCell: TCell;
  AVLNode: TAVLTreeNode;
  CellPosText: string;
  value: Variant;
  fn: String;
begin
  FCurSheetNum := Length(FSSheets);
  SetLength(FSSheets, FCurSheetNum + 1);

  // Create the stream
  if (woSaveMemory in Workbook.WritingOptions) then begin
    fn := IncludeTrailingPathDelimiter(GetTempDir);
    fn := GetTempFileName(fn, Format('fpsSH%d-', [FCurSheetNum+1]));
    FSSheets[FCurSheetNum] := TFileStream.Create(fn, fmCreate);
  end else
    FSSheets[FCurSheetNum] := TMemoryStream.Create;

  // Header
  AppendToStream(FSSheets[FCurSheetNum],
    XML_HEADER);
  AppendToStream(FSSheets[FCurSheetNum], Format(
    '<worksheet xmlns="%s" xmlns:r="%s">', [SCHEMAS_SPREADML, SCHEMAS_DOC_RELS]));
  AppendToStream(FSSheets[FCurSheetNum],
      '<sheetViews>');
  AppendToStream(FSSheets[FCurSheetNum],
        '<sheetView workbookViewId="0" />');
  AppendToStream(FSSheets[FCurSheetNum],
      '</sheetViews>');
  AppendToStream(FSSheets[FCurSheetNum],
      '<sheetData>');

  if (woVirtualMode in Workbook.WritingOptions) and Assigned(Workbook.OnNeedCellData)
  then begin
    for r := 0 to Workbook.VirtualRowCount-1 do begin
      AppendToStream(FSSheets[FCurSheetNum], Format(
        '<row r="%d" spans="1:%d">', [r+1, Workbook.VirtualColCount]));
      for c := 0 to Workbook.VirtualColCount-1 do begin
        FillChar(lCell, SizeOf(lCell), 0);
        CellPosText := CurSheet.CellPosToText(r, c);
        value := varNull;
        Workbook.OnNeedCellData(Workbook, r, c, value);
        lCell.Row := r;
        lCell.Col := c;
        if VarIsNull(value) then
          lCell.ContentType := cctEmpty
        else
        if VarIsNumeric(value) then begin
          lCell.ContentType := cctNumber;
          lCell.NumberValue := value;
        end
        {
        else if VarIsDateTime(value) then begin
          lCell.ContentType := cctNumber;
          lCell.DateTimeValue := value;
        end
        }
        else if VarIsStr(value) then begin
          lCell.ContentType := cctUTF8String;
          lCell.UTF8StringValue := VarToStrDef(value, '');
        end else
        if VarIsBool(value) then begin
          lCell.ContentType := cctBool;
          lCell.BoolValue := value <> 0;
        end;
        WriteCellCallback(@lCell, FSSheets[FCurSheetNum]);
      end;
      AppendToStream(FSSheets[FCurSheetNum],
        '</row>');
    end;
  end else
  begin
    // The cells need to be written in order, row by row, cell by cell
    LastColIndex := CurSheet.GetLastColIndex;
    for r := 0 to CurSheet.GetLastRowIndex do begin
      AppendToStream(FSSheets[FCurSheetNum], Format(
        '<row r="%d" spans="1:%d">', [r+1, LastColIndex+1]));
      // Write cells belonging to this row.
      for c := 0 to LastColIndex do begin
        LCell.Row := r;
        LCell.Col := c;
        AVLNode := CurSheet.Cells.Find(@LCell);
        if Assigned(AVLNode) then
          WriteCellCallback(PCell(AVLNode.Data), nil)
        else begin
          CellPosText := CurSheet.CellPosToText(r, c);
          AppendToStream(FSSheets[FCurSheetNum], Format(
            '<c r="%s">', [CellPosText]),
              '<v></v>',
            '</c>');
        end;
      end;
      AppendToStream(FSSheets[FCurSheetNum],
        '</row>');
    end;
  end;

  // Footer
  AppendToStream(FSSheets[FCurSheetNum],
      '</sheetData>',
    '</worksheet>');
end;

// This is an index to the section cellXfs from the styles.xml file
function TsSpreadOOXMLWriter.GetStyleIndex(ACell: PCell): Cardinal;
begin
  if uffBold in ACell^.UsedFormattingFields then Result := 1
  else Result := 0;
end;

constructor TsSpreadOOXMLWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  // http://en.wikipedia.org/wiki/List_of_spreadsheet_software#Specifications
  FLimitations.MaxCols := 16384;
  FLimitations.MaxRows := 1048576;
end;

procedure TsSpreadOOXMLWriter.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsOOXMLNumFormatList.Create(Workbook);
end;

{ Creates the streams for the individual data files. Will be zipped into a
  single xlsx file. }
procedure TsSpreadOOXMLWriter.CreateStreams;
var
  dir: String;
begin
  if (woSaveMemory in Workbook.WritingOptions) then begin
    dir := IncludeTrailingPathDelimiter(GetTempDir);
    FSContentTypes := TFileStream.Create(GetTempFileName(dir, 'fpsCT'), fmCreate+fmOpenRead);
    FSRelsRels := TFileStream.Create(GetTempFileName(dir, 'fpsRR'), fmCreate+fmOpenRead);
    FSWorkbookRels := TFileStream.Create(GetTempFileName(dir, 'fpsWBR'), fmCreate+fmOpenRead);
    FSWorkbook := TFileStream.Create(GetTempFileName(dir, 'fpsWB'), fmCreate+fmOpenRead);
    FSStyles := TFileStream.Create(GetTempFileName(dir, 'fpsSTY'), fmCreate+fmOpenRead);
    FSSharedStrings := TFileStream.Create(GetTempFileName(dir, 'fpsSST'), fmCreate+fmOpenRead);
    FSSharedStrings_complete := TFileStream.Create(GetTempFileName(dir, 'fpsSSTc'), fmCreate+fmOpenRead);
  end else begin;
    FSContentTypes := TMemoryStream.Create;
    FSRelsRels := TMemoryStream.Create;
    FSWorkbookRels := TMemoryStream.Create;
    FSWorkbook := TMemoryStream.Create;
    FSStyles := TMemoryStream.Create;
    FSSharedStrings := TMemoryStream.Create;
    FSSharedStrings_complete := TMemoryStream.Create;
  end;
  // FSSheets will be created when needed.
end;

{ Destroys the streams that were created by the writer }
procedure TsSpreadOOXMLWriter.DestroyStreams;

  procedure DestroyStream(AStream: TStream);
  var
    fn: String;
  begin
    if AStream is TFileStream then begin
      fn := TFileStream(AStream).Filename;
      DeleteFile(fn);
    end;
    AStream.Free;
  end;

var
  stream: TStream;
begin
  DestroyStream(FSContentTypes);
  DestroyStream(FSRelsRels);
  DestroyStream(FSWorkbookRels);
  DestroyStream(FSWorkbook);
  DestroyStream(FSStyles);
  DestroyStream(FSSharedStrings);
  DestroyStream(FSSharedStrings_complete);
  for stream in FSSheets do DestroyStream(stream);
  SetLength(FSSheets, 0);
end;

{ Is called before zipping the individual file parts. Rewinds the streams. }
procedure TsSpreadOOXMLWriter.ResetStreams;
var
  stream: TStream;
begin
  FSContentTypes.Position := 0;
  FSRelsRels.Position := 0;
  FSWorkbookRels.Position := 0;
  FSWorkbook.Position := 0;
  FSStyles.Position := 0;
  FSSharedStrings_complete.Position := 0;
  for stream in FSSheets do stream.Position := 0;
end;

{
  Writes a string to a file. Helper convenience method.
}
procedure TsSpreadOOXMLWriter.WriteStringToFile(AFileName, AString: string);
var
  TheStream : TFileStream;
  S : String;
begin
  TheStream := TFileStream.Create(AFileName, fmCreate);
  S:=AString;
  TheStream.WriteBuffer(Pointer(S)^,Length(S));
  TheStream.Free;
end;

{
  Writes an OOXML document to the disc
}
procedure TsSpreadOOXMLWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean);
var
  lStream: TFileStream;
  lMode: word;
begin
  if AOverwriteExisting
    then lMode := fmCreate or fmOpenWrite
    else lMode := fmCreate;

  lStream:=TFileStream.Create(AFileName, lMode);
  try
    WriteToStream(lStream);
  finally
    FreeAndNil(lStream);
  end;
end;

procedure TsSpreadOOXMLWriter.WriteToStream(AStream: TStream);
var
  FZip: TZipper;
  i: Integer;
begin
  { Create the streams that will hold the file contents }
  CreateStreams;

  { Fill the streams with the contents of the files }
  WriteGlobalFiles;
  WriteContent;

  { Now compress the files }
  FZip := TZipper.Create;
  try
    FZip.Entries.AddFileEntry(FSContentTypes, OOXML_PATH_TYPES);
    FZip.Entries.AddFileEntry(FSRelsRels, OOXML_PATH_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbookRels, OOXML_PATH_XL_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbook, OOXML_PATH_XL_WORKBOOK);
    FZip.Entries.AddFileEntry(FSStyles, OOXML_PATH_XL_STYLES);
    FZip.Entries.AddFileEntry(FSSharedStrings_complete, OOXML_PATH_XL_STRINGS);

    for i := 0 to Length(FSSheets) - 1 do begin
      FSSheets[i].Position:= 0;
      FZip.Entries.AddFileEntry(FSSheets[i], OOXML_PATH_XL_WORKSHEETS + 'sheet' + IntToStr(i + 1) + '.xml');
    end;

    // Stream position must be at beginning, it was moved to end during adding of xml strings.
    ResetStreams;

    FZip.SaveToStream(AStream);

  finally
    DestroyStreams;
    FZip.Free;
  end;
end;

procedure TsSpreadOOXMLWriter.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  cellPosText: String;
  lStyleIndex: Integer;
begin
  cellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);

  AppendToStream(AStream, Format(
    '<c r="%s" s="%d">', [CellPosText, lStyleIndex]),
      '<v></v>',
    '</c>');
end;


{*******************************************************************
*  TsSpreadOOXMLWriter.WriteLabel ()
*
*  DESCRIPTION:    Writes a string to the sheet
*                  If the string length exceeds 32767 bytes, the string
*                  will be truncated and an exception will be raised as
*                  a warning.
*
*******************************************************************}
procedure TsSpreadOOXMLWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  MaxBytes=32767; //limit for this format
var
  CellPosText: string;
  lStyleIndex: Cardinal;
  TextTooLong: boolean=false;
  ResultingValue: string;
  //S: string;
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);

  // Office 2007-2010 (at least) support no more characters in a cell;
  if Length(AValue)>MaxBytes then
  begin
    TextTooLong:=true;
    ResultingValue:=Copy(AValue,1,MaxBytes); //may chop off multicodepoint UTF8 characters but well...
  end
  else
    ResultingValue:=AValue;

  AppendToStream(FSSharedStrings,
    '<si>',  Format(
      '<t>%s</t>', [UTF8TextToXMLText(ResultingValue)]),
    '</si>' );

  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="s"><v>%d</v></c>', [CellPosText, lStyleIndex, FSharedStringsCount]));
  Inc(FSharedStringsCount);

  {
  //todo: keep a log of errors and show with an exception after writing file or something.
  We can't just do the following

  if TextTooLong then
    Raise Exception.CreateFmt('Text value exceeds %d character limit in cell [%d,%d]. Text has been truncated.',[MaxBytes,ARow,ACol]);
  because the file wouldn't be written.
  }
end;

{
  Writes a number (64-bit IEE 754 floating point) to the sheet
}
procedure TsSpreadOOXMLWriter.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
var
  CellPosText: String;
  CellValueText: String;
  //S: String;
begin
  Unused(AStream, ACell);
  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  CellValueText := Format('%g', [AValue], FPointSeparatorSettings);
  AppendToStream(AStream, Format(
    '<c r="%s" s="0" t="n"><v>%s</v></c>', [CellPosText, CellValueText]));
end;

{*******************************************************************
*  TsSpreadOOXMLWriter.WriteDateTime ()
*
*  DESCRIPTION:    Writes a date/time value as a text
*                  ISO 8601 format is used to preserve interoperability
*                  between locales.
*
*  Note: this should be replaced by writing actual date/time values
*
*******************************************************************}
procedure TsSpreadOOXMLWriter.WriteDateTime(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
begin
  WriteLabel(AStream, ARow, ACol, FormatDateTime(ISO8601Format, AValue), ACell);
end;

{
  Registers this reader / writer on fpSpreadsheet
}
initialization

  RegisterSpreadFormat(TsCustomSpreadReader, TsSpreadOOXMLWriter, sfOOXML);

end.

