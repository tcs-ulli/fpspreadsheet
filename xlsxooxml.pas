{
xlsxooxml.pas

Writes an OOXML (Office Open XML) document

An OOXML document is a compressed ZIP file with the following files inside:

[Content_Types].xml         -
_rels\.rels                 -
xl\_rels\workbook.xml.rels  -
xl\workbook.xml             - Global workbook data and list of worksheets
xl\styles.xml               -
xl\sharedStrings.xml        -
xl\worksheets\sheet1.xml    - Contents of each worksheet
...
xl\worksheets\sheetN.xml

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
  protected
    FPointSeparatorSettings: TFormatSettings;
    { Strings with the contents of files }
    FContentTypes: string;
    FRelsRels: string;
    FWorkbookString, FWorkbookRelsString, FStylesString, FSharedStrings: string;
    FSheets: array of string;
    FSharedStringsCount: Integer;

  protected
    { Helper routines }
    procedure CreateNumFormatList; override;
  protected
    { Streams with the contents of files }
    FSContentTypes: TStringStream;
    FSRelsRels: TStringStream;
    FSWorkbook, FSWorkbookRels, FSStyles, FSSharedStrings: TStringStream;
    FSSheets: array of TStringStream;
    FCurSheetNum: Integer;
  protected
    { Routines to write those files }
    procedure WriteGlobalFiles;
    procedure WriteContent;
    procedure WriteWorksheet(CurSheet: TsWorksheet);
    function GetStyleIndex(ACell: PCell): Cardinal;
  protected
    { Record writing methods }
    //todo: add WriteDate
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal; const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double; ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell); override;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    { General writing methods }
    procedure WriteStringToFile(AFileName, AString: string);
    procedure WriteToFile(const AFileName: string; const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;

implementation

const
  { OOXML general XML constants }
  XML_HEADER           = '<?xml version="1.0" encoding="utf-8" ?>';

  { OOXML Directory structure constants }
  OOXML_PATH_TYPES     = '[Content_Types].xml';
  OOXML_PATH_RELS      = '_rels' + PathDelim;
  OOXML_PATH_RELS_RELS = '_rels' + PathDelim + '.rels';
  OOXML_PATH_XL        = 'xl' + PathDelim;
  OOXML_PATH_XL_RELS   = 'xl' + PathDelim + '_rels' + PathDelim;
  OOXML_PATH_XL_RELS_RELS = 'xl' + PathDelim + '_rels' + PathDelim + 'workbook.xml.rels';
  OOXML_PATH_XL_WORKBOOK = 'xl' + PathDelim + 'workbook.xml';
  OOXML_PATH_XL_STYLES   = 'xl' + PathDelim + 'styles.xml';
  OOXML_PATH_XL_STRINGS  = 'xl' + PathDelim + 'sharedStrings.xml';
  OOXML_PATH_XL_WORKSHEETS = 'xl' + PathDelim + 'worksheets' + PathDelim;

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
//  WriteCellsToStream(AStream, AData.GetFirstWorksheet.FCells);

  FContentTypes :=
   XML_HEADER + LineEnding +
   '<Types xmlns="' + SCHEMAS_TYPES + '">' + LineEnding +
//   '  <Default Extension="xml" ContentType="' + MIME_XML + '" />' + LineEnding +
//   '  <Default Extension="rels" ContentType="' + MIME_RELS + '" />' + LineEnding +
   '  <Override PartName="/_rels/.rels" ContentType="' + MIME_RELS + '" />' + LineEnding +
//   <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
//   <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
   '  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />' + LineEnding +
   '  <Override PartName="/xl/workbook.xml" ContentType="' + MIME_SHEET + '" />' + LineEnding;
  for i := 1 to Workbook.GetWorksheetCount do
  begin
    FContentTypes := FContentTypes +
    Format('  <Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%s" />', [i, MIME_WORKSHEET]) + LineEnding;
  end;
  FContentTypes := FContentTypes +
   '  <Override PartName="/xl/styles.xml" ContentType="' + MIME_STYLES + '" />' + LineEnding +
   '  <Override PartName="/xl/sharedStrings.xml" ContentType="' + MIME_STRINGS + '" />' + LineEnding +
   '</Types>';

  FRelsRels :=
   XML_HEADER + LineEnding +
   '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding +
   '<Relationship Type="' + SCHEMAS_DOCUMENT + '" Target="xl/workbook.xml" Id="rId1" />' + LineEnding +
   '</Relationships>';

  FStylesString :=
   XML_HEADER + LineEnding +
   '<styleSheet xmlns="' + SCHEMAS_SPREADML + '">' + LineEnding +
   '  <fonts count="2">' + LineEnding +
   '    <font><sz val="10" /><name val="Arial" /></font>' + LineEnding +
   '    <font><sz val="10" /><name val="Arial" /><b val="true"/></font>' + LineEnding +
   '  </fonts>' + LineEnding +
   '  <fills count="2">' + LineEnding +
   '    <fill>' + LineEnding +
   '      <patternFill patternType="none" />' + LineEnding +
   '    </fill>' + LineEnding +
   '    <fill>' + LineEnding +
   '      <patternFill patternType="gray125" />' + LineEnding +
   '    </fill>' + LineEnding +
   '  </fills>' + LineEnding +
   '  <borders count="1">' + LineEnding +
   '    <border>' + LineEnding +
   '      <left />' + LineEnding +
   '      <right />' + LineEnding +
   '      <top />' + LineEnding +
   '      <bottom />' + LineEnding +
   '      <diagonal />' + LineEnding +
   '    </border>' + LineEnding +
   '  </borders>' + LineEnding +
   '  <cellStyleXfs count="2">' + LineEnding +
   '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />' + LineEnding +
   '    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" />' + LineEnding +
   '  </cellStyleXfs>' + LineEnding +
   '  <cellXfs count="2">' + LineEnding +
   '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />' + LineEnding +
   '    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" />' + LineEnding +
   '  </cellXfs>' + LineEnding +
   '  <cellStyles count="1">' + LineEnding +
   '    <cellStyle name="Normal" xfId="0" builtinId="0" />' + LineEnding +
   '  </cellStyles>' + LineEnding +
   '  <dxfs count="0" />' + LineEnding +
   '  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />' + LineEnding +
   '</styleSheet>';
end;

procedure TsSpreadOOXMLWriter.WriteContent;
var
  i: Integer;
begin
  { Workbook relations - Mark relation to all sheets }
  FWorkbookRelsString :=
   XML_HEADER + LineEnding +
   '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding +
   '<Relationship Id="rId1" Type="' + SCHEMAS_STYLES + '" Target="styles.xml" />' + LineEnding +
   '<Relationship Id="rId2" Type="' + SCHEMAS_STRINGS + '" Target="sharedStrings.xml" />' + LineEnding;

  for i := 1 to Workbook.GetWorksheetCount do
  begin
    FWorkbookRelsString := FWorkbookRelsString +
      Format('<Relationship Type="%s" Target="worksheets/sheet%d.xml" Id="rId%d" />', [SCHEMAS_WORKSHEET, i, i+2]) + LineEnding;
  end;

  FWorkbookRelsString := FWorkbookRelsString +
   '</Relationships>';

  // Global workbook data - Mark all sheets
  FWorkbookString :=
   XML_HEADER + LineEnding +
   '<workbook xmlns="' + SCHEMAS_SPREADML + '" xmlns:r="' + SCHEMAS_DOC_RELS + '">' + LineEnding +
   '  <fileVersion appName="fpspreadsheet" />' + LineEnding + // lastEdited="4" lowestEdited="4" rupBuild="4505"
   '  <workbookPr defaultThemeVersion="124226" />' + LineEnding +
   '  <bookViews>' + LineEnding +
   '    <workbookView xWindow="480" yWindow="90" windowWidth="15195" windowHeight="12525" />' + LineEnding +
   '  </bookViews>' + LineEnding;

  FWorkbookString := FWorkbookString + '  <sheets>' + LineEnding;
  for i := 1 to Workbook.GetWorksheetCount do
    FWorkbookString := FWorkbookString +
      Format('    <sheet name="Sheet%d" sheetId="%d" r:id="rId%d" />', [i, i, i+2]) + LineEnding;
  FWorkbookString := FWorkbookString + '  </sheets>' + LineEnding;

  FWorkbookString := FWorkbookString +
   '  <calcPr calcId="114210" />' + LineEnding +
   '</workbook>';

  // Preparation for Shared strings
  FSharedStringsCount := 0;
  FSharedStrings := '';

  // Write all worksheets, which fills also FSharedStrings
  SetLength(FSheets, 0);

  for i := 0 to Workbook.GetWorksheetCount - 1 do
    WriteWorksheet(Workbook.GetWorksheetByIndex(i));

  // Finalization of the shared strings document
  FSharedStrings :=
   XML_HEADER + LineEnding +
   '<sst xmlns="' + SCHEMAS_SPREADML + '" count="' + IntToStr(FSharedStringsCount) +
     '" uniqueCount="' + IntToStr(FSharedStringsCount) + '">' + LineEnding +
   FSharedStrings +
   '</sst>';
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
  j, k: Integer;
  LastColIndex: Cardinal;
  LCell: TCell;
  AVLNode: TAVLTreeNode;
  CellPosText: string;
begin
  FCurSheetNum := Length(FSheets);
  SetLength(FSheets, FCurSheetNum + 1);

  LastColIndex := CurSheet.GetLastColIndex;

  // Header
  FSheets[FCurSheetNum] :=
   XML_HEADER + LineEnding +
   '<worksheet xmlns="' + SCHEMAS_SPREADML + '" xmlns:r="' + SCHEMAS_DOC_RELS + '">' + LineEnding +
   '  <sheetViews>' + LineEnding +
   '    <sheetView workbookViewId="0" />' + LineEnding +
   '  </sheetViews>' + LineEnding +
   '  <sheetData>' + LineEnding;

  // The cells need to be written in order, row by row, cell by cell
  for j := 0 to CurSheet.GetLastRowIndex do
  begin
    FSheets[FCurSheetNum] := FSheets[FCurSheetNum] +
     Format('  <row r="%d" spans="1:%d">', [j+1,LastColIndex+1]) + LineEnding;

    // Write cells from this row.
    for k := 0 to LastColIndex do
    begin
      LCell.Row := j;
      LCell.Col := k;
      AVLNode := CurSheet.Cells.Find(@LCell);
      if Assigned(AVLNode) then
        WriteCellCallback(PCell(AVLNode.Data), nil)
      else
      begin
        CellPosText := CurSheet.CellPosToText(j, k);
        FSheets[FCurSheetNum] := FSheets[FCurSheetNum] +
         Format('    <c r="%s">', [CellPosText]) + LineEnding +
         '      <v></v>' + LineEnding +
         '    </c>' + LineEnding;
      end;
    end;

    FSheets[FCurSheetNum] := FSheets[FCurSheetNum] +
     '  </row>' + LineEnding;
  end;

  // Footer
  FSheets[FCurSheetNum] := FSheets[FCurSheetNum] +
   '  </sheetData>' + LineEnding +
   '</worksheet>';
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
end;

destructor TsSpreadOOXMLWriter.Destroy;
begin
  SetLength(FSheets, 0);
  SetLength(FSSheets, 0);

  inherited Destroy;
end;

procedure TsSpreadOOXMLWriter.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsOOXMLNumFormatList.Create(Workbook);
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
  { Fill the strings with the contents of the files }

  WriteGlobalFiles;
  WriteContent;

  { Write the data to streams }

  FSContentTypes := TStringStream.Create(FContentTypes);
  FSRelsRels := TStringStream.Create(FRelsRels);
  FSWorkbookRels := TStringStream.Create(FWorkbookRelsString);
  FSWorkbook := TStringStream.Create(FWorkbookString);
  FSStyles := TStringStream.Create(FStylesString);
  FSSharedStrings := TStringStream.Create(FSharedStrings);

  SetLength(FSSheets, Length(FSheets));

  for i := 0 to Length(FSheets) - 1 do
    FSSheets[i] := TStringStream.Create(FSheets[i]);

  { Now compress the files }

  FZip := TZipper.Create;
  try
    FZip.Entries.AddFileEntry(FSContentTypes, OOXML_PATH_TYPES);
    FZip.Entries.AddFileEntry(FSRelsRels, OOXML_PATH_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbookRels, OOXML_PATH_XL_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbook, OOXML_PATH_XL_WORKBOOK);
    FZip.Entries.AddFileEntry(FSStyles, OOXML_PATH_XL_STYLES);
    FZip.Entries.AddFileEntry(FSSharedStrings, OOXML_PATH_XL_STRINGS);

    for i := 0 to Length(FSheets) - 1 do
      FZip.Entries.AddFileEntry(FSSheets[i], OOXML_PATH_XL_WORKSHEETS + 'sheet' + IntToStr(i + 1) + '.xml');

    FZip.SaveToStream(AStream);
  finally
    FSContentTypes.Free;
    FSRelsRels.Free;
    FSWorkbookRels.Free;
    FSWorkbook.Free;
    FSStyles.Free;
    FSSharedStrings.Free;

    for i := 0 to Length(FSSheets) - 1 do
      FSSheets[i].Free;

    FZip.Free;
  end;
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

  FSharedStrings := FSharedStrings +
          '  <si>' + LineEnding +
   Format('    <t>%s</t>', [ResultingValue]) + LineEnding +
          '  </si>' + LineEnding;

  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  FSheets[FCurSheetNum] := FSheets[FCurSheetNum] +
   Format('    <c r="%s" s="%d" t="s"><v>%d</v></c>', [CellPosText, lStyleIndex, FSharedStringsCount]) + LineEnding;

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
begin
  Unused(AStream, ACell);
  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  CellValueText := Format('%g', [AValue], FPointSeparatorSettings);
  FSheets[FCurSheetNum] := FSheets[FCurSheetNum] +
   Format('    <c r="%s" s="0" t="n"><v>%s</v></c>', [CellPosText, CellValueText]) + LineEnding;
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

