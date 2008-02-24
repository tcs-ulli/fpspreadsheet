{
xlsxooxml.pas

Writes an OOXML (Office Open XML) document

An OOXML document is a compressed ZIP file with the following files inside:

[Content_Types].xml
_rels\.rels
xl\_rels\workbook.xml.rels
xl\workbook.xml
xl\styles.xml
xl\sharedStrings.xml
xl\worksheets\sheet1.xml
...
xl\worksheets\sheetN.xml

Specifications obtained from:

http://openxmldeveloper.org/default.aspx

AUTHORS: Felipe Monteiro de Carvalho

IMPORTANT: This writer doesn't work yet!!! This is just initial code.
}
unit xlsxooxml;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils, {zipper,}
  fpspreadsheet;
  
type

  { TsSpreadOOXMLWriter }

  TsSpreadOOXMLWriter = class(TsCustomSpreadWriter)
  protected
//    FZip: TZipper;
    FContentTypes: string;
    FRelsRels: string;
    FWorkbook, FWorkbookRels, FStyles, FSharedString, FSheet1: string;
  public
    { General writing methods }
    procedure WriteStringToFile(AFileName, AString: string);
    procedure WriteToFile(AFileName: string; AData: TsWorkbook); override;
    procedure WriteToStream(AStream: TStream; AData: TsWorkbook); override;
    { Record writing methods }
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Word; const AValue: string); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double); override;
  end;

implementation

const
  { OOXML general XML constants }
  XML_HEADER           = '<?xml version="1.0" encoding="utf-8" ?>';

  { OOXML Directory structure constants }
  OOXML_PATH_TYPES     = '[Content_Types].xml';
  OOXML_PATH_RELS      = '_rels\';
  OOXML_PATH_RELS_RELS = '_rels\.rels';
  OOXML_PATH_XL        = 'xl\';
  OOXML_PATH_XL_RELS   = 'xl\_rels\';
  OOXML_PATH_XL_RELS_RELS = 'xl\_rels\workbook.xml.rels';
  OOXML_PATH_XL_WORKBOOK = 'xl\workbook.xml';
  OOXML_PATH_XL_STYLES   = 'xl\styles.xml';
  OOXML_PATH_XL_STRINGS  = 'xl\sharedStrings.xml';
  OOXML_PATH_XL_WORKSHEETS = 'xl\worksheets\';

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

{*******************************************************************
*  TsSpreadOOXMLWriter.WriteStringToFile ()
*
*  DESCRIPTION:    Writes a string to a file. Helper convenience method.
*
*******************************************************************}
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

{*******************************************************************
*  TsSpreadOOXMLWriter.WriteToFile ()
*
*  DESCRIPTION:    Writes an OOXML document to the disc
*
*******************************************************************}
procedure TsSpreadOOXMLWriter.WriteToFile(AFileName: string; AData: TsWorkbook);
var
  TempDir: string;
begin
  {FZip := TZipper.Create;
  FZip.ZipFiles(AFileName, x);
  FZip.Free;}
  
  WriteToStream(nil, AData);

  TempDir := IncludeTrailingBackslash(AFileName);

  { files on the root path }

  ForceDirectories(TempDir);

  WriteStringToFile(TempDir + OOXML_PATH_TYPES, FContentTypes);
  
  { _rels directory }

  ForceDirectories(TempDir + OOXML_PATH_RELS);

  WriteStringToFile(TempDir + OOXML_PATH_RELS_RELS, FRelsRels);

  { xl directory }

  ForceDirectories(TempDir + OOXML_PATH_XL_RELS);
  
  WriteStringToFile(TempDir + OOXML_PATH_XL_RELS_RELS, FWorkbookRels);
  
  WriteStringToFile(TempDir + OOXML_PATH_XL_WORKBOOK, FWorkbook);

  WriteStringToFile(TempDir + OOXML_PATH_XL_STYLES, FStyles);

  WriteStringToFile(TempDir + OOXML_PATH_XL_STRINGS, FSharedString);
  
  { xl\worksheets directory }

  ForceDirectories(TempDir + OOXML_PATH_XL_WORKSHEETS);

  WriteStringToFile(TempDir + OOXML_PATH_XL_WORKSHEETS + 'sheet1.xml', FSheet1);
end;

{*******************************************************************
*  TsSpreadOOXMLWriter.WriteToStream ()
*
*  DESCRIPTION:    Writes an Excel 2 file to a stream
*
*                  Excel 2.x files support only one Worksheet per Workbook,
*                  so only the first will be written.
*
*******************************************************************}
procedure TsSpreadOOXMLWriter.WriteToStream(AStream: TStream; AData: TsWorkbook);
begin
//  WriteCellsToStream(AStream, AData.GetFirstWorksheet.FCells);

  FContentTypes :=
   XML_HEADER + LineEnding +
   '<Types xmlns="' + SCHEMAS_TYPES + '">' + LineEnding +
   '  <Default Extension="xml" ContentType="' + MIME_XML + '" />' + LineEnding +
   '  <Default Extension="rels" ContentType="' + MIME_RELS + '" />' + LineEnding +
   '  <Override PartName="/xl/workbook.xml" ContentType="' + MIME_SHEET + '" />' + LineEnding +
   '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="' + MIME_WORKSHEET + '" />' + LineEnding +
   '  <Override PartName="/xl/styles.xml" ContentType="' + MIME_STYLES + '" />' + LineEnding +
   '  <Override PartName="/xl/sharedStrings.xml" ContentType="' + MIME_STRINGS + '" />' + LineEnding +
   '</Types>';
    
  FRelsRels :=
   XML_HEADER + LineEnding +
   '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding +
   '<Relationship Type="' + SCHEMAS_DOCUMENT + '" Target="/xl/workbook.xml" Id="rId1" />' + LineEnding +
   '</Relationships>';
    
  FWorkbookRels :=
   XML_HEADER + LineEnding +
   '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding +
   '<Relationship Type="' + SCHEMAS_WORKSHEET + '" Target="/xl/worksheets/sheet1.xml" Id="rId1" />' + LineEnding +
   '<Relationship Type="' + SCHEMAS_STYLES + '" Target="/xl/styles.xml" Id="rId2" />' + LineEnding +
   '<Relationship Type="' + SCHEMAS_STRINGS + '" Target="/xl/sharedStrings.xml" Id="rId3" />' + LineEnding +
   '</Relationships>';

  FWorkbook :=
   XML_HEADER + LineEnding +
   '<workbook xmlns="' + SCHEMAS_SPREADML + '" xmlns:r="' + SCHEMAS_DOC_RELS + '">' + LineEnding +
   '  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505" />' + LineEnding +
   '  <workbookPr defaultThemeVersion="124226" />' + LineEnding +
   '  <bookViews>' + LineEnding +
   '    <workbookView xWindow="480" yWindow="90" windowWidth="15195" windowHeight="12525" />' + LineEnding +
   '  </bookViews>' + LineEnding +
   '  <sheets>' + LineEnding +
   '    <sheet name="Sheet1" sheetId="1" r:id="rId1" />' + LineEnding +
   '  </sheets>' + LineEnding +
   '  <calcPr calcId="114210" />' + LineEnding +
   '</workbook>';

  FStyles :=
   XML_HEADER + LineEnding +
   '<styleSheet xmlns="' + SCHEMAS_SPREADML + '">' + LineEnding +
   '  <fonts count="1">' + LineEnding +
   '    <font>' + LineEnding +
   '      <sz val="10" />' + LineEnding +
   '      <name val="Arial" />' + LineEnding +
   '    </font>' + LineEnding +
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
   '  <cellStyleXfs count="1">' + LineEnding +
   '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />' + LineEnding +
   '  </cellStyleXfs>' + LineEnding +
   '  <cellXfs count="1">' + LineEnding +
   '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />' + LineEnding +
   '  </cellXfs>' + LineEnding +
   '  <cellStyles count="1">' + LineEnding +
   '    <cellStyle name="Normal" xfId="0" builtinId="0" />' + LineEnding +
   '  </cellStyles>' + LineEnding +
   '  <dxfs count="0" />' + LineEnding +
   '  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />' + LineEnding +
   '</styleSheet>';

  FSharedString :=
   XML_HEADER + LineEnding +
   '<sst xmlns="' + SCHEMAS_SPREADML + '" count="4" uniqueCount="4">' + LineEnding +
   '  <si>' + LineEnding +
   '    <t>First</t>' + LineEnding +
   '  </si>' + LineEnding +
   '  <si>' + LineEnding +
   '    <t>Second</t>' + LineEnding +
   '  </si>' + LineEnding +
   '  <si>' + LineEnding +
   '    <t>Third</t>' + LineEnding +
   '  </si>' + LineEnding +
   '  <si>' + LineEnding +
   '    <t>Fourth</t>' + LineEnding +
   '  </si>' + LineEnding +
   '</sst>';

  FSheet1 :=
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
   '    </row>' + LineEnding +
   '  </sheetData>' + LineEnding +
   '</worksheet>';

end;

{*******************************************************************
*  TsSpreadOOXMLWriter.WriteLabel ()
*
*  DESCRIPTION:    Writes an Excel 2 LABEL record
*
*                  Writes a string to the sheet
*
*******************************************************************}
procedure TsSpreadOOXMLWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Word; const AValue: string);
var
  L: Byte;
begin
  L := Length(AValue);

  { BIFF Record header }
//  AStream.WriteWord(WordToLE(INT_EXCEL_ID_LABEL));
//  AStream.WriteWord(WordToLE(8 + L));

  { BIFF Record data }
//  AStream.WriteWord(WordToLE(ARow));
//  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  AStream.WriteByte($0);
  AStream.WriteByte($0);
  AStream.WriteByte($0);

  { String with 8-bit size }
  AStream.WriteByte(L);
  AStream.WriteBuffer(AValue[1], L);
end;

{*******************************************************************
*  TsSpreadOOXMLWriter.WriteNumber ()
*
*  DESCRIPTION:    Writes an Excel 2 NUMBER record
*
*                  Writes a number (64-bit IEE 754 floating point) to the sheet
*
*******************************************************************}
procedure TsSpreadOOXMLWriter.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double);
begin
  { BIFF Record header }
//  AStream.WriteWord(WordToLE(INT_EXCEL_ID_NUMBER));
//  AStream.WriteWord(WordToLE(15));

  { BIFF Record data }
//  AStream.WriteWord(WordToLE(ARow));
//  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  AStream.WriteByte($0);
  AStream.WriteByte($0);
  AStream.WriteByte($0);

  { IEE 754 floating-point value }
  AStream.WriteBuffer(AValue, 8);
end;

{*******************************************************************
*  Initialization section
*
*  Registers this reader / writer on fpSpreadsheet
*
*******************************************************************}
initialization

  RegisterSpreadFormat(TsCustomSpreadReader, TsSpreadOOXMLWriter, sfOOXML);

end.

