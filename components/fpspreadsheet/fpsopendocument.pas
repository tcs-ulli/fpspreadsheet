{
fpsopendocument.pas

Writes an OpenDocument 1.0 Spreadsheet document

An OpenDocument document is a compressed ZIP file with the following files inside:

content.xml     - Actual contents
meta.xml        - Authoring data
settings.xml    - User persistent viewing information, such as zoom, cursor position, etc.
styles.xml      - Styles, which are the only way to do formatting
mimetype        - application/vnd.oasis.opendocument.spreadsheet
META-INF\manifest.xml  - Describes the other files in the archive

Specifications obtained from:

http://docs.oasis-open.org/office/v1.1/OS/OpenDocument-v1.1.pdf

AUTHORS: Felipe Monteiro de Carvalho
}
unit fpsopendocument;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  fpszipper, {NOTE: fpszipper is the latest zipper.pp Change to standard zipper when FPC 2.4 is released }
  fpspreadsheet;
  
type

  { TsSpreadOpenDocWriter }

  TsSpreadOpenDocWriter = class(TsCustomSpreadWriter)
  protected
    // Strings with the contents of files
    FMeta, FSettings, FStyles, FContent, FMimetype: string;
    FMetaInfManifest: string;
    // Streams with the contents of files
    FSMeta, FSSettings, FSStyles, FSContent, FSMimetype: TStringStream;
    FSMetaInfManifest: TStringStream;
    // Routines to write those files
    procedure WriteGlobalFiles;
    procedure WriteContent(AData: TsWorkbook);
    procedure WriteWorksheet(CurSheet: TsWorksheet);
  public
    { General writing methods }
    procedure WriteStringToFile(AString, AFileName: string);
    procedure WriteToFile(AFileName: string; AData: TsWorkbook); override;
    procedure WriteToStream(AStream: TStream; AData: TsWorkbook); override;
    { Record writing methods }
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Word; const AFormula: TsFormula); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Word; const AValue: string); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double); override;
  end;

implementation

const
  { OpenDocument general XML constants }
  XML_HEADER           = '<?xml version="1.0" encoding="utf-8" ?>';

  { OpenDocument Directory structure constants }
  OPENDOC_PATH_CONTENT   = 'content.xml';
  OPENDOC_PATH_META      = 'meta.xml';
  OPENDOC_PATH_SETTINGS  = 'settings.xml';
  OPENDOC_PATH_STYLES    = 'styles.xml';
  OPENDOC_PATH_MIMETYPE  = 'mimetype';
  OPENDOC_PATH_METAINF = 'META-INF' + '/';
  OPENDOC_PATH_METAINF_MANIFEST = 'META-INF' + '/' + 'manifest.xml';

  { OpenDocument schemas constants }
  SCHEMAS_XMLNS_OFFICE   = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0';
  SCHEMAS_XMLNS_DCTERMS  = 'http://purl.org/dc/terms/';
  SCHEMAS_XMLNS_META     = 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0';
  SCHEMAS_XMLNS          = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
  SCHEMAS_XMLNS_CONFIG   = 'urn:oasis:names:tc:opendocument:xmlns:config:1.0';
  SCHEMAS_XMLNS_OOO      = 'http://openoffice.org/2004/office';
  SCHEMAS_XMLNS_MANIFEST = 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0';
  SCHEMAS_XMLNS_FO       = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0';
  SCHEMAS_XMLNS_STYLE    = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
  SCHEMAS_XMLNS_SVG      = 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0';
  SCHEMAS_XMLNS_TABLE    = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
  SCHEMAS_XMLNS_TEXT     = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
  SCHEMAS_XMLNS_V        = 'urn:schemas-microsoft-com:vml';
  SCHEMAS_XMLNS_NUMBER   = 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0';
  SCHEMAS_XMLNS_CHART    = 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0';
  SCHEMAS_XMLNS_DR3D     = 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0';
  SCHEMAS_XMLNS_MATH     = 'http://www.w3.org/1998/Math/MathML';
  SCHEMAS_XMLNS_FORM     = 'urn:oasis:names:tc:opendocument:xmlns:form:1.0';
  SCHEMAS_XMLNS_SCRIPT   = 'urn:oasis:names:tc:opendocument:xmlns:script:1.0';
  SCHEMAS_XMLNS_OOOW     = 'http://openoffice.org/2004/writer';
  SCHEMAS_XMLNS_OOOC     = 'http://openoffice.org/2004/calc';
  SCHEMAS_XMLNS_DOM      = 'http://www.w3.org/2001/xml-events';
  SCHEMAS_XMLNS_XFORMS   = 'http://www.w3.org/2002/xforms';
  SCHEMAS_XMLNS_XSD      = 'http://www.w3.org/2001/XMLSchema';
  SCHEMAS_XMLNS_XSI      = 'http://www.w3.org/2001/XMLSchema-instance';

{ TsSpreadOpenDocWriter }

procedure TsSpreadOpenDocWriter.WriteGlobalFiles;
begin
  FMimetype := 'application/vnd.oasis.opendocument.spreadsheet';

  FMetaInfManifest :=
   XML_HEADER + LineEnding +
   '<manifest:manifest xmlns:manifest="' + SCHEMAS_XMLNS_MANIFEST + '">' + LineEnding +
   '  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:full-path="/" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml" />' + LineEnding +
   '</manifest:manifest>';

  FMeta :=
   XML_HEADER + LineEnding +
   '<office:document-meta xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
     '" xmlns:dcterms="' + SCHEMAS_XMLNS_DCTERMS +
     '" xmlns:meta="' + SCHEMAS_XMLNS_META +
     '" xmlns="' + SCHEMAS_XMLNS +
     '" xmlns:ex="' + SCHEMAS_XMLNS + '">' + LineEnding +
   '  <office:meta>' + LineEnding +
   '    <meta:generator>FPSpreadsheet Library</meta:generator>' + LineEnding +
   '    <meta:document-statistic />' + LineEnding +
   '  </office:meta>' + LineEnding +
   '</office:document-meta>';

  FSettings :=
   XML_HEADER + LineEnding +
   '<office:document-settings xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
     '" xmlns:config="' + SCHEMAS_XMLNS_CONFIG +
     '" xmlns:ooo="' + SCHEMAS_XMLNS_OOO + '">' + LineEnding +
   '<office:settings>' + LineEnding +
   '  <config:config-item-set config:name="ooo:view-settings">' + LineEnding +
   '    <config:config-item-map-indexed config:name="Views">' + LineEnding +
   '      <config:config-item-map-entry>' + LineEnding +
   '        <config:config-item config:name="ActiveTable" config:type="string">Tabelle1</config:config-item>' + LineEnding +
   '        <config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>' + LineEnding +
   '        <config:config-item config:name="PageViewZoomValue" config:type="int">100</config:config-item>' + LineEnding +
   '        <config:config-item config:name="ShowPageBreakPreview" config:type="boolean">false</config:config-item>' + LineEnding +
   '        <config:config-item config:name="HasColumnRowHeaders" config:type="boolean">true</config:config-item>' + LineEnding +
   '          <config:config-item-map-named config:name="Tables">' + LineEnding +
   '            <config:config-item-map-entry config:name="Tabelle1">' + LineEnding +
   '              <config:config-item config:name="CursorPositionX" config:type="int">3</config:config-item>' + LineEnding +
   '              <config:config-item config:name="CursorPositionY" config:type="int">2</config:config-item>' + LineEnding +
   '            </config:config-item-map-entry>' + LineEnding +
   '          </config:config-item-map-named>' + LineEnding +
   '        </config:config-item-map-entry>' + LineEnding +
   '      </config:config-item-map-indexed>' + LineEnding +
   '    </config:config-item-set>' + LineEnding +
   '  </office:settings>' + LineEnding +
   '</office:document-settings>';

  FStyles :=
   XML_HEADER + LineEnding +
   '<office:document-styles xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
     '" xmlns:fo="' + SCHEMAS_XMLNS_FO +
     '" xmlns:style="' + SCHEMAS_XMLNS_STYLE +
     '" xmlns:svg="' + SCHEMAS_XMLNS_SVG +
     '" xmlns:table="' + SCHEMAS_XMLNS_TABLE +
     '" xmlns:text="' + SCHEMAS_XMLNS_TEXT +
     '" xmlns:v="' + SCHEMAS_XMLNS_V + '">' + LineEnding +
   '<office:font-face-decls>' + LineEnding +
   '<style:font-face style:name="Arial" svg:font-family="Arial" />' + LineEnding +
   '</office:font-face-decls>' + LineEnding +
   '<office:styles>' + LineEnding +
   '<style:style style:name="Default" style:family="table-cell">' + LineEnding +
   '<style:text-properties fo:font-size="10" style:font-name="Arial" />' + LineEnding +
   '</style:style>' + LineEnding +
   '</office:styles>' + LineEnding +
   '<office:automatic-styles>' + LineEnding +
   '<style:page-layout style:name="pm1">' + LineEnding +
   '<style:page-layout-properties fo:margin-top="1.25cm" fo:margin-bottom="1.25cm" fo:margin-left="1.905cm" fo:margin-right="1.905cm" />' + LineEnding +
   '<style:header-style>' + LineEnding +
   '<style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-bottom="0.25cm" fo:margin-top="0cm" />' + LineEnding +
   '</style:header-style>' + LineEnding +
   '<style:footer-style>' + LineEnding +
   '<style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0.25cm" fo:margin-bottom="0cm" />' + LineEnding +
   '</style:footer-style>' + LineEnding +
   '</style:page-layout>' + LineEnding +
   '</office:automatic-styles>' + LineEnding +
   '<office:master-styles>' + LineEnding +
   '<style:master-page style:name="Default" style:page-layout-name="pm1">' + LineEnding +
   '<style:header />' + LineEnding +
   '<style:header-left style:display="false" />' + LineEnding +
   '<style:footer />' + LineEnding +
   '<style:footer-left style:display="false" />' + LineEnding +
   '</style:master-page>' + LineEnding +
   '</office:master-styles>' + LineEnding +
   '</office:document-styles>';
end;

procedure TsSpreadOpenDocWriter.WriteContent(AData: TsWorkbook);
var
  i: Integer;
begin
  FContent :=
   XML_HEADER + LineEnding +
   '<office:document-content xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
     '" xmlns:fo="'     + SCHEMAS_XMLNS_FO +
     '" xmlns:style="'  + SCHEMAS_XMLNS_STYLE +
     '" xmlns:text="'   + SCHEMAS_XMLNS_TEXT +
     '" xmlns:table="'  + SCHEMAS_XMLNS_TABLE +
     '" xmlns:svg="'    + SCHEMAS_XMLNS_SVG +
     '" xmlns:number="' + SCHEMAS_XMLNS_NUMBER +
     '" xmlns:meta="'   + SCHEMAS_XMLNS_META +
     '" xmlns:chart="'  + SCHEMAS_XMLNS_CHART +
     '" xmlns:dr3d="'   + SCHEMAS_XMLNS_DR3D +
     '" xmlns:math="'   + SCHEMAS_XMLNS_MATH +
     '" xmlns:form="'   + SCHEMAS_XMLNS_FORM +
     '" xmlns:script="' + SCHEMAS_XMLNS_SCRIPT +
     '" xmlns:ooo="'    + SCHEMAS_XMLNS_OOO +
     '" xmlns:ooow="'   + SCHEMAS_XMLNS_OOOW +
     '" xmlns:oooc="'   + SCHEMAS_XMLNS_OOOC +
     '" xmlns:dom="'    + SCHEMAS_XMLNS_DOM +
     '" xmlns:xforms="' + SCHEMAS_XMLNS_XFORMS +
     '" xmlns:xsd="'    + SCHEMAS_XMLNS_XSD +
     '" xmlns:xsi="'    + SCHEMAS_XMLNS_XSI + '">' + LineEnding +
   '  <office:scripts />' + LineEnding +

   // Fonts
   '  <office:font-face-decls>' + LineEnding +
   '    <style:font-face style:name="Arial" svg:font-family="Arial" xmlns:v="urn:schemas-microsoft-com:vml" />' + LineEnding +
   '  </office:font-face-decls>' + LineEnding +

   // Automatic styles
  '  <office:automatic-styles>' + LineEnding +
  '    <style:style style:name="co1" style:family="table-column">' + LineEnding +
  '      <style:table-column-properties fo:break-before="auto" style:column-width="2.267cm"/>' + LineEnding +
  '    </style:style>' + LineEnding +
  '    <style:style style:name="ro1" style:family="table-row">' + LineEnding +
  '      <style:table-row-properties style:row-height="0.416cm" fo:break-before="auto" style:use-optimal-row-height="true"/>' + LineEnding +
  '    </style:style>' + LineEnding +
  '    <style:style style:name="ta1" style:family="table" style:master-page-name="Default">' + LineEnding +
  '      <style:table-properties table:display="true" style:writing-mode="lr-tb"/>' + LineEnding +
  '    </style:style>' + LineEnding +
  '  </office:automatic-styles>' + LineEnding +

  // Body
  '  <office:body>' + LineEnding +
  '    <office:spreadsheet>' + LineEnding;

  // Write all worksheets
  for i := 0 to AData.GetWorksheetCount - 1 do
  begin
    WriteWorksheet(Adata.GetWorksheetByIndex(i));
  end;

  FContent :=  FContent +
   '    </office:spreadsheet>' + LineEnding +
   '  </office:body>' + LineEnding +
   '</office:document-content>';
end;

procedure TsSpreadOpenDocWriter.WriteWorksheet(CurSheet: TsWorksheet);
var
  j, k: Integer;
  CurCell: PCell;
  CurRow: array of PCell;
  LastColNum: Cardinal;
begin
  LastColNum := CurSheet.GetLastColNumber;

  // Header
  FContent := FContent +
  '    <table:table table:name="' + CurSheet.Name + '" table:style-name="ta1">' + LineEnding +
  '      <table:table-column table:style-name="co1" table:number-columns-repeated="' +
  IntToStr(LastColNum + 1) + '" table:default-cell-style-name="Default"/>' + LineEnding;

  // The cells need to be written in order, row by row, cell by cell
  for j := 0 to CurSheet.GetLastRowNumber do
  begin
    FContent := FContent +
    '      <table:table-row table:style-name="ro1">' + LineEnding;

    // First make an array with the cells of this row in their respective order
    // nil pointers indicate empty cells, so it's necessary to initialize the array
    SetLength(CurRow, LastColNum + 1);
    for k := 0 to LastColNum do CurRow[k] := nil;

    // Now fill the array with the cells in their proper place
    for k := 0 to CurSheet.FCells.Count - 1 do
    begin
      CurCell := CurSheet.FCells.Items[k];
      if CurCell^.Row = j then CurRow[CurCell^.Col] := CurCell;
    end;

    // And now write all cells from this row
    for k := 0 to LastColNum do
    begin
      CurCell := CurRow[k];

      if CurCell = nil then
        FContent := FContent + '<table:table-cell/>' + LineEnding
      else WriteCellCallback(CurCell, nil);
    end;

    FContent := FContent +
    '      </table:table-row>' + LineEnding;
  end;

  // Clean up
  SetLength(CurRow, 0);

  // Footer
  FContent := FContent +
  '    </table:table>' + LineEnding;
end;

{
  Writes a string to a file. Helper convenience method.
}
procedure TsSpreadOpenDocWriter.WriteStringToFile(AString, AFileName: string);
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
  Writes an OOXML document to the disc.
}
procedure TsSpreadOpenDocWriter.WriteToFile(AFileName: string; AData: TsWorkbook);
var
  FZip: TZipper;
begin
  { Fill the strings with the contents of the files }

  WriteGlobalFiles();
  WriteContent(AData);

  { Write the data to streams }

  FSMeta := TStringStream.Create(FMeta);
  FSSettings := TStringStream.Create(FSettings);
  FSStyles := TStringStream.Create(FStyles);
  FSContent := TStringStream.Create(FContent);
  FSMimetype := TStringStream.Create(FMimetype);
  FSMetaInfManifest := TStringStream.Create(FMetaInfManifest);

  { Now compress the files }

  FZip := TZipper.Create;
  try
    FZip.FileName := AFileName;

    FZip.Entries.AddFileEntry(FSMeta, OPENDOC_PATH_META);
    FZip.Entries.AddFileEntry(FSSettings, OPENDOC_PATH_SETTINGS);
    FZip.Entries.AddFileEntry(FSStyles, OPENDOC_PATH_STYLES);
    FZip.Entries.AddFileEntry(FSContent, OPENDOC_PATH_CONTENT);
    FZip.Entries.AddFileEntry(FSMimetype, OPENDOC_PATH_MIMETYPE);
    FZip.Entries.AddFileEntry(FSMetaInfManifest, OPENDOC_PATH_METAINF_MANIFEST);

    FZip.ZipAllFiles;
  finally
    FZip.Free;
    FSMeta.Free;
    FSSettings.Free;
    FSStyles.Free;
    FSContent.Free;
    FSMimetype.Free;
    FSMetaInfManifest.Free;
  end;
end;


procedure TsSpreadOpenDocWriter.WriteToStream(AStream: TStream; AData: TsWorkbook);
begin
  // Not supported at the moment
  raise Exception.Create('TsSpreadOpenDocWriter.WriteToStream not supported');
end;

procedure TsSpreadOpenDocWriter.WriteFormula(AStream: TStream; const ARow,
  ACol: Word; const AFormula: TsFormula);
begin
{  // The row should already be the correct one
  FContent := FContent +
    '  <table:table-cell office:value-type="string">' + LineEnding +
    '    <text:p>' + AFormula.DoubleValue + '</text:p>' + LineEnding +
    '  </table:table-cell>' + LineEnding;
<table:table-cell table:formula="of:=[.A1]+[.B2]" office:value-type="float" office:value="1833">
<text:p>1833</text:p>
</table:table-cell>}
end;

procedure TsSpreadOpenDocWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Word; const AValue: string);
begin
  // The row should already be the correct one
  FContent := FContent +
    '  <table:table-cell office:value-type="string">' + LineEnding +
    '    <text:p>' + AValue + '</text:p>' + LineEnding +
    '  </table:table-cell>' + LineEnding;
end;

procedure TsSpreadOpenDocWriter.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double);
begin
  // The row should already be the correct one
  FContent := FContent +
    '  <table:table-cell office:value-type="float" office:value="' + FloatToStr(AValue) + '">' + LineEnding +
    '    <text:p>' + FloatToStr(AValue) + '</text:p>' + LineEnding +
    '  </table:table-cell>' + LineEnding;
end;

{
  Registers this reader / writer on fpSpreadsheet
}
initialization

  RegisterSpreadFormat(TsCustomSpreadReader, TsSpreadOpenDocWriter, sfOpenDocument);

end.

