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

AUTHORS: Felipe Monteiro de Carvalho / Jose Luis Jurado Rincon
}


unit fpsopendocument;

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

{.$define FPSPREADDEBUG} //used to be XLSDEBUG
interface

uses
  Classes, SysUtils,
  laz2_xmlread, laz2_DOM,
  AVL_Tree, math, dateutils, contnrs,
 {$IF FPC_FULLVERSION >= 20701}
  zipper,
 {$ELSE}
  fpszipper,
 {$ENDIF}
  fpstypes, fpspreadsheet, fpsReaderWriter, fpsutils, fpsHeaderFooterParser,
  fpsNumFormat, fpsNumFormatParser, fpsxmlcommon;
  
type
  TDateMode=(
    dm1899 {default for ODF; almost same as Excel 1900},
    dm1900 {StarCalc legacy only},
    dm1904 {e.g. Quattro Pro, Mac Excel compatibility}
  );

  { TsSpreadOpenDocHeaderFooterParser }

  TsSpreadOpenDocHeaderFooterParser = class(TsHeaderFooterParser)
  private
    FNode: TDOMNode;
    XMLMode: Boolean;
  protected
    procedure AddNodeElement(ANode: TDOMNode);
    function FindStyle(AStyleName: String): Integer;
    function GetCurrFontIndex: Integer; override;
    procedure Parse; override;
  public
    constructor Create(ANode: TDOMNode; AFontList: TList;
      ADefaultFont: TsHeaderFooterFont); overload;
    function BuildHeaderFooterAsXMLString: String;
  end;

  { TsSpreadOpenDocNumFormatParser }

  TsSpreadOpenDocNumFormatParser = class(TsNumFormatParser)
  protected
    function BuildXMLAsStringFromSection(ASection: Integer;
      AFormatName: String): String;
  public
    function BuildXMLAsString(AFormatName: String): String;
  end;

  { TsSpreadOpenDocReader }

  TsSpreadOpenDocReader = class(TsSpreadXMLReader)
  private
    FColumnStyleList: TFPList;
    FColumnList: TFPList;
    FRowStyleList: TFPList;
    FRowList: TFPList;
    FPageLayoutList: TFPList;
    FMasterPageList: TFPList;
    FHeaderFooterFontList: TObjectList;
    FDateMode: TDateMode;
    // Applies internally stored column widths to current worksheet
    procedure ApplyColWidths;
    // Applies a style to a cell
    function ApplyStyleToCell(ACell: PCell; AStyleName: String): Boolean;
    // Extracts a boolean value from the xml node
    function ExtractBoolFromNode(ANode: TDOMNode): Boolean;
    // Extracts the date/time value from the xml node
    function ExtractDateTimeFromNode(ANode: TDOMNode;
      ANumFormat: TsNumberFormat; const AFormatStr: String): TDateTime;
    // Searches a column style by its column index or its name in the StyleList
    function FindColumnByCol(AColIndex: Integer): Integer;
    function FindColStyleByName(AStyleName: String): integer;
    function FindNumFormatByName(ANumFmtName: String): Integer;
    function FindRowStyleByName(AStyleName: String): Integer;
    procedure ReadColumns(ATableNode: TDOMNode);
    procedure ReadColumnStyle(AStyleNode: TDOMNode);
    // Figures out the base year for times in this file (dates are unambiguous)
    procedure ReadDateMode(SpreadSheetNode: TDOMNode);
    function ReadFont(ANode: TDOMnode; APreferredIndex: Integer = -1): Integer;
    procedure ReadHeaderFooterFont(ANode: TDOMNode; var AFontName: String;
      var AFontSize: Double; var AFontStyle: TsHeaderFooterFontStyles;
      var AFontColor: TsColorValue);
    function ReadHeaderFooterText(ANode: TDOMNode): String;
    procedure ReadRowsAndCells(ATableNode: TDOMNode);
    procedure ReadRowStyle(AStyleNode: TDOMNode);

  protected
    FPointSeparatorSettings: TFormatSettings;
    procedure AddBuiltinNumFormats; override;
    procedure ReadAutomaticStyles(AStylesNode: TDOMNode);
    procedure ReadMasterStyles(AStylesNode: TDOMNode);
    procedure ReadNumFormats(AStylesNode: TDOMNode);
    function ReadPageLayout(AStylesNode: TDOMNode; ATableStyleName: String): PsPageLayout;
    procedure ReadSettings(AOfficeSettingsNode: TDOMNode);
    procedure ReadStyles(AStylesNode: TDOMNode);
    { Record writing methods }
    procedure ReadBlank(ARow, ACol: Cardinal; ACellNode: TDOMNode); reintroduce;
    procedure ReadBoolean(ARow, ACol: Cardinal; ACellNode: TDOMNode);
    procedure ReadComment(ARow, ACol: Cardinal; ACellNode: TDOMNode);
    procedure ReadDateTime(ARow, ACol: Cardinal; ACellNode: TDOMNode);
    procedure ReadFormula(ARow, ACol: Cardinal; ACellNode: TDOMNode); reintroduce;
    procedure ReadLabel(ARow, ACol: Cardinal; ACellNode: TDOMNode); reintroduce;
    procedure ReadNumber(ARow, ACol: Cardinal; ACellNode: TDOMNode); reintroduce;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;

    { General reading methods }
    procedure ReadFromFile(AFileName: string); override;
    procedure ReadFromStream(AStream: TStream); override;
  end;

  { TsSpreadOpenDocWriter }

  TsSpreadOpenDocWriter = class(TsCustomSpreadWriter)
  private
    FColumnStyleList: TFPList;
    FRowStyleList: TFPList;
    FHeaderFooterFontList: TObjectList;

    // Routines to write parts of files
    procedure WriteAutomaticStyles(AStream: TStream);
    procedure WriteCellStyles(AStream: TStream);
    procedure WriteColStyles(AStream: TStream);
    procedure WriteColumns(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteFontNames(AStream: TStream);
    procedure WriteMasterStyles(AStream: TStream);
    procedure WriteNumFormats(AStream: TStream);
    procedure WriteRowStyles(AStream: TStream);
    procedure WriteRowsAndCells(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteTableSettings(AStream: TStream);
    procedure WriteTableStyles(AStream: TStream);
    procedure WriteVirtualCells(AStream: TStream; ASheet: TsWorksheet);

    function WriteBackgroundColorStyleXMLAsString(const AFormat: TsCellFormat): String;
    function WriteBorderStyleXMLAsString(const AFormat: TsCellFormat): String;
    function WriteCommentXMLAsString(AComment: String): String;
    function WriteDefaultFontXMLAsString: String;
    function WriteFontStyleXMLAsString(const AFormat: TsCellFormat): String; overload;
    function WriteFontStyleXMLAsString(AFont: TsFont): String; overload;
    function WriteHeaderFooterFontXMLAsString(AFont: TsHeaderFooterFont): String;
    function WriteHorAlignmentStyleXMLAsString(const AFormat: TsCellFormat): String;
    function WritePageLayoutAsXMLString(AStyleName: String; const APageLayout: TsPageLayout): String;
    function WriteTextRotationStyleXMLAsString(const AFormat: TsCellFormat): String;
    function WriteVertAlignmentStyleXMLAsString(const AFormat: TsCellFormat): String;
    function WriteWordwrapStyleXMLAsString(const AFormat: TsCellFormat): String;

  protected
    FPointSeparatorSettings: TFormatSettings;
    // Streams with the contents of files
    FSMeta, FSSettings, FSStyles, FSContent, FSMimeType, FSMetaInfManifest: TStream;

    { Helpers }
    procedure AddBuiltinNumFormats; override;
    procedure CreateStreams;
    procedure DestroyStreams;
    procedure ListAllColumnStyles;
    procedure ListAllHeaderFooterFonts;
    procedure ListAllNumFormats; override;
    procedure ListAllRowStyles;
    procedure ResetStreams;

    { Routines to write those files }
    procedure WriteContent;
    procedure WriteMimetype;
    procedure WriteMetaInfManifest;
    procedure WriteMeta;
    procedure WriteSettings;
    procedure WriteStyles;
    procedure WriteWorksheet(AStream: TStream; ASheetIndex: Integer);

    { Record writing methods }
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); override;
    procedure WriteError(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TsErrorValue; ACell: PCell); override;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;

    { General writing methods }
    procedure WriteStringToFile(AString, AFileName: string);
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;

implementation

uses
  StrUtils, Variants, LazFileUtils, URIParser,
  fpsPatches, fpsStrings, fpsStreams, fpsExprParser;

const
  { OpenDocument general XML constants }
  XML_HEADER             = '<?xml version="1.0" encoding="utf-8" ?>';

  { OpenDocument Directory structure constants }
  OPENDOC_PATH_CONTENT   = 'content.xml';
  OPENDOC_PATH_META      = 'meta.xml';
  OPENDOC_PATH_SETTINGS  = 'settings.xml';
  OPENDOC_PATH_STYLES    = 'styles.xml';
  OPENDOC_PATH_MIMETYPE  = 'mimetype';
  {%H-}OPENDOC_PATH_METAINF   = 'META-INF' + '/';
  {%H-}OPENDOC_PATH_METAINF_MANIFEST = 'META-INF' + '/' + 'manifest.xml';

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

  { DATEMODE similar to but not the same as XLS format; used in time only values. }
  DATEMODE_1899_BASE=0; //apparently 1899-12-30 for ODF in FPC DateTime;
  // due to Excel's leap year bug, the date floats in the spreadsheets are the same starting
  // 1900-03-01
  DATEMODE_1900_BASE=2; //StarCalc compatibility, 1900-01-01 in FPC DateTime
  DATEMODE_1904_BASE=1462; //1/1/1904 in FPC TDateTime

const
  // lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair)
  BORDER_LINESTYLES: array[TsLineStyle] of string =
    ('solid', 'solid', 'dashed', 'fine-dashed', 'solid', 'double', 'dotted');
  BORDER_LINEWIDTHS: array[TsLinestyle] of string =
    ('0.002cm', '2pt', '0.002cm', '0.002cm', '3pt', '0.039cm', '0.002cm');

  FALSE_TRUE: Array[boolean] of String = ('false', 'true');

  COLWIDTH_EPS  = 1e-2;    // for mm
  ROWHEIGHT_EPS = 1e-2;    // for lines

type

  { Column style items stored in ColStyleList of the reader }
  TColumnStyleData = class
  public
    Name: String;
    ColWidth: Double;             // in mm
  end;

  { Column data items stored in the ColumnList }
  TColumnData = class
  public
    Col: Integer;
    ColStyleIndex: integer;   // index into FColumnStyleList of reader
    DefaultCellStyleIndex: Integer;   // Index of default cell style in FCellStyleList of reader
  end;

  { Row style items stored in RowStyleList of the reader }
  TRowStyleData = class
  public
    Name: String;
    RowHeight: Double;          // in mm
    AutoRowHeight: Boolean;
  end;

  { PageLayout items stored in PageLayoutList }
  TPageLayoutData = class
  public
    Name: String;
    PageLayout: TsPageLayout;
  end;

  { MasterPage items stored in MasterPageList }
  TMasterPageData = class
  public
    Name: String;
    PageLayoutName: String;
  end;

  (* --- presently not used, but this may change... ---

  { Row data items stored in the RowList of the reader }
  TRowData = class
    Row: Integer;
    RowStyleIndex: Integer;   // index into FRowStyleList of reader
    DefaultCellStyleIndex: Integer;  // Index of default row style in FCellStyleList of reader
  end;
  *)

{******************************************************************************}
{                          TXMLHeaderFooterFont                                }
{******************************************************************************}
type
  TXMLHeaderFooterFont = class(TsHeaderFooterFont)
    StyleName: String;
  end;

{******************************************************************************}
{                    TsSpreadOpenDocHeaderFooterParser                         }
{******************************************************************************}

constructor TsSpreadOpenDocHeaderFooterParser.Create(ANode: TDOMNode;
  AFontList: TList; ADefaultFont: TsHeaderFooterFont);
begin
  inherited Create;
  XMLMode := true;
  FNode := ANode; // this is a child of the "<style-header>" etc nodes.
  FFontList := AFontList;
  FDefaultFont := ADefaultFont;
  FFontClass := TXMLHeaderFooterFont;
  Parse;
end;

procedure TsSpreadOpenDocHeaderFooterParser.AddNodeElement(ANode: TDOMNode);
var
  nodeName: String;
begin
  nodeName := ANode.NodeName;
  case nodeName of
    'text:sheet-name':
      AddElement(hftSheetName);
    'text:file-name' :
      case GetAttrValue(ANode, 'text:display') of
        'full': begin
                  AddElement(hftPath);
                  AddElement(hftFileName);
                end;
        'path': AddElement(hftPath);
        else    AddElement(hftFileName);
      end;
    'text:date':
      AddElement(hftDate);
    'text:time':
      AddElement(hftTime);
    'text:page-number':
      AddElement(hftPage);
    'text:page-count':
      AddElement(hftPageCount);
    '#text':
      begin
        FCurrText := ANode.NodeValue; //GetNodeValue(ANode);
        AddCurrTextElement;
      end;
  end;
end;

function TsSpreadOpenDocHeaderFooterParser.BuildHeaderFooterAsXMLString: String;
var
  regionStr: array[TsHeaderFooterSectionIndex] of String;
  sec: TsHeaderFooterSectionIndex;
  element: TsHeaderFooterElement;
  styleName: String;
  s: String;
begin
  for sec := hfsLeft to hfsRight do
  begin
    if Length(FSections[sec]) = 0 then
      Continue;
    regionStr[sec] := '';
    for element in FSections[sec] do
    begin
      stylename := TXMLHeaderFooterFont(FFontList[element.FontIndex]).StyleName;
      case element.Token of
        hftText:
          s := Format('<text:span text:style-name="%s">%s</text:span>', [
                 stylename, UTF8TextToXMLText(element.TextValue)]);
        hftNewLine:
          s := '</text:p><text:p>';
        hftSheetName:
          s := '<text:sheet-name>???</text:sheet-name>';
        hftPath:
          s := '<text:file-name text:display="path">???</text:file-name>';
        hftFileName:
          s := '<text:file-name text:display="name-and-extension">???</text:file-name>';
        hftDate:
          s := Format('<text:date style:data-style-name="N2" text:date-value="%s">%s</text:date>', [
                 FormatDateTime('yyyy"-"mm"-"dd', date()), DateToStr(date())]);
        hftTime:
          s := Format('<text:time>%s</text:time>', [
                 FormatDateTime('hh:nn:ss', time()) ]);
        hftPage:
          s := '<text:page-number>1</text:page-number>';
        hftPageCount:
          s := '<text:page-count>1</text:page-count>';
      end;
      regionStr[sec] := regionStr[sec] + s;
    end; // for element
    if regionStr[sec] <> '' then
      regionStr[sec] := '<text:p>' + regionStr[sec] + '</text:p>';
  end;  // for sec

  Result := '';
  for sec := hfsLeft to hfsRight do
  begin
    case sec of
      hfsLeft   : s := 'style:region-left';
      hfsCenter : s := 'style:region-center';
      hfsRight  : s := 'style:region-right';
    end;
    if regionStr[sec] = '' then
      Result := Result + '<' + s + ' />' else
      Result := Result + '<' + s + '>' + regionStr[sec] + '</' + s + '>';
  end;
end;

function TsSpreadOpenDocHeaderFooterParser.FindStyle(AStyleName: String): Integer;
var
  fnt: TXMLHeaderFooterFont;
begin
  for Result := 0 to FFontList.Count-1 do
  begin
    fnt := TXMLHeaderFooterFont(FFontList[Result]);
    if SameText(fnt.StyleName, AStyleName) then
      exit;
  end;
  Result := -1;
end;

function TsSpreadOpenDocHeaderFooterParser.GetCurrFontIndex: Integer;
begin
  if XMLMode then
    Result := FCurrFontIndex
  else
    Result := inherited GetCurrFontIndex;
end;

procedure TsSpreadOpenDocHeaderFooterParser.Parse;
var
  node, pnode, childpnode, childspannode: TDOMNode;
  nodeName: String;
  s: String;
  firstP: Boolean;
begin
  FFontClass := TXMLHeaderFooterFont;

  if not XMLMode then
  begin
    inherited Parse;
    exit;
  end;

  node := FNode;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    case nodeName of
      'style:region-left'   : FCurrSection := hfsLeft;
      'style:region-center' : FCurrSection := hfsCenter;
      'style:region-right'  : FCurrSection := hfsRight;
    end;
    firstP := true;
    pnode := node.FirstChild;
    while Assigned(pnode) do
    begin
      nodeName := pnode.NodeName;
      if nodeName = 'text:p' then
      begin
        if not firstP then AddElement(hftNewLine);
        childpnode := pnode.FirstChild;
        while Assigned(childpnode) do
        begin
          nodeName := childpnode.NodeName;
          if nodeName = 'text:span' then begin
            s := GetAttrValue(childpnode, 'text:style-name');
            if s <> '' then
              FCurrFontIndex := FindStyle(s) else
              FCurrFontIndex := -1;
            childspannode := childpnode.FirstChild;
            while Assigned(childspannode) do
            begin
              nodeName := childspannode.NodeName;
              AddNodeElement(childspannode);
              childspannode := childspannode.NextSibling;
            end;
          end else
            AddNodeElement(childpnode);
          childpnode := childpnode.NextSibling;
        end;
        firstP := false;
      end;
      pnode := pnode.NextSibling;
    end;
    node := node.NextSibling;
  end;
end;


{******************************************************************************}
{                     TsSpreadOpenDocNumFormatParser                           }
{******************************************************************************}

function TsSpreadOpenDocNumFormatParser.BuildXMLAsString(AFormatName: String): String;
var
  i: Integer;
begin
  Result := '';
  { When there is only one section the next statement is the only one executed.
    When there are several sections the file contains at first the
    positive section (index 0), then the negative section (index 1), and
    finally the zero section (index 2) which contains the style-map. }
  for i:=0 to Length(FSections)-1 do
    Result := Result + BuildXMLAsStringFromSection(i, AFormatName);
end;

function TsSpreadOpenDocNumFormatParser.BuildXMLAsStringFromSection(
  ASection: Integer; AFormatName: String): String;
var
  n: Integer;
  el, nEl: Integer;
  ns: Integer;
  clr: TsColorvalue;
  mask: String;
  timeIntervalStr: String;
  styleMapStr: String;
  int,num,denom: Integer;
begin
  Result := '';

  ns := Length(FSections);
  if (ns = 0) then
    exit;

  styleMapStr := '';
  timeIntervalStr := '';

  if (ns > 1) then
  begin
    // The file corresponding to the last section contains the styleMap.
    if (ASection = ns - 1) then
      case ns of
        2: styleMapStr :=
             '<style:map ' +
               'style:apply-style-name="' + AFormatName + 'P0" ' +
               'style:condition="value()&gt;=0" />';                   // >= 0
        3: styleMapStr :=
             '<style:map '+
               'style:apply-style-name="' + AFormatName + 'P0" ' +     // > 0
               'style:condition="value()&gt;0" />' +
             '<style:map '+
               'style:apply-style-name="' + AFormatName + 'P1" ' +     // < 0
               'style:condition="value()&lt;0" />';
        else
          raise Exception.Create('At most 3 format sections allowed.');
      end
    else
      AFormatName := AFormatName + 'P' + IntToStr(ASection);
  end;

  with FSections[ASection] do
  begin
    nEl := Length(Elements);
    el := 0;
    while (el < nEl) do begin
      case Elements[el].Token of
        nftColor:
          begin
            clr := FWorkbook.GetPaletteColor(Elements[el].IntValue);
            Result := Result + '<style:text-properties fo:color="' + ColorToHTMLColorStr(clr) + '" />';
          end;

        nftSign, nftSignBracket, nftText, nftSpace:
          if Elements[el].TextValue = ' ' then
            Result := Result + '<number:text><![CDATA[ ]]></number:text>'
          else
            Result := Result + '<number:text>' + Elements[el].TextValue + '</number:text>';

        nftPercent:
          Result := Result + '<number:text>%</number:text>';

        nftCurrSymbol:
          Result := Result + '<number:currency-symbol>' + Elements[el].TextValue + '</number:currency-symbol>';

        nftIntTh:
          begin
            Result := Result + '<number:number number:min-integer-digits="1" number:grouping="true"';
            if (el+2 < nel) and (Elements[el+1].Token = nftDecSep) and
               (Elements[el+2].Token in [nftZeroDecs, nftOptDecs, nftSpaceDecs]) then
            begin
              n := IfThen(Elements[el+2].Token = nftZeroDecs, Elements[el+2].IntValue, 1);
              Result := Result + ' number:decimal-places="' + IntToStr(n) + '"';
              inc(el, 2);
            end else
            if (el = nel-1) or (Elements[el+1].Token <> nftDecSep) then
              Result := Result + ' number:decimal-places="0"';
            Result := Result + ' />';
          end;

        nftFracNumZeroDigit, nftFracNumOptDigit, nftFracNumSpaceDigit:
          begin
            num := Elements[el].IntValue;
            inc(el);
            while (el < nel) and (Elements[el].Token in [nftSpace, nftText, nftEscaped]) do
              inc(el);
            if (el < nel) and (Elements[el].Token <> nftFracSymbol) then
              Continue;
            while (el < nel) and (Elements[el].Token in [nftSpace, nftText, nftEscaped]) do
              inc(el);
            if (el < nel) and
               (Elements[el].Token in [nftFracDenomOptDigit, nftFracDenomSpaceDigit, nftFracDenomZeroDigit, nftFracDenom])
            then
              denom := Elements[el].IntValue
            else
              Continue;
            if Elements[el].Token = nftFracDenom then    // fixed denominator
              Result := Result +
                '<number:fraction' +
                ' number:min-numerator-digits="' + IntToStr(num) +
                '" number:min-denominator-digits="' + IntToStr(num) +
                '" number:denominator-value="' + IntToStr(denom) +
                '" />'
            else
              Result := Result +
                '<number:fraction' +
                ' number:min-numerator-digits="' + IntToStr(num) +
                '" number:min-denominator-digits="' + IntToStr(num) +
                '" />'
          end;

        nftIntZeroDigit, nftIntOptDigit, nftIntSpaceDigit:
          begin
            // Mixed fraction
            if nfkFraction in Kind then
            begin
              int := Elements[el].IntValue;
              inc(el);
              while (el < nel) and not
                (Elements[el].Token in [nftFracNumZeroDigit, nftFracNumOptDigit, nftFracNumSpaceDigit])
              do
                inc(el);
              if el = nel then
                Continue;
              num := Elements[el].IntValue;
              while (el < nel) and not
                (Elements[el].Token in [nftFracDenomZeroDigit, nftFracDenomOptDigit, nftFracDenomSpaceDigit, nftFracDenom])
              do
                inc(el);
              if el = nel then
                Continue;
              denom := Elements[el].IntValue;
              if (Elements[el].Token = nftFracDenom) then
                Result := Result +
                  '<number:fraction' +
                  ' number:min-integer-digits="' + IntToStr(int) +
                  '" number:min-numerator-digits="' + IntToStr(num) +
                  '" number:min-denominator-digits="' + IntToStr(num) +
                  '" number:denominator-value="' + IntToStr(denom) +
                  '" />'
              else
                Result := Result +
                  '<number:fraction' +
                  ' number:min-integer-digits="' + IntToStr(int) +
                  '" number:min-numerator-digits="' + IntToStr(num) +
                  '" number:min-denominator-digits="' + IntToStr(denom) +
                  '" />';
            end
            else
            // Scientific, no decimals
            if (el+3 < nel) and (Elements[el+1].Token = nftExpChar) then
            begin
              Result := Result + '<number:scientific-number number:decimal-places="0"';
              n := IfThen(Elements[el].Token = nftIntZeroDigit, Elements[el].IntValue, 1);
              Result := Result + ' number:min-integer-digits="' + IntToStr(n) + '"';
              n := Elements[el+3].IntValue;
              Result := Result + ' number:min-exponent-digits="' + IntToStr(n) + '"';
              Result := Result + ' />';
              inc(el, 3);
            end
            else
            // Scientific, with decimals
            if (el+5 < nel) and (Elements[el+1].Token = nftDecSep) and (Elements[el+3].Token = nftExpChar)
            then begin
              Result := Result + '<number:scientific-number';
              n := IfThen(Elements[el].Token = nftIntZeroDigit, Elements[el].IntValue, 1);
              Result := Result + ' number:min-integer-digits="' + IntToStr(n) + '"';
              n := IfThen(Elements[el+2].Token = nftZeroDecs, Elements[el+2].IntValue, 1);
              Result := Result + ' number:decimal-places="' + IntToStr(n) + '"';
              Result := Result + ' number:min-exponent-digits="' + IntToStr(Elements[el+5].IntValue) + '"';
              Result := Result + ' />';
              inc(el, 5);
            end
            else
            // Standard decimal number
            if (el+2 < nel) and (Elements[el+1].Token = nftDecSep) then
            begin
              Result := Result + '<number:number';
              n := IfThen(Elements[el].Token = nftIntZeroDigit, Elements[el].IntValue, 1);
              Result := Result + ' number:min-integer-digits="' + IntToStr(n) + '"';
              n := IfThen(Elements[el+2].Token = nftZeroDecs, Elements[el+2].IntValue, 1);
              Result := Result + ' number:decimal-places="' + IntToStr(n) + '"';
              Result := Result + ' />';
              inc(el, 2);
            end
            else
            // Standard integer
            if (el = nel-1) or (Elements[el+1].Token <> nftDecSep) then
            begin
              Result := Result + '<number:number number:decimal-places="0"';
              n := IfThen(Elements[el].Token = nftIntZeroDigit, Elements[el].IntValue, 1);
              Result := Result + ' number:min-integer-digits="' + IntToStr(n) + '"';
              Result := Result + ' />';
            end;
          end;

        nftYear:
          if Elements[el].IntValue > 2 then
            Result := Result + '<number:year number:style="long" />'
          else
            Result := Result + '<number:year />';

        nftMonth:
          case Elements[el].IntValue of
            1: Result := Result + '<number:month />';
            2: Result := Result + '<number:month number:style="long" />';
            3: Result := Result + '<number:month number:textual="true" />';
            4: Result := Result + '<number:month number:style="long" number:textual="true" />';
          end;

        nftDay:
          case Elements[el].IntValue of
            1: Result := Result + '<number:day />';
            2: Result := Result + '<number:day number:style="long" />';
            3: Result := Result + '<number:day-of-week />';
            4: Result := Result + '<number:day-of-week number:style="long" />';
          end;

        nftHour:
          begin
            case abs(Elements[el].IntValue) of
              1: Result := Result + '<number:hours />';
              2: Result := Result + '<number:hours number:style="long" />';
            end;
            if Elements[el].IntValue < 0 then
              timeIntervalStr := ' number:truncate-on-overflow="false"';
          end;

        nftMinute:
          begin
            case abs(Elements[el].IntValue) of
              1: Result := Result + '<number:minutes />';
              2: Result := Result + '<number:minutes number:style="long" />';
            end;
            if Elements[el].IntValue < 0 then
              timeIntervalStr := ' number:truncate-on-overflow="false"';
          end;

        nftSecond:
          begin
            case abs(Elements[el].IntValue) of
              1: Result := Result + '<number:seconds />';
              2: Result := Result + '<number:seconds number:style="long" />';
            end;
            if Elements[el].IntValue < 0 then
              timeIntervalStr := ' number:truncate-on-overflow="false"';
          end;

        nftMilliseconds:
          ; // ???

        nftAMPM:
          Result := Result + '<number:am-pm />';

        nftDateTimeSep:
          case Elements[el].TextValue of
            '/': Result := Result + '<number:text>' + FWorkbook.FormatSettings.DateSeparator + '</number:text>';
            ':': Result := Result + '<number:text>' + FWorkbook.FormatSettings.TimeSeparator + '</number:text>';
            ' ': Result := Result + '<number:text><![CDATA[ ]]></number:text>';
            else Result := Result + '<number:text>' + Elements[el].TextValue + '</number:text>';
          end;
      end;

      inc(el);
    end;

    if (nfkPercent in Kind) then
      mask := '<number:percentage-style style:name="%s"%s>%s%s</number:percentage-style>'
    else
    if (nfkCurrency in Kind) then
      mask := '<number:currency-style style:name="%s"%s>%s%s</number:currency-style>'
    else
    if (nfkDate in Kind) then
      mask := '<number:date-style style:name="%s"%s>%s%s</number:date-style>'
    else
    if (Kind * [nfkDate, nfkTime] = [nfkTime]) then
      mask := '<number:time-style style:name="%s"%s>%s%s</number:time-style>'
    else
      mask := '<number:number-style style:name="%s"%s>%s%s</number:number-style>';

    Result := Format(mask, [AFormatName, TimeIntervalStr, Result, StyleMapStr]);
  end;
end;


{******************************************************************************}
{                         TsSpreadOpenDocReader                                }
{******************************************************************************}

constructor TsSpreadOpenDocReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);

  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';
  FPointSeparatorSettings.ListSeparator := ';';  // for formulas

  FCellFormatList := TsCellFormatList.Create(true);
    // true = allow duplicates because style names used in cell records will not be found any more.
  FColumnStyleList := TFPList.Create;
  FColumnList := TFPList.Create;
  FRowStyleList := TFPList.Create;
  FRowList := TFPList.Create;
  FPageLayoutList := TFPList.Create;
  FMasterPageList := TFPList.Create;
  FHeaderFooterFontList := TObjectList.Create;  // frees objects

  // Set up the default palette in order to have the default color names correct.
  Workbook.UseDefaultPalette;
  // Initial base date in case it won't be read from file
  FDateMode := dm1899;
end;

destructor TsSpreadOpenDocReader.Destroy;
var
  j: integer;
begin
  for j := FColumnList.Count-1 downto 0 do TObject(FColumnList[j]).Free;
  FColumnList.Free;

  for j := FColumnStyleList.Count-1 downto 0 do TObject(FColumnStyleList[j]).Free;
  FColumnStyleList.Free;

  for j := FRowList.Count-1 downto 0 do TObject(FRowList[j]).Free;
  FRowList.Free;

  for j := FRowStyleList.Count-1 downto 0 do TObject(FRowStyleList[j]).Free;
  FRowStyleList.Free;

  for j := FPageLayoutList.Count-1 downto 0 do TObject(FPageLayoutList[j]).Free;
  FPageLayoutList.Free;

  for j := FMasterPageList.Count-1 downto 0 do TObject(FMasterPageList[j]).Free;
  FMasterPageList.Free;

  FHeaderFooterFontList.Free;
  inherited Destroy;
end;

procedure TsSpreadOpenDocReader.AddBuiltinNumFormats;
begin
  FNumFormatList.Clear;
  FNumFormatList.Add('N0:');
end;

{ Creates for each non-default column width stored internally in FColumnList
  a TCol record in the current worksheet. }
procedure TsSpreadOpenDocReader.ApplyColWidths;
var
  colIndex: Integer;
  colWidth: Single;
  colStyleIndex: Integer;
  colStyle: TColumnStyleData;
  factor: Double;
  col: PCol;
  i: Integer;
begin
  factor := FWorkbook.GetFont(0).Size/2;
  for i:=0 to FColumnList.Count-1 do
  begin
    colIndex := TColumnData(FColumnList[i]).Col;
    colStyleIndex := TColumnData(FColumnList[i]).ColStyleIndex;
    colStyle := TColumnStyleData(FColumnStyleList[colStyleIndex]);
    { The column width stored in colStyle is in mm (see ReadColumnStyles).
      We convert it to character count by converting it to points and then by
      dividing the points by the approximate width of the '0' character which
      is assumed to be 50% of the default font point size. }
    colWidth := mmToPts(colStyle.ColWidth)/factor;
    { Add only column records to the worksheet if their width is different from
      the default column width. }
    if not SameValue(colWidth, FWorksheet.DefaultColWidth, COLWIDTH_EPS) then
    begin
      col := FWorksheet.GetCol(colIndex);
      col^.Width := colWidth;
    end;
  end;
end;

{ Applies the style data referred to by the style name to the specified cell
  The function result is false if a style with the given name could not be found }
function TsSpreadOpenDocReader.ApplyStyleToCell(ACell: PCell; AStyleName: String): Boolean;
var
  fmt: PsCellFormat;
  styleIndex: Integer;
  i: Integer;
begin
  Result := false;

  if FWorksheet.HasHyperlink(ACell) then
    FWorksheet.WriteFont(ACell, HYPERLINK_FONTINDEX);

  // Is there a style attached to the cell?
  styleIndex := -1;
  if AStyleName <> '' then
    styleIndex := FCellFormatList.FindIndexOfName(AStyleName);
  if (styleIndex = -1) then
  begin
    // No - look for the style attached to the column of the cell and
    // find the cell style by the DefaultCellStyleIndex stored in the column list.
    i := FindColumnByCol(ACell^.Col);
    if i = -1 then
      exit;
    styleIndex := TColumnData(FColumnList[i]).DefaultCellStyleIndex;
  end;
  fmt := FCellFormatList.Items[styleIndex];
  ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt^);

  Result := true;
end;

{ Extracts a boolean value from a "boolean" cell node.
  Is called from ReadBoolean }
function TsSpreadOpenDocReader.ExtractBoolFromNode(ANode: TDOMNode): Boolean;
var
  value: String;
begin
  value := GetAttrValue(ANode, 'office:boolean-value');
  if (lowercase(value) = 'true') then
    Result := true
  else
    Result := false;
end;

{ Extracts a date/time value from a "date-value" or "time-value" cell node.
  Requires the number format and format strings to optimize agreement with
  fpc date/time values.
  Is called from "ReadDateTime". }
function TsSpreadOpenDocReader.ExtractDateTimeFromNode(ANode: TDOMNode;
  ANumFormat: TsNumberFormat; const AFormatStr: String): TDateTime;
var
  Value: String;
  Fmt : TFormatSettings;
  FoundPos : integer;
  Hours, Minutes, Days: integer;
  Seconds: Double;
  HoursPos, MinutesPos, SecondsPos: integer;
begin
  Unused(AFormatStr);

  // Format expects ISO 8601 type date string or
  // time string
  fmt := DefaultFormatSettings;
  fmt.ShortDateFormat := 'yyyy-mm-dd';
  fmt.DateSeparator := '-';
  fmt.LongTimeFormat := 'hh:nn:ss';
  fmt.TimeSeparator := ':';

  Value := GetAttrValue(ANode, 'office:date-value');

  if Value <> '' then
  begin
    // Date or date/time string
    Value := StringReplace(Value,'T',' ',[rfIgnoreCase,rfReplaceAll]);
    // Strip milliseconds?
    FoundPos := Pos('.',Value);
    if (FoundPos > 1) then
       Value := Copy(Value, 1, FoundPos-1);
    Result := StrToDateTime(Value, Fmt);

    // If the date/time is within 1 day of the base date the value is most
    // probably a time-only value (< 1).
    // We need to subtract the datemode offset, otherwise the date/time value
    // would not be < 1 for fpc.
    case FDateMode of
      dm1899: if Result - DATEMODE_1899_BASE < 1 then Result := Result - DATEMODE_1899_BASE;
      dm1900: if Result - DATEMODE_1900_BASE < 1 then Result := Result - DATEMODE_1900_BASE;
      dm1904: if Result - DATEMODE_1904_BASE < 1 then Result := Result - DATEMODE_1904_BASE;
    end;

  end else begin
    // Try time only, e.g. PT23H59M59S
    //                     12345678901
    Value := GetAttrValue(ANode, 'office:time-value');
    if (Value <> '') and (Pos('PT', Value) = 1) then
    begin
      // Get hours
      HoursPos := Pos('H', Value);
      if (HoursPos > 0) then
        Hours := StrToInt(Copy(Value, 3, HoursPos-3))
      else
        Hours := 0;

      // Get minutes
      MinutesPos := Pos('M', Value);
      if (MinutesPos > 0) and (MinutesPos > HoursPos) then
        Minutes := StrToInt(Copy(Value, HoursPos+1, MinutesPos-HoursPos-1))
      else
        Minutes := 0;

      // Get seconds
      SecondsPos := Pos('S', Value);
      if (SecondsPos > 0) and (SecondsPos > MinutesPos) then
        Seconds := StrToFloat(Copy(Value, MinutesPos+1, SecondsPos-MinutesPos-1), FPointSeparatorSettings)
      else
        Seconds := 0;

      Days := Hours div 24;
      Hours := Hours mod 24;
      Result := Days + (Hours + (Minutes + Seconds/60)/60)/24;

      { Values < 1 day are certainly time-only formats --> no datemode correction
        nfTimeInterval formats are differences --> no date mode correction
        In all other case, we have a date part that needs to be corrected for
        the file's datemode. }
      if (ANumFormat <> nfTimeInterval) and (abs(Days) > 0) then
      begin
        case FDateMode of
          dm1899: Result := Result + DATEMODE_1899_BASE;
          dm1900: Result := Result + DATEMODE_1900_BASE;
          dm1904: Result := Result + DATEMODE_1904_BASE;
        end;
      end;
    end;
  end;
end;

function TsSpreadOpenDocReader.FindColumnByCol(AColIndex: Integer): Integer;
begin
  for Result := 0 to FColumnList.Count-1 do
    if TColumnData(FColumnList[Result]).Col = AColIndex then
      exit;
  Result := -1;
end;

function TsSpreadOpenDocReader.FindColStyleByName(AStyleName: String): Integer;
begin
  for Result := 0 to FColumnStyleList.Count-1 do
    if TColumnStyleData(FColumnStyleList[Result]).Name = AStyleName then
      exit;
  Result := -1;
end;

function TsSpreadOpenDocReader.FindNumFormatByName(ANumFmtName: String): Integer;
begin
  for Result := 0 to FNumFormatList.Count-1 do
    if pos(ANumFmtName+':', FNumFormatList[Result]) = 1 then
      exit;
  Result := -1;
end;

function TsSpreadOpenDocReader.FindRowStyleByName(AStyleName: String): Integer;
begin
  for Result := 0 to FRowStyleList.Count-1 do
    if TRowStyleData(FRowStyleList[Result]).Name = AStyleName then
      exit;
  Result := -1;
end;

procedure TsSpreadOpenDocReader.ReadAutomaticStyles(AStylesNode: TDOMNode);
var
  nodeName: String;
  layoutNode, fontNode: TDOMNode;
  node, child: TDOMNode;
  s: String;
  data: TPageLayoutData;
  isHeader: Boolean;
  h: Double;
  fnt: TXMLHeaderFooterFont;
  defFnt: TsFont;
  fntName: String;
  fntSize: Double;
  fntStyle: TsHeaderFooterFontStyles;
  fntColor: TsColorValue;
begin
  if not Assigned(AStylesNode) then
    exit;
  defFnt := Workbook.GetDefaultFont;
  layoutNode := AStylesNode.FirstChild;
  while layoutNode <> nil do
  begin
    nodeName := layoutNode.NodeName;
    if nodeName = 'style:style' then
    begin
      // Read fonts used by page layout's header/footer
      fntName := defFnt.FontName;
      fntSize := defFnt.Size;
      fntColor := defFnt.Color;
      fntStyle := [];
      s := GetAttrValue(layoutNode, 'style:family');
      if s = 'text' then
      begin
        s := GetAttrValue(layoutNode, 'style:name');
        fontNode := layoutNode.FirstChild;
        ReadHeaderFooterFont(fontNode, fntName, fntSize, fntStyle, fntColor);
        fnt := TXMLHeaderFooterFont.Create(fntName, fntSize, fntStyle, fntColor);
        fnt.StyleName := s;
        FHeaderFooterFontList.Add(fnt);
      end;
    end
    else
    if nodeName = 'style:page-layout' then
    begin
      // Read page layout parameters
      data := TPageLayoutData.Create;
      InitPageLayout(data.PageLayout);
      data.Name := GetAttrValue(layoutNode, 'style:name');

      node := layoutNode.FirstChild;
      while node <> nil do
      begin
        nodeName := node.NodeName;
        if nodeName = 'style:page-layout-properties' then
        begin
          s := GetAttrValue(node, 'style:print-orientation');
          if s = 'landscape' then
            data.PageLayout.Orientation := spoLandscape
          else if s = 'portrait' then
            data.PageLayout.Orientation := spoPortrait;

          s := GetAttrValue(node, 'fo:page-width');
          if s <> '' then
            data.PageLayout.PageWidth := PtsToMM(HTMLLengthStrToPts(s));

          s := GetAttrValue(node, 'fo:page-height');
          if s <> '' then
            data.PageLayout.PageHeight := PtsToMM(HTMLLengthStrToPts(s));

          s := GetAttrValue(node, 'fo:margin-top');
          if s <> '' then
            data.PageLayout.TopMargin := PtsToMM(HTMLLengthStrToPts(s));
            // But: if there is a header this value is the headermargin!

          s := GetAttrValue(node, 'fo:margin-bottom');
          if s <> '' then
            data.PageLayout.BottomMargin := PtsToMM(HTMLLengthStrToPts(s));
            // But: if there is a footer this value is the footermargin!

          s := GetAttrValue(node, 'fo:margin-left');
          if s <> '' then
            data.PageLayout.LeftMargin := PtsToMM(HTMLLengthStrToPts(s));

          s := GetAttrValue(node, 'fo:margin-right');
          if s <> '' then
            data.PageLayout.RightMargin := PtsToMM(HTMLLengthStrToPts(s));

          s := GetAttrValue(node, 'style:scale-to');
          if (s <> '') then
          begin
            if s[Length(s)] = '%' then Delete(s, Length(s), 1);
            data.PageLayout.ScalingFactor := round(StrToFloat(s, FPointSeparatorSettings));
            Exclude(data.PageLayout.Options, poFitPages);
          end;

          s := GetAttrValue(node, 'style:scale-to-X');
          if s <> '' then
          begin
            data.PageLayout.FitWidthToPages := StrToInt(s);
            Include(data.PageLayout.Options, poFitPages);
          end;

          s := GetAttrValue(node, 'style:scale-to-Y');
          if s <> '' then
          begin
            data.PageLayout.FitHeightToPages := StrToInt(s);
            Include(data.PageLayout.Options, poFitPages);
          end;

          s := GetAttrValue(node, 'style:table-centering');
          case s of
            'both':
              data.PageLayout.Options := data.PageLayout.Options + [poHorCentered, poVertCentered];
            'horizontal':
              data.PageLayout.Options := data.PageLayout.Options + [poHorCentered] - [poVertCentered];
            'vertical':
              data.PageLayout.Options := data.PageLayout.Options - [poHorCentered] + [poVertCentered];
          end;

          s := GetAttrValue(node, 'style:print');
          if pos('grid', s) > 0 then
            Include(data.PageLayout.Options, poPrintGridLines);
          if pos('headers', s) > 0 then
            Include(data.PageLayout.Options, poPrintHeaders);
          if pos('annotations', s) > 0 then
            Include(data.PageLayout.Options, poPrintCellComments);

          s := GetAttrValue(node, 'style:print-page-order');
          if s = 'ltr' then    // "left-to-right", the other option is "ttb = top-to-bottom"
            Include(data.PageLayout.Options, poPrintPagesByRows);

          s := GetAttrValue(node, 'style:first-page-number');
          if s = 'continue' then
            Exclude(data.PageLayout.Options, poUseStartPageNumber)
          else
          if TryStrToInt(s, data.PageLayout.StartPageNumber) then
            Include(data.PageLayout.Options, poUseStartPageNumber);

          FPageLayoutList.Add(data);
        end else
        if (nodeName = 'style:header-style') or (nodeName = 'style:footer-style')
        then
        begin
          isHeader := nodeName = 'style:header-style';
          child := node.FirstChild;
          while child <> nil do
          begin
            nodeName := child.NodeName;
            if nodeName = 'style:header-footer-properties' then
            begin
              s := GetAttrValue(child, 'fo:min-height');
              if s <> '' then
                h := PtsToMM(HTMLLengthStrToPts(s)) else h := 0;
              if isHeader then
              begin
                data.PageLayout.HeaderMargin := data.PageLayout.TopMargin;
                data.PageLayout.TopMargin := data.PageLayout.HeaderMargin + h;
              end else
              begin
                data.PageLayout.FooterMargin := data.PageLayout.BottomMargin;
                data.PageLayout.BottomMargin := data.PageLayout.FooterMargin + h;
              end;
            end;
            child := child.NextSibling;
          end;
        end;
        node := node.NextSibling;
      end;
    end;
    layoutNode := layoutNode.NextSibling;
  end;
end;

function TsSpreadOpenDocReader.ReadHeaderFooterText(ANode: TDOMNode): String;
var
  parser: TsSpreadOpenDocHeaderFooterParser;
  defFnt: TsHeaderFooterFont;
begin
  defFnt := TsHeaderFooterFont.Create(Workbook.GetDefaultFont);
  parser := TsSpreadOpenDocHeaderFooterParser.Create(ANode.FirstChild,
    FHeaderFooterFontList, defFnt);
  try
    Result := parser.BuildHeaderFooter;
  finally
    parser.Free;
    defFnt.Free;
  end;
end;

{ Reads the master styles nodes which contain the header/footer texts }
procedure TsSpreadOpenDocReader.ReadMasterStyles(AStylesNode: TDOMNode);
var
  masternode, stylenode: TDOMNode;
  nodeName: String;
  s: String;
  data: TMasterPageData;
  pagelayout: PsPageLayout;
  j: Integer;

begin
  if AStylesNode = nil then
    exit;

  masterNode := AStylesNode.FirstChild;
  while (masterNode <> nil) do
  begin
    nodeName := masterNode.NodeName;
    if nodeName = 'style:master-page' then begin
      s := GetAttrvalue(masterNode, 'style:page-layout-name');

      { Find the page layout data belonging to the current node }
      pageLayout := nil;
      for j:=0 to FPageLayoutList.Count-1 do
        if TPageLayoutData(FPageLayoutList[j]).Name = s then
        begin
          pageLayout := @TPageLayoutData(FPageLayoutList[j]).PageLayout;
          break;
        end;
      if pagelayout = nil then
        exit;

      data := TMasterPageData.create;
      data.Name := GetAttrValue(masternode, 'style:name');
      data.PageLayoutName := s;
      FMasterPageList.Add(data);

      styleNode := masterNode.FirstChild;
      while styleNode <> nil do begin
        nodeName := styleNode.NodeName;
        if nodeName = 'style:header' then
        begin
          s := ReadHeaderFooterText(styleNode);
          if s <> '' then
            pageLayout^.Headers[HEADER_FOOTER_INDEX_ODD] := s;
        end else
        if nodeName = 'style:header-left' then
        begin
          s := ReadHeaderFooterText(styleNode);
          if s <> '' then
          begin
            pageLayout^.Headers[HEADER_FOOTER_INDEX_EVEN] := s;
            Include(pageLayout^.Options, poDifferentOddEven);
          end;
          s := GetAttrValue(styleNode, 'style:display');
          if s = 'false' then
            Exclude(pagelayout^.Options, poDifferentOddEven);
        end else
        if nodeName = 'style:footer' then
        begin
          s := ReadHeaderFooterText(styleNode);
          if s <> '' then
            pageLayout^.Footers[HEADER_FOOTER_INDEX_ODD] := s;
        end else
        if nodeName = 'style:footer-left' then
        begin
          s := ReadHeaderFooterText(styleNode);
          if s <> '' then
          begin
            pageLayout^.Footers[HEADER_FOOTER_INDEX_EVEN] := s;
            Include(pageLayout^.Options, poDifferentOddEven);
          end;
          s := GetAttrValue(styleNode, 'style:display');
          if s = 'false' then
            Exclude(pagelayout^.Options, poDifferentOddEven);
        end;
        styleNode := styleNode.NextSibling;
      end;
    end;
    masterNode := masterNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocReader.ReadBlank(ARow, ACol: Cardinal;
  ACellNode: TDOMNode);
var
  styleName: String;
  cell: PCell;
  lCell: TCell;
begin
  // a temporary cell record to store the formatting if there is any
  lCell.Row := ARow;  // to silence a compiler warning...
  InitCell(ARow, ACol, lCell);
  lCell.ContentType := cctEmpty;

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  if not ApplyStyleToCell(@lCell, stylename) then
    exit;
    // No need to store a record for an empty, unformatted cell

  if FIsVirtualMode then
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);
  FWorkSheet.WriteBlank(cell);
  FWorksheet.CopyFormat(@lCell, cell);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadOpenDocReader.ReadBoolean(ARow, ACol: Cardinal;
  ACellNode: TDOMNode);
var
  styleName: String;
  cell: PCell;
  boolValue: Boolean;
begin
  if FIsVirtualMode then
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  boolValue := ExtractBoolFromNode(ACellNode);
  FWorkSheet.WriteBoolValue(cell, boolValue);

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(cell, stylename);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{ Collection columns used in the given table. The columns contain links to
  styles that must be used when cells in that columns are without styles. }
procedure TsSpreadOpenDocReader.ReadColumns(ATableNode: TDOMNode);
var
  col: Integer;
  colNode: TDOMNode;
  s: String;
  defCellStyleIndex: Integer;
  colStyleIndex: Integer;
  colData: TColumnData;
  colsRepeated: Integer;
  j: Integer;
begin
  // clear previous column list (from other sheets)
  for j := FColumnList.Count-1 downto 0 do TObject(FColumnList[j]).Free;
  FColumnList.Clear;

  col := 0;
  colNode := ATableNode.FindNode('table:table-column');
  while Assigned(colNode) do
  begin
    if colNode.NodeName = 'table:table-column' then
    begin;
      s := GetAttrValue(colNode, 'table:style-name');
      colStyleIndex := FindColStyleByName(s);
      if colStyleIndex <> -1 then
      begin
        defCellStyleIndex := -1;
        s := GetAttrValue(ColNode, 'table:default-cell-style-name');
        if s <> '' then
        begin
          defCellStyleIndex := FCellFormatList.FindIndexOfName(s); //FindCellStyleByName(s);
          colData := TColumnData.Create;
          colData.Col := col;
          colData.ColStyleIndex := colStyleIndex;
          colData.DefaultCellStyleIndex := defCellStyleIndex;
          FColumnList.Add(colData);
        end;
        s := GetAttrValue(ColNode, 'table:number-columns-repeated');
        if s = '' then
          inc(col)
        else
        begin
          colsRepeated := StrToInt(s);
          if defCellStyleIndex > -1 then
            for j:=1 to colsRepeated-1 do
            begin
              colData := TColumnData.Create;
              colData.Col := col + j;
              colData.ColStyleIndex := colStyleIndex;
              colData.DefaultCellStyleIndex := defCellStyleIndex;
              FColumnList.Add(colData);
            end;
          inc(col, colsRepeated);
        end;
      end;
    end;
    colNode := colNode.NextSibling;
  end;
end;

{ Reads the column styles and stores them in the FColumnStyleList for later use }
procedure TsSpreadOpenDocReader.ReadColumnStyle(AStyleNode: TDOMNode);
var
  colStyle: TColumnStyleData;
  styleName: String;
  styleChildNode: TDOMNode;
  colWidth: double;
  s: String;
begin
  styleName := GetAttrValue(AStyleNode, 'style:name');
  styleChildNode := AStyleNode.FirstChild;
  colWidth := -1;

  while Assigned(styleChildNode) do
  begin
    if styleChildNode.NodeName = 'style:table-column-properties' then
    begin
      s := GetAttrValue(styleChildNode, 'style:column-width');
      if s <> '' then
      begin
        colWidth := PtsToMM(HTMLLengthStrToPts(s));   // convert to mm
        break;
      end;
    end;
    styleChildNode := styleChildNode.NextSibling;
  end;

  colStyle := TColumnStyleData.Create;
  colStyle.Name := styleName;
  colStyle.ColWidth := colWidth;
  FColumnStyleList.Add(colStyle);
end;

procedure TsSpreadOpenDocReader.ReadComment(ARow, ACol: Cardinal;
  ACellNode: TDOMNode);
var
  cellChildNode, pNode, pChildNode: TDOMNode;
  comment, line: String;
  nodeName: String;
  s: String;
  found: Boolean;
begin
  if ACellNode = nil then
    exit;

  comment := '';
  found := false;

  cellChildNode := ACellNode.FirstChild;
  while cellChildNode <> nil do begin
    nodeName := cellChildNode.NodeName;
    if nodeName = 'office:annotation' then begin
      pNode := cellChildNode.FirstChild;
      while pNode <> nil do begin
        nodeName := pNode.NodeName;
        if nodeName = 'text:p' then
        begin
          line := '';
          pChildNode := pNode.FirstChild;
          while pChildNode <> nil do
          begin
            nodeName := pChildNode.NodeName;
            if nodeName = '#text' then
            begin
              s := pChildNode.NodeValue;
              line := IfThen(line = '', s, line + s);
              found := true;
            end else
            if nodeName = 'text:span' then
            begin
              s := GetNodeValue(pChildNode);
              line := IfThen(line = '', s, line + s);
              found := true;
            end;
            pChildNode := pChildNode.NextSibling;
          end;
          comment := IfThen(comment = '', line, comment + LineEnding + line);
        end;
        pNode := pNode.NextSibling;
      end;
    end;
    cellChildNode := cellChildNode.NextSibling;
  end;
  if found then
    FWorksheet.WriteComment(ARow, ACol, comment);
end;

procedure TsSpreadOpenDocReader.ReadDateTime(ARow, ACol: Cardinal;
  ACellNode : TDOMNode);
var
  dt: TDateTime;
  styleName: String;
  cell: PCell;
  fmt: PsCellFormat;
begin
  if FIsVirtualMode then
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(cell, stylename);
  fmt := FWorkbook.GetPointerToCellFormat(cell^.FormatIndex);;

  dt := ExtractDateTimeFromNode(ACellNode, fmt^.NumberFormat, fmt^.NumberFormatStr);
  FWorkSheet.WriteDateTime(cell, dt, fmt^.NumberFormat, fmt^.NumberFormatStr);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadOpenDocReader.ReadDateMode(SpreadSheetNode: TDOMNode);
var
  CalcSettingsNode, NullDateNode: TDOMNode;
  NullDateSetting: string;
begin
  // Default datemode for ODF:
  NullDateSetting := '1899-12-30';
  CalcSettingsNode := SpreadsheetNode.FindNode('table:calculation-settings');
  if Assigned(CalcSettingsNode) then
  begin
    NullDateNode := CalcSettingsNode.FindNode('table:null-date');
    if Assigned(NullDateNode) then
      NullDateSetting := GetAttrValue(NullDateNode,'table:date-value');
  end;
  if NullDateSetting = '1899-12-30' then
    FDateMode := dm1899
  else if NullDateSetting = '1900-01-01' then
    FDateMode := dm1900
  else if NullDateSetting = '1904-01-01' then
    FDateMode := dm1904
  else
    raise Exception.CreateFmt('Spreadsheet file corrupt: cannot handle null-date format %s', [NullDateSetting]);
end;

{ Reads font data from an xml node, adds the font to the workbooks FontList
  (if not yet contained), and returns the index in the font list.
  If the font is a special font (such as DefaultFont, or HyperlinkFont) then
  APreferredIndex defines the index under which the font should be stored in the
  list. }
function TsSpreadOpenDocReader.ReadFont(ANode: TDOMnode;
  APreferredIndex: Integer = -1): Integer;
var
  fntName: String;
  fntSize: Single;
  fntStyles: TsFontStyles;
  fntColor: TsColor;
  s: String;
begin
  if ANode = nil then
  begin
    Result := 0;
    exit;
  end;

  fntName := GetAttrValue(ANode, 'style:font-name');
  if fntName = '' then
    fntName := FWorkbook.GetFont(0).FontName;

  s := GetAttrValue(ANode, 'fo:font-size');
  if s <> '' then
    fntSize := HTMLLengthStrToPts(s)
  else
    fntSize := FWorkbook.GetDefaultFontSize;

  fntStyles := [];
  if GetAttrValue(ANode, 'fo:font-style') = 'italic' then
    Include(fntStyles, fssItalic);
  if GetAttrValue(ANode, 'fo:font-weight') = 'bold' then
    Include(fntStyles, fssBold);
  s := GetAttrValue(ANode, 'style:text-underline-style');
  if not ((s = '') or (s = 'none')) then
    Include(fntStyles, fssUnderline);
  s := GetAttrValue(ANode, 'style:text-line-through-style');
  if not ((s = '') or (s = 'none')) then
    Include(fntStyles, fssStrikeout);

  s := GetAttrValue(ANode, 'fo:color');
  if s <> '' then
    fntColor := FWorkbook.AddColorToPalette(HTMLColorStrToColor(s))
  else
    fntColor := FWorkbook.GetDefaultFont.Color;

  if APreferredIndex = 0 then
  begin
    FWorkbook.SetDefaultFont(fntName, fntSize);
    Result := 0;
  end else
  if (APreferredIndex > -1) then
  begin
    {  --- wp: No more missing font #4 now !!!
    if (APreferredIndex = 4) then
      raise Exception.Create('Cannot replace font #4');
      }
    FWorkbook.ReplaceFont(APreferredIndex, fntName, fntSize, fntStyles, fntColor);
    Result := APreferredIndex;
  end else
  begin
    Result := FWorkbook.FindFont(fntName, fntSize, fntStyles, fntColor);
    if Result = -1 then
      Result := FWorkbook.AddFont(fntName, fntSize, fntStyles, fntColor);
  end;
end;

procedure TsSpreadOpenDocReader.ReadFormula(ARow, ACol: Cardinal;
  ACellNode : TDOMNode);
var
  cell: PCell;
  formula: String;
  stylename: String;
  floatValue: Double;
  boolValue: Boolean;
  valueType: String;
  valueStr: String;
  node: TDOMNode;
  parser: TsSpreadsheetParser;
  p: Integer;
  fmt: PsCellFormat;
begin
  // Create cell and apply format
  if FIsVirtualMode then
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(cell, stylename);
  fmt := Workbook.GetPointerToCellFormat(cell^.FormatIndex);

  formula := '';
  if (boReadFormulas in FWorkbook.Options) then
  begin
    // Read formula, trim it, ...
    formula := GetAttrValue(ACellNode, 'table:formula');
    if formula <> '' then
    begin
      // formulas written by Spread begin with 'of:=', our's with '=' --> remove that
      p := pos('=', formula);
      Delete(formula, 1, p);
    end;
    // ... convert to Excel dialect used by fps by defailt
    parser := TsSpreadsheetParser.Create(FWorksheet);
    try
      parser.Dialect := fdOpenDocument;
      parser.LocalizedExpression[FPointSeparatorSettings] := formula;
      parser.Dialect := fdExcel;
      formula := parser.Expression;
    finally
      parser.Free;
    end;
    // ... and store in cell's FormulaValue field.
    cell^.FormulaValue := formula;
  end;

  // Read formula results
  valueType := GetAttrValue(ACellNode, 'office:value-type');
  valueStr := GetAttrValue(ACellNode, 'office:value');
  // ODS wants a 0 in the NumberValue field in case of an error. If there is
  // no error, this value will be corrected below.
  cell^.NumberValue := 0.0;
  // (a) number value
  if (valueType = 'float') then
  begin
    if UpperCase(valueStr) = '1.#INF' then
      FWorksheet.WriteNumber(cell, 1.0/0.0)
    else
    begin
      floatValue := StrToFloat(valueStr, FPointSeparatorSettings);
      FWorksheet.WriteNumber(cell, floatValue);
    end;
    if IsDateTimeFormat(fmt^.NumberFormat) then
    begin
      cell^.ContentType := cctDateTime;
      // No datemode correction for intervals and for time-only values
      if (fmt^.NumberFormat = nfTimeInterval) or (cell^.NumberValue < 1) then
        cell^.DateTimeValue := cell^.NumberValue
      else
        case FDateMode of
          dm1899: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1899_BASE;
          dm1900: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1900_BASE;
          dm1904: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1904_BASE;
        end;
    end;
  end else
  // (b) Date/time value
  if (valueType = 'date') or (valueType = 'time') then
  begin
    floatValue := ExtractDateTimeFromNode(ACellNode, fmt^.NumberFormat, fmt^.NumberFormatStr);
    FWorkSheet.WriteDateTime(cell, floatValue);
  end else
  // (c) text
  if (valueType = 'string') then
  begin
    node := ACellNode.FindNode('text:p');
    if (node <> nil) and (node.FirstChild <> nil) then
    begin
      valueStr := node.FirstChild.Nodevalue;
      FWorksheet.WriteUTF8Text(cell, valueStr);
    end;
  end else
  // (d) boolean
  if (valuetype = 'boolean') then
  begin
    boolValue := ExtractBoolFromNode(ACellNode);
    FWorksheet.WriteBoolValue(cell, boolValue);
  end else
  // (e) Text
  if (valueStr <> '') then
    FWorksheet.WriteUTF8Text(cell, valueStr);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadOpenDocReader.ReadFromFile(AFileName: string);
var
  Doc : TXMLDocument;
  FilePath : string;
  UnZip : TUnZipper;
  FileList : TStringList;
  BodyNode, SpreadSheetNode, TableNode: TDOMNode;
  StylesNode: TDOMNode;
  OfficeSettingsNode: TDOMNode;
  nodename: String;
  pageLayout: PsPageLayout;
begin
  //unzip files into AFileName path
  FilePath := GetTempDir(false);
  UnZip := TUnZipper.Create;
  FileList := TStringList.Create;
  try
    FileList.Add('styles.xml');
    FileList.Add('content.xml');
    FileList.Add('settings.xml');
    UnZip.OutputPath := FilePath;
    Unzip.UnZipFiles(AFileName,FileList);
  finally
    FreeAndNil(FileList);
    FreeAndNil(UnZip);
  end; //try

  Doc := nil;
  try
    // process the styles.xml file
    ReadXMLFile(Doc, FilePath+'styles.xml');
    DeleteFile(FilePath+'styles.xml');

    StylesNode := Doc.DocumentElement.FindNode('office:styles');
    ReadNumFormats(StylesNode);
    ReadStyles(StylesNode);

    StylesNode := Doc.DocumentElement.FindNode('office:automatic-styles');
    ReadAutomaticStyles(StylesNode);

    StylesNode := Doc.DocumentElement.FindNode('office:master-styles');
    ReadMasterStyles(StylesNode);

    Doc.Free;

    //process the content.xml file
    ReadXMLFile(Doc, FilePath+'content.xml');
    DeleteFile(FilePath+'content.xml');

    StylesNode := Doc.DocumentElement.FindNode('office:automatic-styles');
    ReadNumFormats(StylesNode);
    ReadStyles(StylesNode);

    BodyNode := Doc.DocumentElement.FindNode('office:body');
    if not Assigned(BodyNode) then Exit;

    SpreadSheetNode := BodyNode.FindNode('office:spreadsheet');
    if not Assigned(SpreadSheetNode) then Exit;

    ReadDateMode(SpreadSheetNode);

    //process each table (sheet)
    TableNode := SpreadSheetNode.FindNode('table:table');
    while Assigned(TableNode) do
    begin
      nodename := TableNode.Nodename;
      // These nodes occur due to leading spaces which are not skipped
      // automatically any more due to PreserveWhiteSpace option applied
      // to ReadXMLFile
      if nodeName <> 'table:table' then
      begin
        TableNode := TableNode.NextSibling;
        continue;
      end;
      FWorkSheet := FWorkbook.AddWorksheet(GetAttrValue(TableNode, 'table:name'), true);
      // Collect column styles used
      ReadColumns(TableNode);
      // Process each row inside the sheet and process each cell of the row
      ReadRowsAndCells(TableNode);
      // Read page layout
      pageLayout := ReadPageLayout(StylesNode, GetAttrValue(TableNode, 'table:style-name'));
      if pageLayout <> nil then
        FWorksheet.PageLayout := pagelayout^;
      // Handle columns and rows
      ApplyColWidths;
      // Page layout
      FixCols(FWorksheet);
      FixRows(FWorksheet);
      // Continue with next table
      TableNode := TableNode.NextSibling;
    end; //while Assigned(TableNode)

    Doc.Free;

    // process the settings.xml file (Note: it does not always exist!)
    if FileExists(FilePath + 'settings.xml') then
    begin
      ReadXMLFile(Doc, FilePath+'settings.xml');
      DeleteFile(FilePath+'settings.xml');

      OfficeSettingsNode := Doc.DocumentElement.FindNode('office:settings');
      ReadSettings(OfficeSettingsNode);
    end;

  finally
    if Assigned(Doc) then Doc.Free;
  end;
end;

procedure TsSpreadOpenDocReader.ReadFromStream(AStream: TStream);
begin
  Unused(AStream);
  raise Exception.Create('[TsSpreadOpenDocReader.ReadFromStream] '+
                         'Method not implemented. Use "ReadFromFile" instead.');
end;

procedure TsSpreadOpenDocReader.ReadHeaderFooterFont(ANode: TDOMNode;
  var AFontName: String; var AFontSize: Double;
  var AFontStyle: TsHeaderFooterFontStyles; var AFontColor: TsColorValue);
var
  s: String;
begin
  if ANode = nil then
    exit;

  AFontName := GetAttrValue(ANode, 'style:font-name');

  s := GetAttrValue(ANode, 'fo:font-size');
  if s <> '' then
    AFontSize := HTMLLengthStrToPts(s);

  AFontStyle := [];

  if GetAttrValue(ANode, 'fo:font-style') = 'italic' then
    Include(AFontStyle, hfsItalic);

  if GetAttrValue(ANode, 'fo:font-weight') = 'bold' then
    Include(AFontStyle, hfsBold);

  s := GetAttrValue(ANode, 'style:text-underline-style');
  if not ((s = '') or (s = 'none')) then
  begin
    if GetAttrValue(ANode, 'style:text-underline-type') = 'double' then
      Include(AFontStyle, hfsDblUnderline)
    else
      Include(AFontStyle, hfsUnderline);
  end;

  s := GetAttrValue(ANode, 'style:text-line-through-style');
  if not ((s = '') or (s = 'none')) then
    Include(AFontStyle, hfsStrikeout);

  if GetAttrValue(ANode, 'style:text-outline') = 'true' then
    Include(AFontStyle, hfsOutline);

  s := GetAttrValue(ANode, 'fo:text-shadow');
  if not ((s = '') or (s = 'none')) then
    Include(AFontStyle, hfsShadow);

  s := GetAttrValue(ANode, 'style:text-position');
  if pos('sub', s) = 1 then
    Include(AFontStyle, hfsSubscript)
  else if pos('super', s) = 1 then
    Include(AFontStyle, hfsSuperscript);

  s := GetAttrValue(ANode, 'fo:color');
  if s <> '' then
    AFontColor := HTMLColorStrToColor(s);
end;

procedure TsSpreadOpenDocReader.ReadLabel(ARow, ACol: Cardinal;
  ACellNode: TDOMNode);
var
  cellText: String;
  styleName: String;
  childnode: TDOMNode;
  subnode: TDOMNode;
  nodeName: String;
  cell: PCell;
  hyperlink: string;

  procedure AddToCellText(AText: String);
  begin
    if cellText = ''
       then cellText := AText
       else cellText := cellText + AText;
  end;

begin
  { We were forced to activate PreserveWhiteSpace in the DOMParser in order to
    catch the spaces inserted in formatting texts. However, this adds lots of
    garbage into the cellText if is is read by means of above statement. Done
    like below is much better: }
  cellText := '';
  hyperlink := '';
  childnode := ACellNode.FirstChild;
  while Assigned(childnode) do
  begin
    nodeName := childNode.NodeName;
    if nodeName = 'text:p' then begin
      // Each 'text:p' node is a paragraph --> we insert a line break after the first paragraph
      if cellText <> '' then
        cellText := cellText + LineEnding;
      subnode := childnode.FirstChild;
      while Assigned(subnode) do
      begin
        nodename := subnode.NodeName;
        case nodename of
          '#text' :
            AddToCellText(subnode.TextContent);
          'text:a':     // "hyperlink anchor"
            begin
              hyperlink := GetAttrValue(subnode, 'xlink:href');
              AddToCellText(subnode.TextContent);
            end;
          'text:span':
            AddToCellText(subnode.TextContent);
        end;
        subnode := subnode.NextSibling;
      end;
    end;
    childnode := childnode.NextSibling;
  end;

  if FIsVirtualMode then
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  FWorkSheet.WriteUTF8Text(cell, cellText);
  if hyperlink <> '' then
  begin
    // ODS sees relative paths relative to the internal own file structure
    // --> we must remove 1 level-up to be at the same level where fps expects
    // the file.
    if pos('../', hyperlink) = 1 then
      Delete(hyperlink, 1, Length('../'));
    FWorksheet.WriteHyperlink(cell, hyperlink);
    FWorksheet.WriteFont(cell, HYPERLINK_FONTINDEX);
  end;

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(cell, stylename);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadOpenDocReader.ReadNumber(ARow, ACol: Cardinal;
  ACellNode : TDOMNode);
var
  Value, Str: String;
  lNumber: Double;
  styleName: String;
  cell: PCell;
  fmt: PsCellFormat;
  numFmt: TsNumFormatParams;
begin
  if FIsVirtualMode then
  begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  Value := GetAttrValue(ACellNode,'office:value');
  if UpperCase(Value)='1.#INF' then
    FWorkSheet.WriteNumber(cell, 1.0/0.0)
  else
  begin
    // Don't merge, or else we can't debug
    Str := GetAttrValue(ACellNode,'office:value');
    lNumber := StrToFloat(Str, FPointSeparatorSettings);
    FWorkSheet.WriteNumber(cell, lNumber);
  end;

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(cell, stylename);
  fmt := Workbook.GetPointerToCellFormat(cell^.FormatIndex);
  numFmt := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);

  // Sometimes date/time cells are stored as "float".
  // We convert them to date/time and also correct the date origin offset if
  // needed.
  if IsDateTimeFormat(numFmt) then
  begin
    cell^.ContentType := cctDateTime;
    // No datemode correction for intervals and for time-only values
    if (numFmt.NumFormat = nfTimeInterval) or (cell^.NumberValue < 1) then
      cell^.DateTimeValue := cell^.NumberValue
    else
      case FDateMode of
        dm1899: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1899_BASE;
        dm1900: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1900_BASE;
        dm1904: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1904_BASE;
      end;
  end;

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadOpenDocReader.ReadNumFormats(AStylesNode: TDOMNode);

  procedure ReadStyleMap(ANode: TDOMNode; var ANumFormat: TsNumberFormat;
    var AFormatStr: String);
  var
    condition: String;
    stylename: String;
    styleindex: Integer;
    fmt: String;
    posfmt, negfmt, zerofmt, currfmt: String;
    nf: TsNumberFormat;
    parser: TsNumFormatParser;
    counter: Integer;
  begin
    posfmt := '';
    negfmt := '';
    zerofmt := '';
    currfmt := AFormatStr;
    counter := 0;

    AFormatStr := '';
    ANumFormat := nfCustom;

    while ANode <> nil do
    begin
      condition := ANode.NodeName;

      if (ANode.NodeName = '#text') or not ANode.HasAttributes then
      begin
        ANode := ANode.NextSibling;
        Continue;
      end;

      condition := GetAttrValue(ANode, 'style:condition');
      stylename := GetAttrValue(ANode, 'style:apply-style-name');
      if (condition = '') or (stylename = '') then
      begin
        ANode := ANode.NextSibling;
        continue;
      end;

      Delete(condition, 1, Length('value()'));
      styleindex := -1;
      styleIndex := FindNumFormatByName(stylename);
      if (styleindex = -1) or (condition = '') then
      begin
        ANode := ANode.NextSibling;
        continue;
      end;

      fmt := NumFormatList[styleIndex];
      fmt := Copy(fmt, pos(':', fmt)+1, Length(fmt));
      parser := TsNumFormatParser.Create(Workbook, fmt);
      try
        nf := parser.NumFormat;
        if (nf = nfCurrency) and (parser.ParsedSections[0].Color = scRed) then
          nf := nfCurrencyRed;
        if nf in [nfCurrency, nfCurrencyRed] then
          ANumFormat := nf;
      finally
        parser.Free;
      end;

      case condition[1] of
        '>': begin
               posfmt := fmt;
               if (Length(condition) > 1) and (condition[2] = '=') then
                 zerofmt := fmt;
             end;
        '<': begin
               negfmt := fmt;
               if (Length(condition) > 1) and (condition[2] = '=') then
                 zerofmt := fmt;
             end;
        '=': begin
               zerofmt := fmt;
             end;
      end;
      ANode := ANode.NextSibling;
      inc(counter);
    end;
    {
    if posfmt = '' then posfmt := currFmt;
    if negfmt = '' then negfmt := currFmt;
    }
//    if posfmt = '' then posfmt := AFormatStr;
//    if negfmt = '' then negfmt := AFormatStr;

    case counter of
      1: begin
           if negfmt = '' then negfmt := currfmt;
           AFormatStr := posfmt + ';' + negfmt;
         end;
      2: begin
           if zerofmt = '' then zerofmt := currfmt;
           AFormatStr := posfmt + ';' + negfmt + ';' + zerofmt;
         end;
      3: AFormatStr := posfmt + ';' + negfmt + ';' + zerofmt;
    end;

    if not (ANumFormat in [nfCurrency, nfCurrencyRed]) then
      ANumFormat := nfCustom;
  end;

  procedure ReadNumberStyle(ANumFormatNode: TDOMNode; ANumFormatName: String);
  var
    node, childNode: TDOMNode;
    nodeName: String;
    nf: TsNumberFormat;
    nfs: String;
    decs: Byte;
    s: String;
    fracInt, fracNum, fracDenom: Integer;
    grouping: Boolean;
    nex: Integer;
    cs: String;
    color: TsColorValue;
    hasColor: Boolean;
    idx: Integer;
  begin
    nfs := '';
    cs := '';
    hasColor := false;
    node := ANumFormatNode.FirstChild;
    while Assigned(node) do
    begin
      nodeName := node.NodeName;
      if nodeName = '#text' then
      begin
        node := node.NextSibling;
        Continue;
      end else
      if nodeName = 'number:number' then
      begin
        s := GetAttrValue(node, 'number:decimal-places');
        if s = '' then s := GetAttrValue(node, 'decimal-places');
        if s <> '' then decs := StrToInt(s) else decs := 0;
        grouping := GetAttrValue(node, 'number:grouping') = 'true';
        nf := IfThen(grouping, nfFixedTh, nfFixed);
        nfs := nfs + BuildNumberFormatString(nf, Workbook.FormatSettings, decs);
      end else
      if nodeName = 'number:fraction' then
      begin
        nf := nfFraction;
        s := GetAttrValue(node, 'number:min-integer-digits');
        if s <> '' then fracInt := StrToInt(s) else fracInt := 0;
        s := GetAttrValue(node, 'number:min-numerator-digits');
        if s <> '' then fracNum := StrToInt(s) else fracNum := 0;
        s := GetAttrValue(node, 'number:min-denominator-digits');
        if s <> '' then fracDenom := StrToInt(s) else fracDenom := 0;
        s := GetAttrValue(node, 'number:denominator-value');
        if s <> '' then fracDenom := -StrToInt(s);
        nfs := nfs + BuildFractionFormatString(fracInt > 0, fracNum, fracDenom);
      end else
      if nodeName = 'number:scientific-number' then
      begin
        nf := nfExp;
        s := GetAttrValue(node, 'number:decimal-places');
        if s <> '' then decs := StrToInt(s) else decs := 0;
        s := GetAttrValue(node, 'number:min-exponent-digits');
        if s <> '' then nex := StrToInt(s) else nex := 1;
        nfs := nfs + BuildNumberFormatString(nfFixed, Workbook.FormatSettings, decs);
        nfs := nfs + 'E+' + DupeString('0', nex);
      end else
      if nodeName = 'number:currency-symbol' then
      begin
        childnode := node.FirstChild;
        while childnode <> nil do
        begin
          cs := cs + childNode.NodeValue;
          nfs := nfs + '"' + childNode.NodeValue + '"';
          childNode := childNode.NextSibling;
        end;
      end else
      if nodeName = 'number:text' then
      begin
        childNode := node.FirstChild;
        while childNode <> nil do
        begin
          nfs := nfs + childNode.NodeValue;
          childNode := childNode.NextSibling;
        end;
      end else
      if nodeName = 'style:text-properties' then
      begin
        s := GetAttrValue(node, 'fo:color');
        if s <> '' then
        begin
          hasColor := true;
         // {                        // currently not needed
          color := HTMLColorStrToColor(s);
          idx := FWorkbook.AddColorToPalette(color);
          if idx < 8 then
            nfs := Format('[%s]%s', [FWorkbook.GetColorName(idx), nfs])
          else
            nfs := Format('[Color%d]%s', [idx, nfs]);
         // }
        end;
      end;
      node := node.NextSibling;
    end;

    node := ANumFormatNode.FindNode('style:map');
    if node <> nil then
      ReadStyleMap(node, nf, nfs);

    if ANumFormatNode.NodeName = 'number:percentage-style' then
      nf := nfPercentage
    else
    if (ANumFormatNode.NodeName = 'number:currency-style') then
      nf := IfThen(hasColor, nfCurrencyRed, nfCurrency);

    NumFormatList.Add(Format('%s:%s', [ANumFormatName, nfs]));
  end;

  procedure ReadDateTimeStyle(ANumFormatNode: TDOMNode; ANumFormatName: String);
  var
    node, childNode: TDOMNode;
    nf: TsNumberFormat;
    nfs: String;
    nodeName: String;
    s, stxt, sovr: String;
    isInterval: Boolean;
  begin
    nfs := '';
    isInterval := false;
    sovr := GetAttrValue(ANumFormatNode, 'number:truncate-on-overflow');
    if (sovr = 'false') then
      isInterval := true;
    node := ANumFormatNode.FirstChild;
    while Assigned(node) do
    begin
      nodeName := node.NodeName;
      if nodeName = '#text' then
      begin
        node := node.NextSibling;
        Continue;
      end else
      if nodeName = 'number:year' then
      begin
        s := GetAttrValue(node, 'number:style');
        nfs := nfs + IfThen(s = 'long', 'yyyy', 'yy');
      end else
      if nodeName = 'number:month' then
      begin
        s := GetAttrValue(node, 'number:style');
        stxt := GetAttrValue(node, 'number:textual');
        if (stxt = 'true') then  // Month as text
          nfs := nfs + IfThen(s = 'long', 'mmmm', 'mmm')
        else                     // Month as number
          nfs := nfs + IfThen(s = 'long', 'mm', 'm');
      end else
      if nodeName = 'number:day' then
      begin
        s := GetAttrValue(node, 'number:style');
        nfs := nfs + IfThen(s = 'long', 'dd', 'd');
      end else
      if nodeName = 'number:day-of-week' then
      begin
        s := GetAttrValue(node, 'number:style');
        nfs := nfs + IfThen(s = 'long', 'dddd', 'ddd');
      end else
      if nodeName = 'number:hours' then
      begin
        s := GetAttrValue(node, 'number:style');
        if (sovr = 'false') then
          nfs := nfs + IfThen(s = 'long', '[hh]', '[h]')
        else
          nfs := nfs + IfThen(s = 'long', 'hh', 'h');
        sovr := '';
      end else
      if nodeName = 'number:minutes' then
      begin
        s := GetAttrValue(node, 'number:style');
        if (sovr = 'false') then
          nfs := nfs + IfThen(s = 'long', '[nn]', '[n]')
        else
          nfs := nfs + IfThen(s = 'long', 'nn', 'n');
        sovr := '';
      end else
      if nodeName = 'number:seconds' then
      begin
        s := GetAttrValue(node, 'number:style');
        if (sovr = 'false') then
          nfs := nfs + IfThen(s = 'long', '[ss]', '[s]')
        else
          nfs := nfs + IfThen(s = 'long', 'ss', 's');
        sovr := '';
        s := GetAttrValue(node, 'number:decimal-places');
        if (s <> '') and (s <> '0') then
          nfs := nfs + '.' + DupeString('0', StrToInt(s));
      end else
      if nodeName = 'number:am-pm' then
        nfs := nfs + 'AM/PM'
      else
      if nodeName = 'number:text' then
      begin
        childnode := node.FirstChild;
        if childnode <> nil then
        begin
          s := childNode.NodeValue;
          if pos(';', s) > 0 then
            nfs := nfs + '"' + s + '"'
            // avoid "misunderstanding" the semicolon as a section separator!
          else
            nfs := nfs + childnode.NodeValue;
        end;
      end;
      node := node.NextSibling;
    end;

    nf := IfThen(isInterval, nfTimeInterval, nfCustom);
    node := ANumFormatNode.FindNode('style:map');
    if node <> nil then
      ReadStyleMap(node, nf, nfs);

    NumFormatList.Add(ANumFormatName + ':' + nfs);
//    NumFormatList.AddFormat(ANumFormatName, nf, nfs);
  end;

  procedure ReadTextStyle(ANumFormatNode: TDOMNode; ANumFormatName: String);
  var
    node, childNode: TDOMNode;
    nf: TsNumberFormat = nfGeneral;
    nfs: String;
    nodeName: String;
  begin
    nfs := '';
    node := ANumFormatNode.FirstChild;
    while Assigned(node) do
    begin
      nodeName := node.NodeName;
      if nodeName = '#text' then
      begin
        node := node.NextSibling;
        Continue;
      end else
      if nodeName = 'number:text-content' then
      begin
        // ???
      end else
      if nodeName = 'number:text' then
      begin
        childnode := node.FirstChild;
        if childnode <> nil then
          nfs := nfs + childnode.NodeValue;
      end;
      node := node.NextSibling;
    end;

    node := ANumFormatNode.FindNode('style:map');
    if node <> nil then
      ReadStyleMap(node, nf, nfs);
    nf := nfCustom;

    NumFormatList.Add(Format('%s:%s', [ANumFormatName, nfs]));

    //NumFormatList.AddFormat(ANumFormatName, nf, nfs);
  end;

var
  NumFormatNode: TDOMNode;
  numfmt_nodename, numfmtname: String;

begin
  if not Assigned(AStylesNode) then
    exit;

  NumFormatNode := AStylesNode.FirstChild;
  while Assigned(NumFormatNode) do
  begin
    numfmt_nodename := NumFormatNode.NodeName;

    if NumFormatNode.HasAttributes then
      numfmtName := GetAttrValue(NumFormatNode, 'style:name') else
      numfmtName := '';

    // Numbers (nfFixed, nfFixedTh, nfExp, nfPercentage)
    if (numfmt_nodename = 'number:number-style') or
       (numfmt_nodename = 'number:percentage-style') or
       (numfmt_nodename = 'number:currency-style')
    then
      ReadNumberStyle(NumFormatNode, numfmtName);

    // Date/time values
    if (numfmt_nodename = 'number:date-style') or (numfmt_nodename = 'number:time-style') then
      ReadDateTimeStyle(NumFormatNode, numfmtName);

    // Text values
    if (numfmt_nodename = 'number:text-style') then
      ReadTextStyle(NumFormatNode, numfmtName);

    // Next node
    NumFormatNode := NumFormatNode.NextSibling;
  end;
end;

{ Finds the PageLayout record for a given TableStyle name in the "styles" nodes.
  First, seeks the TableStyle among the children of the "styles" node in the
  contents.xml - this node contains the name of the used master page.
  Then seeks the FMasterPageList for the entry with the determined master page
  name. This entry contains the name of the associated PageLayoutData stored in
  the PageLayoutList which, finally, contains the requested PageLayout record. }
function TsSpreadOpenDocReader.ReadPageLayout(AStylesNode: TDOMNode;
  ATableStyleName: String): PsPageLayout;
var
  nodeName, s: String;
  node: TDOMNode;
  masterPageName: String;
  masterPageData: TMasterPageData;
  pageLayoutData: TPageLayoutData;
  i, j: Integer;
begin
  Result := nil;

  if AStylesNode = nil then
    exit;

  { Looking through the "styles" node...}
  node := AStylesNode.FirstChild;
  while node <> nil do
  begin
    nodeName := node.NodeName;
    { ... for the node which is named like the requested TableStyle }
    if nodeName = 'style:style' then
    begin
      s := GetAttrValue(node, 'style:name');
      if s = ATableStyleName then
      begin
        { Found: extract the name of the master page }
        masterPageName := GetAttrValue(node, 'style:master-page-name');
        if masterPageName = '' then
          exit;

        { Looking through the MasterPage list...}
        for i:=0 to FMasterPageList.Count-1 do
        begin
          masterPageData := TMasterPageData(FMasterPageList[i]);
          { ... for the entry with the found master page name }
          if masterPageData.Name = masterPageName then
          begin
            { Found: looking through the PageLayout list ...}
            for j:=0 to FPageLayoutList.Count-1 do
            begin
              pageLayoutData := TPageLayoutData(FPageLayoutList[j]);
              { ... for the entry with the name specified by the master page }
              if pageLayoutData.Name = masterPageData.PageLayoutName then
              begin
                { Found: Return a pointer to the PageLayout record stored in the list }
                Result := @pageLayoutData.PageLayout;
                exit;
              end;
            end;
          end;
        end;
      end;
    end;
    { Not found: try next node in the styles list }
    node := node.NextSibling;
  end;
end;

{ Reads the cells in the given table. Loops through all rows, and then finds all
  cells of each row. }
procedure TsSpreadOpenDocReader.ReadRowsAndCells(ATableNode: TDOMNode);
var
  row: Integer;
  col: Integer;
  cellNode, rowNode: TDOMNode;
  nodeName: String;
  paramValueType, paramFormula, tableStyleName: String;
  paramColsSpanned, paramRowsSpanned: String;
  paramColsRepeated, paramRowsRepeated: String;
  rowsRepeated: Integer;
  rowsSpanned: Integer;
  colsSpanned: Integer;
  rowStyleName: String;
  rowStyleIndex: Integer;
  rowStyle: TRowStyleData;
  rowHeight: Single;
  autoRowHeight: Boolean;
  i, n: Integer;
  cell: PCell;
begin
  rowsRepeated := 0;
  row := 0;

  rowNode := ATableNode.FindNode('table:table-row');
  while Assigned(rowNode) do
  begin
    nodename := rowNode.NodeName;

    // Skip all non table-row nodes:
    // Nodes '#text' occur due to indentation spaces which are not skipped
    // automatically any more due to PreserveWhiteSpace option applied
    // to ReadXMLFile
    // And there are other nodes like 'table:named-expression' which we don't
    // need (at the moment)
    if nodeName <> 'table:table-row' then
    begin
      rowNode := rowNode.NextSibling;
      Continue;
    end;

    // Read rowstyle
    rowStyleName := GetAttrValue(rowNode, 'table:style-name');
    rowStyleIndex := FindRowStyleByName(rowStyleName);
    if rowStyleIndex > -1 then        // just for safety
    begin
      rowStyle := TRowStyleData(FRowStyleList[rowStyleIndex]);
      rowHeight := rowStyle.RowHeight;           // in mm (see ReadRowStyles)
      rowHeight := mmToPts(rowHeight) / Workbook.GetDefaultFontSize;
      if rowHeight > ROW_HEIGHT_CORRECTION
        then rowHeight := rowHeight - ROW_HEIGHT_CORRECTION  // in "lines"
        else rowHeight := 0;
      autoRowHeight := rowStyle.AutoRowHeight;
    end else
      autoRowHeight := true;

    col := 0;

    //process each cell of the row
    cellNode := rowNode.FindNode('table:table-cell');
    while Assigned(cellNode) do
    begin
      nodeName := cellNode.NodeName;
      if nodeName = 'table:table-cell' then
      begin
        // select this cell value's type
        paramValueType := GetAttrValue(CellNode, 'office:value-type');
        paramFormula := GetAttrValue(CellNode, 'table:formula');
        tableStyleName := GetAttrValue(CellNode, 'table:style-name');

        if paramValueType = 'string' then
          ReadLabel(row, col, cellNode)
        else
        if (paramValueType = 'float') or (paramValueType = 'percentage') or
           (paramValueType = 'currency')
        then
          ReadNumber(row, col, cellNode)
        else if (paramValueType = 'date') or (paramValueType = 'time') then
          ReadDateTime(row, col, cellNode)
        else if (paramValueType = 'boolean') then
          ReadBoolean(row, col, cellNode)
        else if (paramValueType = '') and (tableStyleName <> '') then
          ReadBlank(row, col, cellNode);
        { NOTE: Empty cells having no cell format, but a column format only,
          are skipped here. --> Currently the reader does not detect the format
          of empty cells correctly.
          It would work if the "(tableStyleName <> '')" would be omitted, but
          then the reader would create a record for all 1E9 cells prepared by
          the Excel2007 export --> crash!
          The column format is available in the FColumnList, but since the usage
          of colsSpanned in the row it is possible to miss the correct column format.
          Pretty nasty situation! }

        if ParamFormula <> '' then
          ReadFormula(row, col, cellNode);

        // Read cell comment
        ReadComment(row, col, cellNode);

        paramColsSpanned := GetAttrValue(cellNode, 'table:number-columns-spanned');
        if paramColsSpanned <> '' then
          colsSpanned := StrToInt(paramColsSpanned) - 1
        else
          colsSpanned := 0;

        paramRowsSpanned := GetAttrValue(cellNode, 'table:number-rows-spanned');
        if paramRowsSpanned <> '' then
          rowsSpanned := StrToInt(paramRowsSpanned) - 1
        else
          rowsSpanned := 0;

        if (colsSpanned <> 0) or (rowsSpanned <> 0) then
          FWorksheet.MergeCells(row, col, row+rowsSpanned, col+colsSpanned);

        paramColsRepeated := GetAttrValue(cellNode, 'table:number-columns-repeated');
        if paramColsRepeated = '' then paramColsRepeated := '1';
        n := StrToInt(paramColsRepeated);
        if n > 1 then
        begin
          cell := FWorksheet.FindCell(row, col);
          if cell <> nil then
            for i:=1 to n-1 do
              FWorksheet.CopyCell(row, col, row, col+i);
        end;
      end
      else
      if nodeName = 'table:covered-table-cell' then
      begin
        paramColsRepeated := GetAttrValue(cellNode, 'table:number-columns-repeated');
        if paramColsRepeated = '' then paramColsRepeated := '1';
      end else
        paramColsRepeated := '0';

      col := col + StrToInt(paramColsRepeated);
      cellNode := cellNode.NextSibling;
    end; //while Assigned(cellNode)

    paramRowsRepeated := GetAttrValue(RowNode, 'table:number-rows-repeated');
    if paramRowsRepeated = '' then
      rowsRepeated := 1
    else
      rowsRepeated := StrToInt(paramRowsRepeated);

    // Transfer non-default row heights to sheet's rows
    if not autoRowHeight then
      for i:=1 to rowsRepeated do
        FWorksheet.WriteRowHeight(row + i - 1, rowHeight);

    row := row + rowsRepeated;

    rowNode := rowNode.NextSibling;
  end; // while Assigned(rowNode)
end;

procedure TsSpreadOpenDocReader.ReadRowStyle(AStyleNode: TDOMNode);
var
  styleName: String;
  styleChildNode: TDOMNode;
  rowHeight: Double;
  auto: Boolean;
  s: String;
  rowStyle: TRowStyleData;
begin
  styleName := GetAttrValue(AStyleNode, 'style:name');
  styleChildNode := AStyleNode.FirstChild;
  rowHeight := -1;
  auto := false;

  while Assigned(styleChildNode) do
  begin
    if styleChildNode.NodeName = 'style:table-row-properties' then
    begin
      s := GetAttrValue(styleChildNode, 'style:row-height');
      if s <> '' then
        rowHeight := PtsToMm(HTMLLengthStrToPts(s));  // convert to mm
      s := GetAttrValue(styleChildNode, 'style:use-optimal-row-height');
      if s = 'true' then
        auto := true;
    end;
    styleChildNode := styleChildNode.NextSibling;
  end;

  rowStyle := TRowStyleData.Create;
  rowStyle.Name := styleName;
  rowStyle.RowHeight := rowHeight;
  rowStyle.AutoRowHeight := auto;
  FRowStyleList.Add(rowStyle);
end;

procedure TsSpreadOpenDocReader.ReadSettings(AOfficeSettingsNode: TDOMNode);
var
  cfgItemSetNode, cfgItemNode, cfgItemMapEntryNode, cfgEntryItemNode, cfgTableItemNode, node: TDOMNode;
  nodeName, cfgName, cfgValue, tblName: String;
  sheet: TsWorksheet;
  vsm, hsm, hsp, vsp: Integer;
  showGrid, showHeaders: Boolean;
  i: Integer;
begin
  showGrid := true;
  showHeaders := true;
  cfgItemSetNode := AOfficeSettingsNode.FirstChild;
  while Assigned(cfgItemSetNode) do
  begin
    if (cfgItemSetNode.NodeName <> '#text') and
       (GetAttrValue(cfgItemSetNode, 'config:name') = 'ooo:view-settings') then
    begin
      cfgItemNode := cfgItemSetNode.FirstChild;
      while Assigned(cfgItemNode) do begin
        if (cfgItemNode.NodeName <> '#text') and
           (cfgItemNode.NodeName = 'config:config-item-map-indexed') and
           (GetAttrValue(cfgItemNode, 'config:name') = 'Views') then
        begin
          cfgItemMapEntryNode := cfgItemNode.FirstChild;
          while Assigned(cfgItemMapEntryNode) do
          begin
            cfgEntryItemNode := cfgItemMapEntryNode.FirstChild;
            while Assigned(cfgEntryItemNode) do
            begin
              nodeName := cfgEntryItemNode.NodeName;
              if (nodeName = 'config:config-item') then
              begin
                cfgName := lowercase(GetAttrValue(cfgEntryItemNode, 'config:name'));
                if cfgName = 'showgrid' then
                begin
                  cfgValue := GetNodeValue(cfgEntryItemNode);
                  if cfgValue = 'false' then showGrid := false;
                end else
                if cfgName = 'hascolumnrowheaders' then
                begin
                  cfgValue := GetNodeValue(cfgEntryItemNode);
                  if cfgValue = 'false' then showHeaders := false;
                end;
              end else
              if (nodeName = 'config:config-item-map-named') and
                 (GetAttrValue(cfgEntryItemNode, 'config:name') = 'Tables') then
              begin
                cfgTableItemNode := cfgEntryItemNode.FirstChild;
                while Assigned(cfgTableItemNode) do
                begin
                  nodeName := cfgTableItemNode.NodeName;
                  if nodeName <> '#text' then
                  begin
                    tblName := GetAttrValue(cfgTableItemNode, 'config:name');
                    if tblName <> '' then
                    begin
                      hsm := 0; vsm := 0;
                      sheet := Workbook.GetWorksheetByName(tblName);
                      if sheet <> nil then
                      begin
                        node := cfgTableItemNode.FirstChild;
                        while Assigned(node) do
                        begin
                          nodeName := node.NodeName;
                          if nodeName <> '#text' then
                          begin
                            cfgName := GetAttrValue(node, 'config:name');
                            cfgValue := GetNodeValue(node);
                            if cfgName = 'VerticalSplitMode' then
                              vsm := StrToInt(cfgValue)
                            else if cfgName = 'HorizontalSplitMode' then
                              hsm := StrToInt(cfgValue)
                            else if cfgName = 'VerticalSplitPosition' then
                              vsp := StrToInt(cfgValue)
                            else if cfgName = 'HorizontalSplitPosition' then
                              hsp := StrToInt(cfgValue);
                          end;
                          node := node.NextSibling;
                        end;
                        if (hsm = 2) or (vsm = 2) then
                        begin
                          sheet.Options := sheet.Options + [soHasFrozenPanes];
                          sheet.LeftPaneWidth := hsp;
                          sheet.TopPaneHeight := vsp;
                        end else
                          sheet.Options := sheet.Options - [soHasFrozenPanes];
                      end;
                    end;
                  end;
                  cfgTableItemNode := cfgTableItemNode.NextSibling;
                end;
              end;
              cfgEntryItemNode := cfgEntryItemNode.NextSibling;
            end;
            cfgItemMapEntryNode := cfgItemMapEntryNode.NextSibling;
          end;
        end;
        cfgItemNode := cfgItemNode.NextSibling;
      end;
    end;
    cfgItemSetNode := cfgItemSetNode.NextSibling;
  end;

  { Now let's apply the showGrid and showHeader values to all sheets - they
    are document-wide settings (although there is a ShowGrid in the Tables node) }
  for i:=0 to Workbook.GetWorksheetCount-1 do
  begin
    sheet := Workbook.GetWorksheetByIndex(i);
    if not showGrid then sheet.Options := sheet.Options - [soShowGridLines];
    if not showHeaders then sheet.Options := sheet.Options - [soShowHeaders];
  end;
end;

procedure TsSpreadOpenDocReader.ReadStyles(AStylesNode: TDOMNode);
var
  styleNode: TDOMNode;
  styleChildNode: TDOMNode;
  nodeName: String;
  family: String;
  styleName: String;
  fmt: TsCellFormat;
  numFmtIndexDefault: Integer;
  numFmtName: String;
  numFmtStr: String;
  numFmtIndex: Integer;
  numFmtParams: TsNumFormatParams;
  clr: TsColorValue;
  s: String;

  procedure SetBorderStyle(ABorder: TsCellBorder; AStyleValue: String);
  const
    EPS = 0.1;  // takes care of rounding errors for line widths
  var
    L: TStringList;
    i: Integer;
    s: String;
    wid: Double;
    linestyle: String;
    rgb: TsColorValue;
    p: Integer;
  begin
    L := TStringList.Create;
    try
      L.Delimiter := ' ';
      L.StrictDelimiter := true;
      L.DelimitedText := AStyleValue;
      wid := 0;
      rgb := TsColorValue(-1);
      linestyle := '';
      for i:=0 to L.Count-1 do
      begin
        s := L[i];
        if (s = 'solid') or (s = 'dashed') or (s = 'fine-dashed') or (s = 'dotted') or (s = 'double')
        then begin
          linestyle := s;
          continue;
        end;
        p := pos('pt', s);
        if p = Length(s)-1 then
        begin
          wid := StrToFloat(copy(s, 1, p-1), FPointSeparatorSettings);
          continue;
        end;
        p := pos('mm', s);
        if p = Length(s)-1 then
        begin
          wid := mmToPts(StrToFloat(copy(s, 1, p-1), FPointSeparatorSettings));
          Continue;
        end;
        p := pos('cm', s);
        if p = Length(s)-1 then
        begin
          wid := cmToPts(StrToFloat(copy(s, 1, p-1), FPointSeparatorSettings));
          Continue;
        end;
        rgb := HTMLColorStrToColor(s);
      end;
      fmt.BorderStyles[ABorder].LineStyle := lsThin;
      if (linestyle = 'solid') then
      begin
        if (wid >= 3 - EPS) then fmt.BorderStyles[ABorder].LineStyle := lsThick
        else if (wid >= 2 - EPS) then fmt.BorderStyles[ABorder].LineStyle := lsMedium
      end else
      if (linestyle = 'dotted') then
        fmt.BorderStyles[ABorder].LineStyle := lsHair
      else
      if (linestyle = 'dashed') then
        fmt.BorderStyles[ABorder].LineStyle := lsDashed
      else
      if (linestyle = 'fine-dashed') then
        fmt.BorderStyles[ABorder].LineStyle := lsDotted
      else
      if (linestyle = 'double') then
        fmt.BorderStyles[ABorder].LineStyle := lsDouble;
      fmt.BorderStyles[ABorder].Color := IfThen(rgb = TsColorValue(-1),
        scBlack, Workbook.AddColorToPalette(rgb));
    finally
      L.Free;
    end;
  end;

begin
  if not Assigned(AStylesNode) then
    exit;

  numFmtIndexDefault := FindNumFormatByName('N0');

  styleNode := AStylesNode.FirstChild;
  while Assigned(styleNode) do begin
    nodeName := styleNode.NodeName;
    if nodeName = 'style:default-style' then
    begin
      ReadFont(styleNode.FindNode('style:text-properties'), DEFAULT_FONTINDEX);
    end else
    if nodeName = 'style:style' then
    begin
      family := GetAttrValue(styleNode, 'style:family');

      // Column styles
      if family = 'table-column' then
        ReadColumnStyle(styleNode);

      // Row styles
      if family = 'table-row' then
        ReadRowStyle(styleNode);

      // Cell styles
      if family = 'table-cell' then
      begin
        styleName := GetAttrValue(styleNode, 'style:name');

        InitFormatRecord(fmt);
        fmt.Name := styleName;

        numFmtIndex := -1;
        numFmtName := GetAttrValue(styleNode, 'style:data-style-name');
        if numFmtName <> '' then numFmtIndex := FindNumFormatByName(numFmtName);
        if numFmtIndex = -1 then numFmtIndex := numFmtIndexDefault;
        numFmtStr := NumFormatList[numFmtIndex];
        numFmtStr := Copy(numFmtStr, pos(':', numFmtStr)+1, Length(numFmtStr));
        fmt.NumberFormatIndex := Workbook.AddNumberFormat(numFmtStr);
        numFmtParams := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
        if numFmtParams <> nil then begin
          fmt.NumberFormat := numFmtParams.NumFormat;
          fmt.NumberFormatStr := numFmtStr;
          Include(fmt.UsedFormattingFields, uffNumberFormat);
        end;

        styleChildNode := styleNode.FirstChild;
        while Assigned(styleChildNode) do
        begin
          nodeName := styleChildNode.NodeName;
          if nodeName = 'style:text-properties' then
          begin
            if SameText(stylename, 'Default') then
              fmt.FontIndex := ReadFont(styleChildNode, DEFAULT_FONTINDEX)
            else
            if SameText(stylename, 'Excel_20_Built-in_20_Hyperlink') then
              fmt.FontIndex := ReadFont(styleChildNode, HYPERLINK_FONTINDEX)
            else
              fmt.FontIndex := ReadFont(styleChildNode);
            {
            if fmt.FontIndex = BOLD_FONTINDEX  then
              Include(fmt.UsedFormattingFields, uffBold)
            else }
            if fmt.FontIndex > 0 then
              Include(fmt.UsedFormattingFields, uffFont);
          end else
          if nodeName = 'style:table-cell-properties' then
          begin
            // Background color
            s := GetAttrValue(styleChildNode, 'fo:background-color');
            if (s <> '') and (s <> 'transparent') then begin
              clr := HTMLColorStrToColor(s);
              // ODS does not support background fill patterns!
              fmt.Background.FgColor := IfThen(clr = TsColorValue(-1),
                scTransparent, Workbook.AddColorToPalette(clr));
              fmt.Background.BgColor := fmt.Background.FgColor;
              if (fmt.Background.BgColor <> scTransparent) then
              begin
                fmt.Background.Style := fsSolidFill;
                Include(fmt.UsedFormattingFields, uffBackground);
              end;
            end;
            // Borders
            s := GetAttrValue(styleChildNode, 'fo:border');
            if (s <> '') and (s <> 'none') then
            begin
              fmt.Border := fmt.Border + [cbNorth, cbSouth, cbEast, cbWest];
              SetBorderStyle(cbNorth, s);
              SetBorderStyle(cbSouth, s);
              SetBorderStyle(cbEast, s);
              SetBorderStyle(cbWest, s);
              Include(fmt.UsedFormattingFields, uffBorder);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-top');
            if (s <> '') and (s <> 'none') then
            begin
              Include(fmt.Border, cbNorth);
              SetBorderStyle(cbNorth, s);
              Include(fmt.UsedFormattingFields, uffBorder);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-right');
            if (s <> '') and (s <> 'none') then
            begin
              Include(fmt.Border, cbEast);
              SetBorderStyle(cbEast, s);
              Include(fmt.UsedFormattingFields, uffBorder);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-bottom');
            if (s <> '') and (s <> 'none') then
            begin
              Include(fmt.Border, cbSouth);
              SetBorderStyle(cbSouth, s);
              Include(fmt.UsedFormattingFields, uffBorder);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-left');
            if (s <> '') and (s <> 'none') then
            begin
              Include(fmt.Border, cbWest);
              SetBorderStyle(cbWest, s);
              Include(fmt.UsedFormattingFields, uffBorder);
            end;
            s := GetAttrValue(styleChildNode, 'style:diagonal-bl-tr');
            if (s <> '') and (s <> 'none') then
            begin
              Include(fmt.Border, cbDiagUp);
              SetBorderStyle(cbDiagUp, s);
              Include(fmt.UsedFormattingFields, uffBorder);
            end;
            s := GetAttrValue(styleChildNode, 'style:diagonal-tl-br');
            if (s <> '') and (s <>'none') then
            begin
              Include(fmt.Border, cbDiagDown);
              SetBorderStyle(cbDiagDown, s);
              Include(fmt.UsedFormattingFields, uffBorder);
            end;

            // Text wrap
            s := GetAttrValue(styleChildNode, 'fo:wrap-option');
            if (s='wrap') then
              Include(fmt.UsedFormattingFields, uffWordwrap);

            // Test rotation
            s := GetAttrValue(styleChildNode, 'style:rotation-angle');
            if s = '90' then
              fmt.TextRotation := rt90DegreeCounterClockwiseRotation
            else if s = '270' then
              fmt.TextRotation := rt90DegreeClockwiseRotation;
            s := GetAttrValue(styleChildNode, 'style:direction');
            if s = 'ttb' then
              fmt.TextRotation := rtStacked;
            if fmt.TextRotation <> trHorizontal then
              Include(fmt.UsedFormattingFields, uffTextRotation);

            // Vertical text alignment
            s := GetAttrValue(styleChildNode, 'style:vertical-align');
            if s = 'top' then
              fmt.VertAlignment := vaTop
            else if s = 'middle' then
              fmt.VertAlignment := vaCenter
            else if s = 'bottom' then
              fmt.VertAlignment := vaBottom;
            if fmt.VertAlignment <> vaDefault then
              Include(fmt.UsedFormattingFields, uffVertAlign);
          end
          else
          if nodeName = 'style:paragraph-properties' then
          begin
            // Horizontal text alignment
            s := GetAttrValue(styleChildNode, 'fo:text-align');
            if s = 'start' then
              fmt.HorAlignment := haLeft
            else if s = 'end' then
              fmt.HorAlignment := haRight
            else if s = 'center' then
              fmt.HorAlignment := haCenter;
            if fmt.HorAlignment <> haDefault then
              Include(fmt.UsedFormattingFields, uffHorAlign);
          end;
          styleChildNode := styleChildNode.NextSibling;
        end;

        FCellFormatList.Add(fmt);
      end;
    end;
    styleNode := styleNode.NextSibling;
  end;
end;


{ TsSpreadOpenDocWriter }

procedure TsSpreadOpenDocWriter.AddBuiltinNumFormats;
begin
  FNumFormatList.Clear;
  FNumFormatList.Add('N0:');
end;

{ Creates the streams for the individual data files. Will be zipped into a
  single xlsx file. }
procedure TsSpreadOpenDocWriter.CreateStreams;
begin
  if (boBufStream in Workbook.Options) then
  begin
    FSMeta := TBufStream.Create(GetTempFileName('', 'fpsM'));
    FSSettings := TBufStream.Create(GetTempFileName('', 'fpsS'));
    FSStyles := TBufStream.Create(GetTempFileName('', 'fpsSTY'));
    FSContent := TBufStream.Create(GetTempFileName('', 'fpsC'));
    FSMimeType := TBufStream.Create(GetTempFileName('', 'fpsMT'));
    FSMetaInfManifest := TBufStream.Create(GetTempFileName('', 'fpsMIM'));
  end else
  begin
    FSMeta := TMemoryStream.Create;
    FSSettings := TMemoryStream.Create;
    FSStyles := TMemoryStream.Create;
    FSContent := TMemoryStream.Create;
    FSMimeType := TMemoryStream.Create;
    FSMetaInfManifest := TMemoryStream.Create;
  end;
  // FSSheets will be created when needed.
end;

{ Destroys the streams that were created by the writer }
procedure TsSpreadOpenDocWriter.DestroyStreams;

  procedure DestroyStream(AStream: TStream);
  var
    fn: String;
  begin
    if AStream is TFileStream then
    begin
      fn := TFileStream(AStream).Filename;
      DeleteFile(fn);
    end;
    AStream.Free;
  end;

begin
  DestroyStream(FSMeta);
  DestroyStream(FSSettings);
  DestroyStream(FSStyles);
  DestroyStream(FSContent);
  DestroyStream(FSMimeType);
  DestroyStream(FSMetaInfManifest);
end;

procedure TsSpreadOpenDocWriter.ListAllColumnStyles;
var
  i, j, c: Integer;
  sheet: TsWorksheet;
  found: Boolean;
  colstyle: TColumnStyleData;
  w: Double;
  multiplier: Double;
begin
  { At first, add the default column width }
  colStyle := TColumnStyleData.Create;
  colStyle.Name := 'co1';
  colStyle.ColWidth := 12; //Workbook.DefaultColWidth;
  FColumnStyleList.Add(colStyle);

  for i:=0 to Workbook.GetWorksheetCount-1 do
  begin
    sheet := Workbook.GetWorksheetByIndex(i);
//    colStyle.ColWidth := sheet.DefaultColWidth;
    for c:=0 to sheet.GetLastColIndex do
    begin
      w := sheet.GetColWidth(c);
      // Look for this width in the current ColumnStyleList
      found := false;
      for j := 0 to FColumnStyleList.Count-1 do
        if SameValue(TColumnStyleData(FColumnStyleList[j]).ColWidth, w, COLWIDTH_EPS)
        then begin
          found := true;
          break;
        end;
      // Not found? Then add the column as new column style
      if not found then
      begin
        colStyle := TColumnStyleData.Create;
        colStyle.Name := Format('co%d', [FColumnStyleList.Count+1]);
        colStyle.ColWidth := w;
        FColumnStyleList.Add(colStyle);
      end;
    end;
  end;

  { fpspreadsheet's column width is the count of '0' characters of the
    default font. On average, the width of the '0' is about half of the
    point size of the font. --> we can convert the fps col width to pts and
    then to millimeters. }
  multiplier := Workbook.GetFont(0).Size / 2;
  for i:=0 to FColumnStyleList.Count-1 do
  begin
    w := TColumnStyleData(FColumnStyleList[i]).ColWidth * multiplier;
    TColumnStyleData(FColumnStyleList[i]).ColWidth := PtsToMM(w);
  end;
end;

{ Collects the fonts used by headers and footers in the FHeaderFooterFontList }
procedure TsSpreadOpenDocWriter.ListAllHeaderFooterFonts;

  { Add the fonts used in the specified header/footer line to the
    HeaderFooterFontList. This is done while the HeaderFooterParser is created. }
  procedure AddFontsOfHeaderFooter(AText: String; ADefaultFont: TsHeaderFooterFont);
  begin
    TsSpreadOpenDocHeaderFooterParser.Create(AText, FHeaderFooterFontList, ADefaultFont).Free;
  end;

var
  defFnt: TsHeaderFooterFont;
  i: Integer;
  sheet: TsWorksheet;
begin
  defFnt := TsHeaderFooterFont.Create(Workbook.GetDefaultFont);
  try
    for i:=0 to Workbook.GetWorksheetCount-1 do
    begin
      sheet := Workbook.GetWorksheetByIndex(i);
      AddFontsOfHeaderFooter(sheet.pageLayout.Headers[1], defFnt);
      AddFontsOfHeaderFooter(sheet.PageLayout.Headers[2], defFnt);
      AddFontsOfHeaderFooter(sheet.PageLayout.Footers[1], defFnt);
      AddFontsOfHeaderFooter(sheet.PageLayout.Footers[2], defFnt);
    end;
  finally
    defFnt.Free;
  end;
end;

{ Contains all number formats used in the workbook. Overrides the inherited
  method to assign a unique name according to the OpenDocument syntax ("N<number>"
  to the format items. }
procedure TsSpreadOpenDocWriter.ListAllNumFormats;
const
  FMT_BASE = 1000;  // Format number to start with. Not clear if this is correct...
var
  i: Integer;
  nfparams: TsNumFormatParams;
begin
  // The default format has already been added.
  for i:=0 to Workbook.GetNumberFormatCount - 1 do
  begin
    nfParams := Workbook.GetNumberFormat(i);
    if nfParams <> nil then
      FNumFormatList.Add(Format('N%d:%s', [FMT_BASE+i, nfParams.NumFormatStr]));
  end;
end;

procedure TsSpreadOpenDocWriter.ListAllRowStyles;
var
  i, j, r: Integer;
  sheet: TsWorksheet;
  row: PRow;
  found: Boolean;
  rowstyle: TRowStyleData;
  h, multiplier: Double;
begin
  { At first, add the default row height }
  { Initially, row height units will be the same as in the sheet, i.e. in "lines" }
  rowStyle := TRowStyleData.Create;
  rowStyle.Name := 'ro1';
  rowStyle.RowHeight := 1; //Workbook.DefaultRowHeight;
  rowStyle.AutoRowHeight := true;
  FRowStyleList.Add(rowStyle);

  for i:=0 to Workbook.GetWorksheetCount-1 do
  begin
    sheet := Workbook.GetWorksheetByIndex(i);
    for r:=0 to sheet.GetLastRowIndex do
    begin
      row := sheet.FindRow(r);
      if row <> nil then
      begin
        h := sheet.GetRowHeight(r);
        // Look for this height in the current RowStyleList
        found := false;
        for j:=0 to FRowStyleList.Count-1 do
          if SameValue(TRowStyleData(FRowStyleList[j]).RowHeight, h, ROWHEIGHT_EPS) and
             (not TRowStyleData(FRowStyleList[j]).AutoRowHeight) then
          begin
            found := true;
            break;
          end;
        // Not found? Then add the row as a new row style
        if not found then
        begin
          rowStyle := TRowStyleData.Create;
          rowStyle.Name := Format('ro%d', [FRowStyleList.Count+1]);
          rowStyle.RowHeight := h;
          rowStyle.AutoRowHeight := false;
          FRowStyleList.Add(rowStyle);
        end;
      end;
    end;
  end;

  { fpspreadsheet's row heights are measured as line count of the default font.
    Using the default font size (which is in points) we convert the line count
    to points and then to millimeters as needed by ods. }
  multiplier := Workbook.GetDefaultFontSize;;
  for i:=0 to FRowStyleList.Count-1 do
  begin
    h := (TRowStyleData(FRowStyleList[i]).RowHeight + ROW_HEIGHT_CORRECTION) * multiplier;
    TRowStyleData(FRowStyleList[i]).RowHeight := PtsToMM(h);
  end;
end;

{ Is called before zipping the individual file parts. Rewinds the streams. }
procedure TsSpreadOpenDocWriter.ResetStreams;
begin
  FSMeta.Position := 0;
  FSSettings.Position := 0;
  FSStyles.Position := 0;
  FSContent.Position := 0;
  FSMimeType.Position := 0;
  FSMetaInfManifest.Position := 0;
end;

procedure TsSpreadOpenDocWriter.WriteAutomaticStyles(AStream: TStream);
var
  i: Integer;
  sheet: TsWorksheet;
  fnt: TXMLHeaderFooterFont;

begin
  AppendToStream(AStream,
    '<office:automatic-styles>');

  AppendToStream(AStream,
      '<style:page-layout style:name="Mpm1">' +

        '<style:page-layout-properties '+
             'fo:margin-top="1.25cm" '+
             'fo:margin-bottom="1.25cm" '+
             'fo:margin-left="1.905cm" '+
             'fo:margin-right="1.905cm" />' +

        '<style:header-style>' +
          '<style:header-footer-properties '+
             'fo:min-height="0.751cm" '+
             'fo:margin-left="0cm" '+
             'fo:margin-right="0cm" '+
             'fo:margin-bottom="0.25cm" '+
             'fo:margin-top="0cm" />' +
        '</style:header-style>' +

        '<style:footer-style>' +
          '<style:header-footer-properties '+
             'fo:min-height="0.751cm" '+
             'fo:margin-left="0cm" '+
             'fo:margin-right="0cm" '+
             'fo:margin-top="0.25cm" '+
             'fo:margin-bottom="0cm" />' +
        '</style:footer-style>' +

      '</style:page-layout>');

  for i:=0 to FWorkbook.GetWorksheetCount-1 do begin
    sheet := FWorkbook.GetWorksheetByIndex(i);
    AppendToStream(AStream,
      WritePageLayoutAsXMLString('Mpm' + IntToStr(3+i), sheet.PageLayout));
  end;

  for i:=0 to FHeaderFooterFontList.Count-1 do
  begin
    fnt := TXMLHeaderFooterFont(FHeaderFooterFontList[i]);
    fnt.StyleName := 'MT' + IntToStr(i+1);
    AppendToStream(AStream, Format(
      '<style:style style:name="%s" style:family="text">' +
        '<style:text-properties %s />' +
      '</style:style>', [
      fnt.StyleName, WriteHeaderFooterFontXMLAsString(fnt)
    ]));
  end;

  AppendToStream(AStream,
    '</office:automatic-styles>');
end;

procedure TsSpreadOpenDocWriter.WriteMimetype;
begin
  AppendToStream(FSMimeType,
    'application/vnd.oasis.opendocument.spreadsheet'
  );
end;

procedure TsSpreadOpenDocWriter.WriteMetaInfManifest;
begin
  AppendToStream(FSMetaInfManifest,
    '<manifest:manifest xmlns:manifest="' + SCHEMAS_XMLNS_MANIFEST + '">');
  AppendToStream(FSMetaInfManifest,
      '<manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:full-path="/" />');
  AppendToStream(FSMetaInfManifest,
      '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml" />');
  AppendToStream(FSMetaInfManifest,
      '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml" />');
  AppendToStream(FSMetaInfManifest,
      '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml" />');
  AppendToStream(FSMetaInfManifest,
      '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml" />');
  AppendToStream(FSMetaInfManifest,
    '</manifest:manifest>');
end;

procedure TsSpreadOpenDocWriter.WriteMeta;
begin
  AppendToStream(FSMeta,
    XML_HEADER);
  AppendToStream(FSMeta,
    '<office:document-meta xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
      '" xmlns:dcterms="' + SCHEMAS_XMLNS_DCTERMS +
      '" xmlns:meta="' + SCHEMAS_XMLNS_META +
      '" xmlns="' + SCHEMAS_XMLNS +
      '" xmlns:ex="' + SCHEMAS_XMLNS + '">');
  AppendToStream(FSMeta,
      '<office:meta>',
        '<meta:generator>FPSpreadsheet Library</meta:generator>' +
        '<meta:document-statistic />',
      '</office:meta>');
  AppendToStream(FSMeta,
    '</office:document-meta>');
end;

procedure TsSpreadOpenDocWriter.WriteSettings;
var
  i: Integer;
  showGrid, showHeaders: Boolean;
  sheet: TsWorksheet;
begin
  // Open/LibreOffice allow to change showGrid and showHeaders only globally.
  // As a compromise, we check whether there is at least one page with these
  // settings off. Then we assume it to be valid also for the other sheets.
  showGrid := true;
  showHeaders := true;
  for i:=0 to Workbook.GetWorksheetCount-1 do
  begin
    sheet := Workbook.GetWorksheetByIndex(i);
    if not (soShowGridLines in sheet.Options) then showGrid := false;
    if not (soShowHeaders in sheet.Options) then showHeaders := false;
  end;

  AppendToStream(FSSettings,
    XML_HEADER);
  AppendToStream(FSSettings,
    '<office:document-settings xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
     '" xmlns:config="' + SCHEMAS_XMLNS_CONFIG +
     '" xmlns:ooo="' + SCHEMAS_XMLNS_OOO + '">');
  AppendToStream(FSSettings,
      '<office:settings>' +
        '<config:config-item-set config:name="ooo:view-settings">' +
         '<config:config-item-map-indexed config:name="Views">' +
            '<config:config-item-map-entry>' +
              '<config:config-item config:name="ActiveTable" config:type="string">Tabelle1</config:config-item>' +
              '<config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>' +
              '<config:config-item config:name="PageViewZoomValue" config:type="int">100</config:config-item>' +
              '<config:config-item config:name="ShowPageBreakPreview" config:type="boolean">false</config:config-item>' +
              '<config:config-item config:name="ShowGrid" config:type="boolean">'+FALSE_TRUE[showGrid]+'</config:config-item>' +
              '<config:config-item config:name="HasColumnRowHeaders" config:type="boolean">'+FALSE_TRUE[showHeaders]+'</config:config-item>' +
              '<config:config-item-map-named config:name="Tables">');

                WriteTableSettings(FSSettings);

  AppendToStream(FSSettings,
            '</config:config-item-map-named>' +
          '</config:config-item-map-entry>' +
        '</config:config-item-map-indexed>' +
      '</config:config-item-set>' +
    '</office:settings>' +
   '</office:document-settings>');
end;

procedure TsSpreadOpenDocWriter.WriteStyles;
begin
  AppendToStream(FSStyles,
     XML_HEADER);

  AppendToStream(FSStyles,
    '<office:document-styles xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
      '" xmlns:fo="' + SCHEMAS_XMLNS_FO +
      '" xmlns:style="' + SCHEMAS_XMLNS_STYLE +
      '" xmlns:svg="' + SCHEMAS_XMLNS_SVG +
      '" xmlns:table="' + SCHEMAS_XMLNS_TABLE +
      '" xmlns:text="' + SCHEMAS_XMLNS_TEXT +
      '" xmlns:v="' + SCHEMAS_XMLNS_V + '">');

  AppendToStream(FSStyles,
      '<office:font-face-decls>');
  WriteFontNames(FSStyles);
  AppendToStream(FSStyles,
      '</office:font-face-decls>');

  AppendToStream(FSStyles,
      '<office:styles>');
  AppendToStream(FSStyles,
        '<style:style style:name="Default" style:family="table-cell">',
           WriteDefaultFontXMLAsString,
        '</style:style>');
  AppendToStream(FSStyles,
      '</office:styles>');

  WriteAutomaticStyles(FSStyles);
  WriteMasterStyles(FSStyles);

  AppendToStream(FSStyles,
    '</office:document-styles>');
end;

procedure TsSpreadOpenDocWriter.WriteContent;
var
  i: Integer;
begin
  AppendToStream(FSContent,
    XML_HEADER);
  AppendToStream(FSContent,
    '<office:document-content xmlns:office="' + SCHEMAS_XMLNS_OFFICE +
        '" xmlns:fo="'     + SCHEMAS_XMLNS_FO +
        '" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/' +
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
        '" xmlns:xsi="'    + SCHEMAS_XMLNS_XSI + '">' +
      '<office:scripts />');

  // Fonts
  AppendToStream(FSContent,
      '<office:font-face-decls>');
         WriteFontNames(FSContent);
  AppendToStream(FSContent,
      '</office:font-face-decls>');

  // Automatic styles
  AppendToStream(FSContent,
      '<office:automatic-styles>');

  WriteNumFormats(FSContent);
  WriteColStyles(FSContent);
  WriteRowStyles(FSContent);
  WriteTableStyles(FSContent);
  WriteCellStyles(FSContent);

  AppendToStream(FSContent,
      '</office:automatic-styles>');

  // Body
  AppendToStream(FSContent,
      '<office:body>' +
        '<office:spreadsheet>');

  // Write all worksheets
  for i := 0 to Workbook.GetWorksheetCount - 1 do
    WriteWorksheet(FSContent, i);

  AppendToStream(FSContent,
        '</office:spreadsheet>' +
      '</office:body>' +
    '</office:document-content>'
  );
end;

procedure TsSpreadOpenDocWriter.WriteWorksheet(AStream: TStream;
  ASheetIndex: Integer);
begin
  FWorksheet := FWorkbook.GetWorksheetByIndex(ASheetIndex);

  // Header
  AppendToStream(AStream, Format(
    '<table:table table:name="%s" table:style-name="ta%d">', [
    FWorkSheet.Name, ASheetIndex+1
  ]));

  // columns
  WriteColumns(AStream, FWorkSheet);

  // rows and cells
  // The cells need to be written in order, row by row, cell by cell
  if (boVirtualMode in Workbook.Options) then
  begin
    if Assigned(Workbook.OnWriteCellData) then
      WriteVirtualCells(AStream, FWorksheet)
  end else
    WriteRowsAndCells(AStream, FWorksheet);

  // Footer
  AppendToStream(AStream,
    '</table:table>');
end;

procedure TsSpreadOpenDocWriter.WriteCellStyles(AStream: TStream);
var
  i, j, p: Integer;
  s: String;
  nfidx: Integer;
  nfs: String;
  fmt: TsCellFormat;
  nfParams: TsNumFormatParams;
begin
  for i := 0 to FWorkbook.GetNumCellFormats - 1 do
  begin
    fmt := FWorkbook.GetCellFormat(i);
    nfs := '';
    nfidx := fmt.NumberFormatIndex;
    if nfidx <> -1 then
    begin
      nfParams := FWorkbook.GetNumberFormat(nfidx);
      if nfParams <> nil then
      begin
        nfs := nfParams.NumFormatStr;
        for j:=0 to NumFormatList.Count-1 do
        begin
          s := NumFormatList[j];
          p := pos(':', s);
          if SameText(Copy(s, p+1, Length(s)), nfs) then
          begin
            nfs := Format('style:data-style-name="%s"', [copy(s, 1, p-1)]);
            break;
          end;
          p := 0;
        end;
        if p = 0 then // not found
          nfs := '';
      end;
    end;

    // Start and name
    AppendToStream(AStream,
      '<style:style style:name="ce' + IntToStr(i) + '" style:family="table-cell" ' +
                   'style:parent-style-name="Default" '+ nfs + '>');

    // style:text-properties
    s := WriteFontStyleXMLAsString(fmt);
    if s <> '' then
      AppendToStream(AStream,
        '<style:text-properties '+ s + '/>');

    s := WriteBorderStyleXMLAsString(fmt) +
         WriteBackgroundColorStyleXMLAsString(fmt) +
         WriteWordwrapStyleXMLAsString(fmt) +
         WriteTextRotationStyleXMLAsString(fmt) +
         WriteVertAlignmentStyleXMLAsString(fmt);
    if s <> '' then
      AppendToStream(AStream,
        '<style:table-cell-properties ' + s + '/>');

    // style:paragraph-properties
    s := WriteHorAlignmentStyleXMLAsString(fmt);
    if s <> '' then
      AppendToStream(AStream,
        '<style:paragraph-properties ' + s + '/>');

    // End
    AppendToStream(AStream,
      '</style:style>');
  end;
end;

procedure TsSpreadOpenDocWriter.WriteColStyles(AStream: TStream);
var
  i: Integer;
  colstyle: TColumnStyleData;
begin
  if FColumnStyleList.Count = 0 then
  begin
    AppendToStream(AStream,
      '<style:style style:name="co1" style:family="table-column">',
        '<style:table-column-properties fo:break-before="auto" style:column-width="2.267cm"/>',
      '</style:style>');
    exit;
  end;

  for i := 0 to FColumnStyleList.Count-1 do
  begin
    colStyle := TColumnStyleData(FColumnStyleList[i]);

    // Start and Name
    AppendToStream(AStream, Format(
      '<style:style style:name="%s" style:family="table-column">', [colStyle.Name]));

    // Column width
    AppendToStream(AStream, Format(
        '<style:table-column-properties style:column-width="%.3fmm" fo:break-before="auto"/>',
          [colStyle.ColWidth], FPointSeparatorSettings));

    // End
    AppendToStream(AStream,
      '</style:style>');
  end;
end;

procedure TsSpreadOpenDocWriter.WriteColumns(AStream: TStream;
  ASheet: TsWorksheet);
var
  lastCol: Integer;
  j, k: Integer;
  w, w_mm: Double;
  widthMultiplier: Double;
  styleName: String;
  colsRepeated: Integer;
  colsRepeatedStr: String;
begin
  widthMultiplier := Workbook.GetFont(0).Size / 2;
  lastCol := ASheet.GetLastColIndex;

  j := 0;
  while (j <= lastCol) do
  begin
    w := ASheet.GetColWidth(j);
    // Convert to mm
    w_mm := PtsToMM(w * widthMultiplier);

    // Find width in ColumnStyleList to retrieve corresponding style name
    styleName := '';
    for k := 0 to FColumnStyleList.Count-1 do
      if SameValue(TColumnStyleData(FColumnStyleList[k]).ColWidth, w_mm, COLWIDTH_EPS) then begin
        styleName := TColumnStyleData(FColumnStyleList[k]).Name;
        break;
      end;
    if stylename = '' then
      raise Exception.Create(rsColumnStyleNotFound);

    // Determine value for "number-columns-repeated"
    colsRepeated := 1;
    k := j+1;
    while (k <= lastCol) do
    begin
      if ASheet.GetColWidth(k) = w then
        inc(colsRepeated)
      else
        break;
      inc(k);
    end;
    colsRepeatedStr := IfThen(colsRepeated = 1, '', Format(' table:number-columns-repeated="%d"', [colsRepeated]));

    AppendToStream(AStream, Format(
      '<table:table-column table:style-name="%s"%s table:default-cell-style-name="Default" />',
        [styleName, colsRepeatedStr]));

    j := j + colsRepeated;
  end;
end;

function TsSpreadOpenDocWriter.WriteCommentXMLAsString(AComment: String): String;
var
  L: TStringList;
  s: String;
  err: Boolean;
  i: Integer;
begin
  Result := '';
  if AComment = '' then exit;

  result := '<office:annotation office:display="false">';
  err := false;
  L := TStringList.Create;
  try
    L.Text := AComment;
    for i:=0 to L.Count-1 do begin
      s := L[i];
      if not ValidXMLText(s) then begin
        if not err then
          Workbook.AddErrorMsg(rsInvalidCharacterInCellComment, [AComment]);
        err := true;
      end;
      Result := Result + '<text:p>' + s + '</text:p>';
    end;
  finally
    L.Free;
  end;

  Result := Result + '</office:annotation>';
end;

{@@ ----------------------------------------------------------------------------
  Writes the declaration of the font faces used in the workbook.
  Is used in styles.xml and content.xml.

  Procedure must be enclosed by
    <office:font-face-decls> ... </office:font-face-decls>
-------------------------------------------------------------------------------}
procedure TsSpreadOpenDocWriter.WriteFontNames(AStream: TStream);
var
  L: TStringList;
  fnt: TsFont;
  hfFnt: TXMLHeaderFooterFont;
  i: Integer;
begin
  // Collect all unique font names in a string list
  L := TStringList.Create;
  try
    // First collect the font names from the workbook's FontList
    for i:=0 to Workbook.GetFontCount-1 do
    begin
      fnt := Workbook.GetFont(i);
      if (fnt <> nil) and (L.IndexOf(fnt.FontName) = -1) then
        L.Add(fnt.FontName);
    end;
    // Then collect the header/footer font names from the HeaderFooterFontList
    for i:=0 to FHeaderFooterFontList.Count-1 do
    begin
      hfFnt := TXMLHeaderFooterFont(FHeaderFooterFontList[i]);
      if (hfFnt <> nil) and (L.Indexof(hfFnt.FontName) = -1) then
        L.Add(hfFnt.FontName);
    end;
    // Done. Now write all font names as xml nodes to the stream
    for i:=0 to L.Count-1 do
      AppendToStream(AStream, Format(
        '<style:font-face style:name="%s" svg:font-family="%s" />', [L[i], L[i]]));
  finally
    L.Free;
  end;
end;

procedure TsSpreadOpenDocWriter.WriteMasterStyles(AStream: TStream);
var
  defFnt: TsHeaderFooterFont;
  i: Integer;
  sheet: TsWorksheet;

  function HeaderFooterAsString(AIndex: Integer; AIsHeader: Boolean;
    const APageLayout: TsPageLayout): String;
  var
    parser: TsSpreadOpenDocHeaderFooterParser;
    str: String;
  begin
    if AIsHeader then
      str := APageLayout.Headers[AIndex] else
      str := APageLayout.Footers[AIndex];
    if str = '' then
      exit;
    parser := TsSpreadOpenDocHeaderFooterParser.Create(str, FHeaderFooterFontList,
      defFnt);
    try
      Result := parser.BuildHeaderFooterAsXMLString;
    finally
      parser.Free;
    end;
  end;

  function MasterPageAsString(AStyleName, ADisplayName, APageLayoutName: String;
    const APageLayout: TsPageLayout): String;
  const
    IS_HEADER = true;
    IS_FOOTER = false;
  begin
    Result := Format(
      '<style:master-page style:name="%s" ' +
                        'style:display-name="%s" ' +
                        'style:page-layout-name="%s">', [
      AStyleName, ADisplayName, APageLayoutName
    ]);

    Result := Result +
        '<style:header>' +
          HeaderFooterAsString(1, IS_HEADER, APageLayout) +
        '</style:header>' +
        '<style:footer>' +
          HeaderFooterAsString(1, IS_FOOTER, APageLayout) +
        '</style:footer>';

    if poDifferentOddEven in APageLayout.Options then
      Result := Result +
        '<style:header-left>' +
          HeaderFooterAsString(2, IS_HEADER, APageLayout) +
        '</style:header-left>' +
        '<style:footer-left>' +
          HeaderFooterAsString(2, IS_FOOTER, APageLayout) +
        '</style:footer-left>';

    Result := Result + '</style:master-page>';
  end;

begin
  defFnt := TsHeaderFooterFont.Create(Workbook.GetDefaultFont);

  AppendToStream(AStream,
    '<office:master-styles>');

  AppendToStream(AStream,
      '<style:master-page style:name="Default" style:page-layout-name="Mpm1">' +
        '<style:header />' +
        '<style:header-left style:display="false" />' +
        '<style:footer />' +
        '<style:footer-left style:display="false" />' +
      '</style:master-page>');

  for i:=0 to FWorkbook.GetWorksheetCount-1 do begin
    sheet := FWorkbook.GetWorksheetByIndex(i);
    AppendToStream(AStream,
      MasterPageAsString('PageStyle_5f_' + sheet.Name, 'PageStyle_' + sheet.Name,
        'Mpm' + IntToStr(3+i), sheet.PageLayout));
  end;

  AppendToStream(AStream,
    '</office:master-styles>');

  defFnt.Free;
end;

procedure TsSpreadOpenDocWriter.WriteNumFormats(AStream: TStream);
var
  i, p: Integer;
  numFmtXML: String;
  numFmtStr: String;
  numFmtName: String;
  parser: TsSpreadOpenDocNumFormatParser;
begin
  for i:=0 to NumFormatList.Count-1 do
  begin
    numFmtStr := NumFormatList[i];
    p := pos(':', numFmtStr);
    numFmtName := Copy(numFmtStr, 1, p-1);
    numFmtStr := Copy(numFmtStr, p+1, Length(numFmtStr));
    parser := TsSpreadOpenDocNumFormatParser.Create(Workbook, numFmtStr);
    try
      numFmtXML := parser.BuildXMLAsString(numFmtName);
      if numFmtXML <> '' then
        AppendToStream(AStream, numFmtXML);
    finally
      parser.Free;
    end;
  end;
end;

procedure TsSpreadOpenDocWriter.WriteRowsAndCells(AStream: TStream; ASheet: TsWorksheet);
var
  r, rr: Cardinal;  // row index in sheet
  c, cc: Cardinal;  // column index in sheet
  row: PRow;        // sheet row record
  cell: PCell;      // current cell
  styleName: String;
  k: Integer;
  h, h_mm: Single;  // row height in "lines" and millimeters, respectively
  h1: Single;
  colsRepeated: Cardinal;
  rowsRepeated: Cardinal;
  colsRepeatedStr: String;
  rowsRepeatedStr: String;
  firstCol, firstRow, lastCol, lastRow: Cardinal;
  rowStyleData: TRowStyleData;
  defFontSize: Single;
  emptyRowsAbove: Boolean;
begin
  // some abbreviations...
  defFontSize := Workbook.GetFont(0).Size;
  GetSheetDimensions(ASheet, firstRow, lastRow, firstCol, lastCol);
  emptyRowsAbove := firstRow > 0;

  // Now loop through all rows
  r := firstRow;
  while (r <= lastRow) do
  begin
    rowsRepeated := 1;
    // Look for the row style of the current row (r)
    row := ASheet.FindRow(r);
    if row = nil then
      styleName := 'ro1'
    else
    begin
      styleName := '';

      h := row^.Height;   // row height in "lines"
      h_mm := PtsToMM((h + ROW_HEIGHT_CORRECTION) * defFontSize);  // in mm
      for k := 0 to FRowStyleList.Count-1 do begin
        rowStyleData := TRowStyleData(FRowStyleList[k]);
        // Compare row heights, but be aware of rounding errors
        if SameValue(rowStyleData.RowHeight, h_mm, 1E-3) then
        begin
          styleName := rowStyleData.Name;
          break;
        end;
      end;
      if styleName = '' then
        raise Exception.Create(rsRowStyleNotFound);
    end;

    // Take care of empty rows above the first row
    if (r = firstRow) and emptyRowsAbove then
    begin
      rowsRepeated := r;
      rowsRepeatedStr := IfThen(rowsRepeated = 1, '',
        Format('table:number-rows-repeated="%d"', [rowsRepeated]));
      colsRepeated := lastCol + 1;
      colsRepeatedStr := IfThen(colsRepeated = 1, '',
        Format('table:number-columns-repeated="%d"', [colsRepeated]));
      AppendToStream(AStream, Format(
        '<table:table-row table:style-name="%s" %s>' +
          '<table:table-cell %s/>' +
        '</table:table-row>',
        [styleName, rowsRepeatedStr, colsRepeatedStr]));
      rowsRepeated := 1;
    end
    else
    // Look for empty rows with the same style, they need the "number-rows-repeated" element.
    if (ASheet.Cells.GetFirstCellOfRow(r) = nil) then
    begin
      rr := r + 1;
      while (rr <= lastRow) do
      begin
        if ASheet.Cells.GetFirstCellOfRow(rr) <> nil then
          break;
        h1 := ASheet.GetRowHeight(rr);
        if not SameValue(h, h1, ROWHEIGHT_EPS) then
          break;
        inc(rr);
      end;
      rowsRepeated := rr - r;
      rowsRepeatedStr := IfThen(rowsRepeated = 1, '',
        Format('table:number-rows-repeated="%d"', [rowsRepeated]));
      colsRepeated := lastCol - firstCol + 1;
      colsRepeatedStr := IfThen(colsRepeated = 1, '',
        Format('table:number-columns-repeated="%d"', [colsRepeated]));

      AppendToStream(AStream, Format(
        '<table:table-row table:style-name="%s" %s>' +
          '<table:table-cell %s/>' +
        '</table:table-row>',
        [styleName, rowsRepeatedStr, colsRepeatedStr]));

      r := rr;
      continue;
    end;

    // Now we know that there are cells.
    // Write the row XML
    AppendToStream(AStream, Format(
        '<table:table-row table:style-name="%s">', [styleName]));

    // Loop along the row and find the cells.
    c := 0;
    while c <= lastCol do
    begin
      // Get the cell from the sheet
      cell := ASheet.FindCell(r, c);

      // Belongs to merged block?
      if (cell <> nil) and not FWorksheet.IsMergeBase(cell) and FWorksheet.IsMerged(cell) then
      // this means: all cells of a merged block except for the merge base
      begin
        AppendToStream(AStream,
          '<table:covered-table-cell />');
        inc(c);
        continue;
      end;

      // Empty cell? Need to count how many to add "table:number-columns-repeated"
      colsRepeated := 1;
      if cell = nil then
      begin
        cc := c + 1;
        while (cc <= lastCol) do
        begin
          cell := ASheet.FindCell(r, cc);
          if cell <> nil then
            break;
          inc(cc)
        end;
        colsRepeated := cc - c;
        colsRepeatedStr := IfThen(colsRepeated = 1, '',
          Format('table:number-columns-repeated="%d"', [colsRepeated]));
        AppendToStream(AStream, Format(
          '<table:table-cell %s/>', [colsRepeatedStr]));
      end else
        WriteCellToStream(AStream, cell);
      inc(c, colsRepeated);
    end;

    AppendToStream(AStream,
        '</table:table-row>');

    // Next row
    inc(r, rowsRepeated);
  end;
end;

procedure TsSpreadOpenDocWriter.WriteRowStyles(AStream: TStream);
var
  i: Integer;
  rowstyle: TRowStyleData;
begin
  if FRowStyleList.Count = 0 then
  begin
    AppendToStream(AStream,
      '<style:style style:name="ro1" style:family="table-row">' +
        '<style:table-row-properties style:row-height="0.416cm" fo:break-before="auto" style:use-optimal-row-height="true"/>' +
      '</style:style>');
    exit;
  end;

  for i := 0 to FRowStyleList.Count-1 do
  begin
    rowStyle := TRowStyleData(FRowStyleList[i]);

    // Start and Name
    AppendToStream(AStream, Format(
      '<style:style style:name="%s" style:family="table-row">', [rowStyle.Name]));

    // Column width
    AppendToStream(AStream, Format(
      '<style:table-row-properties style:row-height="%.3fmm" ', [rowStyle.RowHeight], FPointSeparatorSettings));
    if rowStyle.AutoRowHeight then
      AppendToStream(AStream, 'style:use-optimal-row-height="true" ');
    AppendToStream(AStream, 'fo:break-before="auto"/>');

    // End
    AppendToStream(AStream,
      '</style:style>');
  end;
end;


constructor TsSpreadOpenDocWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);

  FColumnStyleList := TFPList.Create;
  FRowStyleList := TFPList.Create;
  FHeaderFooterFontList := TObjectList.Create;

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';
  FPointSeparatorSettings.ListSeparator := ';';   // for formulas

  // http://en.wikipedia.org/wiki/List_of_spreadsheet_software#Specifications
  FLimitations.MaxColCount := 1024;
  FLimitations.MaxRowCount := 1048576;
end;

destructor TsSpreadOpenDocWriter.Destroy;
var
  j: Integer;
begin
  for j:=FColumnStyleList.Count-1 downto 0 do TObject(FColumnStyleList[j]).Free;
  FColumnStyleList.Free;

  for j:=FRowStyleList.Count-1 downto 0 do TObject(FRowStyleList[j]).Free;
  FRowStyleList.Free;

  FHeaderFooterFontList.Free;

  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Writes a string to a file. Helper convenience method.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Writes an OOXML document to a file.
-------------------------------------------------------------------------------}
procedure TsSpreadOpenDocWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean);
var
  lStream: TStream;
  lMode: word;
begin
  if AOverwriteExisting
    then lMode := fmCreate or fmOpenWrite
    else lMode := fmCreate;

  if (boBufStream in Workbook.Options) then
    lStream := TBufStream.Create(AFileName, lMode)
  else
    lStream := TFileStream.Create(AFileName, lMode);

  try
    WriteToStream(lStream);
  finally
    FreeAndNil(lStream);
  end;
end;

procedure TsSpreadOpenDocWriter.WriteToStream(AStream: TStream);
var
  FZip: TZipper;
begin
  { Analyze the workbook and collect all information needed }
  ListAllNumFormats;
  ListAllColumnStyles;
  ListAllRowStyles;
  ListAllHeaderFooterFonts;

  { Create the streams that will hold the file contents }
  CreateStreams;

  { Fill the strings with the contents of the files }
  WriteMimetype();
  WriteMetaInfManifest();
  WriteMeta();
  WriteSettings();
  WriteStyles();
  WriteContent;

  { Now compress the files }
  FZip := TZipper.Create;
  try
    FZip.FileName := '__temp__.tmp';

    FZip.Entries.AddFileEntry(FSMeta, OPENDOC_PATH_META);
    FZip.Entries.AddFileEntry(FSSettings, OPENDOC_PATH_SETTINGS);
    FZip.Entries.AddFileEntry(FSStyles, OPENDOC_PATH_STYLES);
    FZip.Entries.AddFileEntry(FSContent, OPENDOC_PATH_CONTENT);
    FZip.Entries.AddFileEntry(FSMimetype, OPENDOC_PATH_MIMETYPE);
    FZip.Entries.AddFileEntry(FSMetaInfManifest, OPENDOC_PATH_METAINF_MANIFEST);

    ResetStreams;

    FZip.SaveToStream(AStream);

  finally
    DestroyStreams;
    FZip.Free;
  end;
end;

{ Writes an empty cell to the stream }
procedure TsSpreadOpenDocWriter.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  colsSpannedStr: String;
  rowsSpannedStr: String;
  spannedStr: String;
  comment: String;
  r1,c1,r2,c2: Cardinal;
  fmt: TsCellFormat;
begin
  Unused(ARow, ACol);

  // Hyperlink
  if FWorksheet.HasHyperlink(ACell) then
    FWorkbook.AddErrorMsg(rsODSHyperlinksOfTextCellsOnly, [GetCellString(ARow, ACol)]);

  // Comment
  comment := WriteCommentXMLAsString(FWorksheet.ReadComment(ACell));

  // Merged?
  if FWorksheet.IsMergeBase(ACell) then
  begin
    FWorksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    rowsSpannedStr := Format('table:number-rows-spanned="%d"', [r2 - r1 + 1]);
    colsSpannedStr := Format('table:number-columns-spanned="%d"', [c2 - c1 + 1]);
    spannedStr := colsSpannedStr + ' ' + rowsSpannedStr;
  end else
    spannedStr := '';

  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  if fmt.UsedFormattingFields <> [] then
    AppendToStream(AStream, Format(
      '<table:table-cell table:style-name="ce%d" %s>', [ACell^.FormatIndex, spannedStr]),
      comment,
      '</table:table-cell>')
  else
  if comment <> '' then
    AppendToStream(AStream,
      '<table:table-cell ' + spannedStr + '>' + comment + '</table:table-cell>')
  else
    AppendToStream(AStream,
      '<table:table-cell ' + spannedStr + '/>');
end;

{@@ ----------------------------------------------------------------------------
  Writes a boolean cell to the stream
-------------------------------------------------------------------------------}
procedure TsSpreadOpenDocWriter.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: Boolean; ACell: PCell);
var
  lStyle, valType: String;
  r1,c1,r2,c2: Cardinal;
  rowsSpannedStr, colsSpannedStr, spannedStr: String;
  comment: String;
  strValue: String;
  displayStr: String;
  fmt: TsCellFormat;
begin
  Unused(ARow, ACol);

  valType := 'boolean';

  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  if fmt.UsedFormattingFields <> [] then
    lStyle := ' table:style-name="ce' + IntToStr(ACell^.FormatIndex) + '" '
  else
    lStyle := '';

  // Comment
  comment := WriteCommentXMLAsString(FWorksheet.ReadComment(ACell));

  // Merged?
  if FWorksheet.IsMergeBase(ACell) then
  begin
    FWorksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    rowsSpannedStr := Format('table:number-rows-spanned="%d"', [r2 - r1 + 1]);
    colsSpannedStr := Format('table:number-columns-spanned="%d"', [c2 - c1 + 1]);
    spannedStr := colsSpannedStr + ' ' + rowsSpannedStr;
  end else
    spannedStr := '';

  // Displayed value
  if AValue then
  begin
    StrValue := 'true';
    DisplayStr := rsTRUE;
  end else
  begin
    strValue := 'false';
    DisplayStr := rsFALSE;
  end;

  // Hyperlink
  if FWorksheet.HasHyperlink(ACell) then
    FWorkbook.AddErrorMsg(rsODSHyperlinksOfTextCellsOnly, [GetCellString(ARow, ACol)]);

  AppendToStream(AStream, Format(
    '<table:table-cell office:value-type="%s" office:boolean-value="%s" %s %s >' +
      comment +
      '<text:p>%s</text:p>' +
    '</table:table-cell>', [
    valType, StrValue, lStyle, spannedStr,
    DisplayStr
  ]));
end;

{@@ ----------------------------------------------------------------------------
  Creates an XML string for inclusion of the background color into the
  written file from the backgroundcolor setting in the given format record.
  Is called from WriteStyles (via WriteStylesXMLAsString).

  NOTE: ODS does not support fill patterns. Fill patterns are converted to
  solid fills by mixing pattern and background colors in the ratio defined
  by the fill pattern. Result agrees with that what LO/OO show for an imported
  xls file.
-------------------------------------------------------------------------------}
function TsSpreadOpenDocWriter.WriteBackgroundColorStyleXMLAsString(
  const AFormat: TsCellFormat): String;
type
  TRgb = record r,g,b,a: byte; end;
const  // fraction of pattern color in fill pattern
  FRACTION: array[TsFillStyle] of Double = (
    0.0, 1.0, 0.75, 0.50, 0.25, 0.125, 0.0625,  // fsNoFill..fsGray6
    0.5, 0.5, 0.5, 0.5,                         // fsStripeHor..fsStripeDiagDown
    0.25, 0.25, 0.25, 0.25,                     // fsThinStripeHor..fsThinStripeDiagDown
    0.5, 6.0/16, 0.75, 7.0/16);                 // fsHatchDiag..fsThinHatchHor
var
  fc,bc: TsColorValue;
  mix: TRgb;
  fraction_fc, fraction_bc: Double;
begin
  Result := '';

  if not (uffBackground in AFormat.UsedFormattingFields) then
    exit;

  // Foreground and background colors
  fc := Workbook.GetPaletteColor(AFormat.Background.FgColor);
  if Aformat.Background.BgColor = scTransparent then
    bc := Workbook.GetPaletteColor(scWhite)
  else
    bc := Workbook.GetPaletteColor(AFormat.Background.BgColor);
  // Mixing fraction
  fraction_fc := FRACTION[AFormat.Background.Style];
  fraction_bc := 1.0 - fraction_fc;
  // Mixed color
  mix.r := Min(round(fraction_fc*TRgb(fc).r + fraction_bc*TRgb(bc).r), 255);
  mix.g := Min(round(fraction_fc*TRgb(fc).g + fraction_bc*TRgb(bc).g), 255);
  mix.b := Min(round(fraction_fc*TRgb(fc).b + fraction_bc*TRgb(bc).b), 255);

  Result := Format('fo:background-color="%s" ', [
    ColorToHTMLColorStr(TsColorValue(mix))
  ]);
end;

{@@ ----------------------------------------------------------------------------
  Creates an XML string for inclusion of borders and border styles into the
  written file from the border settings in the given format record.
  Is called from WriteStyles (via WriteStylesXMLAsString).
-------------------------------------------------------------------------------}
function TsSpreadOpenDocWriter.WriteBorderStyleXMLAsString(
  const AFormat: TsCellFormat): String;
begin
  Result := '';

  if not (uffBorder in AFormat.UsedFormattingFields) then
    exit;

  if cbSouth in AFormat.Border then
  begin
    Result := Result + Format('fo:border-bottom="%s %s %s" ', [
      BORDER_LINEWIDTHS[AFormat.BorderStyles[cbSouth].LineStyle],
      BORDER_LINESTYLES[AFormat.BorderStyles[cbSouth].LineStyle],
      Workbook.GetPaletteColorAsHTMLStr(AFormat.BorderStyles[cbSouth].Color)
    ]);
    if AFormat.BorderStyles[cbSouth].LineStyle = lsDouble then
      Result := Result + 'style:border-linewidth-bottom="0.002cm 0.035cm 0.002cm" ';
  end
  else
    Result := Result + 'fo:border-bottom="none" ';

  if cbWest in AFormat.Border then
  begin
    Result := Result + Format('fo:border-left="%s %s %s" ', [
      BORDER_LINEWIDTHS[AFormat.BorderStyles[cbWest].LineStyle],
      BORDER_LINESTYLES[AFormat.BorderStyles[cbWest].LineStyle],
      Workbook.GetPaletteColorAsHTMLStr(AFormat.BorderStyles[cbWest].Color)
    ]);
    if AFormat.BorderStyles[cbWest].LineStyle = lsDouble then
      Result := Result + 'style:border-linewidth-left="0.002cm 0.035cm 0.002cm" ';
  end
  else
    Result := Result + 'fo:border-left="none" ';

  if cbEast in AFormat.Border then
  begin
    Result := Result + Format('fo:border-right="%s %s %s" ', [
      BORDER_LINEWIDTHS[AFormat.BorderStyles[cbEast].LineStyle],
      BORDER_LINESTYLES[AFormat.BorderStyles[cbEast].LineStyle],
      Workbook.GetPaletteColorAsHTMLStr(AFormat.BorderStyles[cbEast].Color)
    ]);
    if AFormat.BorderStyles[cbSouth].LineStyle = lsDouble then
      Result := Result + 'style:border-linewidth-right="0.002cm 0.035cm 0.002cm" ';
  end
  else
    Result := Result + 'fo:border-right="none" ';

  if cbNorth in AFormat.Border then
  begin
    Result := Result + Format('fo:border-top="%s %s %s" ', [
      BORDER_LINEWIDTHS[AFormat.BorderStyles[cbNorth].LineStyle],
      BORDER_LINESTYLES[AFormat.BorderStyles[cbNorth].LineStyle],
      Workbook.GetPaletteColorAsHTMLStr(AFormat.BorderStyles[cbNorth].Color)
    ]);
    if AFormat.BorderStyles[cbSouth].LineStyle = lsDouble then
      Result := Result + 'style:border-linewidth-top="0.002cm 0.035cm 0.002cm" ';
  end else
    Result := Result + 'fo:border-top="none" ';

  if cbDiagUp in AFormat.Border then
  begin
    Result := Result + Format('style:diagonal-bl-tr="%s %s %s" ', [
      BORDER_LINEWIDTHS[AFormat.BorderStyles[cbDiagUp].LineStyle],
      BORDER_LINESTYLES[AFormat.BorderStyles[cbDiagUp].LineStyle],
      Workbook.GetPaletteColorAsHTMLStr(AFormat.BorderStyles[cbDiagUp].Color)
    ]);
  end;

  if cbDiagDown in AFormat.Border then
  begin
    Result := Result + Format('style:diagonal-tl-br="%s %s %s" ', [
      BORDER_LINEWIDTHS[AFormat.BorderStyles[cbDiagDown].LineStyle],
      BORDER_LINESTYLES[AFormat.BorderStyles[cbDiagDown].LineStyle],
      Workbook.GetPaletteColorAsHTMLStr(AFormat.BorderStyles[cbDiagDown].Color)
    ]);
  end;
end;

function TsSpreadOpenDocWriter.WriteDefaultFontXMLAsString: String;
var
  fnt: TsFont;
begin
  fnt := Workbook.GetDefaultFont;
  Result := Format(
    '<style:text-properties style:font-name="%s" fo:font-size="%.1fpt" />',
    [fnt.FontName, fnt.Size], FPointSeparatorSettings
  );
end;

procedure TsSpreadOpenDocWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol);
  Unused(AValue, ACell);
  // ??
end;

function TsSpreadOpenDocWriter.WriteFontStyleXMLAsString(AFont: TsFont): String;
var
  defFnt: TsFont;
begin
  Result := '';

  defFnt := Workbook.GetDefaultFont;
  if AFont = nil then AFont := defFnt;

  if AFont.FontName <> defFnt.FontName then
    Result := Result + Format('style:font-name="%s" ', [AFont.FontName]);

  if AFont.Size <> defFnt.Size then
    Result := Result + Format('fo:font-size="%.1fpt" style:font-size-asian="%.1fpt" style:font-size-complex="%.1fpt" ',
      [AFont.Size, AFont.Size, AFont.Size], FPointSeparatorSettings);

  if fssBold in AFont.Style then
    Result := Result + 'fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold" ';

  if fssItalic in AFont.Style then
    Result := Result + 'fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic" ';

  if fssUnderline in AFont.Style then
    Result := Result + 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" ';

  if fssStrikeout in AFont.Style then
    Result := Result + 'style:text-line-through-style="solid" ';

  if AFont.Color <> defFnt.Color then
    Result := Result + Format('fo:color="%s" ', [Workbook.GetPaletteColorAsHTMLStr(AFont.Color)]);
end;

function TsSpreadOpenDocWriter.WriteFontStyleXMLAsString(
  const AFormat: TsCellFormat): String;
begin
  Result := '';
  if (uffFont in AFormat.UsedFormattingFields) then
    Result := WriteFontStyleXMLAsString(Workbook.GetFont(AFormat.FontIndex));
end;

function TsSpreadOpenDocWriter.WriteHeaderFooterFontXMLAsString(
  AFont: TsHeaderFooterFont): String;
begin
  Result := Format('style:font-name="%s" fo:font-size="%dpt" ', [
    AFont.FontName, round(AFont.Size)
  ]);

  if hfsBold in AFont.Style then
    Result := Result + 'fo:font-weight="bold" ';

  if hfsItalic in AFont.Style then
    Result := Result + 'fo:font-style="italic" ';

  if hfsUnderline in AFont.Style then
    Result := Result + 'style:text-underline-style="solid" '+
                       'style:text-underline-width="auto" '+
                       'style:text-underline-color="font-color" ';

  if hfsDblUnderline in AFont.Style then
    Result := Result + 'style:text-underline-style="solid" '+
                       'style:text-underline-type="double" ' +
                       'style:text-underline-width="auto" '+
                       'style:text-underline-color="font-color" ';

  if hfsStrikeout in AFont.Style then
    Result := Result + 'style:text-line-through-style="solid" ';

  if hfsOutline in AFont.Style then
    Result := Result + 'style:text-outline="true" ';

  if hfsShadow in AFont.Style then
    Result := Result + 'fo:text-shadow="1pt 1pt" ' +
                       'style:text-outline="none" ';

  if hfsSubscript in AFont.Style then
    Result := Result + 'style:text-position="sub 58%" ';

  if hfsSuperscript in AFont.Style then
    Result := Result + 'style:text-position="super 58%" ';

  if AFont.Color <> 0 then
    Result := Result + Format('fo:color="%s" ', [ColorToHTMLColorStr(AFont.Color)]);
end;


{@@ ----------------------------------------------------------------------------
  Creates an XML string for inclusion of the horizontal alignment into the
  written file from the horizontal alignment setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString).
-------------------------------------------------------------------------------}
function TsSpreadOpenDocWriter.WriteHorAlignmentStyleXMLAsString(
  const AFormat: TsCellFormat): String;
begin
  Result := '';
  if not (uffHorAlign in AFormat.UsedFormattingFields) then
    exit;
  case AFormat.HorAlignment of
    haLeft   : Result := 'fo:text-align="start" ';
    haCenter : Result := 'fo:text-align="center" ';
    haRight  : Result := 'fo:text-align="end" ';
  end;
end;

function TsSpreadOpenDocWriter.WritePageLayoutAsXMLString(AStyleName: String;
  const APageLayout: TsPageLayout): String;
var
  pageLayoutStr: String;
  headerStyleStr: String;
  footerStyleStr: String;
  options: String;
  i: Integer;
  hasHeader, hasFooter: Boolean;
  topmargin, bottommargin: Double;
  h: Double;
begin
  hasHeader := false;
  hasFooter := false;
  for i:=0 to High(APageLayout.Headers) do
  begin
    if APageLayout.Headers[i] <> '' then hasHeader := true;
    if APageLayout.Footers[i] <> '' then hasFooter := true;
  end;

  if hasHeader then
    topMargin := APageLayout.HeaderMargin
  else
    topMargin := APageLayout.TopMargin;

  if hasFooter then
    bottomMargin := APageLayout.FooterMargin
  else
    bottomMargin := APageLayout.BottomMargin;

  pageLayoutStr := Format(
      'fo:page-width="%.2fmm" fo:page-height="%.2fmm" '+
      'fo:margin-top="%.2fmm" fo:margin-bottom="%.2fmm" '+
      'fo:margin-left="%.2fmm" fo:margin-right="%.2fmm" ', [
      APageLayout.PageWidth, APageLayout.PageHeight,
      topmargin, bottommargin,
      APageLayout.LeftMargin, APageLayout.RightMargin
    ], FPointSeparatorSettings);

  if APageLayout.Orientation = spoLandscape then
    pageLayoutStr := pageLayoutStr + 'style:print-orientation="landscape" ';

  if poPrintPagesByRows in APageLayout.Options then
    pageLayoutStr := pageLayoutStr + 'style:print-page-order="ltr" ';

  if poUseStartPageNumber in APageLayout.Options then
    pageLayoutStr := pageLayoutStr + 'style:first-page-number="' + IntToStr(APageLayout.StartPageNumber) +'" '
  else
    pageLayoutStr := pageLayoutStr + 'style:first-page-number="continue" ';

  if APageLayout.Options * [poHorCentered, poVertCentered] = [poHorCentered, poVertCentered] then
    pageLayoutStr := pageLayoutStr + 'style:table-centering="both" '
  else if poHorCentered in APageLayout.Options then
    pageLayoutStr := pageLayoutStr + 'style:table-centering="horizontal" '
  else if poVertCentered in APageLayout.Options then
    pageLayoutStr := pageLayoutStr + 'style:table-centering="vertical" ';

  if poFitPages in APageLayout.Options then
  begin
    if APageLayout.FitWidthToPages > 0 then
      pageLayoutStr := pageLayoutStr + 'style:scale-to-X="' + IntToStr(APageLayout.FitWidthToPages) + '" ';
    if APageLayout.FitHeightToPages > 0 then
      pageLayoutStr := pageLayoutStr + 'style:scale-to-Y="' + IntToStr(APageLayout.FitHeightToPages) + '" ';
  end else
    pageLayoutStr := pageLayoutStr + 'style:scale-to="' + IntToStr(APageLayout.ScalingFactor) + '%" ';

  options := 'charts drawings objects zero-values';
  if poPrintGridLines in APageLayout.Options then
    options := options + ' grid';
  if poPrintHeaders in APageLayout.Options then
    options := options + ' headers';
  if poPrintCellComments in APageLayout.Options then
    options := options + ' annotations';

  pageLayoutStr := pageLayoutStr + 'style:print="' + options + '" ';

  h := PtsToMM(FWorkbook.GetDefaultFontSize);

  if hasHeader then
    headerStyleStr := Format(
      '<style:header-style>'+
        '<style:header-footer-properties ' +
          'fo:margin-left="0mm" fo:margin-right="0mm" '+
          'fo:min-height="%.2fmm" fo:margin-bottom="%.2fmm" '+
        '/>'+
      '</style:header-style>', [
      APageLayout.TopMargin - APageLayout.HeaderMargin,
      APageLayout.TopMargin - APageLayout.HeaderMargin - h], FPointSeparatorSettings)
  else
    headerStyleStr := '';

  if hasFooter then
    footerStyleStr := Format(
      '<style:footer-style>'+
        '<style:header-footer-properties ' +
          'fo:margin-left="0mm" fo:margin-right="0mm" '+
          'fo:min-height="%.2fmm" fo:margin-top="%.2fmm" '+
        '/>'+
      '</style:footer-style>', [
      APageLayout.BottomMargin - APageLayout.FooterMargin,
      APageLayout.BottomMargin - APageLayout.FooterMargin - h], FPointSeparatorSettings)
  else
    footerStyleStr := '';

  Result := '<style:page-layout style:name="' + AStyleName + '">' +
              '<style:page-layout-properties ' + pageLayoutStr + '/>'+
              headerStyleStr +
              footerStyleStr +
            '</style:page-layout>';
end;

procedure TsSpreadOpenDocWriter.WriteTableSettings(AStream: TStream);
var
  i: Integer;
  sheet: TsWorkSheet;
  hsm: Integer;  // HorizontalSplitMode
  vsm: Integer;  // VerticalSplitMode
  asr: Integer;  // ActiveSplitRange
begin
  for i:=0 to Workbook.GetWorksheetCount-1 do
  begin
    sheet := Workbook.GetWorksheetByIndex(i);

    AppendToStream(AStream,
      '<config:config-item-map-entry config:name="' + sheet.Name + '">');

    hsm := 0; vsm := 0; asr := 2;
    if (soHasFrozenPanes in sheet.Options) then
    begin
      if (sheet.LeftPaneWidth > 0) and (sheet.TopPaneHeight > 0) then
      begin
        hsm := 2; vsm := 2; asr := 3;
      end else
      if (sheet.LeftPaneWidth > 0) then
      begin
        hsm := 2; vsm := 0; asr := 3;
      end else if (sheet.TopPaneHeight > 0) then
      begin
        hsm := 0; vsm := 2; asr := 2;
      end;
    end;
    {showGrid := (soShowGridLines in sheet.Options);}

    AppendToStream(AStream,
        '<config:config-item config:name="CursorPositionX" config:type="int">'+IntToStr(sheet.LeftPaneWidth)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="CursorPositionY" config:type="int">'+IntToStr(sheet.TopPaneHeight)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="HorizontalSplitMode" config:type="short">'+IntToStr(hsm)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="VerticalSplitMode" config:type="short">'+IntToStr(vsm)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="HorizontalSplitPosition" config:type="int">'+IntToStr(sheet.LeftPaneWidth)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="VerticalSplitPosition" config:type="int">'+IntToStr(sheet.TopPaneHeight)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="ActiveSplitRange" config:type="short">'+IntToStr(asr)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="PositionLeft" config:type="int">0</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="PositionRight" config:type="int">'+IntToStr(sheet.LeftPaneWidth)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="PositionTop" config:type="int">0</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="PositionBottom" config:type="int">'+IntToStr(sheet.TopPaneHeight)+'</config:config-item>');
    AppendToStream(AStream,
        '<config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item>');
       // this "ShowGrid" overrides the global setting. But Open/LibreOffice do not allow to change ShowGrid per sheet.
    AppendToStream(AStream,
      '</config:config-item-map-entry>');
  end;
end;

procedure TsSpreadOpenDocWriter.WriteTableStyles(AStream: TStream);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  for i:=0 to FWorkbook.GetWorksheetCount-1 do
  begin
    sheet := FWorkbook.GetWorksheetByIndex(i);
    AppendToStream(AStream, Format(
      '<style:style style:name="ta%d" style:family="table" style:master-page-name="PageStyle_5f_%s">' +
        '<style:table-properties table:display="true" style:writing-mode="lr-tb"/>' +
      '</style:style>', [
      i+1, sheet.Name
    ]));
  end;
end;


{@@ ----------------------------------------------------------------------------
  Creates an XML string for inclusion of the textrotation style option into the
  written file from the textrotation setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString).
-------------------------------------------------------------------------------}
function TsSpreadOpenDocWriter.WriteTextRotationStyleXMLAsString(
  const AFormat: TsCellFormat): String;
begin
  Result := '';
  if not (uffTextRotation in AFormat.UsedFormattingFields) then
    exit;

  case AFormat.TextRotation of
    rt90DegreeClockwiseRotation        : Result := 'style:rotation-angle="270" ';
    rt90DegreeCounterClockwiseRotation : Result := 'style:rotation-angle="90" ';
    rtStacked                          : Result := 'style:direction="ttb" ';
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates an XML string for inclusion of the vertical alignment into the
  written file from the vertical alignment setting in the given format record.
  Is called from WriteStyles (via WriteStylesXMLAsString).
-------------------------------------------------------------------------------}
function TsSpreadOpenDocWriter.WriteVertAlignmentStyleXMLAsString(
  const AFormat: TsCellFormat): String;
begin
  Result := '';
  if not (uffVertAlign in AFormat.UsedFormattingFields) then
    exit;
  case AFormat.VertAlignment of
    vaTop    : Result := 'style:vertical-align="top" ';
    vaCenter : Result := 'style:vertical-align="middle" ';
    vaBottom : Result := 'style:vertical-align="bottom" ';
  end;
end;

procedure TsSpreadOpenDocWriter.WriteVirtualCells(AStream: TStream;
  ASheet: TsWorksheet);
var
  r, c, cc: Cardinal;
  lCell: TCell;
  row: PRow;
  value: variant;
  styleCell: PCell;
  styleName: String;
  h, h_mm: Single;      // row height in "lines" and millimeters, respectively
  k: Integer;
  rowStyleData: TRowStyleData;
  rowsRepeated: Cardinal;
  colsRepeated: Cardinal;
  colsRepeatedStr: String;
  defFontSize: Single;
  lastCol, lastRow: Cardinal;
begin
  // some abbreviations...
  lastCol := Workbook.VirtualColCount - 1;
  lastRow := Workbook.VirtualRowCount - 1;
  defFontSize := Workbook.GetFont(0).Size;

  rowsRepeated := 1;
  r := 0;
  while (r <= lastRow) do
  begin
    // Look for the row style of the current row (r)
    row := ASheet.FindRow(r);
    if row = nil then
      styleName := 'ro1'
    else
    begin
      styleName := '';

      h := row^.Height;   // row height in "lines"
      h_mm := PtsToMM((h + ROW_HEIGHT_CORRECTION) * defFontSize);  // in mm
      for k := 0 to FRowStyleList.Count-1 do
      begin
        rowStyleData := TRowStyleData(FRowStyleList[k]);
        // Compare row heights, but be aware of rounding errors
        if SameValue(rowStyleData.RowHeight, h_mm, 1E-3) then
        begin
          styleName := rowStyleData.Name;
          break;
        end;
      end;
      if styleName = '' then
        raise Exception.Create(rsRowStyleNotFound);
    end;

    // No empty rows allowed here for the moment!


    // Write the row XML
    AppendToStream(AStream, Format(
        '<table:table-row table:style-name="%s">', [styleName]));

    // Loop along the row and write the cells.
    c := 0;
    while c <= lastCol do
    begin
      // Empty cell? Need to count how many "table:number-columns-repeated" to be added
      colsRepeated := 1;

      lCell.Row := r;  // to silence a compiler hint...
      InitCell(r, c, lCell);
      value := varNull;
      styleCell := nil;

      Workbook.OnWriteCellData(Workbook, r, c, value, styleCell);

      if VarIsNull(value) then
      begin
        // Local loop to count empty cells
        cc := c + 1;
        while (cc <= lastCol) do
        begin
          InitCell(r, cc, lCell);
          value := varNull;
          styleCell := nil;
          Workbook.OnWriteCellData(Workbook, r, cc, value, styleCell);
          if not VarIsNull(value) then
            break;
          inc(cc);
        end;
        colsRepeated := cc - c;
        colsRepeatedStr := IfThen(colsRepeated = 1, '',
          Format('table:number-columns-repeated="%d"', [colsRepeated]));
        AppendToStream(AStream, Format(
          '<table:table-cell %s />', [colsRepeatedStr]));
      end else begin
        if VarIsNumeric(value) then
        begin
          lCell.ContentType := cctNumber;
          lCell.NumberValue := value;
        end else
        if VarType(value) = varDate then
        begin
          lCell.ContentType := cctDateTime;
          lCell.DateTimeValue := StrToDateTime(VarToStr(value), Workbook.FormatSettings);  // was: StrToDate
        end else
        if VarIsStr(value) then
        begin
          lCell.ContentType := cctUTF8String;
          lCell.UTF8StringValue := VarToStrDef(value, '');
        end else
        if VarIsBool(value) then
        begin
          lCell.ContentType := cctBool;
          lCell.BoolValue := value <> 0;
        end else
          lCell.ContentType := cctEmpty;
        WriteCellToStream(AStream, @lCell);
      end;
      inc(c, colsRepeated);
    end;

    AppendToStream(AStream,
        '</table:table-row>');

    // Next row
    inc(r, rowsRepeated);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates an XML string for inclusion of the wordwrap option into the
  written file from the wordwrap setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString).
-------------------------------------------------------------------------------}
function TsSpreadOpenDocWriter.WriteWordwrapStyleXMLAsString(
  const AFormat: TsCellFormat): String;
begin
  if (uffWordWrap in AFormat.UsedFormattingFields) then
    Result := 'fo:wrap-option="wrap" '
  else
    Result := '';
end;

{@@ ----------------------------------------------------------------------------
  Writes a string formula
-------------------------------------------------------------------------------}
procedure TsSpreadOpenDocWriter.WriteFormula(AStream: TStream; const ARow,
  ACol: Cardinal; ACell: PCell);
var
  lStyle: String = '';
  parser: TsExpressionParser;
  formula: String;
  valuetype: String;
  value: string;
  valueStr: String;
  colsSpannedStr: String;
  rowsSpannedStr: String;
  spannedStr: String;
  comment: String;
  r1,c1,r2,c2: Cardinal;
  fmt: TsCellFormat;
begin
  Unused(ARow, ACol);

  // Style
  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  if fmt.UsedFormattingFields <> [] then
    lStyle := ' table:style-name="ce' + IntToStr(ACell^.FormatIndex) + '" '
  else
    lStyle := '';

  // Comment
  comment := WriteCommentXMLAsString(FWorksheet.ReadComment(ACell));

  // Merged?
  if FWorksheet.IsMergeBase(ACell) then
  begin
    FWorksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    rowsSpannedStr := Format('table:number-rows-spanned="%d"', [r2 - r1 + 1]);
    colsSpannedStr := Format('table:number-columns-spanned="%d"', [c2 - c1 + 1]);
    spannedStr := colsSpannedStr + ' ' + rowsSpannedStr;
  end else
    spannedStr := '';

  // Hyperlink
  if FWorksheet.HasHyperlink(ACell) then
    FWorkbook.AddErrorMsg(rsODSHyperlinksOfTextCellsOnly, [GetCellString(ARow, ACol)]);

  // Convert string formula to the format needed by ods: semicolon list separators!
  parser := TsSpreadsheetParser.Create(FWorksheet);
  try
    parser.Dialect := fdOpenDocument;
    parser.Expression := ACell^.FormulaValue;
    formula := Parser.LocalizedExpression[FPointSeparatorSettings];
  finally
    parser.Free;
  end;

  valueStr := '';
  case ACell^.ContentType of
    cctNumber:
      begin
        valuetype := 'float';
        value := 'office:value="' + Format('%g', [ACell^.NumberValue], FPointSeparatorSettings) + '"';
      end;
    cctDateTime:
      if trunc(ACell^.DateTimeValue) = 0 then
      begin
        valuetype := 'time';
        value := 'office:time-value="' + FormatDateTime(ISO8601FormatTimeOnly, ACell^.DateTimeValue) + '"';
      end
      else
      begin
        valuetype := 'date';
        if frac(ACell^.DateTimeValue) = 0.0 then
          value := 'office:date-value="' + FormatDateTime(ISO8601FormatDateOnly, ACell^.DateTimeValue) + '"'
        else
          value := 'office:date-value="' + FormatDateTime(ISO8601FormatExtended, ACell^.DateTimeValue) + '"';
      end;
    cctUTF8String:
      begin
        valuetype := 'string';
        value := 'office:string-value="' + ACell^.UTF8StringValue +'"';
        valueStr := '<text:p>' + ACell^.UTF8StringValue + '</text:p>';
      end;
    cctBool:
      begin
        valuetype := 'boolean';
        value := 'office:boolean-value="' + BoolToStr(ACell^.BoolValue, 'true', 'false') + '"';
      end;
    cctError:
      begin
        // Strange: but in case of an error, Open/LibreOffice always writes a
        // float value 0 to the cell
        valuetype := 'float';
        value := 'office:value="0"';
      end;
  end;

  { Fix special xml characters }
  formula := UTF8TextToXMLText(formula);

  { We are writing a very rudimentary formula here without result and result
    data type. Seems to work... }
  if FWorksheet.GetCalcState(ACell) = csCalculated then
    AppendToStream(AStream, Format(
      '<table:table-cell table:formula="=%s" office:value-type="%s" %s %s %s>' +
        comment +
        valueStr +
      '</table:table-cell>', [
      formula, valuetype, value, lStyle, spannedStr
    ]))
  else
  begin
    AppendToStream(AStream, Format(
      '<table:table-cell table:formula="=%s" %s %s', [
        formula, lStyle, spannedStr]));
    if comment <> '' then
      AppendToStream(AStream, '>' + comment + '</table:table-cell>')
    else
      AppendToStream(AStream, '/>');
  end;
end;


{@@ ----------------------------------------------------------------------------
  Writes a cell with text content

  The UTF8 Text needs to be converted, because some chars are invalid in XML
  See bug with patch 19422
-------------------------------------------------------------------------------}
procedure TsSpreadOpenDocWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
var
  lStyle: string = '';
  colsSpannedStr: String;
  rowsSpannedStr: String;
  spannedStr: String;
  r1,c1,r2,c2: Cardinal;
  txt: ansistring;
  textp, target, bookmark, comment: String;
  fmt: TsCellFormat;
  hyperlink: PsHyperlink;
  u: TUri;
begin
  Unused(ARow, ACol);

  // Style
  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  if fmt.UsedFormattingFields <> [] then
    lStyle := ' table:style-name="ce' + IntToStr(ACell^.FormatIndex) + '" '
  else
    lStyle := '';

  // Comment
  comment := WriteCommentXMLAsString(FWorksheet.ReadComment(ACell));

  // Merged?
  if FWorksheet.IsMergeBase(ACell) then
  begin
    FWorksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    rowsSpannedStr := Format('table:number-rows-spanned="%d"', [r2 - r1 + 1]);
    colsSpannedStr := Format('table:number-columns-spanned="%d"', [c2 - c1 + 1]);
    spannedStr := colsSpannedStr + ' ' + rowsSpannedStr;
  end else
    spannedStr := '';

  // Check for invalid characters
  txt := AValue;
  if not ValidXMLText(txt) then
    Workbook.AddErrorMsg(
      rsInvalidCharacterInCell, [
      GetCellString(ARow, ACol)
    ]);

  // Hyperlink?
  if FWorksheet.HasHyperlink(ACell) then
  begin
    hyperlink := FWorksheet.FindHyperlink(ACell);
    SplitHyperlink(hyperlink^.Target, target, bookmark);

    if (target <> '') and (pos('file:', target) = 0) then
    begin
      u := ParseURI(target);
      if u.Protocol = '' then
        target := '../' + target;
    end;

    // ods absolutely wants "/" path delimiters in the file uri!
    FixHyperlinkPathdelims(target);

    if (bookmark <> '') then
      target := target + '#' + bookmark;

    textp := Format(
      '<text:p>'+
        '<text:a xlink:href="%s" xlink:type="simple">%s</text:a>'+
      '</text:p>', [target, txt]);

  end else
    // No hyperlink, normal text only
    textp := '<text:p>' + txt + '</text:p>';

  // Write it ...
  AppendToStream(AStream, Format(
    '<table:table-cell office:value-type="string" %s %s>' +
      comment +
      textp +
    '</table:table-cell>', [
    lStyle, spannedStr
  ]));
end;

procedure TsSpreadOpenDocWriter.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
var
  StrValue: string;
  DisplayStr: string;
  lStyle: string = '';
  valType: String;
  colsSpannedStr: String;
  rowsSpannedStr: String;
  spannedStr: String;
  comment: String;
  r1,c1,r2,c2: Cardinal;
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams;
  nfSection: TsNumFormatSection;
begin
  Unused(ARow, ACol);

  valType := 'float';
  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  if fmt.UsedFormattingFields <> [] then
  begin
    numFmt := FWorkbook.GetNumberFormat(fmt.NumberFormatIndex);
    if (numFmt <> nil) then begin
      if (Length(numFmt.Sections) > 1) and (AValue < 0) then
        nfSection := numFmt.Sections[1]
      else
      if (Length(numFmt.Sections) > 2) and (AValue = 0) then
        nfSection := NumFmt.Sections[2]
      else
        nfSection := numFmt.Sections[0];
      if (nfkPercent in nfSection.Kind) then
        valType := 'percentage'
      else
      if (nfkCurrency in nfSection.Kind) then
        valtype := 'currency'
    end;
    lStyle := ' table:style-name="ce' + IntToStr(ACell^.FormatIndex) + '" ';
  end else
    lStyle := '';

  // Comment
  comment := WriteCommentXMLAsString(FWorksheet.ReadComment(ACell));

  // Merged?
  if FWorksheet.IsMergeBase(ACell) then
  begin
    FWorksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    rowsSpannedStr := Format('table:number-rows-spanned="%d"', [r2 - r1 + 1]);
    colsSpannedStr := Format('table:number-columns-spanned="%d"', [c2 - c1 + 1]);
    spannedStr := colsSpannedStr + ' ' + rowsSpannedStr;
  end else
    spannedStr := '';

  // Displayed value
  if IsInfinite(AValue) then
  begin
    StrValue := '1.#INF';
    DisplayStr := '1.#INF';
  end else begin
    StrValue := FloatToStr(AValue, FPointSeparatorSettings); // Uses '.' as decimal separator
    DisplayStr := FWorksheet.ReadAsUTF8Text(ACell); //FloatToStr(AValue); // Uses locale decimal separator
  end;

  // Hyperlink
  if FWorksheet.HasHyperlink(ACell) then
    FWorkbook.AddErrorMsg(rsODSHyperlinksOfTextCellsOnly, [GetCellString(ARow, ACol)]);

  AppendToStream(AStream, Format(
    '<table:table-cell office:value-type="%s" office:value="%s" %s %s >' +
      comment +
      '<text:p>%s</text:p>' +
    '</table:table-cell>', [
    valType, StrValue, lStyle, spannedStr,
    DisplayStr
  ]));
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value
-------------------------------------------------------------------------------}
procedure TsSpreadOpenDocWriter.WriteDateTime(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
const
  DATE_FMT: array[boolean] of string = (ISO8601FormatExtended, ISO8601FormatTimeOnly);
  DT: array[boolean] of string = ('date', 'time');
  // Index "boolean" is to be understood as "isTimeOnly"
var
  lStyle: string;
  strValue: String;
  displayStr: String;
  isTimeOnly: Boolean;
  colsSpannedStr: String;
  rowsSpannedStr: String;
  spannedStr: String;
  comment: String;
  r1,c1,r2,c2: Cardinal;
  fmt: TsCellFormat;
  numFmtParams: TsNumFormatParams;
begin
  Unused(ARow, ACol);

  // Merged?
  if FWorksheet.IsMergeBase(ACell) then
  begin
    FWorksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    rowsSpannedStr := Format('table:number-rows-spanned="%d"', [r2 - r1 + 1]);
    colsSpannedStr := Format('table:number-columns-spanned="%d"', [c2 - c1 + 1]);
    spannedStr := colsSpannedStr + ' ' + rowsSpannedStr;
  end else
    spannedStr := '';

  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  numFmtParams := FWorkbook.GetNumberFormat(fmt.NumberFormatIndex);
  if fmt.UsedFormattingFields <> [] then
    lStyle := ' table:style-name="ce' + IntToStr(ACell^.FormatIndex) + '" '
  else
    lStyle := '';

  // Comment
  comment := WriteCommentXMLAsString(FWorksheet.ReadComment(ACell));

  // Hyperlink
  if FWorksheet.HasHyperlink(ACell) then
    FWorkbook.AddErrorMsg(rsODSHyperlinksOfTextCellsOnly, [GetCellString(ARow, ACol)]);

  // nfTimeInterval is a special case - let's handle it first:

  if IsTimeIntervalformat(numFmtParams) then
  begin
    strValue := FormatDateTime(ISO8601FormatHoursOverflow, AValue, [fdoInterval]);
    displayStr := FWorksheet.ReadAsUTF8Text(ACell);
//    displayStr := FormatDateTime(fmt.NumberFormatStr, AValue, [fdoInterval]);
    AppendToStream(AStream, Format(
      '<table:table-cell office:value-type="time" office:time-value="%s" %s %s>' +
        comment +
        '<text:p>%s</text:p>' +
      '</table:table-cell>', [
      strValue, lStyle, spannedStr,
      displayStr
    ]));
  end else
  begin
    // We have to distinguish between time-only values and values that contain date parts.
    if (numFmtParams <> nil) then
      isTimeOnly := Assigned(numFmtParams) and (numFmtParams.Sections[0].Kind * [nfkDate, nfkTime] = [nfkTime])
    else
      isTimeOnly := false;
    strValue := FormatDateTime(DATE_FMT[isTimeOnly], AValue);
    displayStr := FWorksheet.ReadAsUTF8Text(ACell);
    AppendToStream(AStream, Format(
      '<table:table-cell office:value-type="%s" office:%s-value="%s" %s %s>' +
        comment +
        '<text:p>%s</text:p> ' +
      '</table:table-cell>', [
      DT[isTimeOnly], DT[isTimeOnly], strValue, lStyle, spannedStr,
      displayStr
    ]));
  end;
end;

initialization

{@@ ----------------------------------------------------------------------------
  Registers this reader / writer on fpSpreadsheet
-------------------------------------------------------------------------------}
  RegisterSpreadFormat(TsSpreadOpenDocReader, TsSpreadOpenDocWriter, sfOpenDocument);

end.

