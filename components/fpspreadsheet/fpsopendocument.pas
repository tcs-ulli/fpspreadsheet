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
  {$mode delphi}
{$endif}

{.$define FPSPREADDEBUG} //used to be XLSDEBUG
interface

uses
  Classes, SysUtils,
  {$IFDEF FPC_FULLVERSION >= 20701}
  zipper,
  {$ELSE}
  fpszipper,
  {$ENDIF}
  fpspreadsheet,
  laz2_xmlread, laz2_DOM,
  AVL_Tree, math, dateutils,
  fpsutils;
  
type
  TDateMode=(
    dm1899 {default for ODF; almost same as Excel 1900},
    dm1900 {StarCalc legacy only},
    dm1904 {e.g. Quattro Pro,Mac Excel compatibility}
  );

  { TsSpreadOpenDocNumFormatList }
  TsSpreadOpenDocNumFormatList = class(TsCustomNumFormatList)
  protected
    procedure AddBuiltinFormats; override;
  public
//    function FormatStringForWriting(AIndex: Integer): String; override;
  end;

  { TsSpreadOpenDocReader }

  TsSpreadOpenDocReader = class(TsCustomSpreadReader)
  private
    FCellStyleList: TFPList;
    FColumnStyleList: TFPList;
    FColumnList: TFPList;
    FRowStyleList: TFPList;
    FRowList: TFPList;
    FVolatileNumFmtList: TsCustomNumFormatList;
    FDateMode: TDateMode;
    // Applies internally stored column widths to current worksheet
    procedure ApplyColWidths;
    // Applies a style to a cell
    procedure ApplyStyleToCell(ARow, ACol: Cardinal; AStyleName: String); overload;
    procedure ApplyStyleToCell(ACell: PCell; AStyleName: String); overload;
    // Extracts the date/time value from the xml node
    function ExtractDateTimeFromNode(ANode: TDOMNode;
      ANumFormat: TsNumberFormat; const AFormatStr: String): TDateTime;
    // Searches a style by its name in the CellStyleList
    function FindCellStyleByName(AStyleName: String): integer;
    // Searches a column style by its column index or its name in the StyleList
    function FindColumnByCol(AColIndex: Integer): Integer;
    function FindColStyleByName(AStyleName: String): integer;
    function FindRowStyleByName(AStyleName: String): Integer;
    // Gets value for the specified attribute. Returns empty string if attribute
    // not found.
    function GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
    procedure ReadColumns(ATableNode: TDOMNode);
    procedure ReadColumnStyle(AStyleNode: TDOMNode);
    // Figures out the base year for times in this file (dates are unambiguous)
    procedure ReadDateMode(SpreadSheetNode: TDOMNode);
    function ReadFont(ANode: TDOMnode; IsDefaultFont: Boolean): Integer;
    procedure ReadRowsAndCells(ATableNode: TDOMNode);
    procedure ReadRowStyle(AStyleNode: TDOMNode);
  protected
    procedure CreateNumFormatList; override;
    procedure ReadNumFormats(AStylesNode: TDOMNode);
    procedure ReadStyles(AStylesNode: TDOMNode);
    { Record writing methods }
    procedure ReadBlank(ARow, ACol: Word; ACellNode: TDOMNode);
    procedure ReadDateTime(ARow, ACol: Word; ACellNode: TDOMNode);
    procedure ReadFormula(ARow, ACol: Word; ACellNode: TDOMNode);
    procedure ReadLabel(ARow, ACol: Word; ACellNode: TDOMNode);
    procedure ReadNumber(ARow, ACol: Word; ACellNode: TDOMNode);
  public
    { General reading methods }
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); override;
  end;

  { TsSpreadOpenDocWriter }

  TsSpreadOpenDocWriter = class(TsCustomSpreadWriter)
  private
    FColumnStyleList: TFPList;
    FRowStyleList: TFPList;

    // Routines to write parts of files
    function WriteCellStylesXMLAsString: string;
    function WriteColStylesXMLAsString: String;
    function WriteRowStylesXMLAsString: String;

    function WriteColumnsXMLAsString(ASheet: TsWorksheet): String;
    function WriteRowsAndCellsXMLAsString(ASheet: TsWorksheet): String;

    function WriteBackgroundColorStyleXMLAsString(const AFormat: TCell): String;
    function WriteBorderStyleXMLAsString(const AFormat: TCell): String;
    function WriteDefaultFontXMLAsString: String;
    function WriteFontNamesXMLAsString: String;
    function WriteFontStyleXMLAsString(const AFormat: TCell): String;
    function WriteHorAlignmentStyleXMLAsString(const AFormat: TCell): String;
    function WriteTextRotationStyleXMLAsString(const AFormat: TCell): String;
    function WriteVertAlignmentStyleXMLAsString(const AFormat: TCell): String;
    function WriteWordwrapStyleXMLAsString(const AFormat: TCell): String;

  protected
    FPointSeparatorSettings: TFormatSettings;
    // Strings with the contents of files
    FMeta, FSettings, FStyles, FContent, FCellContent, FMimetype: string;
    FMetaInfManifest: string;
    // Streams with the contents of files
    FSMeta, FSSettings, FSStyles, FSContent, FSMimetype: TStringStream;
    FSMetaInfManifest: TStringStream;
    // Helpers
    procedure CreateNumFormatList; override;
    procedure ListAllColumnStyles;
    procedure ListAllRowStyles;
    // Routines to write those files
    procedure WriteMimetype;
    procedure WriteMetaInfManifest;
    procedure WriteMeta;
    procedure WriteSettings;
    procedure WriteStyles;
    procedure WriteContent;
    procedure WriteWorksheet(CurSheet: TsWorksheet);
    { Record writing methods }
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      const AFormula: TsFormula; ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); override;
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
  StrUtils;

const
  { OpenDocument general XML constants }
  XML_HEADER           = '<?xml version="1.0" encoding="utf-8" ?>';

  { OpenDocument Directory structure constants }
  OPENDOC_PATH_CONTENT   = 'content.xml';
  OPENDOC_PATH_META      = 'meta.xml';
  OPENDOC_PATH_SETTINGS  = 'settings.xml';
  OPENDOC_PATH_STYLES    = 'styles.xml';
  OPENDOC_PATH_MIMETYPE  = 'mimetype';
  OPENDOC_PATH_METAINF   = 'META-INF' + '/';
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

  COLWIDTH_EPS = 1e-2;   // for mm
  ROWHEIGHT_EPS = 1e-2;    // for lines

type
  { Cell style items relevant to FPSpreadsheet. Stored in the CellStyleList of the reader. }
  TCellStyleData = class
  public
    Name: String;
    FontIndex: Integer;
    NumFormatIndex: Integer;
    HorAlignment: TsHorAlignment;
    VertAlignment: TsVertAlignment;
    WordWrap: Boolean;
    TextRotation: TsTextRotation;
    Borders: TsCellBorders;
    BorderStyles: TsCellBorderStyles;
    BackgroundColor: TsColor;
  end;

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

  { Row data items stored in the RowList of the reader }
  TRowData = class
    Row: Integer;
    RowStyleIndex: Integer;   // index into FRowStyleList of reader
    DefaultCellStyleIndex: Integer;  // Index of default row style in FCellStyleList of reader
  end;


{ TsSpreadOpenDocNumFormatList }

procedure TsSpreadOpenDocNumFormatList.AddBuiltinFormats;
begin
  // there are no built-in number formats which are silently assumed to exist.
end;


{ TsSpreadOpenDocReader }

constructor TsSpreadOpenDocReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FCellStyleList := TFPList.Create;
  FColumnStyleList := TFPList.Create;
  FColumnList := TFPList.Create;
  FRowStyleList := TFPList.Create;
  FRowList := TFPList.Create;
  FVolatileNumFmtList := TsCustomNumFormatList.Create(Workbook);
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

  for j := FCellStyleList.Count-1 downto 0 do TObject(FCellStyleList[j]).Free;
  FCellStyleList.Free;

  FVolatileNumFmtList.Free;  // automatically destroys its items.

  inherited Destroy;
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
  for i:=0 to FColumnList.Count-1 do begin
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
    if not SameValue(colWidth, Workbook.DefaultColWidth, COLWIDTH_EPS) then begin
      col := FWorksheet.GetCol(colIndex);
      col^.Width := colWidth;
    end;
  end;
end;

{ Applies the style data referred to by the style name to the specified cell }
procedure TsSpreadOpenDocReader.ApplyStyleToCell(ARow, ACol: Cardinal;
  AStyleName: String);
var
  cell: PCell;
begin
  cell := FWorksheet.GetCell(ARow, ACol);
  if Assigned(cell) then
    ApplyStyleToCell(cell, AStyleName);
end;

procedure TsSpreadOpenDocReader.ApplyStyleToCell(ACell: PCell; AStyleName: String);
var
  styleData: TCellStyleData;
  styleIndex: Integer;
  numFmtData: TsNumFormatData;
  i: Integer;
begin
  // Is there a style attached to the cell?
  styleIndex := -1;
  if AStyleName <> '' then
    styleIndex := FindCellStyleByName(AStyleName);
  if (styleIndex = -1) then begin
    // No - look for the style attached to the column of the cell and
    // find the cell style by the DefaultCellStyleIndex stored in the column list.
    i := FindColumnByCol(ACell^.Col);
    if i = -1 then
      exit;
    styleIndex := TColumnData(FColumnList[i]).DefaultCellStyleIndex;
  end;

  styleData := TCellStyleData(FCellStyleList[styleIndex]);

  // Now copy all style parameters from the styleData to the cell.

  // Font
  if styleData.FontIndex = 1 then
    Include(ACell^.UsedFormattingFields, uffBold)
  else
  if styleData.FontIndex > 1 then
    Include(ACell^.UsedFormattingFields, uffFont);
  ACell^.FontIndex := styleData.FontIndex;

  // Word wrap
  if styleData.WordWrap then
    Include(ACell^.UsedFormattingFields, uffWordWrap)
  else
    Exclude(ACell^.UsedFormattingFields, uffWordWrap);

  // Text rotation
  if styleData.TextRotation > trHorizontal then
    Include(ACell^.UsedFormattingFields, uffTextRotation)
  else
    Exclude(ACell^.UsedFormattingFields, uffTextRotation);
  ACell^.TextRotation := styledata.TextRotation;

  // Text alignment
  if styleData.HorAlignment <> haDefault then begin
    Include(ACell^.UsedFormattingFields, uffHorAlign);
    ACell^.HorAlignment := styleData.HorAlignment;
  end else
    Exclude(ACell^.UsedFormattingFields, uffHorAlign);
  if styleData.VertAlignment <> vaDefault then begin
    Include(ACell^.UsedFormattingFields, uffVertAlign);
    ACell^.VertAlignment := styleData.VertAlignment;
  end else
    Exclude(ACell^.UsedFormattingFields, uffVertAlign);

  // Borders
  ACell^.BorderStyles := styleData.BorderStyles;
  if styleData.Borders <> [] then begin
    Include(ACell^.UsedFormattingFields, uffBorder);
    ACell^.Border := styleData.Borders;
  end else
    Exclude(ACell^.UsedFormattingFields, uffBorder);

  // Background color
  if styleData.BackgroundColor <> scNotDefined then begin
    Include(ACell^.UsedFormattingFields, uffBackgroundColor);
    ACell^.BackgroundColor := styleData.BackgroundColor;
  end;

  // Number format
  if styleData.NumFormatIndex > -1 then begin
    numFmtData := NumFormatList[styleData.NumFormatIndex];
    if numFmtData <> nil then begin
      Include(ACell^.UsedFormattingFields, uffNumberFormat);
      ACell^.NumberFormat := numFmtData.NumFormat;
      ACell^.NumberFormatStr := numFmtData.FormatString;
    end;
  end;
end;

{ Creates the correct version of the number format list
  suited for ODS file formats. }
procedure TsSpreadOpenDocReader.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsSpreadOpenDocNumFormatList.Create(Workbook);
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
  Hours, Minutes, Seconds, Days: integer;
  HoursPos, MinutesPos, SecondsPos: integer;
begin
  // Format expects ISO 8601 type date string or
  // time string
  fmt := DefaultFormatSettings;
  fmt.ShortDateFormat := 'yyyy-mm-dd';
  fmt.DateSeparator := '-';
  fmt.LongTimeFormat := 'hh:nn:ss';
  fmt.TimeSeparator := ':';

  Value := GetAttrValue(ANode, 'office:date-value');

  if Value <> '' then begin
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
    if (Value <> '') and (Pos('PT', Value) = 1) then begin
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
        Seconds := StrToInt(Copy(Value, MinutesPos+1, SecondsPos-MinutesPos-1))
      else
        Seconds := 0;

      Days := Hours div 24;
      Hours := Hours mod 24;
      Result := Days + (Hours + (Minutes + Seconds/60)/60)/24;

      { Values < 1 day are certainly time-only formats --> no datemode correction
        nfTimeInterval formats are differences --> no date mode correction
        In all other case, we have a date part that needs to be corrected for
        the file's datemode. }
      if (ANumFormat <> nfTimeInterval) and (abs(Days) > 0) then begin
        case FDateMode of
          dm1899: Result := Result + DATEMODE_1899_BASE;
          dm1900: Result := Result + DATEMODE_1900_BASE;
          dm1904: Result := Result + DATEMODE_1904_BASE;
        end;
      end;
    end;
  end;
end;

function TsSpreadOpenDocReader.FindCellStyleByName(AStyleName: String): Integer;
begin
  for Result:=0 to FCellStyleList.Count-1 do begin
    if TCellStyleData(FCellStyleList[Result]).Name = AStyleName then
      exit;
  end;
  Result := -1;
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

function TsSpreadOpenDocReader.FindRowStyleByName(AStyleName: String): Integer;
begin
  for Result := 0 to FRowStyleList.Count-1 do
    if TRowStyleData(FRowStyleList[Result]).Name = AStyleName then
      exit;
  Result := -1;
end;

function TsSpreadOpenDocReader.GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
var
  i : integer;
  Found : Boolean;
begin
  Found:=false;
  i:=0;
  Result:='';
  while not Found and (i<ANode.Attributes.Length) do begin
    if ANode.Attributes.Item[i].NodeName=AAttrName then begin
      Found:=true;
      Result:=ANode.Attributes.Item[i].NodeValue;
    end;
    inc(i);
  end;
end;

procedure TsSpreadOpenDocReader.ReadBlank(ARow, ACol: Word; ACellNode: TDOMNode);
var
  styleName: String;
begin
  FWorkSheet.WriteBlank(ARow, ACol);

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(ARow, ACol, stylename);
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
  colStyleData: TColumnStyleData;
  colData: TColumnData;
  colsRepeated: Integer;
  j: Integer;
begin
  // clear previous column list (from other sheets)
  for j:=FColumnList.Count-1 downto 0 do TObject(FColumnList[j]).Free;
  FColumnList.Clear;

  col := 0;
  colNode := ATableNode.FindNode('table:table-column');
  while Assigned(colNode) do begin
    if colNode.NodeName = 'table:table-column' then begin;
      s := GetAttrValue(colNode, 'table:style-name');
      colStyleIndex := FindColStyleByName(s);
      if colStyleIndex <> -1 then begin
        defCellStyleIndex := -1;
        colStyleData := TColumnStyleData(FColumnStyleList[colStyleIndex]);
        s := GetAttrValue(ColNode, 'table:default-cell-style-name');
        if s <> '' then begin
          defCellStyleIndex := FindCellStyleByName(s);
          colData := TColumnData.Create;
          colData.Col := col;
          colData.ColStyleIndex := colStyleIndex;
          colData.DefaultCellStyleIndex := defCellStyleIndex;
          FColumnList.Add(colData);
        end;
        s := GetAttrValue(ColNode, 'table:number-columns-repeated');
        if s = '' then
          inc(col)
        else begin
          colsRepeated := StrToInt(s);
          if defCellStyleIndex > -1 then
            for j:=1 to colsRepeated-1 do begin
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

  while Assigned(styleChildNode) do begin
    if styleChildNode.NodeName = 'style:table-column-properties' then begin
      s := GetAttrValue(styleChildNode, 'style:column-width');
      if s <> '' then begin
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

procedure TsSpreadOpenDocReader.ReadDateMode(SpreadSheetNode: TDOMNode);
var
  CalcSettingsNode, NullDateNode: TDOMNode;
  NullDateSetting: string;
begin
  // Default datemode for ODF:
  NullDateSetting:='1899-12-30';
  CalcSettingsNode:=SpreadsheetNode.FindNode('table:calculation-settings');
  if Assigned(CalcSettingsNode) then
  begin
    NullDateNode:=CalcSettingsNode.FindNode('table:null-date');
    if Assigned(NullDateNode) then
      NullDateSetting:=GetAttrValue(NullDateNode,'table:date-value');
  end;
  if NullDateSetting='1899-12-30' then
    FDateMode := dm1899
  else if NullDateSetting='1900-01-01' then
    FDateMode := dm1900
  else if NullDateSetting='1904-01-01' then
    FDateMode := dm1904
  else
    raise Exception.CreateFmt('Spreadsheet file corrupt: cannot handle null-date format %s', [NullDateSetting]);
end;

{ Reads font data from an xml node, adds the font to the workbooks FontList
  (if not yet contained), and returns the index in the font list.
  If "IsDefaultFont" is true the first FontList entry (DefaultFont) is replaced. }
function TsSpreadOpenDocReader.ReadFont(ANode: TDOMnode;
  IsDefaultFont: Boolean): Integer;
var
  fntName: String;
  fntSize: Single;
  fntStyles: TsFontStyles;
  fntColor: TsColor;
  s: String;
begin
  if ANode = nil then begin
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
  if GetAttrValue(ANode, 'style:text-underline-style') <> '' then
    Include(fntStyles, fssUnderline);
  if GetAttrValue(ANode, 'style:text-strike-through-style') <> '' then
    Include(fntStyles, fssStrikeout);

  s := GetAttrValue(ANode, 'fo:color');
  if s <> '' then
    fntColor := FWorkbook.AddColorToPalette(HTMLColorStrToColor(s))
  else
    fntColor := FWorkbook.GetFont(0).Color;

  if IsDefaultFont then begin
    FWorkbook.SetDefaultFont(fntName, fntSize);
    Result := 0;
  end
  else begin
    Result := FWorkbook.FindFont(fntName, fntSize, fntStyles, fntColor);
    if Result = -1 then
      Result := FWorkbook.AddFont(fntName, fntSize, fntStyles, fntColor);
  end;
end;

procedure TsSpreadOpenDocReader.ReadFromFile(AFileName: string; AData: TsWorkbook);
var
  Doc : TXMLDocument;
  FilePath : string;
  UnZip : TUnZipper;
  FileList : TStringList;
  BodyNode, SpreadSheetNode, TableNode: TDOMNode;
  StylesNode: TDOMNode;

  { We have to use our own ReadXMLFile procedure (there is one in xmlread)
    because we have to preserve spaces in element text for date/time separator.
    As a side-effect we have to skip leading spaces by ourselves. }
  procedure ReadXMLFile(out ADoc: TXMLDocument; AFileName: String);
  var
    parser: TDOMParser;
    src: TXMLInputSource;
    stream: TStream;
  begin
    stream := TFileStream.Create(AFileName, fmOpenRead + fmShareDenyWrite);
    try
      parser := TDOMParser.Create;
      try
        parser.Options.PreserveWhiteSpace := true;    // This preserves spaces!
        src := TXMLInputSource.Create(stream);
        try
          parser.Parse(src, ADoc);
        finally
          src.Free;
        end;
      finally
        parser.Free;
      end;
    finally
      stream.Free;
    end;
  end;

begin
  //unzip content.xml into AFileName path
  FilePath := GetTempDir(false);
  UnZip := TUnZipper.Create;
  UnZip.OutputPath := FilePath;
  FileList := TStringList.Create;
  FileList.Add('styles.xml');
  FileList.Add('content.xml');
  try
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
    while Assigned(TableNode) do begin
      // These nodes occur due to leading spaces which are not skipped
      // automatically any more due to PreserveWhiteSpace option applied
      // to ReadXMLFile
      if TableNode.NodeName = '#text' then begin
        TableNode := TableNode.NextSibling;
        continue;
      end;
      FWorkSheet := aData.AddWorksheet(GetAttrValue(TableNode,'table:name'));
      // Collect column styles used
      ReadColumns(TableNode);
      // Process each row inside the sheet and process each cell of the row
      ReadRowsAndCells(TableNode);
      ApplyColWidths;
      // Continue with next table
      TableNode := TableNode.NextSibling;
    end; //while Assigned(TableNode)

  finally
    if Assigned(Doc) then Doc.Free;
  end;
end;

procedure TsSpreadOpenDocReader.ReadFormula(ARow: Word; ACol : Word; ACellNode : TDOMNode);
var
  cell: PCell;
  formula: String;
  stylename: String;
  txtValue: String;
  floatValue: Double;
  fs: TFormatSettings;
  valueType: String;
  valueStr: String;
  node: TDOMNode;
begin
  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';

  // Create cell and apply format
  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(ARow, ACol, stylename);
  cell := FWorksheet.FindCell(ARow, ACol);

  // Read formula, store in the cell's FormulaValue.FormulaStr
  formula := GetAttrValue(ACellNode, 'table:formula');
  if formula <> '' then Delete(formula, 1, 3);  // delete "of:"
  cell^.FormulaValue.FormulaStr := formula;

  // Read formula results
  // ... number value
  valueType := GetAttrValue(ACellNode, 'office:value-type');
  valueStr := GetAttrValue(ACellNode, 'office:value');
  if (valueType = 'float') then begin
    if UpperCase(valueStr) = '1.#INF' then
      FWorksheet.WriteNumber(cell, 1.0/0.0)
    else begin
      floatValue := StrToFloat(valueStr, fs);
      FWorksheet.WriteNumber(cell, floatValue);
    end;
    if IsDateTimeFormat(cell^.NumberFormat) then begin
      cell^.ContentType := cctDateTime;
      // No datemode correction for intervals and for time-only values
      if (cell^.NumberFormat = nfTimeInterval) or (cell^.NumberValue < 1) then
        cell^.DateTimeValue := cell^.NumberValue
      else
        case FDateMode of
          dm1899: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1899_BASE;
          dm1900: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1900_BASE;
          dm1904: cell^.DateTimeValue := cell^.NumberValue + DATEMODE_1904_BASE;
        end;
    end;
  end else
  // Date/time value
  if (valueType = 'date') or (valueType = 'time') then begin
    floatValue := ExtractDateTimeFromNode(ACellNode, cell^.NumberFormat, cell^.NumberFormatStr);
    FWorkSheet.WriteDateTime(cell, floatValue);
  end else
  // text
  if (valueType = 'string') then begin
    node := ACellNode.FindNode('text:p');
    if (node <> nil) and (node.FirstChild <> nil) then begin
      valueStr := node.FirstChild.Nodevalue;
      FWorksheet.WriteUTF8Text(cell, valueStr);
    end;
  end else
  // Text
    FWorksheet.WriteUTF8Text(cell, valueStr);
end;

procedure TsSpreadOpenDocReader.ReadLabel(ARow: Word; ACol: Word; ACellNode: TDOMNode);
var
  cellText: String;
  styleName: String;
begin
  cellText := ACellNode.TextContent;
  FWorkSheet.WriteUTF8Text(ARow, ACol, cellText);

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(ARow, ACol, stylename);
end;

procedure TsSpreadOpenDocReader.ReadNumber(ARow: Word; ACol : Word; ACellNode : TDOMNode);
var
  FSettings: TFormatSettings;
  Value, Str: String;
  lNumber: Double;
  styleName: String;
  lCell: PCell;
begin
  FSettings := DefaultFormatSettings;
  FSettings.DecimalSeparator:='.';

  Value := GetAttrValue(ACellNode,'office:value');
  if UpperCase(Value)='1.#INF' then
    FWorkSheet.WriteNumber(Arow,ACol,1.0/0.0)
  else
  begin
    // Don't merge, or else we can't debug
    Str := GetAttrValue(ACellNode,'office:value');
    lNumber := StrToFloat(Str,FSettings);
    FWorkSheet.WriteNumber(ARow,ACol,lNumber);
  end;

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(ARow, ACol, stylename);

  // Sometimes date/time cells are stored as "float".
  // We convert them to date/time and also correct the date origin offset if
  // needed.
  lCell := FWorksheet.FindCell(ARow, ACol);
  if IsDateTimeFormat(lCell^.NumberFormat) or IsDateTimeFormat(lCell^.NumberFormatStr)
  then begin
    lCell^.ContentType := cctDateTime;
    // No datemode correction for intervals and for time-only values
    if (lCell^.NumberFormat = nfTimeInterval) or (lCell^.NumberValue < 1) then
      lCell^.DateTimeValue := lCell^.NumberValue
    else
      case FDateMode of
        dm1899: lCell^.DateTimeValue := lCell^.NumberValue + DATEMODE_1899_BASE;
        dm1900: lCell^.DateTimeValue := lCell^.NumberValue + DATEMODE_1900_BASE;
        dm1904: lCell^.DateTimeValue := lCell^.NumberValue + DATEMODE_1904_BASE;
      end;
  end;
end;

procedure TsSpreadOpenDocReader.ReadDateTime(ARow: Word; ACol: Word;
  ACellNode : TDOMNode);
var
  dt: TDateTime;
  styleName: String;
  cell: PCell;
begin
  cell := FWorksheet.GetCell(ARow, ACol);

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(cell, stylename);

  dt := ExtractDateTimeFromNode(ACellNode, cell^.NumberFormat, cell^.NumberFormatStr);
  FWorkSheet.WriteDateTime(cell, dt, cell^.NumberFormat, cell^.NumberFormatStr);
end;

procedure TsSpreadOpenDocReader.ReadNumFormats(AStylesNode: TDOMNode);

  procedure ReadStyleMap(ANode: TDOMNode; var ANumFormat: TsNumberFormat;
    var AFormatStr: String);
  var
    condition: String;
    stylename: String;
    styleindex: Integer;
    fmt: String;
    posfmt, negfmt, zerofmt: String;
  begin
    posfmt := '';
    negfmt := '';
    zerofmt := '';

    while ANode <> nil do begin
      condition := ANode.NodeName;

      if (ANode.NodeName = '#text') or not ANode.HasAttributes then begin
        ANode := ANode.NextSibling;
        Continue;
      end;

      condition := GetAttrValue(ANode, 'style:condition');
      stylename := GetAttrValue(ANode, 'style:apply-style-name');
      if (condition = '') or (stylename = '') then begin
        ANode := ANode.NextSibling;
        continue;
      end;

      Delete(condition, 1, Length('value()'));
      styleindex := -1;
      styleindex := FNumFormatList.FindByName(stylename);
      if (styleindex = -1) or (condition = '') then begin
        ANode := ANode.NextSibling;
        continue;
      end;

      fmt := FNumFormatList[styleindex].FormatString;
      case condition[1] of
        '<': begin
               negfmt := fmt;
               if (Length(condition) > 1) and (condition[2] = '=') then
                 zerofmt := fmt;
             end;
        '>': begin
               posfmt := fmt;
               if (Length(condition) > 1) and (condition[2] = '=') then
                 zerofmt := fmt;
             end;
        '=': begin
               zerofmt := fmt;
             end;
      end;
      ANode := ANode.NextSibling;
    end;
    if posfmt = '' then posfmt := AFormatStr;
    if negfmt = '' then negfmt := AFormatStr;

    AFormatStr := posFmt;
    if negfmt <> '' then AFormatStr := AFormatStr + ';' + negfmt;
    if zerofmt <> '' then AFormatStr := AFormatStr + ';' + zerofmt;

//    if ANumFormat <> nfFmtDateTime then
      ANumFormat := nfCustom;
  end;

  procedure ReadNumberStyle(ANumFormatNode: TDOMNode; ANumFormatName: String);
  var
    node, childNode: TDOMNode;
    fmtName, nodeName: String;
    fmt: String;
    nf: TsNumberFormat;
    decs: Byte;
    s: String;
    grouping: Boolean;
    nex: Integer;
    cs: String;
  begin
    fmt := '';
    node := ANumFormatNode.FirstChild;
    while Assigned(node) do begin
      nodeName := node.NodeName;
      if nodeName = '#text' then begin
        node := node.NextSibling;
        Continue;
      end else
      if nodeName = 'number:number' then begin
        s := GetAttrValue(node, 'number:decimal-places');
        if s <> '' then decs := StrToInt(s) else decs := 0;
        grouping := GetAttrValue(node, 'number:grouping') = 'true';
        nf := IfThen(grouping, nfFixedTh, nfFixed);
        fmt := fmt + BuildNumberFormatString(nf, Workbook.FormatSettings, decs);
      end else
      if nodeName = 'number:scientific-number' then begin
        nf := nfExp;
        s := GetAttrValue(node, 'number:decimal-places');
        if s <> '' then decs := StrToInt(s) else decs := 0;
        s := GetAttrValue(node, 'number:min-exponent-digits');
        if s <> '' then nex := StrToInt(s) else nex := 1;
        fmt := fmt + BuildNumberFormatString(nfFixed, Workbook.FormatSettings, decs);
        fmt := fmt + 'E+' + DupeString('0', nex);
      end else
      if nodeName = 'number:currency-symbol' then begin
        childnode := node.FirstChild;
        while childnode <> nil do begin
          cs := cs + childNode.NodeValue;
          fmt := fmt + childNode.NodeValue;
          childNode := childNode.NextSibling;
        end;
      end else
      if nodeName = 'number:text' then begin
        childNode := node.FirstChild;
        while childNode <> nil do begin
          fmt := fmt + childNode.NodeValue;
          childNode := childNode.NextSibling;
        end;
      end;
      node := node.NextSibling;
    end;

    node := ANumFormatNode.FindNode('style:map');
    if node <> nil then
      ReadStyleMap(node, nf, fmt);

    if ANumFormatNode.NodeName = 'number:percentage-style' then
      nf := nfPercentage
    else if ANumFormatNode.NodeName = 'number:currency-style' then
      nf := nfCurrency;

    NumFormatList.AddFormat(ANumFormatName, fmt, nf);
  end;

  procedure ReadDateTimeStyle(ANumFormatNode: TDOMNode; ANumFormatName: String);
  var
    node, childNode: TDOMNode;
    nf: TsNumberFormat;
    fmt: String;
    nodeName: String;
    s, stxt, sovr: String;
    isInterval: Boolean;
  begin
    fmt := '';
    isInterval := false;
    sovr := GetAttrValue(ANumFormatNode, 'number:truncate-on-overflow');
    if (sovr = 'false') then
      isInterval := true;
    node := ANumFormatNode.FirstChild;
    while Assigned(node) do begin
      nodeName := node.NodeName;
      if nodeName = '#text' then begin
        node := node.NextSibling;
        Continue;
      end else
      if nodeName = 'number:year' then begin
        s := GetAttrValue(node, 'number:style');
        if s = 'long' then fmt := fmt + 'yyyy'
        else if s = '' then fmt := fmt + 'yy';
      end else
      if nodeName = 'number:month' then begin
        s := GetAttrValue(node, 'number:style');
        stxt := GetAttrValue(node, 'number:textual');
        if (stxt = 'true') then begin // Month as text
          if (s = 'long') then fmt := fmt + 'mmmm' else fmt := fmt + 'mmm';
        end else begin   // Month as number
          if (s = 'long') then fmt := fmt + 'mm' else fmt := fmt + 'm';
        end;
      end else
      if nodeName = 'number:day' then begin
        s := GetAttrValue(node, 'number:style');
        stxt := GetAttrValue(node, 'number:textual');
        if (stxt = 'true') then begin   // day as text
          if (s = 'long') then fmt := fmt + 'dddd' else fmt := fmt + 'ddd';
        end else begin   // day as number
          if (s = 'long') then fmt := fmt + 'dd' else fmt := fmt + 'd';
        end;
      end;
      if nodeName = 'number:day-of-week' then begin
        s := GetAttrValue(node, 'number:stye');
        if (s = 'long') then fmt := fmt + 'dddddd' else fmt := fmt + 'ddddd';
      end else
      if nodeName = 'number:hours' then begin
        s := GetAttrValue(node, 'number:style');
        if (sovr = 'false') then begin
          if (s = 'long') then fmt := fmt + '[hh]' else fmt := fmt + '[h]';
        end else begin
          if (s = 'long') then fmt := fmt + 'hh' else fmt := fmt + 'h';
        end;
        sovr := '';
      end else
      if nodeName = 'number:minutes' then begin
        s := GetAttrValue(node, 'number:style');
        if (sovr = 'false') then begin
          if (s = 'long') then fmt := fmt + '[nn]' else fmt := fmt + '[n]';
        end else begin
          if (s = 'long') then fmt := fmt + 'nn' else fmt := fmt + 'n';
        end;
        sovr := '';
      end else
      if nodeName = 'number:seconds' then begin
        s := GetAttrValue(node, 'number:style');
        if (sovr = 'false') then begin
          if (s = 'long') then fmt := fmt + '[ss]' else fmt := fmt + '[s]';
        end else begin
          if (s = 'long') then fmt := fmt + 'ss' else fmt := fmt + 's';
          sovr := '';
        end;
        s := GetAttrValue(node, 'number:decimal-places');
        if (s <> '') and (s <> '0') then
          fmt := fmt + '.' + DupeString('0', StrToInt(s));
      end else
      if nodeName = 'number:am-pm' then
        fmt := fmt + 'AM/PM'
      else
      if nodeName = 'number:text' then begin
        childnode := node.FirstChild;
        if childnode <> nil then
          fmt := fmt + childnode.NodeValue;
      end;
      node := node.NextSibling;
    end;

//    nf := IfThen(isInterval, nfTimeInterval, nfFmtDateTime);
    nf := IfThen(isInterval, nfTimeInterval, nfCustom);
    node := ANumFormatNode.FindNode('style:map');
    if node <> nil then
      ReadStyleMap(node, nf, fmt);

    NumFormatList.AddFormat(ANumFormatName, fmt, nf);
  end;

  procedure ReadTextStyle(ANumFormatNode: TDOMNode; ANumFormatName: String);
  var
    node, childNode: TDOMNode;
    nf: TsNumberFormat;
    fmt: String;
    nodeName: String;
    s: String;
  begin
    fmt := '';
    node := ANumFormatNode.FirstChild;
    while Assigned(node) do begin
      nodeName := node.NodeName;
      if nodeName = '#text' then begin
        node := node.NextSibling;
        Continue;
      end else
      if nodeName = 'number:text-content' then begin
        // ???
      end else
      if nodeName = 'number:text' then begin
        childnode := node.FirstChild;
        if childnode <> nil then
          fmt := fmt + childnode.NodeValue;
      end;
      node := node.NextSibling;
    end;

    node := ANumFormatNode.FindNode('style:map');
    if node <> nil then
      ReadStyleMap(node, nf, fmt);
    {
    if IsDateTimeFormat(fmt) then
      nf := nfCustom
    else
    }
      nf := nfCustom;

    NumFormatList.AddFormat(ANumFormatName, fmt, nf);
  end;

var
  NumFormatNode: TDOMNode;
  numfmt_nodename, numfmtname: String;

begin
  if not Assigned(AStylesNode) then
    exit;

  NumFormatNode := AStylesNode.FirstChild;
  while Assigned(NumFormatNode) do begin
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

{ Reads the cells in the given table. Loops through all rows, and then finds all
  cells of each row. }
procedure TsSpreadOpenDocReader.ReadRowsAndCells(ATableNode: TDOMNode);
var
  row: Integer;
  col: Integer;
  cellNode, rowNode: TDOMNode;
  paramValueType, paramFormula, tableStyleName: String;
  paramColsRepeated, paramRowsRepeated: String;
  colsRepeated, rowsRepeated: Integer;
  rowStyleName: String;
  rowStyleIndex: Integer;
  rowStyle: TRowStyleData;
  rowHeight: Single;
  autoRowHeight: Boolean;
  i: Integer;
  lRow: PRow;
  s: String;
begin
  rowsRepeated := 0;
  row := 0;

  rowNode := ATableNode.FindNode('table:table-row');
  while Assigned(rowNode) do begin
    // These nodes occur due to indentation spaces which are not skipped
    // automatically any more due to PreserveWhiteSpace option applied
    // to ReadXMLFile
    if rowNode.NodeName = '#text' then begin
      rowNode := rowNode.NextSibling;
      Continue;
    end;

    // Read rowstyle
    rowStyleName := GetAttrValue(rowNode, 'table:style-name');
    rowStyleIndex := FindRowStyleByName(rowStyleName);
    rowStyle := TRowStyleData(FRowStyleList[rowStyleIndex]);
    rowHeight := rowStyle.RowHeight;           // in mm (see ReadRowStyles)
    rowHeight := mmToPts(rowHeight) / Workbook.GetDefaultFontSize;
    if rowHeight > ROW_HEIGHT_CORRECTION
      then rowHeight := rowHeight - ROW_HEIGHT_CORRECTION  // in "lines"
      else rowHeight := 0;
    autoRowHeight := rowStyle.AutoRowHeight;

    col := 0;

    //process each cell of the row
    cellNode := rowNode.FindNode('table:table-cell');
    while Assigned(cellNode) do begin

      // These nodes occur due to indentation spaces which are not skipped
      // automatically any more due to PreserveWhiteSpace option applied
      // to ReadXMLFile
      if cellNode.NodeName = '#text' then begin
        cellNode := cellNode.NextSibling;
        Continue;
      end;

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
      else if (paramValueType = '') and (tableStyleName <> '') then
        ReadBlank(row, col, cellNode);

      if ParamFormula <> '' then
        ReadFormula(row, col, cellNode);
//        ReadLabel(row, col, cellNode);

      paramColsRepeated := GetAttrValue(cellNode, 'table:number-columns-repeated');
      if paramColsRepeated = '' then paramColsRepeated := '1';
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

  while Assigned(styleChildNode) do begin
    if styleChildNode.NodeName = 'style:table-row-properties' then begin
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

procedure TsSpreadOpenDocReader.ReadStyles(AStylesNode: TDOMNode);
var
  fs: TFormatSettings;
  style: TCellStyleData;
  styleNode: TDOMNode;
  styleChildNode: TDOMNode;
  colStyle: TColumnStyleData;
  colWidth: Double;
  family: String;
  styleName: String;
  styleIndex: Integer;
  numFmtName: String;
  numFmtIndex: Integer;
  numFmtIndexDefault: Integer;
  wrap: Boolean;
  txtRot: TsTextRotation;
  vertAlign: TsVertAlignment;
  horAlign: TsHorAlignment;
  borders: TsCellBorders;
  borderStyles: TsCellBorderStyles;
  bkClr: TsColorValue;
  fntIndex: Integer;
  s: String;

  procedure SetBorderStyle(ABorder: TsCellBorder; AStyleValue: String);
  const
    EPS = 0.1;  // takes care of rounding errors for line widths
  var
    L: TStringList;
    i: Integer;
    isSolid: boolean;
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
      for i:=0 to L.Count-1 do begin
        s := L[i];
        if (s = 'solid') or (s = 'dashed') or (s = 'fine-dashed') or (s = 'dotted') or (s = 'double')
        then begin
          linestyle := s;
          continue;
        end;
        p := pos('pt', s);
        if p = Length(s)-1 then begin
          wid := StrToFloat(copy(s, 1, p-1), fs);
          continue;
        end;
        p := pos('mm', s);
        if p = Length(s)-1 then begin
          wid := mmToPts(StrToFloat(copy(s, 1, p-1), fs));
          Continue;
        end;
        p := pos('cm', s);
        if p = Length(s)-1 then begin
          wid := cmToPts(StrToFloat(copy(s, 1, p-1), fs));
          Continue;
        end;
        rgb := HTMLColorStrToColor(s);
      end;
      borderStyles[ABorder].LineStyle := lsThin;
      if (linestyle = 'solid') then begin
        if (wid >= 3 - EPS) then borderStyles[ABorder].LineStyle := lsThick
        else if (wid >= 2 - EPS) then borderStyles[ABorder].LineStyle := lsMedium
      end else
      if (linestyle = 'dotted') then
        borderStyles[ABorder].LineStyle := lsHair
      else
      if (linestyle = 'dashed') then
        borderStyles[ABorder].LineStyle := lsDashed
      else
      if (linestyle = 'fine-dashed') then
        borderStyles[ABorder].LineStyle := lsDotted
      else
      if (linestyle = 'double') then
        borderStyles[ABorder].LineStyle := lsDouble;
      borderStyles[ABorder].Color := IfThen(rgb = TsColorValue(-1),
        scBlack, Workbook.AddColorToPalette(rgb));
    finally
      L.Free;
    end;
  end;

begin
  if not Assigned(AStylesNode) then
    exit;

  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';

  numFmtIndexDefault := NumFormatList.FindByName('N0');

  styleNode := AStylesNode.FirstChild;
  while Assigned(styleNode) do begin
    if styleNode.NodeName = 'style:default-style' then begin
      ReadFont(styleNode.FindNode('style:text-properties'), true);
    end else
    if styleNode.NodeName = 'style:style' then begin
      family := GetAttrValue(styleNode, 'style:family');

      // Column styles
      if family = 'table-column' then
        ReadColumnStyle(styleNode);

      // Row styles
      if family = 'table-row' then
        ReadRowStyle(styleNode);

      // Cell styles
      if family = 'table-cell' then begin
        styleName := GetAttrValue(styleNode, 'style:name');
        numFmtName := GetAttrValue(styleNode, 'style:data-style-name');
        numFmtIndex := -1;
        if numFmtName <> '' then numFmtIndex := NumFormatList.FindByName(numFmtName);
        if numFmtIndex = -1 then numFmtIndex := numFmtIndexDefault;

        borders := [];
        wrap := false;
        bkClr := TsColorValue(-1);
        txtRot := trHorizontal;
        horAlign := haDefault;
        vertAlign := vaDefault;
        fntIndex := 0;

        styleChildNode := styleNode.FirstChild;
        while Assigned(styleChildNode) do begin
          if styleChildNode.NodeName = 'style:text-properties' then
            fntIndex := ReadFont(styleChildNode, false)
          else
          if styleChildNode.NodeName = 'style:table-cell-properties' then begin
            // Background color
            s := GetAttrValue(styleChildNode, 'fo:background-color');
            if (s <> '') and (s <> 'transparent') then
              bkClr := HTMLColorStrToColor(s);
            // Borders
            s := GetAttrValue(styleChildNode, 'fo:border');
            if (s <>'') then begin
              borders := borders + [cbNorth, cbSouth, cbEast, cbWest];
              SetBorderStyle(cbNorth, s);
              SetBorderStyle(cbSouth, s);
              SetBorderStyle(cbEast, s);
              SetBorderStyle(cbWest, s);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-top');
            if (s <> '') and (s <> 'none') then begin
              Include(borders, cbNorth);
              SetBorderStyle(cbNorth, s);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-right');
            if (s <> '') and (s <> 'none') then begin
              Include(borders, cbEast);
              SetBorderStyle(cbEast, s);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-bottom');
            if (s <> '') and (s <> 'none') then begin
              Include(borders, cbSouth);
              SetBorderStyle(cbSouth, s);
            end;
            s := GetAttrValue(styleChildNode, 'fo:border-left');
            if (s <> '') and (s <> 'none') then begin
              Include(borders, cbWest);
              SetBorderStyle(cbWest, s);
            end;

            // Text wrap
            s := GetAttrValue(styleChildNode, 'fo:wrap-option');
            wrap := (s='wrap');

            // Test rotation
            s := GetAttrValue(styleChildNode, 'style:rotation-angle');
            if s = '90' then
              txtRot := rt90DegreeCounterClockwiseRotation
            else if s = '270' then
              txtRot := rt90DegreeClockwiseRotation;
            s := GetAttrValue(styleChildNode, 'style:direction');
            if s = 'ttb' then
              txtRot := rtStacked;

            // Vertical text alignment
            s := GetAttrValue(styleChildNode, 'style:vertical-align');
            if s = 'top' then
              vertAlign := vaTop
            else if s = 'middle' then
              vertAlign := vaCenter
            else if s = 'bottom' then
              vertAlign := vaBottom;

          end else
          if styleChildNode.NodeName = 'style:paragraph-properties' then begin
            // Horizontal text alignment
            s := GetAttrValue(styleChildNode, 'fo:text-align');
            if s = 'start' then
              horAlign := haLeft
            else if s = 'end' then
              horAlign := haRight
            else if s = 'center' then
              horAlign := haCenter;
          end;
          styleChildNode := styleChildNode.NextSibling;
        end;

        style := TCellStyleData.Create;
        style.Name := stylename;
        style.FontIndex := fntIndex;
        style.NumFormatIndex := numFmtIndex;
        style.HorAlignment := horAlign;
        style.VertAlignment := vertAlign;
        style.WordWrap := wrap;
        style.TextRotation := txtRot;
        style.Borders := borders;
        style.BorderStyles := borderStyles;
        style.BackgroundColor := IfThen(bkClr = TsColorValue(-1), scNotDefined,
          Workbook.AddColorToPalette(bkClr));

        styleIndex := FCellStyleList.Add(style);
      end;
    end;
    styleNode := styleNode.NextSibling;
  end;
end;


{ TsSpreadOpenDocWriter }

procedure TsSpreadOpenDocWriter.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsSpreadOpenDocNumFormatList.Create(Workbook);
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
  colStyle.ColWidth := Workbook.DefaultColWidth;
  FColumnStyleList.Add(colStyle);

  for i:=0 to Workbook.GetWorksheetCount-1 do begin
    sheet := Workbook.GetWorksheetByIndex(i);
    for c:=0 to sheet.GetLastColIndex do begin
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
      if not found then begin
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
  for i:=0 to FColumnStyleList.Count-1 do begin
    w := TColumnStyleData(FColumnStyleList[i]).ColWidth * multiplier;
    TColumnStyleData(FColumnStyleList[i]).ColWidth := PtsToMM(w);
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
  rowStyle.RowHeight := Workbook.DefaultRowHeight;
  rowStyle.AutoRowHeight := true;
  FRowStyleList.Add(rowStyle);

  for i:=0 to Workbook.GetWorksheetCount-1 do begin
    sheet := Workbook.GetWorksheetByIndex(i);
    for r:=0 to sheet.GetLastRowIndex do begin
      row := sheet.FindRow(r);
      if row <> nil then begin
        h := sheet.GetRowHeight(r);
        // Look for this height in the current RowStyleList
        found := false;
        for j:=0 to FRowStyleList.Count-1 do
          if SameValue(TRowStyleData(FRowStyleList[j]).RowHeight, h, ROWHEIGHT_EPS) and
             (not TRowStyleData(FRowStyleList[j]).AutoRowHeight)
          then begin
            found := true;
            break;
          end;
        // Not found? Then add the row as a new row style
        if not found then begin
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
  for i:=0 to FRowStyleList.Count-1 do begin
    h := (TRowStyleData(FRowStyleList[i]).RowHeight + ROW_HEIGHT_CORRECTION) * multiplier;
    TRowStyleData(FRowStyleList[i]).RowHeight := PtsToMM(h);
  end;
end;

procedure TsSpreadOpenDocWriter.WriteMimetype;
begin
  FMimetype := 'application/vnd.oasis.opendocument.spreadsheet';
end;

procedure TsSpreadOpenDocWriter.WriteMetaInfManifest;
begin
  FMetaInfManifest :=
   XML_HEADER + LineEnding +
   '<manifest:manifest xmlns:manifest="' + SCHEMAS_XMLNS_MANIFEST + '">' + LineEnding +
   '  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:full-path="/" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml" />' + LineEnding +
   '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml" />' + LineEnding +
   '</manifest:manifest>';
end;

procedure TsSpreadOpenDocWriter.WriteMeta;
begin
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
end;

procedure TsSpreadOpenDocWriter.WriteSettings;
begin
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
end;

procedure TsSpreadOpenDocWriter.WriteStyles;
begin
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
   '  '+WriteFontNamesXMLAsString + LineEnding +
//   '  <style:font-face style:name="Arial" svg:font-family="Arial" />' + LineEnding +
   '</office:font-face-decls>' + LineEnding +
   '<office:styles>' + LineEnding +
   '  <style:style style:name="Default" style:family="table-cell">' + LineEnding +
   '    ' + WriteDefaultFontXMLAsString + LineEnding +
   '  </style:style>' + LineEnding +
   '</office:styles>' + LineEnding +
   '<office:automatic-styles>' + LineEnding +
   '  <style:page-layout style:name="pm1">' + LineEnding +
   '    <style:page-layout-properties fo:margin-top="1.25cm" fo:margin-bottom="1.25cm" fo:margin-left="1.905cm" fo:margin-right="1.905cm" />' + LineEnding +
   '    <style:header-style>' + LineEnding +
   '    <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-bottom="0.25cm" fo:margin-top="0cm" />' + LineEnding +
   '    </style:header-style>' + LineEnding +
   '    <style:footer-style>' + LineEnding +
   '    <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0.25cm" fo:margin-bottom="0cm" />' + LineEnding +
   '    </style:footer-style>' + LineEnding +
   '  </style:page-layout>' + LineEnding +
   '</office:automatic-styles>' + LineEnding +
   '<office:master-styles>' + LineEnding +
   '  <style:master-page style:name="Default" style:page-layout-name="pm1">' + LineEnding +
   '    <style:header />' + LineEnding +
   '    <style:header-left style:display="false" />' + LineEnding +
   '    <style:footer />' + LineEnding +
   '    <style:footer-left style:display="false" />' + LineEnding +
   '  </style:master-page>' + LineEnding +
   '</office:master-styles>' + LineEnding +
   '</office:document-styles>';
end;

procedure TsSpreadOpenDocWriter.WriteContent;
var
  i: Integer;
  lCellStylesCode: string;
  lColStylesCode: String;
  lRowStylesCode: String;
begin
  ListAllColumnStyles;
  ListAllRowStyles;
  ListAllFormattingStyles;

  lColStylesCode := WriteColStylesXMLAsString;
  if lColStylesCode = '' then lColStylesCode :=
  '    <style:style style:name="co1" style:family="table-column">' + LineEnding +
  '      <style:table-column-properties fo:break-before="auto" style:column-width="2.267cm"/>' + LineEnding +
  '    </style:style>' + LineEnding;

  lRowStylesCode := WriteRowStylesXMLAsString;
  if lRowStylesCode = '' then lRowStylesCode :=
  '    <style:style style:name="ro1" style:family="table-row">' + LineEnding +
  '      <style:table-row-properties style:row-height="0.416cm" fo:break-before="auto" style:use-optimal-row-height="true"/>' + LineEnding +
  '    </style:style>' + LineEnding;

  lCellStylesCode := WriteCellStylesXMLAsString;

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
   '    ' + WriteFontNamesXMLAsString + LineEnding +
//   '    <style:font-face style:name="Arial" svg:font-family="Arial" xmlns:v="urn:schemas-microsoft-com:vml" />' + LineEnding +
   '  </office:font-face-decls>' + LineEnding +

   // Automatic styles
  '  <office:automatic-styles>' + LineEnding +
  lColStylesCode +
  lRowStylesCode +
  '    <style:style style:name="ta1" style:family="table" style:master-page-name="Default">' + LineEnding +
  '      <style:table-properties table:display="true" style:writing-mode="lr-tb"/>' + LineEnding +
  '    </style:style>' + LineEnding +

  // Automatically Generated Styles
  lCellStylesCode +
  '  </office:automatic-styles>' + LineEnding +

  // Body
  '  <office:body>' + LineEnding +
  '    <office:spreadsheet>' + LineEnding;

  // Write all worksheets
  for i := 0 to Workbook.GetWorksheetCount - 1 do
    WriteWorksheet(Workbook.GetWorksheetByIndex(i));

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
  LastColIndex: Cardinal;
  LCell: TCell;
  AVLNode: TAVLTreeNode;
  defFontSize: Single;
  h, h_mm: Double;
  styleName: String;
  rowStyleData: TRowStyleData;
  row: PRow;
begin
  LastColIndex := CurSheet.GetLastColIndex;
  defFontSize := Workbook.GetFont(0).Size;

  // Header
  FContent := FContent +
  '    <table:table table:name="' + CurSheet.Name + '" table:style-name="ta1">' + LineEnding;

  // columns
  FContent := FContent + WriteColumnsXMLAsString(CurSheet);

  // rows and cells
  // The cells need to be written in order, row by row, cell by cell
  FContent := FContent + WriteRowsAndCellsXMLAsString(CurSheet);

  // Footer
  FContent := FContent +
  '    </table:table>' + LineEnding;
end;

function TsSpreadOpenDocWriter.WriteCellStylesXMLAsString: string;
var
  i: Integer;
  s: String;
begin
  Result := '';

  for i := 0 to Length(FFormattingStyles) - 1 do
  begin
    // Start and Name
    Result := Result +
    '    <style:style style:name="ce' + IntToStr(i) + '" style:family="table-cell" style:parent-style-name="Default">' + LineEnding;

    // Fields

    // style:text-properties
    if uffBold in FFormattingStyles[i].UsedFormattingFields then
      Result := Result +
    '      <style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/>' + LineEnding;

    s := WriteFontStyleXMLAsString(FFormattingStyles[i]);
    if s <> '' then
      Result := Result +
    '      <style:text-properties '+ s + '/>' + LineEnding;

    // style:table-cell-properties
    s :=  WriteBorderStyleXMLAsString(FFormattingStyles[i]) +
          WriteBackgroundColorStyleXMLAsString(FFormattingStyles[i]) +
          WriteWordwrapStyleXMLAsString(FFormattingStyles[i]) +
          WriteTextRotationStyleXMLAsString(FFormattingStyles[i]) +
          WriteVertAlignmentStyleXMLAsString(FFormattingStyles[i]);
    if s <> '' then
      Result := Result +
    '      <style:table-cell-properties ' + s + '/>' + LineEnding;

    // style:paragraph-properties
    s := WriteHorAlignmentStyleXMLAsString(FFormattingStyles[i]);
    if s <> '' then
      Result := Result +
    '      <style:paragraph-properties ' + s + '/>' + LineEnding;

    // End
    Result := Result +
    '    </style:style>' + LineEnding;
  end;
end;

function TsSpreadOpenDocWriter.WriteColStylesXMLAsString: string;
var
  i: Integer;
  s: String;
  colstyle: TColumnStyleData;
begin
  Result := '';

  for i := 0 to FColumnStyleList.Count-1 do begin
    colStyle := TColumnStyleData(FColumnStyleList[i]);

    // Start and Name
    Result := Result +
    '    <style:style style:name="%s" style:family="table-column">' + LineEnding;

    // Column width
    Result := Result +
    '      <style:table-column-properties style:column-width="%.3fmm" fo:break-before="auto"/>' + LineEnding;

    // End
    Result := Result +
    '    </style:style>' + LineEnding;

    Result := Format(Result, [colStyle.Name, colStyle.ColWidth], FPointSeparatorSettings);
  end;
end;

function TsSpreadOpenDocWriter.WriteColumnsXMLAsString(ASheet: TsWorksheet): String;
var
  lastCol: Integer;
  j, k: Integer;
  w, w_mm: Double;
  widthMultiplier: Double;
  styleName: String;
  colsRepeated: Integer;
  colsRepeatedStr: String;
begin
  Result := '';

  widthMultiplier := Workbook.GetFont(0).Size / 2;
  lastCol := ASheet.GetLastColIndex;

  j := 0;
  while (j <= lastCol) do begin
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
      raise Exception.Create('Column style not found.');

    // Determine value for "number-columns-repeated"
    colsRepeated := 1;
    k := j+1;
    while (k <= lastCol) do begin
      if ASheet.GetColWidth(k) = w then
        inc(colsRepeated)
      else
        break;
      inc(k);
    end;
    colsRepeatedStr := IfThen(colsRepeated = 1, '', Format(' table:number-columns-repeated="%d"', [colsRepeated]));

    Result := Result + Format(
    '      <table:table-column table:style-name="%s"%s table:default-cell-style-name="Default"/>',
           [styleName, colsRepeatedStr]) + LineEnding;

    j := j + colsRepeated;
  end;
end;

function TsSpreadOpenDocWriter.WriteRowsAndCellsXMLAsString(ASheet: TsWorksheet): String;
var
  r, rr: Cardinal;  // row index in sheet
  c, cc: Cardinal;  // column index in sheet
  row: PRow;        // sheet row record
  cell: PCell;      // current cell
  styleName: String;
  k: Integer;
  h, h_mm: Single;  // row height in "lines" and millimeters, respectively
  h1: Single;
  colsRepeated: Integer;
  rowsRepeated: Integer;
  colsRepeatedStr: String;
  rowsRepeatedStr: String;
  lastCol, lastRow: Cardinal;
  rowStyleData: TRowStyleData;
  colData: TColumnData;
  colStyleData: TColumnStyleData;
  defFontSize: Single;
  sameRowStyle: Boolean;
begin
  Result := '';

  // some abbreviations...
  lastCol := ASheet.GetLastColIndex;
  lastRow := ASheet.GetLastRowIndex;
  defFontSize := Workbook.GetFont(0).Size;

  // Now loop through all rows
  r := 0;
  while (r <= lastRow) do begin
    // Look for the row style of the current row (r)
    row := ASheet.FindRow(r);
    if row = nil then
      styleName := 'ro1'
    else begin
      styleName := '';

      h := row^.Height;   // row height in "lines"
      h_mm := PtsToMM((h + ROW_HEIGHT_CORRECTION) * defFontSize);  // in mm
      for k := 0 to FRowStyleList.Count-1 do begin
        rowStyleData := TRowStyleData(FRowStyleList[k]);
        // Compare row heights, but be aware of rounding errors
        if SameValue(rowStyleData.RowHeight, h_mm, 1E-3) then begin
          styleName := rowStyleData.Name;
          break;
        end;
      end;
      if styleName = '' then
        raise Exception.Create('Row style not found.');
    end;

    // Look for empty rows with the same style, they need the "number-rows-repeated" element.
    rowsRepeated := 1;
    if ASheet.GetCellCountInRow(r) = 0 then begin
      rr := r + 1;
      while (rr <= lastRow) do begin
        if ASheet.GetCellCountInRow(rr) > 0 then begin
          break;
        end;
        h1 := ASheet.GetRowHeight(rr);
        if not SameValue(h, h1, ROWHEIGHT_EPS) then
          break;
        inc(rr);
      end;
      rowsRepeated := rr - r;
      rowsRepeatedStr := IfThen(rowsRepeated = 1, '',
        Format('table:number-rows-repeated="%d"', [rowsRepeated]));
      colsRepeated := lastCol+1;
      colsRepeatedStr := IfThen(colsRepeated = 1, '',
        Format('table:number-columns-repeated="%d"', [colsRepeated]));
      Result := Result + Format(
        '      <table:table-row table:style-name="%s" %s>' + LineEnding +
        '        <table:table-cell %s/>'                   + LineEnding +
        '      </table:table-row>'                         + LineEnding,
              [styleName, rowsRepeatedStr, colsRepeatedStr]);
      r := rr;
      continue;
    end;

    // Now we know that there are cells.
    // Write the row XML
    Result := Result + Format(
        '      <table:table-row table:style-name="%s">', [styleName]) + LineEnding;

    // Loop along the row and find the cells.
    c := 0;
    while c <= lastCol do begin
      // Get the cell from the sheet
      cell := ASheet.FindCell(r, c);
      // Empty cell? Need to count how many to add "table:number-columns-repeated"
      colsRepeated := 1;
      if cell = nil then begin
        cc := c + 1;
        while (cc <= lastCol) do begin
          cell := ASheet.FindCell(r, cc);
          if cell <> nil then
            break;
          inc(cc)
        end;
        colsRepeated := cc - c;
        colsRepeatedStr := IfThen(colsRepeated = 1, '',
          Format('table:number-columns-repeated="%d"', [colsRepeated]));
        Result := Result + Format(
        '        <table:table-cell %s/>', [colsRepeatedStr]) + LineEnding;
      end
      else begin
        WriteCellCallback(cell, nil);
        Result := Result + FCellContent;
      end;
      inc(c, colsRepeated);
    end;

    Result := Result +
        '      </table:table-row>' + LineEnding;

    // Next row
    inc(r, rowsRepeated);
  end;
end;

function TsSpreadOpenDocWriter.WriteRowStylesXMLAsString: string;
const
  FALSE_TRUE: array[boolean] of string = ('false', 'true');
var
  i: Integer;
  s: String;
  rowstyle: TRowStyleData;
begin
  Result := '';

  for i := 0 to FRowStyleList.Count-1 do begin
    rowStyle := TRowStyleData(FRowStyleList[i]);

    // Start and Name
    Result := Result +
    '    <style:style style:name="%s" style:family="table-row">' + LineEnding;

    // Column width
    Result := Result +
    '      <style:table-row-properties style:row-height="%.3gmm" style:use-optimal-row-height="%s" fo:break-before="auto"/>' + LineEnding;

    // End
    Result := Result +
    '    </style:style>' + LineEnding;

    Result := Format(Result,
      [rowStyle.Name, rowStyle.RowHeight, FALSE_TRUE[rowStyle.AutoRowHeight]],
      FPointSeparatorSettings
    );
  end;
end;


constructor TsSpreadOpenDocWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);

  FColumnStyleList := TFPList.Create;
  FRowStyleList := TFPList.Create;

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';
end;

destructor TsSpreadOpenDocWriter.Destroy;
var
  j: Integer;
begin
  for j:=FColumnStyleList.Count-1 downto 0 do TObject(FColumnStyleList[j]).Free;
  FColumnStyleList.Free;

  for j:=FRowStyleList.Count-1 downto 0 do TObject(FRowStyleList[j]).Free;
  FRowStyleList.Free;
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
procedure TsSpreadOpenDocWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean);
var
  FZip: TZipper;
begin
  { Fill the strings with the contents of the files }

  WriteMimetype();
  WriteMetaInfManifest();
  WriteMeta();
  WriteSettings();
  WriteStyles();
  WriteContent;

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


procedure TsSpreadOpenDocWriter.WriteToStream(AStream: TStream);
begin
  // Not supported at the moment
  raise Exception.Create('TsSpreadOpenDocWriter.WriteToStream not supported');
end;

procedure TsSpreadOpenDocWriter.WriteFormula(AStream: TStream; const ARow,
  ACol: Cardinal; const AFormula: TsFormula; ACell: PCell);
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

{
  Writes an empty cell

  Not clear whether this is needed for ods, but the inherited procedure is abstract.
}
procedure TsSpreadOpenDocWriter.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  lStyle: String = '';
  lIndex: Integer;
begin
  // Write empty cell only if it has formatting
  if ACell^.UsedFormattingFields <> [] then begin
    lIndex := FindFormattingInList(ACell);
    lStyle := ' table:style-name="ce' + IntToStr(lIndex) + '" ';
    FCellContent :=
      '  <table:table-cell ' + lStyle + '>' + LineEnding +
      '  </table:table-cell>' + LineEnding;
  end else
    FCellContent := '';
end;

{ Creates an XML string for inclusion of the background color into the
  written file from the backgroundcolor setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString). }
function TsSpreadOpenDocWriter.WriteBackgroundColorStyleXMLAsString(
  const AFormat: TCell): String;
begin
  Result := '';

  if not (uffBackgroundColor in AFormat.UsedFormattingFields) then
    exit;

  Result := Format('fo:background-color="%s" ', [
    Workbook.GetPaletteColorAsHTMLStr(AFormat.BackgroundColor)
  ]);
//          + Workbook.FPSColorToHexString(FFormattingStyles[i].BackgroundColor, FFormattingStyles[i].RGBBackgroundColor) +'" ';
end;

{ Creates an XML string for inclusion of borders and border styles into the
  written file from the border settings in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString). }
function TsSpreadOpenDocWriter.WriteBorderStyleXMLAsString(const AFormat: TCell): String;
begin
  Result := '';

  if not (uffBorder in AFormat.UsedFormattingFields) then
    exit;

  if cbSouth in AFormat.Border then begin
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

  if cbWest in AFormat.Border then begin
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

  if cbEast in AFormat.Border then begin
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

  if cbNorth in AFormat.Border then begin
    Result := Result + Format('fo:border-top="%s %s %s" ', [
      BORDER_LINEWIDTHS[AFormat.BorderStyles[cbNorth].LineStyle],
      BORDER_LINESTYLES[AFormat.BorderStyles[cbNorth].LineStyle],
      Workbook.GetPaletteColorAsHTMLStr(AFormat.BorderStyles[cbNorth].Color)
    ]);
    if AFormat.BorderStyles[cbSouth].LineStyle = lsDouble then
      Result := Result + 'style:border-linewidth-top="0.002cm 0.035cm 0.002cm" ';
  end else
    Result := Result + 'fo:border-top="none" ';
end;

function TsSpreadOpenDocWriter.WriteDefaultFontXMLAsString: String;
var
  fnt: TsFont;
begin
  fnt := Workbook.GetFont(0);
  Result := Format(
    '<style:text-properties style:font-name="%s" fo:font-size="%.1f" />',
    [fnt.FontName, fnt.Size], FPointSeparatorSettings
  );
end;

function TsSpreadOpenDocWriter.WriteFontNamesXMLAsString: String;
var
  L: TStringList;
  fnt: TsFont;
  i: Integer;
begin
  Result := '';
  L := TStringList.Create;
  try
    for i:=0 to Workbook.GetFontCount-1 do begin
      fnt := Workbook.GetFont(i);
      if (fnt <> nil) and (L.IndexOf(fnt.FontName) = -1) then
        L.Add(fnt.FontName);
    end;
    for i:=0 to L.Count-1 do
      Result := Format(
        '<style:font-face style:name="%s" svg:font-family="%s" />',
        [ L[i], L[i] ]
      );
  finally
    L.Free;
  end;
end;

function TsSpreadOpenDocWriter.WriteFontStyleXMLAsString(const AFormat: TCell): String;
var
  fnt: TsFont;
  defFnt: TsFont;
begin
  Result := '';

  if not (uffFont in AFormat.UsedFormattingFields) then
    exit;

  fnt := Workbook.GetFont(AFormat.FontIndex);
  defFnt := Workbook.GetFont(0);  // Defaultfont

  if fnt.FontName <> defFnt.FontName then
    Result := Result + Format('style:font-name="%s" ', [fnt.FontName]);

  if fnt.Size <> defFnt.Size then
    Result := Result + Format('fo:font-size="%.1fpt" style:font-size-asian="%.1fpt" style:font-size-complex="%.1fpt" ',
      [fnt.Size, fnt.Size, fnt.Size], FPointSeparatorSettings);

  if fssBold in fnt.Style then
    Result := Result + 'fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold" ';

  if fssItalic in fnt.Style then
    Result := Result + 'fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic" ';

  if fssUnderline in fnt.Style then
    Result := Result + 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" ';

  if fssStrikeout in fnt.Style then
    Result := Result + 'style:text-line-through-style="solid" ';

  if fnt.Color <> defFnt.Color then
    Result := Result + Format('fo:color="%s" ', [Workbook.GetPaletteColorAsHTMLStr(fnt.Color)]);
end;

{ Creates an XML string for inclusion of the horizontal alignment into the
  written file from the horizontal alignment setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString). }
function TsSpreadOpenDocWriter.WriteHorAlignmentStyleXMLAsString(
  const AFormat: TCell): String;
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

{ Creates an XML string for inclusion of the textrotation style option into the
  written file from the textrotation setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString). }
function TsSpreadOpenDocWriter.WriteTextRotationStyleXMLAsString(
  const AFormat: TCell): String;
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

{ Creates an XML string for inclusion of the vertical alignment into the
  written file from the vertical alignment setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString). }
function TsSpreadOpenDocWriter.WriteVertAlignmentStyleXMLAsString(
  const AFormat: TCell): String;
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

{ Creates an XML string for inclusion of the wordwrap option into the
  written file from the wordwrap setting in the format cell.
  Is called from WriteStyles (via WriteStylesXMLAsString). }
function TsSpreadOpenDocWriter.WriteWordwrapStyleXMLAsString(
  const AFormat: TCell): String;
begin
  if (uffWordWrap in AFormat.UsedFormattingFields) then
    Result := 'fo:wrap-option="wrap" '
  else
    Result := '';
end;

{
  Writes a cell with text content

  The UTF8 Text needs to be converted, because some chars are invalid in XML
  See bug with patch 19422
}
procedure TsSpreadOpenDocWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
var
  lStyle: string = '';
  lIndex: Integer;
begin
  if ACell^.UsedFormattingFields <> [] then begin
    lIndex := FindFormattingInList(ACell);
    lStyle := ' table:style-name="ce' + IntToStr(lIndex) + '" ';
  end else
    lStyle := '';

  // The row should already be the correct one
  FCellContent :=
    '  <table:table-cell office:value-type="string"' + lStyle + '>' + LineEnding +
    '    <text:p>' + UTF8TextToXMLText(AValue) + '</text:p>' + LineEnding +
    '  </table:table-cell>' + LineEnding;
end;

procedure TsSpreadOpenDocWriter.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
var
  StrValue: string;
  DisplayStr: string;
  lStyle: string = '';
  lIndex: Integer;
begin
  if ACell^.UsedFormattingFields <> [] then begin
    lIndex := FindFormattingInList(ACell);
    lStyle := ' table:style-name="ce' + IntToStr(lIndex) + '" ';
  end else
    lStyle := '';

  // The row should already be the correct one
  if IsInfinite(AValue) then begin
    StrValue:='1.#INF';
    DisplayStr:='1.#INF';
  end else begin
    StrValue:=FloatToStr(AValue,FPointSeparatorSettings); //Uses '.' as decimal separator
    DisplayStr:=FloatToStr(AValue); // Uses locale decimal separator
  end;
  FCellContent :=
    '  <table:table-cell office:value-type="float" office:value="' + StrValue + '"' + lStyle + '>' + LineEnding +
    '    <text:p>' + DisplayStr + '</text:p>' + LineEnding +
    '  </table:table-cell>' + LineEnding;
end;

{*******************************************************************
*  TsSpreadOpenDocWriter.WriteDateTime ()
*
*  DESCRIPTION:    Writes a date/time value
*
*
*******************************************************************}
procedure TsSpreadOpenDocWriter.WriteDateTime(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
var
  lStyle: string = '';
  lIndex: Integer;
begin
  if ACell^.UsedFormattingFields <> [] then begin
    lIndex := FindFormattingInList(ACell);
    lStyle := ' table:style-name="ce' + IntToStr(lIndex) + '" ';
  end else
    lStyle := '';

  // The row should already be the correct one
  FCellContent :=
    '  <table:table-cell office:value-type="date" office:date-value="' + FormatDateTime(ISO8601FormatExtended, AValue) + '"' + lStyle + '>' + LineEnding +
    '  </table:table-cell>' + LineEnding;
end;

{
  Registers this reader / writer on fpSpreadsheet
}
initialization

  RegisterSpreadFormat(TsSpreadOpenDocReader, TsSpreadOpenDocWriter, sfOpenDocument);

end.

