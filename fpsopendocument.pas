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
  xmlread, DOM, AVL_Tree,
  math,
  dateutils,
  fpsutils;
  
type
  TDateMode=(dm1899 {default for ODF; almost same as Excel 1900},
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
    FDateMode: TDateMode;
    // Applies a style to a cell
    procedure ApplyStyleToCell(ARow, ACol: Cardinal; AStyleName: String);
    // Searches a style by its name in the StyleList
    function FindCellStyleByName(AStyleName: String): integer;
    // Searches a column style by its column index or its name in the StyleList
    function FindColumnByCol(AColIndex: Integer): Integer;
    function FindColStyleByName(AStyleName: String): integer;
    // Gets value for the specified attribute. Returns empty string if attribute
    // not found.
    function GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
    procedure ReadCells(ATableNode: TDOMNode);
    procedure ReadColumns(ATableNode: TDOMNode);
    // Figures out the base year for times in this file (dates are unambiguous)
    procedure ReadDateMode(SpreadSheetNode: TDOMNode);
  protected
    procedure CreateNumFormatList; override;
    procedure ReadNumFormats(AStylesNode: TDOMNode);
    procedure ReadStyles(AStylesNode: TDOMNode);
    { Record writing methods }
    procedure ReadBlank(ARow, ACol: Word; ACellNode: TDOMNode);
    procedure ReadFormula(ARow : Word; ACol : Word; ACellNode: TDOMNode);
    procedure ReadLabel(ARow : Word; ACol : Word; ACellNode: TDOMNode);
    procedure ReadNumber(ARow : Word; ACol : Word; ACellNode: TDOMNode);
    procedure ReadDate(ARow : Word; ACol : Word; ACellNode: TDOMNode);
  public
    { General reading methods }
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); override;
  end;

  { TsSpreadOpenDocWriter }

  TsSpreadOpenDocWriter = class(TsCustomSpreadWriter)
  private
    function WriteBackgroundColorStyleXMLAsString(const AFormat: TCell): String;
    function WriteBorderStyleXMLAsString(const AFormat: TCell): String;
    function WriteHorAlignmentStyleXMLAsString(const AFormat: TCell): String;
    function WriteTextRotationStyleXMLAsString(const AFormat: TCell): String;
    function WriteVertAlignmentStyleXMLAsString(const AFormat: TCell): String;
    function WriteWordwrapStyleXMLAsString(const AFormat: TCell): String;
  protected
    FPointSeparatorSettings: TFormatSettings;
    // Strings with the contents of files
    FMeta, FSettings, FStyles, FContent, FMimetype: string;
    FMetaInfManifest: string;
    // Streams with the contents of files
    FSMeta, FSSettings, FSStyles, FSContent, FSMimetype: TStringStream;
    FSMetaInfManifest: TStringStream;
    // Helpers
    procedure CreateNumFormatList; override;
    // Routines to write those files
    procedure WriteMimetype;
    procedure WriteMetaInfManifest;
    procedure WriteMeta;
    procedure WriteSettings;
    procedure WriteStyles;
    procedure WriteContent;
    procedure WriteWorksheet(CurSheet: TsWorksheet);
    // Routines to write parts of those files
    function WriteStylesXMLAsString: string;
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
    ColWidth: Double;
  end;

  { Column data items stored in the ColList of the reader  }
  TColumnData = class
  public
    Col: Integer;
    ColStyleIndex: integer;   // index into FColumnStyleList of reader
    DefaultCellStyleIndex: Integer;   // Index of default cell style in FCellStyleList of reader
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

  for j := FCellStyleList.Count-1 downto 0 do TObject(FCellStyleList[j]).Free;
  FCellStyleList.Free;

  inherited Destroy;
end;

{ Applies the style data referred to by the style name to the specified cell }
procedure TsSpreadOpenDocReader.ApplyStyleToCell(ARow, ACol: Cardinal;
  AStyleName: String);
var
  cell: PCell;
  styleData: TCellStyleData;
  styleIndex: Integer;
  numFmtData: TsNumFormatData;
  i: Integer;
begin
  cell := FWorksheet.GetCell(ARow, ACol);
  if not Assigned(cell) then
    exit;

  // Is there a style attached to the cell?
  styleIndex := -1;
  if AStyleName <> '' then
    styleIndex := FindCellStyleByName(AStyleName);
  if (styleIndex = -1) then begin
    // No - look for the style attached to the column of the cell and
    // find the cell style by the DefaultCellStyleIndex stored in the column list.
    i := FindColumnByCol(ACol);
    if i = -1 then
      exit;
    styleIndex := TColumnData(FColumnList[i]).DefaultCellStyleIndex;
  end;

  styleData := TCellStyleData(FCellStyleList[styleIndex]);

  // Now copy all style parameters from the styleData to the cell.

  // Font
  {
  if style.FontIndex = 1 then
    Include(cell^.UsedFormattingFields, uffBold)
  else
  if XFData.FontIndex > 1 then
    Include(cell^.UsedFormattingFields, uffFont);
  cell^.FontIndex := styleData.FontIndex;
   }

  // Alignment
  cell^.HorAlignment := styleData.HorAlignment;
  cell^.VertAlignment := styleData.VertAlignment;
   // Word wrap
  if styleData.WordWrap then
    Include(cell^.UsedFormattingFields, uffWordWrap)
  else
    Exclude(cell^.UsedFormattingFields, uffWordWrap);
   // Text rotation
  if styleData.TextRotation > trHorizontal then
    Include(cell^.UsedFormattingFields, uffTextRotation)
  else
    Exclude(cell^.UsedFormattingFields, uffTextRotation);
  cell^.TextRotation := styledata.TextRotation;
  // Text alignment
  if styleData.HorAlignment <> haDefault then begin
    Include(cell^.UsedFormattingFields, uffHorAlign);
    cell^.HorAlignment := styleData.HorAlignment;
  end else
    Exclude(cell^.UsedFormattingFields, uffHorAlign);
  if styleData.VertAlignment <> vaDefault then begin
    Include(cell^.UsedFormattingFields, uffVertAlign);
    cell^.VertAlignment := styleData.VertAlignment;
  end else
    Exclude(cell^.UsedFormattingFields, uffVertAlign);
  // Borders
  cell^.BorderStyles := styleData.BorderStyles;
  if styleData.Borders <> [] then begin
    Include(cell^.UsedFormattingFields, uffBorder);
    cell^.Border := styleData.Borders;
  end else
    Exclude(cell^.UsedFormattingFields, uffBorder);
  // Background color
  if styleData.BackgroundColor <> scNotDefined then begin
    Include(cell^.UsedFormattingFields, uffBackgroundColor);
    cell^.BackgroundColor := styleData.BackgroundColor;
  end;

  // Number format
  if styleData.NumFormatIndex > -1 then
    if cell^.ContentType = cctNumber then begin
      numFmtData := NumFormatList[styleData.NumFormatIndex];
      if numFmtData <> nil then begin
        Include(cell^.UsedFormattingFields, uffNumberFormat);
        cell^.NumberFormat := numFmtData.NumFormat;
        cell^.NumberFormatStr := numFmtData.FormatString;
        cell^.Decimals := numFmtData.Decimals;
        cell^.CurrencySymbol := numFmtData.CurrencySymbol;
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

{ Reads the cells in the given table. Loops through all rows, and then finds all
  cells of each row. }
procedure TsSpreadOpenDocReader.ReadCells(ATableNode: TDOMNode);
var
  row: Integer;
  col: Integer;
  cellNode, rowNode: TDOMNode;
  paramValueType, paramFormula, tableStyleName: String;
  paramColsRepeated, paramRowsRepeated: String;
begin
  row := 0;
  rowNode := ATableNode.FindNode('table:table-row');
  while Assigned(rowNode) do begin
    col := 0;

    //process each cell of the row
    cellNode := rowNode.FindNode('table:table-cell');
    while Assigned(cellNode) do begin
      // select this cell value's type
      paramValueType := GetAttrValue(CellNode, 'office:value-type');
      paramFormula := GetAttrValue(CellNode, 'table:formula');
      tableStyleName := GetAttrValue(CellNode, 'table:style-name');

      if paramValueType = 'string' then
        ReadLabel(row, col, cellNode)
      else if (paramValueType = 'float') or (paramValueType = 'percentage') then
        ReadNumber(row, col, cellNode)
      else if (paramValueType = 'date') or (paramValueType = 'time') then
        ReadDate(row, col, cellNode)
      else if (paramValueType = '') and (tableStyleName <> '') then
        ReadBlank(row, col, cellNode)
      else if ParamFormula <> '' then
        ReadLabel(row, col, cellNode);

      paramColsRepeated := GetAttrValue(cellNode, 'table:number-columns-repeated');
      if paramColsRepeated = '' then paramColsRepeated := '1';
      col := col + StrToInt(paramColsRepeated);

      cellNode := cellNode.NextSibling;
    end; //while Assigned(cellNode)

    paramRowsRepeated := GetAttrValue(RowNode, 'table:number-rows-repeated');
    if paramRowsRepeated = '' then paramRowsRepeated := '1';
    row := row + StrToInt(paramRowsRepeated);

    rowNode := rowNode.NextSibling;
  end; // while Assigned(rowNode)
end;

{ Collection columns used in the given table. The columns contain links to
  styles that must be used when cells in that columns are without styles. }
procedure TsSpreadOpenDocReader.ReadColumns(ATableNode: TDOMNode);
var
  col: Integer;
  colNode: TDOMNode;
  s: String;
  colStyleIndex: Integer;
  colStyleData: TColumnStyleData;
  colData: TColumnData;
begin
  col := 0;
  colNode := ATableNode.FindNode('table:table-column');
  while Assigned(colNode) do begin
    if colNode.NodeName = 'table:table-column' then begin;
      s := GetAttrValue(colNode, 'table:style-name');
      colStyleIndex := FindColStyleByName(s);
      if colStyleIndex <> -1 then begin
        colStyleData := TColumnStyleData(FColumnStyleList[colStyleIndex]);
        s := GetAttrValue(ColNode, 'table:default-cell-style-name');
        if s <> '' then begin
          colData := TColumnData.Create;
          colData.Col := col;
          colData.ColStyleIndex := colStyleIndex;
          colData.DefaultCellStyleIndex := FindCellStyleByName(s);
          FColumnList.Add(colData);
        end;
      end;
      s := GetAttrValue(ColNode, 'table:number-columns-repeated');
      if s = '' then
        inc(col)
      else
        inc(col, StrToInt(s));
    end;
    colNode := colNode.NextSibling;
  end;
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

procedure TsSpreadOpenDocReader.ReadFromFile(AFileName: string; AData: TsWorkbook);
var
  Doc : TXMLDocument;
  FilePath : string;
  UnZip : TUnZipper;
  FileList : TStringList;
  BodyNode, SpreadSheetNode, TableNode: TDOMNode;
  StylesNode: TDOMNode;
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
      FWorkSheet := aData.AddWorksheet(GetAttrValue(TableNode,'table:name'));
      // Collect column styles used
      ReadColumns(TableNode);
      // Process each row inside the sheet and process each cell of the row
      ReadCells(TableNode);
      // Continue with next table
      TableNode := TableNode.NextSibling;
    end; //while Assigned(TableNode)

  finally
    Doc.Free;
  end;
end;

procedure TsSpreadOpenDocReader.ReadFormula(ARow: Word; ACol : Word; ACellNode : TDOMNode);
begin
  // For now just read the number
  ReadNumber(ARow, ACol, ACellNode);
end;

procedure TsSpreadOpenDocReader.ReadLabel(ARow: Word; ACol: Word; ACellNode: TDOMNode);
var
  cellText: String;
  styleName: String;
begin
  cellText := UTF8Encode(ACellNode.TextContent);
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
begin
  FSettings := DefaultFormatSettings;
  FSettings.DecimalSeparator:='.';
  Value := GetAttrValue(ACellNode,'office:value');
  if UpperCase(Value)='1.#INF' then
  begin
    FWorkSheet.WriteNumber(Arow,ACol,1.0/0.0);
  end
  else
  begin
    // Don't merge, or else we can't debug
    Str := GetAttrValue(ACellNode,'office:value');
    lNumber := StrToFloat(Str,FSettings);
    FWorkSheet.WriteNumber(ARow,ACol,lNumber);
  end;

  styleName := GetAttrValue(ACellNode, 'table:style-name');
  ApplyStyleToCell(ARow, ACol, stylename);
end;

procedure TsSpreadOpenDocReader.ReadDate(ARow: Word; ACol : Word; ACellNode : TDOMNode);
var
  dt: TDateTime;
  Value: String;
  Fmt : TFormatSettings;
  FoundPos : integer;
  Hours, Minutes, Seconds: integer;
  HoursPos, MinutesPos, SecondsPos: integer;
begin
  // Format expects ISO 8601 type date string or
  // time string
  fmt := DefaultFormatSettings;
  fmt.ShortDateFormat:='yyyy-mm-dd';
  fmt.DateSeparator:='-';
  fmt.LongTimeFormat:='hh:nn:ss';
  fmt.TimeSeparator:=':';
  Value:=GetAttrValue(ACellNode,'office:date-value');
  if Value<>'' then
  begin
    {$IFDEF FPSPREADDEBUG}
        end;
    writeln('Row (1based): ',ARow+1,'office:date-value: '+Value);
    {$ENDIF}
    // Date or date/time string
    Value:=StringReplace(Value,'T',' ',[rfIgnoreCase,rfReplaceAll]);
    // Strip milliseconds?
    FoundPos:=Pos('.',Value);
    if (FoundPos>1) then
    begin
       Value:=Copy(Value,1,FoundPos-1);
    end;
    dt:=StrToDateTime(Value,Fmt);
    FWorkSheet.WriteDateTime(Arow,ACol,dt);
  end
  else
  begin
    // Try time only, e.g. PT23H59M59S
    //                     12345678901
    Value:=GetAttrValue(ACellNode,'office:time-value');
    {$IFDEF FPSPREADDEBUG}
    writeln('Row (1based): ',ARow+1,'office:time-value: '+Value);
    {$ENDIF}
    if (Value<>'') and (Pos('PT',Value)=1) then
    begin
      // Get hours
      HoursPos:=Pos('H',Value);
      if (HoursPos>0) then
        Hours:=StrToInt(Copy(Value,3,HoursPos-3))
      else
        Hours:=0;

      // Get minutes
      MinutesPos:=Pos('M',Value);
      if (MinutesPos>0) and (MinutesPos>HoursPos) then
        Minutes:=StrToInt(Copy(Value,HoursPos+1,MinutesPos-HoursPos-1))
      else
        Minutes:=0;

      // Get seconds
      SecondsPos:=Pos('S',Value);
      if (SecondsPos>0) and (SecondsPos>MinutesPos) then
        Seconds:=StrToInt(Copy(Value,MinutesPos+1,SecondsPos-MinutesPos-1))
      else
        Seconds:=0;

      // Times smaller than a day can be taken as is
      // Times larger than a day depend on the file's date mode.
      // Convert to date/time via Unix timestamp so avoiding limits for number of
      // hours etc in EncodeDateTime. Perhaps there's a faster way of doing this?
      if (Hours>-24) and (Hours<24) then
      begin
        dt:=UnixToDateTime(
          Hours*(MinsPerHour*SecsPerMin)+
          Minutes*(SecsPerMin)+
          Seconds
          )-UnixEpoch;
      end
      else
      begin
        // A day or longer
        case FDateMode of
        dm1899:
        dt:=DATEMODE_1899_BASE+UnixToDateTime(
          Hours*(MinsPerHour*SecsPerMin)+
          Minutes*(SecsPerMin)+
          Seconds
          )-UnixEpoch;
        dm1900:
        dt:=DATEMODE_1900_BASE+UnixToDateTime(
          Hours*(MinsPerHour*SecsPerMin)+
          Minutes*(SecsPerMin)+
          Seconds
          )-UnixEpoch;
        dm1904:
        dt:=DATEMODE_1904_BASE+UnixToDateTime(
          Hours*(MinsPerHour*SecsPerMin)+
          Minutes*(SecsPerMin)+
          Seconds
          )-UnixEpoch;
        end;

      end;
      FWorkSheet.WriteDateTime(Arow,ACol,dt);
    end;
  end;
end;

procedure TsSpreadOpenDocReader.ReadNumFormats(AStylesNode: TDOMNode);
var
  NumFormatNode, node: TDOMNode;
  decs: Integer;
  fmtName: String;
  grouping: boolean;
  fmt: String;
  nf: TsNumberFormat;
  nex: Integer;
  s, s1, s2: String;
begin
  if not Assigned(AStylesNode) then
    exit;

  NumFormatNode := AStylesNode.FirstChild;
  while Assigned(NumFormatNode) do begin
    // Numbers (nfFixed, nfFixedTh, nfExp)
    if NumFormatNode.NodeName = 'number:number-style' then begin
      fmtName := GetAttrValue(NumFormatNode, 'style:name');
      node := NumFormatNode.FindNode('number:number');
      if node <> nil then begin
        s := GetAttrValue(node, 'number:decimal-places');
        if s = '' then
          nf := nfGeneral
        else begin
          decs := StrToInt(s);
          grouping := GetAttrValue(node, 'grouping') = 'true';
          nf := IfThen(grouping, nfFixedTh, nfFixed);
        end;
        fmt := BuildNumberFormatString(nf, Workbook.FormatSettings, decs);
        NumFormatList.AddFormat(fmtName, fmt, nf, decs);
      end;
      node := NumFormatNode.FindNode('number:scientific-number');
      if node <> nil then begin
        nf := nfExp;
        decs := StrToInt(GetAttrValue(node, 'number:decimal-places'));
        nex := StrToInt(GetAttrValue(node, 'number:min-exponent-digits'));
        fmt := BuildNumberFormatString(nfFixed, Workbook.FormatSettings, decs);
        fmt := fmt + 'E+' + DupeString('0', nex);
        NumFormatList.AddFormat(fmtName, fmt, nf, decs);
      end;
    end else
    // Percentage
    if NumFormatNode.NodeName = 'number:percentage-style' then begin
      fmtName := GetAttrValue(NumFormatNode, 'style:name');
      node := NumFormatNode.FindNode('number:number');
      if node <> nil then begin
        nf := nfPercentage;
        decs := StrToInt(GetAttrValue(node, 'number:decimal-places'));
        fmt := BuildNumberFormatString(nf, Workbook.FormatSettings, decs);
        NumFormatList.AddFormat(fmtName, fmt, nf, decs);
      end;
    end else
    // Date/Time
    if (NumFormatNode.NodeName = 'number:date-style') or
       (NumFormatNode.NodeName = 'number:time-style')
    then begin
      fmtName := GetAttrValue(NumFormatNode, 'style:name');
      fmt := '';
      node := NumFormatNode.FirstChild;
      while Assigned(node) do begin
        if node.NodeName = 'number:year' then begin
          s := GetAttrValue(node, 'number:style');
          if s = 'long' then fmt := fmt + 'yyyy'
          else if s = '' then fmt := fmt + 'yy';
        end else
        if node.NodeName = 'number:month' then begin
          s := GetAttrValue(node, 'number:style');
          s1 := GetAttrValue(node, 'number:textual');
          if (s = 'long') and (s1 = 'text') then fmt := fmt + 'mmmm'
          else if (s = '') and (s1 = 'text') then fmt := fmt + 'mmm'
          else if (s = 'long') and (s1 = '') then fmt := fmt + 'mm'
          else if (s = '') and (s1 = '') then fmt := fmt + 'm';
        end else
        if node.NodeName = 'number:day' then begin
          s := GetAttrValue(node, 'number:style');
          s1 := GetAttrValue(node, 'number:textual');
          if (s='long') and (s1 = 'text') then fmt := fmt + 'dddd'
          else if (s='') and (s1 = 'text') then fmt := fmt + 'ddd'
          else if (s='long') and (s1 = '') then fmt := fmt + 'dd'
          else if (s='') and (s1='') then fmt := Fmt + 'd';
        end else
        if node.NodeName = 'number:day-of-week' then
          fmt := fmt + 'ddddd'
        else
        if node.NodeName = 'number:hours' then begin
          s := GetAttrValue(node, 'number:style');
          s1 := GetAttrValue(node, 'number:truncate-on-overflow');
          if (s='long') and (s1='false') then fmt := fmt + '[hh]'
          else if (s='long') and (s1='') then fmt := fmt + 'hh'
          else if (s='') and (s1='false') then fmt := fmt + '[h]'
          else if (s='') and (s1='') then fmt := fmt + 'h';
        end else
        if node.NodeName = 'number:minutes' then begin
          s := GetAttrValue(node, 'number:style');
          s1 := GetAttrValue(node, 'number:truncate-on-overflow');
          if (s='long') and (s1='false') then fmt := fmt + '[nn]'
          else if (s='long') and (s1='') then fmt := fmt + 'nn'
          else if (s='') and (s1='false') then fmt := fmt + '[n]'
          else if (s='') and (s1='') then fmt := fmt + 'n';
        end else
        if node.NodeName = 'number:seconds' then begin
          s := GetAttrValue(node, 'number:style');
          s1 := GetAttrValue(node, 'number:truncate-on-overflow');
          s2 := GetAttrValue(node, 'number:decimal-places');
          if (s='long') and (s1='false') then fmt := fmt + '[ss]'
          else if (s='long') and (s1='') then fmt := fmt + 'ss'
          else if (s='') and (s1='false') then fmt := fmt + '[s]'
          else if (s='') and (s1='') then fmt := fmt + 's';
          if (s2 <> '') and (s2 <> '0') then fmt := fmt + '.' + DupeString('0', StrToInt(s2));
        end else
        if node.NodeName = 'number:am-pm' then
          fmt := fmt + 'AM/PM'
        else
        if node.NodeName = 'number:text' then
          fmt := fmt + node.TextContent;
        node := node.NextSibling;
      end;
      NumFormatList.AddFormat(fmtName, fmt, nfFmtDateTime);
    end;
    NumFormatNode := NumFormatNode.NextSibling;
  end;
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
    if styleNode.NodeName = 'style:style' then begin
      family := GetAttrValue(styleNode, 'style:family');

      // Column styles
      if family = 'table-column' then begin
        styleName := GetAttrValue(styleNode, 'style:name');
        styleChildNode := styleNode.FirstChild;
        colWidth := -1;
        while Assigned(styleChildNode) do begin
          if styleChildNode.NodeName = 'style:table-column-properties' then begin
            s := GetAttrValue(styleChildNode, 'style:column-width');
            if s <> '' then begin
              s := Copy(s, 1, Length(s)-2);    // TO DO: use correct units!
              colWidth := StrToFloat(s, fs);
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

        styleChildNode := styleNode.FirstChild;
        while Assigned(styleChildNode) do begin
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
        style.FontIndex := 0;
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
   '  <style:font-face style:name="Arial" svg:font-family="Arial" />' + LineEnding +
   '</office:font-face-decls>' + LineEnding +
   '<office:styles>' + LineEnding +
   '  <style:style style:name="Default" style:family="table-cell">' + LineEnding +
   '    <style:text-properties fo:font-size="10" style:font-name="Arial" />' + LineEnding +
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
  lStylesCode: string;
begin
  ListAllFormattingStyles;

  lStylesCode := WriteStylesXMLAsString;

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
  // Automatically Generated Styles
  lStylesCode +
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
begin
  LastColIndex := CurSheet.GetLastColIndex;

  // Header
  FContent := FContent +
  '    <table:table table:name="' + CurSheet.Name + '" table:style-name="ta1">' + LineEnding +
  '      <table:table-column table:style-name="co1" table:number-columns-repeated="' +
  IntToStr(LastColIndex + 1) + '" table:default-cell-style-name="Default"/>' + LineEnding;

  // The cells need to be written in order, row by row, cell by cell
  for j := 0 to CurSheet.GetLastRowIndex do
  begin
    FContent := FContent +
    '      <table:table-row table:style-name="ro1">' + LineEnding;

    // Write cells from this row.
    for k := 0 to LastColIndex do
    begin
      LCell.Row := j;
      LCell.Col := k;
      AVLNode := CurSheet.Cells.Find(@LCell);
      if Assigned(AVLNode) then
        WriteCellCallback(PCell(AVLNode.Data), nil)
      else
        FContent := FContent + '<table:table-cell/>' + LineEnding;
    end;

    FContent := FContent +
    '      </table:table-row>' + LineEnding;
  end;

  // Footer
  FContent := FContent +
  '    </table:table>' + LineEnding;
end;

function TsSpreadOpenDocWriter.WriteStylesXMLAsString: string;
var
  i: Integer;
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

    // style:table-cell-properties
    if (FFormattingStyles[i].UsedFormattingFields *
      [uffBorder, uffBackgroundColor, uffWordWrap, uffTextRotation, uffVertAlign] <> [])
    then begin
      Result := Result +
    '      <style:table-cell-properties ' +
                WriteBorderStyleXMLAsString(FFormattingStyles[i]) +
                WriteBackgroundColorStyleXMLAsString(FFormattingStyles[i]) +
                WriteWordwrapStyleXMLAsString(FFormattingStyles[i]) +
                WriteTextRotationStyleXMLAsString(FFormattingStyles[i]) +
                WriteVertAlignmentStyleXMLAsString(FFormattingStyles[i]) +
                '/>' + LineEnding;
    end;

    // style:paragraph-properties
    if (uffHorAlign in FFormattingStyles[i].UsedFormattingFields) and
       (FFormattingStyles[i].HorAlignment <> haDefault)
    then begin
      Result := Result +
    '      <style:paragraph-properties ' +
              WriteHorAlignmentStyleXMLAsString(FFormattingStyles[i]) +
              '/>' + LineEnding;
    end;


    // End
    Result := Result +
    '    </style:style>' + LineEnding;
  end;
end;

constructor TsSpreadOpenDocWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';
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
    FContent := FContent +
      '  <table:table-cell ' + lStyle + '>' + LineEnding +
      '  </table:table-cell>' + LineEnding;
  end;
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
  if ACell^.UsedFormattingFields <> [] then
  begin
    lIndex := FindFormattingInList(ACell);
    lStyle := ' table:style-name="ce' + IntToStr(lIndex) + '" ';
  end;

  // The row should already be the correct one
  FContent := FContent +
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
  if ACell^.UsedFormattingFields <> [] then
  begin
    lIndex := FindFormattingInList(ACell);
    lStyle := ' table:style-name="ce' + IntToStr(lIndex) + '" ';
  end;

  // The row should already be the correct one
  if IsInfinite(AValue) then begin
    StrValue:='1.#INF';
    DisplayStr:='1.#INF';
  end else begin
    StrValue:=FloatToStr(AValue,FPointSeparatorSettings); //Uses '.' as decimal separator
    DisplayStr:=FloatToStr(AValue); // Uses locale decimal separator
  end;
  FContent := FContent +
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
  if ACell^.UsedFormattingFields <> [] then
  begin
    lIndex := FindFormattingInList(ACell);
    lStyle := ' table:style-name="ce' + IntToStr(lIndex) + '" ';
  end;

  // The row should already be the correct one
  FContent := FContent +
    '  <table:table-cell office:value-type="date" office:date-value="' + FormatDateTime(ISO8601FormatExtended, AValue) + '"' + lStyle + '>' + LineEnding +
    '  </table:table-cell>' + LineEnding;
end;

{
  Registers this reader / writer on fpSpreadsheet
}
initialization

  RegisterSpreadFormat(TsSpreadOpenDocReader, TsSpreadOpenDocWriter, sfOpenDocument);

end.

