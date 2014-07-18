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

also:
http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx#BMworksheetworkbook

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
    procedure AddBuiltinFormats; override;
  public
    procedure ConvertBeforeWriting(var AFormatString: String;
      var ANumFormat: TsNumberFormat); override;
  end;

  { TsSpreadOOXMLWriter }

  TsSpreadOOXMLWriter = class(TsCustomSpreadWriter)
  private
  protected
    FPointSeparatorSettings: TFormatSettings;
    FSharedStringsCount: Integer;
    FFillList: array of PCell;
    FBorderList: array of PCell;
  protected
    { Helper routines }
    procedure AddDefaultFormats; override;
    procedure CreateNumFormatList; override;
    procedure CreateStreams;
    procedure DestroyStreams;
    function  FindBorderInList(ACell: PCell): Integer;
    function  FindFillInList(ACell: PCell): Integer;
    function GetStyleIndex(ACell: PCell): Cardinal;
    procedure ListAllBorders;
    procedure ListAllFills;
    procedure ResetStreams;
    procedure WriteBorderList(AStream: TStream);
    procedure WriteCols(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteFillList(AStream: TStream);
    procedure WriteFontList(AStream: TStream);
    procedure WriteNumFormatList(AStream: TStream);
    procedure WriteStyleList(AStream: TStream; ANodeName: String);
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
  variants, fpsStreams, fpsNumFormatParser, xlscommon;

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


{ TsOOXMLNumFormatList }

{ These are the built-in number formats as expected in the biff spreadsheet file.
  Identical to BIFF8. These formats are not written to file but they are used
  for lookup of the number format that Excel used. They are specified here in
  fpc dialect. }
procedure TsOOXMLNumFormatList.AddBuiltinFormats;
var
  fs: TFormatSettings;
  cs: String;
begin
  fs := Workbook.FormatSettings;
  cs := AnsiToUTF8(Workbook.FormatSettings.CurrencyString);

  AddFormat( 0, '', nfGeneral);
  AddFormat( 1, '0', nfFixed);
  AddFormat( 2, '0.00', nfFixed);
  AddFormat( 3, '#,##0', nfFixedTh);
  AddFormat( 4, '#,##0.00', nfFixedTh);
  AddFormat( 5, '"'+cs+'"#,##0_);("'+cs+'"#,##0)', nfCurrency);
  AddFormat( 6, '"'+cs+'"#,##0_);[Red]("'+cs+'"#,##0)', nfCurrencyRed);
  AddFormat( 7, '"'+cs+'"#,##0.00_);("'+cs+'"#,##0.00)', nfCurrency);
  AddFormat( 8, '"'+cs+'"#,##0.00_);[Red]("'+cs+'"#,##0.00)', nfCurrencyRed);
  AddFormat( 9, '0%', nfPercentage);
  AddFormat(10, '0.00%', nfPercentage);
  AddFormat(11, '0.00E+00', nfExp);
  // fraction formats 12 ('# ?/?') and 13 ('# ??/??') not supported
  AddFormat(14, fs.ShortDateFormat, nfShortDate);                       // 'M/D/YY'
  AddFormat(15, fs.LongDateFormat, nfLongDate);                         // 'D-MMM-YY'
  AddFormat(16, 'd/mmm', nfCustom);                                     // 'D-MMM'
  AddFormat(17, 'mmm/yy', nfCustom);                                    // 'MMM-YY'
  AddFormat(18, AddAMPM(fs.ShortTimeFormat, fs), nfShortTimeAM);        // 'h:mm AM/PM'
  AddFormat(19, AddAMPM(fs.LongTimeFormat, fs), nfLongTimeAM);          // 'h:mm:ss AM/PM'
  AddFormat(20, fs.ShortTimeFormat, nfShortTime);                       // 'h:mm'
  AddFormat(21, fs.LongTimeFormat, nfLongTime);                         // 'h:mm:ss'
  AddFormat(22, fs.ShortDateFormat + ' ' + fs.ShortTimeFormat, nfShortDateTime);  // 'M/D/YY h:mm' (localized)
  // 23..36 not supported
  AddFormat(37, '_(#,##0_);(#,##0)', nfCurrency);
  AddFormat(38, '_(#,##0_);[Red](#,##0)', nfCurrencyRed);
  AddFormat(39, '_(#,##0.00_);(#,##0.00)', nfCurrency);
  AddFormat(40, '_(#,##0.00_);[Red](#,##0.00)', nfCurrencyRed);
  AddFormat(41, '_("'+cs+'"* #,##0_);_("'+cs+'"* (#,##0);_("'+cs+'"* "-"_);_(@_)', nfCustom);
  AddFormat(42, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', nfCustom);
  AddFormat(43, '_("'+cs+'"* #,##0.00_);_("'+cs+'"* (#,##0.00);_("'+cs+'"* "-"??_);_(@_)', nfCustom);
  AddFormat(44, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', nfCustom);
  AddFormat(45, 'nn:ss', nfCustom);
  AddFormat(46, '[h]:nn:ss', nfTimeInterval);
  AddFormat(47, 'nn:ss.z', nfCustom);
  AddFormat(48, '##0.0E+00', nfCustom);
  // 49 ("Text") not supported

  // All indexes from 0 to 163 are reserved for built-in formats.
  // The first user-defined format starts at 164.
  FFirstFormatIndexInFile := 164;
  FNextFormatIndex := 164;
end;

procedure TsOOXMLNumFormatList.ConvertBeforeWriting(var AFormatString: String;
  var ANumFormat: TsNumberFormat);
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(Workbook, AFormatString, ANumFormat);
  try
    if parser.Status = psOK then begin
      // For writing, we have to convert the fpc format string to Excel dialect
      AFormatString := parser.FormatString[nfdExcel];
      ANumFormat := parser.NumFormat;
    end;
  finally
    parser.Free;
  end;
end;


{ TsSpreadOOXMLWriter }

{ Adds built-in styles:
  - Default style for cells having no specific formatting
  - Bold styles for cells having UsedFormattingFileds = [uffBold]
  All other styles will be added by "ListAllFormattingStyles".
}
procedure TsSpreadOOXMLWriter.AddDefaultFormats();
// We store the index of the XF record that will be assigned to this style in
// the "row" of the style. Will be needed when writing the XF record.
// --- This is needed for BIFF. Not clear if it is important here as well...
var
  len: Integer;
begin
  SetLength(FFormattingStyles, 2);

  // Default style
  FillChar(FFormattingStyles[0], SizeOf(TCell), 0);
  FFormattingStyles[0].BorderStyles := DEFAULT_BORDERSTYLES;
  FFormattingStyles[0].Row := 0;

  // Bold style
  FillChar(FFormattingStyles[1], SizeOf(TCell), 0);
  FFormattingStyles[1].UsedFormattingFields := [uffBold];
  FFormattingStyles[1].FontIndex := 1;  // this is the "bold" font
  FFormattingStyles[1].Row := 1;

  NextXFIndex := 2;
end;

{ Looks for the combination of border attributes of the given cell in the
  FBorderList and returns its index. }
function TsSpreadOOXMLWriter.FindBorderInList(ACell: PCell): Integer;
var
  i: Integer;
  styleCell: PCell;
begin
  // No cell, or border-less --> index 0
  if (ACell = nil) or not (uffBorder in ACell^.UsedFormattingFields) then begin
    Result := 0;
    exit;
  end;

  for i:=0 to High(FBorderList) do begin
    styleCell := FBorderList[i];
    if SameCellBorders(styleCell, ACell) then begin
      Result := i;
      exit;
    end;
  end;

  // Not found --> return -1
  Result := -1;
end;

{ Looks for the combination of fill attributes of the given cell in the
  FFillList and returns its index. }
function TsSpreadOOXMLWriter.FindFillInList(ACell: PCell): Integer;
var
  i: Integer;
  styleCell: PCell;
begin
  if (ACell = nil) or not (uffBackgroundColor in ACell^.UsedFormattingFields)
  then begin
    Result := 0;
    exit;
  end;

  // Index 0 is "no fill" which already has been handled.
  for i:=2 to High(FFillList) do begin
    styleCell := FFillList[i];
    if (uffBackgroundColor in styleCell^.UsedFormattingFields) then
      if (styleCell^.BackgroundColor = ACell^.BackgroundColor) then begin
        Result := i;
        exit;
      end;
  end;

  // Not found --> return -1
  Result := -1;
end;

{ Determines the formatting index which a given cell has in list of
  "FormattingStyles" which correspond to the section cellXfs of the styles.xml
  file. }
function TsSpreadOOXMLWriter.GetStyleIndex(ACell: PCell): Cardinal;
begin
  Result := FindFormattingInList(ACell);
  if Result = -1 then
    Result := 0;
end;

{ Creates a list of all border styles found in the workbook.
  The list contains indexes into the array FFormattingStyles for each unique
  combination of border attributes.
  To be used for the styles.xml. }
procedure TsSpreadOOXMLWriter.ListAllBorders;
var
  styleCell: PCell;
  i, n : Integer;
begin
  // first list entry is a no-border cell
  SetLength(FBorderList, 1);
  FBorderList[0] := nil;

  n := 1;
  for i := 0 to High(FFormattingStyles) do begin
    styleCell := @FFormattingStyles[i];
    if FindBorderInList(styleCell) = -1 then begin
      SetLength(FBorderList, n+1);
      FBorderList[n] := styleCell;
      inc(n);
    end;
  end;
end;

{ Creates a list of all fill styles found in the workbook.
  The list contains indexes into the array FFormattingStyles for each unique
  combination of fill attributes.
  Currently considers only backgroundcolor, fill style is always "solid".
  To be used for styles.xml. }
procedure TsSpreadOOXMLWriter.ListAllFills;
var
  styleCell: PCell;
  i, n: Integer;
begin
  // Add built-in fills first.
  SetLength(FFillList, 2);
  FFillList[0] := nil;  // built-in "no fill"
  FFillList[1] := nil;  // built-in "gray125"

  n := 2;
  for i := 0 to High(FFormattingStyles) do begin
    styleCell := @FFormattingStyles[i];
    if FindFillInList(styleCell) = -1 then begin
      SetLength(FFillList, n+1);
      FFillList[n] := styleCell;
      inc(n);
    end;
  end;
end;

procedure TsSpreadOOXMLWriter.WriteBorderList(AStream: TStream);

  procedure WriteBorderStyle(AStream: TStream; ACell: PCell; ABorder: TsCellBorder);
  { border names found in xlsx files for Excel selections:
    "thin", "hair", "dotted", "dashed", "dashDotDot", "dashDot", "mediumDashDotDot",
    "slantDashDot", "mediumDashDot", "mediumDashed", "medium", "thick", "double" }
  var
    borderName: String;
    styleName: String;
    colorName: String;
    rgb: TsColorValue;
  begin
    // Border line location
    case ABorder of
      cbWest  : borderName := 'left';
      cbEast  : borderName := 'right';
      cbNorth : borderName := 'top';
      cbSouth : borderName := 'bottom';
    end;
    if (ABorder in ACell^.Border) then begin
      // Line style
      case ACell.BorderStyles[ABorder].LineStyle of
        lsThin   : styleName := 'thin';
        lsMedium : styleName := 'medium';
        lsDashed : styleName := 'dashed';
        lsDotted : styleName := 'dotted';
        lsThick  : styleName := 'thick';
        lsDouble : styleName := 'double';
        lsHair   : styleName := 'hair';
        else       raise Exception.Create('TsOOXMLWriter.WriteBorderList: LineStyle not supported.');
      end;
      // Border color
      rgb := Workbook.GetPaletteColor(ACell^.BorderStyles[ABorder].Color);
      colorName := Copy(ColorToHTMLColorStr(rgb), 2, 255);
      AppendToStream(AStream, Format(
        '<%s style="%s"><color rgb="%s" /></%s>',
          [borderName, styleName, colorName, borderName]
        ));
    end else
      AppendToStream(AStream, Format(
        '<%s />', [borderName]));
  end;

var
  i: Integer;
  styleCell: PCell;
begin
  AppendToStream(AStream, Format(
    '<borders count="%d">', [Length(FBorderList)]));

  // index 0 -- build-in "no borders"
  AppendToStream(AStream,
      '<border>',
        '<left /><right /><top /><bottom /><diagonal />',
      '</border>');

  for i:=1 to High(FBorderList) do begin
    styleCell := FBorderList[i];
    AppendToStream(AStream,
      '<border>');
    WriteBorderStyle(AStream, styleCell, cbWest);
    WriteBorderStyle(AStream, styleCell, cbEast);
    WriteBorderStyle(AStream, styleCell, cbNorth);
    WriteBorderStyle(AStream, styleCell, cbSouth);
    AppendToStream(AStream,
        '<diagonal />',
      '</border>');
  end;

  AppendToStream(AStream,
    '</borders>');
end;

procedure TsSpreadOOXMLWriter.WriteCols(AStream: TStream; ASheet: TsWorksheet);
var
  col: PCol;
  c: Integer;
begin
  if ASheet.Cols.Count = 0 then
    exit;

  AppendToStream(AStream,
    '<cols>');

  for c:=0 to ASheet.GetLastColIndex do begin
    col := ASheet.FindCol(c);
    if col <> nil then
      AppendToStream(AStream, Format(
        '<col min="%d" max="%d" width="%g" customWidth="1" />',
        [c+1, c+1, col.Width])
      );
  end;

  AppendToStream(AStream,
    '</cols>');
end;

procedure TsSpreadOOXMLWriter.WriteFillList(AStream: TStream);
var
  i: Integer;
  styleCell: PCell;
  rgb: TsColorValue;
begin
  AppendToStream(AStream, Format(
    '<fills count="%d">', [Length(FFillList)]));

  // index 0 -- built-in empty fill
  AppendToStream(AStream,
      '<fill>',
        '<patternFill patternType="none" />',
      '</fill>');

  // index 1 -- built-in gray125 pattern
  AppendToStream(AStream,
      '<fill>',
        '<patternFill patternType="gray125" />',
      '</fill>');

  // user-defined fills
  for i:=2 to High(FFillList) do begin
    styleCell := FFillList[i];
    rgb := Workbook.GetPaletteColor(styleCell^.BackgroundColor);
    AppendToStream(AStream,
      '<fill>',
        '<patternFill patternType="solid">');
    AppendToStream(AStream, Format(
          '<fgColor rgb="%s" />', [Copy(ColorToHTMLColorStr(rgb), 2, 255)]),
          '<bgColor indexed="64" />');
    AppendToStream(AStream,
        '</patternFill>',
      '</fill>');
  end;

  AppendToStream(FSStyles,
    '</fills>');
end;

{ Writes the fontlist of the workbook to the stream. The font id used in xf
  records is given by the index of a font in the list. Therefore, we have
  to write an empty record for font #4 which is nil due to compatibility with BIFF }
procedure TsSpreadOOXMLWriter.WriteFontList(AStream: TStream);
var
  i: Integer;
  font: TsFont;
  s: String;
  rgb: TsColorValue;
begin
  AppendToStream(FSStyles, Format(
      '<fonts count="%d">', [Workbook.GetFontCount]));
  for i:=0 to Workbook.GetFontCount-1 do begin
    font := Workbook.GetFont(i);
    if font = nil then
      AppendToStream(AStream, '<font />')
      // Font #4 is missing in fpspreadsheet due to BIFF compatibility. We write
      // an empty node to keep the numbers in sync with the stored font index.
    else begin
      s := Format('<sz val="%g" /><name val="%s" />', [font.Size, font.FontName]);
      if (fssBold in font.Style) then
        s := s + '<b />';
      if (fssItalic in font.Style) then
        s := s + '<i />';
      if (fssUnderline in font.Style) then
        s := s + '<u />';
      if (fssStrikeout in font.Style) then
        s := s + '<strike />';
      if font.Color <> scBlack then begin
        rgb := Workbook.GetPaletteColor(font.Color);
        s := s + Format('<color rgb="%s" />', [Copy(ColorToHTMLColorStr(rgb), 2, 255)]);
      end;
      AppendToStream(AStream,
        '<font>', s, '</font>');
    end;
  end;
  AppendToStream(AStream,
      '</fonts>');
end;

{ Writes all number formats to the stream. Saving starts at the item with the
  FirstFormatIndexInFile. }
procedure TsSpreadOOXMLWriter.WriteNumFormatList(AStream: TStream);
var
  i: Integer;
  item: TsNumFormatData;
  s: String;
  n: Integer;
  fmt: String;
begin
  s := '';
  n := 0;
  i := NumFormatList.FindByIndex(NumFormatList.FirstFormatIndexInFile);
  if i > -1 then begin
    while i < NumFormatList.Count do begin
      item := NumFormatList[i];
      if item <> nil then begin
        s := s + Format('<numFmt numFmtId="%d" formatCode="%s" />',
          [item.Index, UTF8TextToXMLText(NumFormatList.FormatStringForWriting(i))]);
        inc(n);
      end;
      inc(i);
    end;
    if n > 0 then
      AppendToStream(AStream, Format(
        '<numFmts count="%d">', [n]),
          s,
        '</numFmts>'
      );
  end;
end;

{ Writes the style list which the writer has collected in FFormattingStyles. }
procedure TsSpreadOOXMLWriter.WriteStyleList(AStream: TStream; ANodeName: String);
var
  styleCell: TCell;
  s, sAlign: String;
  fontID: Integer;
  numFmtId: Integer;
  fillId: Integer;
  borderId: Integer;
  idx: Integer;
begin
  AppendToStream(AStream, Format(
    '<%s count="%d">', [ANodeName, Length(FFormattingStyles)]));

  for styleCell in FFormattingStyles do begin
    s := '';
    sAlign := '';

    { Number format }
    if (uffNumberFormat in styleCell.UsedFormattingFields) then begin
      idx := NumFormatList.FindFormatOf(@styleCell);
      if idx > -1 then begin
        numFmtID := NumFormatList[idx].Index;
        s := s + Format('numFmtId="%d" applyNumberFormat="1" ', [numFmtId]);
      end;
    end;

    { Font }
    fontId := 0;
    if (uffBold in styleCell.UsedFormattingFields) then
      fontId := 1;
    if (uffFont in styleCell.UsedFormattingFields) then
      fontId := styleCell.FontIndex;
    s := s + Format('fontId="%d" ', [fontId]);
    if fontID > 0 then s := s + 'applyFont="1" ';

    if ANodeName = 'cellXfs' then s := s + 'xfId="0" ';

    { Text rotation }
    if (uffTextRotation in styleCell.UsedFormattingFields) or (styleCell.TextRotation <> trHorizontal)
    then
      case styleCell.TextRotation of
        rt90DegreeClockwiseRotation       : sAlign := sAlign + Format('textRotation="%d" ', [180]);
        rt90DegreeCounterClockwiseRotation: sAlign := sAlign + Format('textRotation="%d" ',  [90]);
        rtStacked                         : sAlign := sAlign + Format('textRotation="%d" ', [255]);
      end;

    { Text alignment }
    if (uffHorAlign in styleCell.UsedFormattingFields) or (styleCell.HorAlignment <> haDefault)
    then
      case styleCell.HorAlignment of
        haLeft  : sAlign := sAlign + 'horizontal="left" ';
        haCenter: sAlign := sAlign + 'horizontal="center" ';
        haRight : sAlign := sAlign + 'horizontal="right" ';
      end;

    if (uffVertAlign in styleCell.UsedformattingFields) or (styleCell.VertAlignment <> vaDefault)
    then
      case styleCell.VertAlignment of
        vaTop   : sAlign := sAlign + 'vertical="top" ';
        vaCenter: sAlign := sAlign + 'vertical="center" ';
        vaBottom: sAlign := sAlign + 'vertical="bottom" ';
      end;

    if (uffWordWrap in styleCell.UsedFormattingFields) then
      sAlign := sAlign + 'wrapText="1" ';

    { Fill }
    if (uffBackgroundColor in styleCell.UsedFormattingFields) then begin
      fillID := FindFillInList(@styleCell);
      if fillID = -1 then fillID := 0;
      s := s + Format('fillId="%d" applyFill="1" ', [fillID]);
    end;

    { Border }
    if (uffBorder in styleCell.UsedFormattingFields) then begin
      borderID := FindBorderInList(@styleCell);
      if borderID = -1 then borderID := 0;
      s := s + Format('borderId="%d" applyBorder="1" ', [borderID]);
    end;

    { Write everything to stream }
    if sAlign = '' then
      AppendToStream(AStream,
        '<xf ' + s + '/>')
    else
      AppendToStream(AStream,
       '<xf ' + s + 'applyAlignment="1">',
         '<alignment ' + sAlign + ' />',
       '</xf>');
  end;

  AppendToStream(FSStyles, Format(
    '</%s>', [ANodeName]));
end;

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

  // Number formats
  WriteNumFormatList(FSStyles);

  // Fonts
  WriteFontList(FSStyles);

  // Fill patterns
  WriteFillList(FSStyles);

  // Borders
  WriteBorderList(FSStyles);

  // Style records
  AppendToStream(FSStyles,
      '<cellStyleXfs count="1">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />' +
      '</cellStyleXfs>'
  );
  WriteStyleList(FSStyles, 'cellXfs');

  // Cell style records
  AppendToStream(FSStyles,
      '<cellStyles count="1">' +
        '<cellStyle name="Normal" xfId="0" builtinId="0" />' +
      '</cellStyles>');

  // Misc
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

  AppendToStream(FSWorkbookRels,
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
  ResetStream(FSSharedStrings);
  FSSharedStrings_complete.CopyFrom(FSSharedStrings, FSSharedStrings.Size);
  AppendToStream(FSSharedStrings_complete,
    '</sst>');
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
  styleCell: PCell;
  row: PRow;
  rh: String;
  h0: Single;
begin
  FCurSheetNum := Length(FSSheets);
  SetLength(FSSheets, FCurSheetNum + 1);
  h0 := Workbook.GetDefaultFontSize;  // Point size of default font

  // Create the stream
  if (woSaveMemory in Workbook.WritingOptions) then
    FSSheets[FCurSheetNum] := TBufStream.Create(GetTempFileName('', Format('fpsSH%d', [FCurSheetNum])))
  else
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

  WriteCols(FSSheets[FCurSheetNum], CurSheet);

  AppendToStream(FSSheets[FCurSheetNum],
      '<sheetData>');

  if (woVirtualMode in Workbook.WritingOptions) and Assigned(Workbook.OnNeedCellData)
  then begin
    for r := 0 to Workbook.VirtualRowCount-1 do begin
      row := CurSheet.FindRow(r);
      if row <> nil then
        rh := Format(' ht="%g" customHeight="1"', [
          (row^.Height + ROW_HEIGHT_CORRECTION)*h0])
      else
        rh := '';
      AppendToStream(FSSheets[FCurSheetNum], Format(
        '<row r="%d" spans="1:%d"%s>', [r+1, Workbook.VirtualColCount, rh]));
      for c := 0 to Workbook.VirtualColCount-1 do begin
        FillChar(lCell, SizeOf(lCell), 0);
        CellPosText := CurSheet.CellPosToText(r, c);
        value := varNull;
        styleCell := nil;
        Workbook.OnNeedCellData(Workbook, r, c, value, styleCell);
        if styleCell <> nil then
          lCell := styleCell^;
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
      // If the row has a custom height add this value to the <row> specification
      row := CurSheet.FindRow(r);
      if row <> nil then
        rh := Format(' ht="%g" customHeight="1"', [
          (row^.Height + ROW_HEIGHT_CORRECTION)*h0])
      else
        rh := '';
      AppendToStream(FSSheets[FCurSheetNum], Format(
        '<row r="%d" spans="1:%d"%s>', [r+1, LastColIndex+1, rh]));
      // Write cells belonging to this row.
      for c := 0 to LastColIndex do begin
        LCell.Row := r;
        LCell.Col := c;
        AVLNode := CurSheet.Cells.Find(@LCell);
        if Assigned(AVLNode) then
          WriteCellCallback(PCell(AVLNode.Data), FSSheets[FCurSheetNum])
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
begin
  if (woSaveMemory in Workbook.WritingOptions) then begin
    FSContentTypes := TBufStream.Create(GetTempFileName('', 'fpsCT'));
    FSRelsRels := TBufStream.Create(GetTempFileName('', 'fpsRR'));
    FSWorkbookRels := TBufStream.Create(GetTempFileName('', 'fpsWBR'));
    FSWorkbook := TBufStream.Create(GetTempFileName('', 'fpsWB'));
    FSStyles := TBufStream.Create(GetTempFileName('', 'fpsSTY'));
    FSSharedStrings := TBufStream.Create(GetTempFileName('', 'fpsSS'));
    FSSharedStrings_complete := TBufStream.Create(GetTempFileName('', 'fpsSSC'));
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
  i: Integer;
begin
  ResetStream(FSContentTypes);
  ResetStream(FSRelsRels);
  ResetStream(FSWorkbookRels);
  ResetStream(FSWorkbook);
  ResetStream(FSStyles);
  ResetStream(FSSharedStrings_complete);
  for i := 0 to High(FSSheets) do
    ResetStream(FSSheets[i]);
    {
  FSContentTypes.Position := 0;
  FSRelsRels.Position := 0;
  FSWorkbookRels.Position := 0;
  FSWorkbook.Position := 0;
  FSStyles.Position := 0;
  FSSharedStrings_complete.Position := 0;
  for stream in FSSheets do stream.Position := 0;
  }
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
  { Analyze the workbook and collect all information needed }
  ListAllNumFormats;
  ListAllFormattingStyles;
  ListAllFills;
  ListAllBorders;

  { Create the streams that will hold the file contents }
  CreateStreams;

  { Fill the streams with the contents of the files }
  WriteGlobalFiles;
  WriteContent;

  // Stream position must be at beginning, it was moved to end during adding of xml strings.
  ResetStreams;

  { Now compress the files }
  FZip := TZipper.Create;
  try
    FZip.FileName := '__temp__.tmp';
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
  MAXBYTES = 32767; //limit for this format
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
  if Length(AValue) > MAXBYTES then begin
    TextTooLong := true;
    ResultingValue := Copy(AValue, 1, MAXBYTES); //may chop off multicodepoint UTF8 characters but well...
  end
  else
    ResultingValue:=AValue;

  AppendToStream(FSSharedStrings,
    '<si>' +
      '<t>' + UTF8TextToXMLText(ResultingValue) + '</t>' +
    '</si>');

  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="s"><v>%d</v></c>', [CellPosText, lStyleIndex, FSharedStringsCount]));

  inc(FSharedStringsCount);

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
  lStyleIndex: Integer;
begin
  Unused(AStream, ACell);
  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  CellValueText := Format('%g', [AValue], FPointSeparatorSettings);
  lStyleIndex := GetStyleIndex(ACell);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="n"><v>%s</v></c>', [CellPosText, lStyleIndex, CellValueText]));
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
var
  ExcelDateSerial: double;
begin
  ExcelDateSerial := ConvertDateTimeToExcelDateTime(AValue, dm1900); //FDateMode);
  WriteNumber(AStream, ARow, ACol, ExcelDateSerial, ACell);
//  WriteLabel(AStream, ARow, ACol, FormatDateTime(ISO8601Format, AValue), ACell);
end;

{
  Registers this reader / writer on fpSpreadsheet
}
initialization

  RegisterSpreadFormat(TsCustomSpreadReader, TsSpreadOOXMLWriter, sfOOXML);

end.

