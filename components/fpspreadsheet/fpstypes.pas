unit fpsTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpimage;

type
  {@@ File formats supported by fpspreadsheet }
  TsSpreadsheetFormat = (sfExcel2, sfExcel5, sfExcel8,
   sfOOXML, sfOpenDocument, sfCSV, sfWikiTable_Pipes, sfWikiTable_WikiMedia);

  {@@ Record collection limitations of a particular file format }
  TsSpreadsheetFormatLimitations = record
    MaxRowCount: Cardinal;
    MaxColCount: Cardinal;
    MaxPaletteSize: Integer;
  end;

const
  {@@ Default binary <b>Excel</b> file extension}
  STR_EXCEL_EXTENSION = '.xls';
  {@@ Default xml <b>Excel</b> file extension (>= Excel 2007) }
  STR_OOXML_EXCEL_EXTENSION = '.xlsx';
  {@@ Default <b>OpenDocument</b> spreadsheet file extension }
  STR_OPENDOCUMENT_CALC_EXTENSION = '.ods';
  {@@ Default extension of <b>comma-separated-values</b> file }
  STR_COMMA_SEPARATED_EXTENSION = '.csv';
  {@@ Default extension of <b>wikitable files</b> in <b>pipes</b> format}
  STR_WIKITABLE_PIPES = '.wikitable_pipes';
  {@@ Default extension of <b>wikitable files</b> in <b>wikimedia</b> format }
  STR_WIKITABLE_WIKIMEDIA = '.wikitable_wikimedia';

  {@@ Maximum count of worksheet columns}
  MAX_COL_COUNT = 65535;

  {@@ Name of the default font}
  DEFAULT_FONTNAME = 'Arial';
  {@@ Size of the default font}
  DEFAULT_FONTSIZE = 10;
  {@@ Index of the default font in workbook's font list }
  DEFAULT_FONTINDEX = 0;
  {@@ Index of the hyperlink font in workbook's font list }
  HYPERLINK_FONTINDEX = 6;

  {@@ Takes account of effect of cell margins on row height by adding this
      value to the nominal row height. Note that this is an empirical value
      and may be wrong. }
  ROW_HEIGHT_CORRECTION = 0.2;


type
                       (*
  {@@ Possible encodings for a non-unicode encoded text }
  TsEncoding = (
    seLatin1,
    seLatin2,
    seCyrillic,
    seGreek,
    seTurkish,
    seHebrew,
    seArabic,
    seUTF16
    );            *)

  {@@ Tokens to identify the <b>elements in an expanded formula</b>.

   NOTE: When adding or rearranging items
   * make sure that the subtypes TOperandTokens and TBasicOperationTokens
     are complete
   * make sure to keep the table "TokenIDs" in unit xlscommon in sync
  }
  TFEKind = (
    { Basic operands }
    fekCell, fekCellRef, fekCellRange, fekCellOffset, fekNum, fekInteger,
    fekString, fekBool, fekErr, fekMissingArg,
    { Basic operations }
    fekAdd, fekSub, fekMul, fekDiv, fekPercent, fekPower, fekUMinus, fekUPlus,
    fekConcat,  // string concatenation
    fekEqual, fekGreater, fekGreaterEqual, fekLess, fekLessEqual, fekNotEqual,
    fekParen,   // show parenthesis around expression node
    { Functions - they are identified by their name }
    fekFunc
  );

  {@@ These tokens identify operands in RPN formulas. }
  TOperandTokens = fekCell..fekMissingArg;

  {@@ These tokens identify basic operations in RPN formulas. }
  TBasicOperationTokens = fekAdd..fekParen;

type
  {@@ Flags to mark the address or a cell or a range of cells to be <b>absolute</b>
      or <b>relative</b>. They are used in the set TsRelFlags. }
  TsRelFlag = (rfRelRow, rfRelCol, rfRelRow2, rfRelCol2);

  {@@ Flags to mark the address of a cell or a range of cells to be <b>absolute</b>
      or <b>relative</b>. It is a set consisting of TsRelFlag elements. }
  TsRelFlags = set of TsRelFlag;

const
  {@@ Abbreviation of all-relative cell reference flags }
  rfAllRel = [rfRelRow, rfRelCol, rfRelRow2, rfRelCol2];

  {@@ Separator between worksheet name and cell (range) reference in an address }
  SHEETSEPARATOR = '!';

type
  {@@ Elements of an expanded formula.
    Note: If ElementKind is fekCellOffset, "Row" and "Col" have to be cast
          to signed integers! }
  TsFormulaElement = record
    ElementKind: TFEKind;
    Row, Row2: Cardinal;   // zero-based
    Col, Col2: Cardinal;   // zero-based
//    Param1, Param2: Word;  // Extra parameters
    DoubleValue: double;
    IntValue: Word;
    StringValue: String;
    RelFlags: TsRelFlags;  // store info on relative/absolute addresses
    FuncName: String;
    ParamsNum: Byte;
  end;

  {@@ RPN formula. Similar to the expanded formula, but in RPN notation.
      Simplifies the task of format writers which need RPN }
  TsRPNFormula = array of TsFormulaElement;

  {@@ Describes the <b>type of content</b> in a cell of a TsWorksheet }
  TCellContentType = (cctEmpty, cctFormula, cctNumber, cctUTF8String,
    cctDateTime, cctBool, cctError, cctHyperlink);

  {@@ The record TsComment describes a comment attached to a cell.
     @param   Row        (0-based) row index of the cell
     @param   Col        (0-based) column index of the cell
     @param   Text       Comment text }
  TsComment = record
    Row, Col: Cardinal;
    Text: String;
  end;

  {@@ Pointer to a TsComment record }
  PsComment = ^TsComment;

  {@@ Specifies whether a hyperlink refers to a cell address within the current
      workbook, an external file, or a URL }
  TsHyperlinkKind = (hkNone, hkCell, hkFile, hkURL);

  {@@ The record TsHyperlink contains info on a hyperlink in a cell
    @param   Row          Row index of the cell containing the hyperlink
    @param   Col          Column index of the cell containing the hyperlink
    @param   Kind         Specifies whether clicking on the hyperlink results in
                          jumping the a cell address within the current workbook,
                          opens a file, or opens a URL
    @param   Destination  Hyperlink (cell address, filename, URL)
    @param   Note         Text displayed as a popup hint by Excel }
  TsHyperlink = record
    Row, Col: Cardinal;
    Kind: TsHyperlinkKind;
    Destination: String;
    Note: String;
  end;

  {@@ Pointer to a TsHyperlink record }
  PsHyperlink = ^TsHyperlink;

  {@@ Callback function, e.g. for iterating the internal AVL trees of the workbook/sheet}
  TsCallback = procedure (data, arg: Pointer) of object;

  {@@ Error code values }
  TsErrorValue = (
    errOK,                 // no error
    errEmptyIntersection,  // #NULL!
    errDivideByZero,       // #DIV/0!
    errWrongType,          // #VALUE!
    errIllegalRef,         // #REF!
    errWrongName,          // #NAME?
    errOverflow,           // #NUM!
    errArgError,           // #N/A
    // --- no Excel errors --
    errFormulaNotSupported
  );

  {@@ List of possible formatting fields }
  TsUsedFormattingField = (uffTextRotation, uffFont, uffBold, uffBorder,
    uffBackground, uffNumberFormat, uffWordWrap, uffHorAlign, uffVertAlign
  );
  { NOTE: "uffBackgroundColor" of older versions replaced by "uffBackground" }

  {@@ Describes which formatting fields are active }
  TsUsedFormattingFields = set of TsUsedFormattingField;

  {@@ Number/cell formatting. Only uses a subset of the default formats,
      enough to be able to read/write date/time values.
      nfCustom allows to apply a format string directly. }
  TsNumberFormat = (
    // general-purpose for all numbers
    nfGeneral,
    // numbers
    nfFixed, nfFixedTh, nfExp, nfPercentage,
    // currency
    nfCurrency, nfCurrencyRed,
    // dates and times
    nfShortDateTime, {nfFmtDateTime, }nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfTimeInterval,
    // other (format string goes directly into the file)
    nfCustom);

  {@@ Identifies which "dialect" is used in the format strings:
    nfdDefault is the dialect used by fpc
    fndExcel is the dialect used by Excel }
  TsNumFormatDialect = (nfdDefault, nfdExcel);

const
  { @@ Codes for curreny format according to FormatSettings.CurrencyFormat:
       "C" = currency symbol, "V" = currency value, "S" = space character
       For the negative value formats, we use also:
       "B" = bracket, "M" = Minus

       The order of these characters represents the order of these items.

       Example: 1000 dollars  --> "$1000"  for pCV,   or "1000 $"  for pVsC
               -1000 dollars --> "($1000)" for nbCVb, or "-$ 1000" for nMCSV

       Assignment taken from "sysstr.inc" }
  pcfDefault = -1;   // use value from Worksheet.FormatSettings.CurrencyFormat
  pcfCV      = 0;    // $1000
  pcfVC      = 1;    // 1000$
  pcfCSV     = 2;    // $ 1000
  pcfVSC     = 3;    // 1000 $

  ncfDefault = -1;   // use value from Worksheet.FormatSettings.NegCurrFormat
  ncfBCVB    = 0;    // ($1000)
  ncfMCV     = 1;    // -$1000
  ncfCMV     = 2;    // $-1000
  ncfCVM     = 3;    // $1000-
  ncfBVCB    = 4;    // (1000$)
  ncfMVC     = 5;    // -1000$
  ncfVMC     = 6;    // 1000-$
  ncfVCM     = 7;    // 1000$-
  ncfMVSC    = 8;    // -1000 $
  ncfMCSV    = 9;    // -$ 1000
  ncfVSCM    = 10;   // 1000 $-
  ncfCSVM    = 11;   // $ 1000-
  ncfCSMV    = 12;   // $ -1000
  ncfVMSC    = 13;   // 1000- $
  ncfBCSVB   = 14;   // ($ 1000)
  ncfBVSCB   = 15;   // (1000 $)

type
  {@@ Text rotation formatting. The text is rotated relative to the standard
      orientation, which is from left to right horizontal:
      <pre>
       --->
       ABC </pre>

      So 90 degrees clockwise means that the text will be:
      <pre>
       |  A
       |  B
       v  C </pre>

      And 90 degree counter clockwise will be:
      <pre>
       ^  C
       |  B
       |  A</pre>

      Due to limitations of the text mode the characters are not rotated here.
      There is, however, also a "stacked" variant which looks exactly like
      the 90-degrees-clockwise case.
  }
  TsTextRotation = (trHorizontal, rt90DegreeClockwiseRotation,
    rt90DegreeCounterClockwiseRotation, rtStacked);

  {@@ Indicates horizontal text alignment in cells }
  TsHorAlignment = (haDefault, haLeft, haCenter, haRight);

  {@@ Indicates vertical text alignment in cells }
  TsVertAlignment = (vaDefault, vaTop, vaCenter, vaBottom);

  {@@
    Colors in fpspreadsheet are given as indices into a palette.
    Use the workbook's GetPaletteColor to determine the color rgb value as
    little-endian (with "r" being the low-value byte, in agreement with TColor).
    The data type for rgb values is TsColorValue. }
  TsColor = Word;

{@@
  These are some constants for color indices into the default palette.
  Note, however, that if a different palette is used there may be more colors,
  and the names of the color constants may no longer be correct.
}
const
  {@@ Index of <b>black</b> color in the standard color palettes }
  scBlack = $00;
  {@@ Index of <b>white</b> color in the standard color palettes }
  scWhite = $01;
  {@@ Index of <b>red</b> color in the standard color palettes }
  scRed = $02;
  {@@ Index of <b>green</b> color in the standard color palettes }
  scGreen = $03;
  {@@ Index of <b>blue</b> color in the standard color palettes }
  scBlue = $04;
  {@@ Index of <b>yellow</b> color in the standard color palettes }
  scYellow = $05;
  {@@ Index of <b>magenta</b> color in the standard color palettes }
  scMagenta = $06;
  {@@ Index of <b>cyan</b> color in the standard color palettes }
  scCyan = $07;
  {@@ Index of <b>dark red</b> color in the standard color palettes }
  scDarkRed = $08;
  {@@ Index of <b>dark green</b> color in the standard color palettes }
  scDarkGreen = $09;
  {@@ Index of <b>dark blue</b> color in the standard color palettes }
  scDarkBlue = $0A;
  {@@ Index of <b>"navy"</b> color (dark blue) in the standard color palettes }
  scNavy = $0A;
  {@@ Index of <b>olive</b> color in the standard color palettes }
  scOlive = $0B;
  {@@ Index of <b>purple</b> color in the standard color palettes }
  scPurple = $0C;
  {@@ Index of <b>teal</b> color in the standard color palettes }
  scTeal = $0D;
  {@@ Index of <b>silver</b> color in the standard color palettes }
  scSilver = $0E;
  {@@ Index of <b>grey</b> color in the standard color palettes }
  scGrey = $0F;
  {@@ Index of <b>gray</b> color in the standard color palettes }
  scGray = $0F;       // redefine to allow different spelling
  {@@ Index of a <b>10% grey</b> color in the standard color palettes }
  scGrey10pct = $10;
  {@@ Index of a <b>10% gray</b> color in the standard color palettes }
  scGray10pct = $10;
  {@@ Index of a <b>20% grey</b> color in the standard color palettes }
  scGrey20pct = $11;
  {@@ Index of a <b>20% gray</b> color in the standard color palettes }
  scGray20pct = $11;
  {@@ Index of <b>orange</b> color in the standard color palettes }
  scOrange = $12;
  {@@ Index of <b>dark brown</b> color in the standard color palettes }
  scDarkbrown = $13;
  {@@ Index of <b>brown</b> color in the standard color palettes }
  scBrown = $14;
  {@@ Index of <b>beige</b> color in the standard color palettes }
  scBeige = $15;
  {@@ Index of <b>"wheat"</b> color (yellow-orange) in the standard color palettes }
  scWheat = $16;

  // not sure - but I think the mechanism with scRGBColor is not working...
  // Will be removed sooner or later...
  scRGBColor = $FFFD;

  {@@ Identifier for transparent color }
  scTransparent = $FFFE;
  {@@ Identifier for not-defined color }
  scNotDefined = $FFFF;

type
  {@@ Data type for rgb color values }
  TsColorValue = DWord;

  {@@ Palette of color values. A "color value" is a DWord value containing
      rgb colors. }
  TsPalette = array[0..0] of TsColorValue;
  PsPalette = ^TsPalette;

  {@@ Font style (redefined to avoid usage of "Graphics" }
  TsFontStyle = (fssBold, fssItalic, fssStrikeOut, fssUnderline);

  {@@ Set of font styles }
  TsFontStyles = set of TsFontStyle;

  {@@ Font record used in fpspreadsheet. Contains the font name, the font size
      (in points), the font style, and the font color. }
  TsFont = class
    {@@ Name of the font face, such as 'Arial' or 'Times New Roman' }
    FontName: String;
    {@@ Size of the font in points }
    Size: Single;   // in "points"
    {@@ Font style, such as bold, italics etc. - see TsFontStyle}
    Style: TsFontStyles;
    {@@ Text color given by the index into the workbook's color palette }
    Color: TsColor;
  end;

  {@@ Indicates the border for a cell. If included in the CellBorders set the
      corresponding border is drawn in the style defined by the CellBorderStyle. }
  TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth, cbDiagUp, cbDiagDown);

  {@@ Indicates the border for a cell }
  TsCellBorders = set of TsCellBorder;

  {@@ Line style (for cell borders) }
  TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair);

  {@@ The Cell border style reocrd contains the linestyle and color of a cell
      border. There is a CellBorderStyle for each border. }
  TsCellBorderStyle = record
    LineStyle: TsLineStyle;
    Color: TsColor;
  end;

  {@@ The cell border styles of each cell border are collected in this array. }
  TsCellBorderStyles = array[TsCellBorder] of TsCellBorderStyle;

  {@@ Border styles for each cell border used by default: a thin, black, solid line }
const
  DEFAULT_BORDERSTYLES: TsCellBorderStyles = (
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack)
  );

type
  {@@ Style of fill pattern for cell backgrounds }
  TsFillStyle = (fsNoFill, fsSolidFill,
    fsGray75, fsGray50, fsGray25, fsGray12, fsGray6,
    fsStripeHor, fsStripeVert, fsStripeDiagUp, fsStripeDiagDown,
    fsThinStripeHor, fsThinStripeVert, fsThinStripeDiagUp, fsThinStripeDiagDown,
    fsHatchDiag, fsThinHatchDiag, fsThickHatchDiag, fsThinHatchHor);

  {@@ Fill pattern record }
  TsFillPattern = record
    Style: TsFillStyle;  // pattern type
    FgColor: TsColor;    // pattern color
    BgColor: TsColor;    // background color
  end;

const
  {@@ Parameters for a non-filled cell background }
  EMPTY_FILL: TsFillPattern = (
    Style: fsNoFill;
    FgColor: scTransparent;
    BgColor: scTransparent;
  );

type
  {@@ Identifier for a compare operation }
  TsCompareOperation = (coNotUsed,
    coEqual, coNotEqual, coLess, coGreater, coLessEqual, coGreaterEqual
  );

  {@@ Cell calculation state }
  TsCalcState = (csNotCalculated, csCalculating, csCalculated);

  {@@ Cell flag }
  TsCellFlag = (cfCalculating, cfCalculated, cfHasComment, cfMerged);

  {@@ Set of cell flags }
  TsCellFlags = set of TsCellFlag;

  {@@ Record combining a cell's row and column indexes }
  TsCellCoord = record
    Row, Col: Cardinal;
  end;

  {@@ Record combining row and column cornder indexes of a range of cells }
  TsCellRange = record
    Row1, Col1, Row2, Col2: Cardinal;
  end;
  PsCellRange = ^TsCellRange;

  {@@ Array with cell ranges }
  TsCellRangeArray = array of TsCellRange;

  {@@ Options for sorting }
  TsSortOption = (ssoDescending, ssoCaseInsensitive);
  {@@ Set of options for sorting }
  TsSortOptions = set of TsSortOption;

  {@@ Sort priority }
  TsSortPriority = (spNumAlpha, spAlphaNum);   // spNumAlpha: Number < Text

  {@@ Sort key: sorted column or row index and sort direction }
  TsSortKey = record
    ColRowIndex: Integer;
    Options: TsSortOptions;
  end;

  {@@ Array of sort keys for multiple sorting criteria }
  TsSortKeys = array of TsSortKey;

  {@@ Complete set of sorting parameters
    @param SortByCols  If true sorting is top-down, otherwise left-right
    @param Priority    Determines whether numbers are before or after text.
    @param SortKeys    Array of sorting indexes and sorting directions }
  TsSortParams = record
    SortByCols: Boolean;
    Priority: TsSortPriority;
    Keys: TsSortKeys;
  end;

  {@@ Record containing all details for cell formatting }
  TsCellFormat = record
    Name: String;
    ID: Integer;
    UsedFormattingFields: TsUsedFormattingFields;
    FontIndex: Integer;
    TextRotation: TsTextRotation;
    HorAlignment: TsHorAlignment;
    VertAlignment: TsVertAlignment;
    Border: TsCellBorders;
    BorderStyles: TsCelLBorderStyles;
    Background: TsFillPattern;
    NumberFormat: TsNumberFormat;
    NumberFormatStr: String;
  end;

  {@@ Pointer to a format record }
  PsCellFormat = ^TsCellFormat;

  {@@ Specialized list for format records }
  TsCellFormatList = class(TFPList)
  private
    FAllowDuplicates: Boolean;
    function GetItem(AIndex: Integer): PsCellFormat;
    procedure SetItem(AIndex: Integer; const AValue: PsCellFormat);
  public
    constructor Create(AAllowDuplicates: Boolean);
    destructor Destroy; override;
    function Add(const AItem: TsCellFormat): Integer; overload;
    function Add(AItem: PsCellFormat): Integer; overload;
    procedure Clear;
    procedure Delete(AIndex: Integer);
    function FindIndexOfID(ID: Integer): Integer;
    function FindIndexOfName(AName: String): Integer;
    function IndexOf(const AItem: TsCellFormat): Integer; overload;
    property Items[AIndex: Integer]: PsCellFormat read GetItem write SetItem; default;
  end;

procedure InitFormatRecord(out AValue: TsCellFormat);


implementation

{ Utilities }

procedure InitFormatRecord(out AValue: TsCellFormat);
begin
  AValue.Name := '';
  AValue.NumberFormatStr := '';
  FillChar(AValue, SizeOf(AValue), 0);
  AValue.BorderStyles := DEFAULT_BORDERSTYLES;
  AValue.Background := EMPTY_FILL;
end;


{ TsCellFormatList }

constructor TsCellFormatList.Create(AAllowDuplicates: Boolean);
begin
  inherited Create;
  FAllowDuplicates := AAllowDuplicates;
end;

destructor TsCellFormatList.Destroy;
begin
  Clear;
  inherited;
end;

function TsCellFormatList.Add(const AItem: TsCellFormat): Integer;
var
  P: PsCellFormat;
begin
  if FAllowDuplicates then
    Result := -1
  else
    Result := IndexOf(AItem);
  if Result = -1 then begin
    New(P);
    P^.Name := AItem.Name;
    P^.ID := AItem.ID;
    P^.UsedFormattingFields := AItem.UsedFormattingFields;
    P^.FontIndex := AItem.FontIndex;
    P^.TextRotation := AItem.TextRotation;
    P^.HorAlignment := AItem.HorAlignment;
    P^.VertAlignment := AItem.VertAlignment;
    P^.Border := AItem.Border;
    P^.BorderStyles := AItem.BorderStyles;
    P^.Background := AItem.Background;
    P^.NumberFormat := AItem.NumberFormat;
    P^.NumberFormatStr := AItem.NumberFormatStr;
    Result := inherited Add(P);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds a pointer to a FormatRecord to the list. Allows nil for the predefined
  formats which are not stored in the file.
-------------------------------------------------------------------------------}
function TsCellFormatList.Add(AItem: PsCellFormat): Integer;
begin
  if AItem = nil then
    Result := inherited Add(AItem)
  else
    Result := Add(AItem^);
end;

procedure TsCellFormatList.Clear;
var
  i: Integer;
begin
  for i:=Count-1 downto 0 do
    Delete(i);
  inherited;
end;

procedure TsCellFormatList.Delete(AIndex: Integer);
var
  P: PsCellFormat;
begin
  P := GetItem(AIndex);
  if P <> nil then
    Dispose(P);
  inherited Delete(AIndex);
end;

function TsCellFormatList.GetItem(AIndex: Integer): PsCellFormat;
begin
  Result := inherited Items[AIndex];
end;

function TsCellFormatList.FindIndexOfID(ID: Integer): Integer;
var
  P: PsCellFormat;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P <> nil) and (P^.ID = ID) then
      exit;
  end;
  Result := -1;
end;

function TsCellFormatList.FindIndexOfName(AName: String): Integer;
var
  P: PsCellFormat;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P <> nil) and (P^.Name = AName) then
      exit;
  end;
  Result := -1;
end;

function TsCellFormatList.IndexOf(const AItem: TsCellFormat): Integer;
var
  P: PsCellFormat;
  equ: Boolean;
  b: TsCellBorder;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P = nil) then continue;

    if (P^.UsedFormattingFields <> AItem.UsedFormattingFields) then continue;

    if (uffFont in AItem.UsedFormattingFields) then
      if (P^.FontIndex) <> (AItem.FontIndex) then continue;

    if (uffTextRotation in AItem.UsedFormattingFields) then
      if (P^.TextRotation <> AItem.TextRotation) then continue;

    if (uffHorAlign in AItem.UsedFormattingFields) then
      if (P^.HorAlignment <> AItem.HorAlignment) then continue;

    if (uffVertAlign in AItem.UsedFormattingFields) then
      if (P^.VertAlignment <> AItem.VertAlignment) then continue;

    if (uffBorder in AItem.UsedFormattingFields) then begin
      if (P^.Border <> AItem.Border) then continue;
      equ := true;
      for b in AItem.Border do begin
        if (P^.BorderStyles[b].LineStyle <> AItem.BorderStyles[b].LineStyle) or
           (P^.BorderStyles[b].Color <> Aitem.BorderStyles[b].Color)
        then begin
          equ := false;
          break;
        end;
      end;
      if not equ then continue;
    end;

    if (uffBackground in AItem.UsedFormattingFields) then begin
      if (P^.Background.Style <> AItem.Background.Style) then continue;
      if (P^.Background.BgColor <> AItem.Background.BgColor) then continue;
      if (P^.Background.FgColor <> AItem.Background.FgColor) then continue;
    end;

    if (uffNumberFormat in AItem.UsedFormattingFields) then begin
      if (P^.NumberFormat <> AItem.NumberFormat) then continue;
      if (P^.NumberFormatStr <> AItem.NumberFormatStr) then continue;
    end;

    // If we arrive here then the format records match.
    exit;
  end;

  // We get here if no record matches
  Result := -1;
end;

procedure TsCellFormatList.SetItem(AIndex: Integer; const AValue: PsCellFormat);
begin
  inherited Items[AIndex] := AValue;
end;


end.

