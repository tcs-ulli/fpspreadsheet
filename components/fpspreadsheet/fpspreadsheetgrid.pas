{
fpspreadsheetgrid.pas

Grid component which can load and write data from / to FPSpreadsheet documents

AUTHORS: Felipe Monteiro de Carvalho, Werner Pamler
}

{ To do:
 - When Lazarus 1.4 comes out remove the workaround for the RGB2HLS bug in
   FindNearestPaletteIndex.
 - Arial bold is not shown as such if loaded from ods
 - Background color of first cell is ignored.
}

unit fpspreadsheetgrid;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs, Grids,
  fpspreadsheet;

type

  { TsCustomWorksheetGrid }

  {@@
    TsCustomWorksheetGrid is the ancestor of TsWorkseetGrid and is able to
    display spreadsheet data along with their formatting.
  }
  TsCustomWorksheetGrid = class(TCustomDrawGrid)
  private
    { Private declarations }
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FHeaderCount: Integer;
    FInitColCount: Integer;
    FInitRowCount: Integer;
    FFrozenCols: Integer;
    FFrozenRows: Integer;
    FEditText: String;
    FOldEditText: String;
    FLockCount: Integer;
    FEditing: Boolean;
    FCellFont: TFont;
    FAutoCalc: Boolean;
    FTextOverflow: Boolean;
    FReadFormulas: Boolean;
    FDrawingCell: PCell;
    FTextOverflowing: Boolean;
    function CalcAutoRowHeight(ARow: Integer): Integer;
    function CalcColWidth(AWidth: Single): Integer;
    function CalcRowHeight(AHeight: Single): Integer;
    procedure ChangedCellHandler(ASender: TObject; ARow, ACol: Cardinal);
    procedure ChangedFontHandler(ASender: TObject; ARow, ACol: Cardinal);
    procedure FixNeighborCellBorders(ACol, ARow: Integer);
    function GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer;
      out ABorderStyle: TsCellBorderStyle): Boolean;

    // Setter/Getter
    function GetBackgroundColor(ACol, ARow: Integer): TsColor;
    function GetBackgroundColors(ARect: TGridRect): TsColor;
    function GetCellBorder(ACol, ARow: Integer): TsCellBorders;
    function GetCellBorders(ARect: TGridRect): TsCellBorders;
    function GetCellBorderStyle(ACol, ARow: Integer; ABorder: TsCellBorder): TsCellBorderStyle;
    function GetCellBorderStyles(ARect: TGridRect; ABorder: TsCellBorder): TsCellBorderStyle;
    function GetCellFont(ACol, ARow: Integer): TFont;
    function GetCellFonts(ARect: TGridRect): TFont;
    function GetCellFontColor(ACol, ARow: Integer): TsColor;
    function GetCellFontColors(ARect: TGridRect): TsColor;
    function GetCellFontName(ACol, ARow: Integer): String;
    function GetCellFontNames(ARect: TGridRect): String;
    function GetCellFontSize(ACol, ARow: Integer): Single;
    function GetCellFontSizes(ARect: TGridRect): Single;
    function GetCellFontStyle(ACol, ARow: Integer): TsFontStyles;
    function GetCellFontStyles(ARect: TGridRect): TsFontStyles;
    function GetHorAlignment(ACol, ARow: Integer): TsHorAlignment;
    function GetHorAlignments(ARect: TGridRect): TsHorAlignment;
    function GetShowGridLines: Boolean;
    function GetShowHeaders: Boolean;
    function GetTextRotation(ACol, ARow: Integer): TsTextRotation;
    function GetTextRotations(ARect: TGridRect): TsTextRotation;
    function GetVertAlignment(ACol, ARow: Integer): TsVertAlignment;
    function GetVertAlignments(ARect: TGridRect): TsVertAlignment;
    function GetWordwrap(ACol, ARow: Integer): Boolean;
    function GetWordwraps(ARect: TGridRect): Boolean;
    procedure SetAutoCalc(AValue: Boolean);
    procedure SetBackgroundColor(ACol, ARow: Integer; AValue: TsColor);
    procedure SetBackgroundColors(ARect: TGridRect; AValue: TsColor);
    procedure SetCellBorder(ACol, ARow: Integer; AValue: TsCellBorders);
    procedure SetCellBorders(ARect: TGridRect; AValue: TsCellBorders);
    procedure SetCellBorderStyle(ACol, ARow: Integer; ABorder: TsCellBorder; AValue: TsCellBorderStyle);
    procedure SetCellBorderStyles(ARect: TGridRect; ABorder: TsCellBorder; AValue: TsCellBorderStyle);
    procedure SetCellFont(ACol, ARow: Integer; AValue: TFont);
    procedure SetCellFonts(ARect: TGridRect; AValue: TFont);
    procedure SetCellFontColor(ACol, ARow: Integer; AValue: TsColor);
    procedure SetCellFontColors(ARect: TGridRect; AValue: TsColor);
    procedure SetCellFontName(ACol, ARow: Integer; AValue: String);
    procedure SetCellFontNames(ARect: TGridRect; AValue: String);
    procedure SetCellFontStyle(ACol, ARow: Integer; AValue: TsFontStyles);
    procedure SetCellFontStyles(ARect: TGridRect; AValue: TsFontStyles);
    procedure SetCellFontSize(ACol, ARow: Integer; AValue: Single);
    procedure SetCellFontSizes(ARect: TGridRect; AValue: Single);
    procedure SetFrozenCols(AValue: Integer);
    procedure SetFrozenRows(AValue: Integer);
    procedure SetHorAlignment(ACol, ARow: Integer; AValue: TsHorAlignment);
    procedure SetHorAlignments(ARect: TGridRect; AValue: TsHorAlignment);
    procedure SetShowGridLines(AValue: Boolean);
    procedure SetShowHeaders(AValue: Boolean);
    procedure SetTextRotation(ACol, ARow: Integer; AValue: TsTextRotation);
    procedure SetTextRotations(ARect: TGridRect; AValue: TsTextRotation);
    procedure SetVertAlignment(ACol, ARow: Integer; AValue: TsVertAlignment);
    procedure SetVertAlignments(ARect: TGridRect; AValue: TsVertAlignment);
    procedure SetWordwrap(ACol, ARow: Integer; AValue: boolean);
    procedure SetWordwraps(ARect: TGridRect; AValue: boolean);

  protected
    { Protected declarations }
    procedure AutoAdjustColumn(ACol: Integer); override;
    procedure AutoAdjustRow(ARow: Integer); virtual;
    function CellOverflow(ACol, ARow: Integer; AState: TGridDrawState;
      out ACol1, ACol2: Integer; var ARect: TRect): Boolean;
    procedure CreateNewWorkbook;
    procedure DblClick; override;
    procedure DoPrepareCanvas(ACol, ARow: Integer; AState: TGridDrawState); override;
    procedure DrawAllRows; override;
    procedure DrawCellBorders; overload;
    procedure DrawCellBorders(ACol, ARow: Integer; ARect: TRect); overload;
    procedure DrawFocusRect(aCol,aRow:Integer; ARect:TRect); override;
    procedure DrawFrozenPaneBorders(ARect: TRect);
    procedure DrawRow(aRow: Integer); override;
    procedure DrawSelection;
    procedure DrawTextInCell(ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState); override;
    function GetCellHeight(ACol, ARow: Integer): Integer;
    function GetCellText(ACol, ARow: Integer): String;
    function GetEditText(ACol, ARow: Integer): String; override;
    function HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
    procedure HeaderSized(IsColumn: Boolean; AIndex: Integer); override;
    procedure InternalDrawTextInCell(AText, AMeasureText: String; ARect: TRect;
      AJustification: Byte; ACellHorAlign: TsHorAlignment;
      ACellVertAlign: TsVertAlignment; ATextRot: TsTextRotation;
      ATextWrap, ReplaceTooLong: Boolean);
    procedure KeyDown(var Key : Word; Shift : TShiftState); override;
    procedure Loaded; override;
    procedure LoadFromWorksheet(AWorksheet: TsWorksheet);
    procedure MoveSelection; override;
    procedure SelectEditor; override;
    procedure SetEditText(ACol, ARow: Longint; const AValue: string); override;
    procedure Setup;
    procedure UpdateColWidths(AStartIndex: Integer = 0);
    procedure UpdateRowHeights(AStartIndex: Integer = 0);
    {@@ Automatically recalculate formulas whenever a cell value changes, }
    property AutoCalc: Boolean read FAutoCalc write SetAutoCalc default false;
    {@@ Displays column and row headers in the fixed col/row style of the grid.
        Deprecated. Use ShowHeaders instead. }
    property DisplayFixedColRow: Boolean read GetShowHeaders write SetShowHeaders default true;
    {@@ This number of columns at the left is "frozen", i.e. it is not possible to
        scroll these columns }
    property FrozenCols: Integer read FFrozenCols write SetFrozenCols;
    {@@ This number of rows at the top is "frozen", i.e. it is not possible to
        scroll these rows. }
    property FrozenRows: Integer read FFrozenRows write SetFrozenRows;
    {@@ Activates reading of RPN formulas. Should be turned off when reading
        not implemented formulas crashes reading of the spreadsheet file. }
    property ReadFormulas: Boolean read FReadFormulas write FReadFormulas;
    {@@ Shows/hides vertical and horizontal grid lines }
    property ShowGridLines: Boolean read GetShowGridLines write SetShowGridLines default true;
    {@@ Shows/hides column and row headers in the fixed col/row style of the grid. }
    property ShowHeaders: Boolean read GetShowHeaders write SetShowHeaders default true;
    {@@ Activates text overflow (cells reaching into neighbors) }
    property TextOverflow: Boolean read FTextOverflow write FTextOverflow default false;

  public
    { public methods }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure BeginUpdate;
    procedure DefaultDrawCell(ACol, ARow: Integer; var ARect: TRect; AState: TGridDrawState); override;
    procedure DeleteCol(AGridCol: Integer);
    procedure DeleteRow(AGridRow: Integer);
    procedure EditingDone; override;
    procedure EndUpdate;
    procedure GetSheets(const ASheets: TStrings);
    function GetGridCol(ASheetCol: Cardinal): Integer;
    function GetGridRow(ASheetRow: Cardinal): Integer;
    function GetWorksheetCol(AGridCol: Integer): Cardinal;
    function GetWorksheetRow(AGridRow: Integer): Cardinal;
    procedure InsertCol(AGridCol: Integer);
    procedure InsertRow(AGridRow: Integer);
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AWorksheetIndex: Integer = 0); overload;
    procedure NewWorkbook(AColCount, ARowCount: Integer);
    procedure SaveToSpreadsheetFile(AFileName: string;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat;
      AOverwriteExisting: Boolean = true); overload;
    procedure SelectSheetByIndex(AIndex: Integer);

    procedure MergeCells;
    procedure UnmergeCells;

    { Utilities related to Workbooks }
    procedure Convert_sFont_to_Font(sFont: TsFont; AFont: TFont);
    procedure Convert_Font_to_sFont(AFont: TFont; sFont: TsFont);
    function FindNearestPaletteIndex(AColor: TColor): TsColor;

    { public properties }
    {@@ Currently selected worksheet of the workbook }
    property Worksheet: TsWorksheet read FWorksheet;
    {@@ Workbook displayed in the grid }
    property Workbook: TsWorkbook read FWorkbook;
    {@@ Count of header lines - for conversion between grid- and workbook-based
     row and column indexes. Either 1 if row and column headers are shown or 0 if not}
    property HeaderCount: Integer read FHeaderCount;

    { maybe these should become published ... }

    {@@ Background color of the cell at the given column and row. Expressed as
        index into the workbook's color palette. }
    property BackgroundColor[ACol, ARow: Integer]: TsColor
        read GetBackgroundColor write SetBackgroundColor;
    {@@ Common background color of the cells covered by the given rectangle.
        Expressed as index into the workbook's color palette. }
    property BackgroundColors[ARect: TGridRect]: TsColor
        read GetBackgroundColors write SetBackgroundColors;
    {@@ Set of flags indicating at which cell border a border line is drawn. }
    property CellBorder[ACol, ARow: Integer]: TsCellBorders
        read GetCellBorder write SetCellBorder;
    {@@ Set of flags indicating at which border of a range of cells a border
        line is drawn }
    property CellBorders[ARect: TGridRect]: TsCellBorders
        read GetCellBorders write SetCellBorders;
    {@@ Style of the border line at the given border of the cell at column ACol
        and row ARow. Requires the cellborder flag of the border to be set
        for the border line to be shown }
    property CellBorderStyle[ACol, ARow: Integer; ABorder: TsCellBorder]: TsCellBorderStyle
        read GetCellBorderStyle write SetCellBorderStyle;
    {@@ Style of the border line at the given border of the cells within the
        range of colum/row indexes defined by the rectangle. Requires the cellborder
        flag of the border to be set for the border line to be shown }
    property CellBorderStyles[ARect: TGridRect; ABorder: TsCellBorder]: TsCellBorderStyle
        read GetCellBorderStyles write SetCellBorderStyles;
    {@@ Font to be used for text in the cell at column ACol and row ARow. }
    property CellFont[ACol, ARow: Integer]: TFont
        read GetCellFont write SetCellFont;
    {@@ Font to be used for the cells in the column/row index range
        given by the rectangle }
    property CellFonts[ARect: TGridRect]: TFont
        read GetCellFonts write SetCellFonts;
    {@@ Name of the font used for the cell on column ACol and row ARow }
    property CellFontName[ACol, ARow: Integer]: String
        read GetCellFontName write SetCellFontName;
    {@@ Name of the font used for the cells within the range
        of column/row indexes defined by the rectangle. }
    property CellFontNames[ARect: TGridRect]: String
        read GetCellFontNames write SetCellFontNames;
    {@@ Style of the font (bold, italic, ...) used for text in the
        cell at column ACol and row ARow. }
    property CellFontStyle[ACol, ARow: Integer]: TsFontStyles
        read GetCellFontStyle write SetCellFontStyle;
    {@@ Style of the font (bold, italic, ...) used for the cells within
        the range of column/row indexes defined by the rectangle. }
    property CellFontStyles[ARect: TGridRect]: TsFontStyles
        read GetCellFontStyles write SetCellFontStyles;
    {@@ Size of the font (in points) used for the cell at column ACol
        and row ARow }
    property CellFontSize[ACol, ARow: Integer]: Single
        read GetCellFontSize write SetCellFontSize;
    {@@ Size of the font (in points) used for the cells within the
        range of column/row indexes defined by the rectangle. }
    property CellFontSizes[ARect: TGridRect]: Single
        read GetCellFontSizes write SetCellFontSizes;
    {@@ Parameter for horizontal text alignment within the cell at column ACol
        and row ARow }
    property HorAlignment[ACol, ARow: Integer]: TsHorAlignment
        read GetHorAlignment write SetHorAlignment;
    {@@ Parameter for the horizontal text alignments in all cells within the
        range cf column/row indexes defined by the rectangle. }
    property HorAlignments[ARect: TGridRect]: TsHorAlignment
        read GetHorAlignments write SetHorAlignments;
    {@@ Rotation of the text in the cell at column ACol and row ARow. }
    property TextRotation[ACol, ARow: Integer]: TsTextRotation
        read GetTextRotation write SetTextRotation;
    {@@ Rotation of the text in the cells within the range of column/row indexes
        defined by the rectangle. }
    property TextRotations[ARect: TGridRect]: TsTextRotation
        read GetTextRotations write SetTextRotations;
    {@@ Parameter for vertical text alignment in the cell at column ACol and
        row ARow. }
    property VertAlignment[ACol, ARow: Integer]: TsVertAlignment
        read GetVertAlignment write SetVertAlignment;
    {@@ Parameter for vertical text alignment in the cells having column/row
        indexes defined by the rectangle. }
    property VertAlignments[ARect: TGridRect]: TsVertAlignment
        read GetVertAlignments write SetVertAlignments;
    {@@ If true, word-wrapping of text within the cell at column ACol and row ARow
        is activated. }
    property Wordwrap[ACol, ARow: Integer]: Boolean
        read GetWordwrap write SetWordwrap;
    {@@ If true, word-wrapping of text within all cells within the range defined
        by the rectangle is activated. }
    property Wordwraps[ARect: TGridRect]: Boolean
        read GetWordwraps write SetWordwraps;
  end;

  { TsWorksheetGrid }

  {@@
    TsWorksheetGrid is a grid which displays spreadsheet data along with
    formatting. As it is linked to an instance of TsWorkbook, it provides
    methods for reading data from or writing to spreadsheet files. It has the
    same funtionality as TsCustomWorksheetGrid, but publishes has all properties.
  }
  TsWorksheetGrid = class(TsCustomWorksheetGrid)
  published
    // inherited from TsCustomWorksheetGrid
    {@@ Automatically recalculates the worksheet of a cell value changes. }
    property AutoCalc;
    {@@ Displays column and row headers in the fixed col/row style of the grid.
        Deprecated. Use ShowHeaders instead. }
    property DisplayFixedColRow; deprecated 'Use ShowHeaders';
    {@@ This number of columns at the left is "frozen", i.e. it is not possible to
        scroll these columns. }
    property FrozenCols;
    {@@ This number of rows at the top is "frozen", i.e. it is not possible to
        scroll these rows. }
    property FrozenRows;
    {@@ Activates reading of RPN formulas. Should be turned off when reading of
        not implemented formulas crashes reading of the spreadsheet file. }
    property ReadFormulas;
    {@@ Shows/hides vertical and horizontal grid lines. }
    property ShowGridLines;
    {@@ Shows/hides column and row headers in the fixed col/row style of the grid. }
    property ShowHeaders;
    {@@ Activates text overflow (cells reaching into neighbors) }
    property TextOverflow;

    {@@ inherited from ancestors}
    property Align;
    {@@ inherited from ancestors}
    property AlternateColor;
    {@@ inherited from ancestors}
    property Anchors;
    {@@ inherited from ancestors}
    property AutoAdvance;
    {@@ inherited from ancestors}
    property AutoEdit;
    {@@ inherited from ancestors}
    property AutoFillColumns;
    //property BiDiMode;
    {@@ inherited from ancestors}
    property BorderSpacing;
    {@@ inherited from ancestors}
    property BorderStyle;
    {@@ inherited from ancestors}
    property Color;
    {@@ inherited from ancestors}
    property ColCount;
    //property Columns;
    {@@ inherited from ancestors}
    property Constraints;
    {@@ inherited from ancestors}
    property DefaultColWidth;
    {@@ inherited from ancestors}
    property DefaultDrawing;
    {@@ inherited from ancestors}
    property DefaultRowHeight;
    {@@ inherited from ancestors}
    property DragCursor;
    {@@ inherited from ancestors}
    property DragKind;
    {@@ inherited from ancestors}
    property DragMode;
    {@@ inherited from ancestors}
    property Enabled;
    {@@ inherited from ancestors}
    property ExtendedSelect default true;
    {@@ inherited from ancestors}
    property FixedColor;
    {@@ inherited from ancestors}
    property Flat;
    {@@ inherited from ancestors}
    property Font;
    {@@ inherited from ancestors}
    property GridLineWidth;
    {@@ inherited from ancestors}
    property HeaderHotZones;
    {@@ inherited from ancestors}
    property HeaderPushZones;
    {@@ inherited from ancestors}
    property MouseWheelOption;
    {@@ inherited from TCustomGrid. Select the option goEditing to make the grid editable! }
    property Options;
    //property ParentBiDiMode;
    {@@ inherited from ancestors}
    property ParentColor default false;
    {@@ inherited from ancestors}
    property ParentFont;
    {@@ inherited from ancestors}
    property ParentShowHint;
    {@@ inherited from ancestors}
    property PopupMenu;
    {@@ inherited from ancestors}
    property RowCount;
    {@@ inherited from ancestors}
    property ScrollBars;
    {@@ inherited from ancestors}
    property ShowHint;
    {@@ inherited from ancestors}
    property TabOrder;
    {@@ inherited from ancestors}
    property TabStop;
    {@@ inherited from ancestors}
    property TitleFont;
    {@@ inherited from ancestors}
    property TitleImageList;
    {@@ inherited from ancestors}
    property TitleStyle;
    {@@ inherited from ancestors}
    property UseXORFeatures;
    {@@ inherited from ancestors}
    property Visible;
    {@@ inherited from ancestors}
    property VisibleColCount;
    {@@ inherited from ancestors}
    property VisibleRowCount;

    {@@ inherited from ancestors}
    property OnBeforeSelection;
    {@@ inherited from ancestors}
    property OnChangeBounds;
    {@@ inherited from ancestors}
    property OnClick;
    {@@ inherited from ancestors}
    property OnColRowDeleted;
    {@@ inherited from ancestors}
    property OnColRowExchanged;
    {@@ inherited from ancestors}
    property OnColRowInserted;
    {@@ inherited from ancestors}
    property OnColRowMoved;
    {@@ inherited from ancestors}
    property OnCompareCells;
    {@@ inherited from ancestors}
    property OnDragDrop;
    {@@ inherited from ancestors}
    property OnDragOver;
    {@@ inherited from ancestors}
    property OnDblClick;
    {@@ inherited from ancestors}
    property OnDrawCell;
    {@@ inherited from ancestors}
    property OnEditButtonClick;
    {@@ inherited from ancestors}
    property OnEditingDone;
    {@@ inherited from ancestors}
    property OnEndDock;
    {@@ inherited from ancestors}
    property OnEndDrag;
    {@@ inherited from ancestors}
    property OnEnter;
    {@@ inherited from ancestors}
    property OnExit;
    {@@ inherited from ancestors}
    property OnGetEditMask;
    {@@ inherited from ancestors}
    property OnGetEditText;
    {@@ inherited from ancestors}
    property OnHeaderClick;
    {@@ inherited from ancestors}
    property OnHeaderSized;
    {@@ inherited from ancestors}
    property OnKeyDown;
    {@@ inherited from ancestors}
    property OnKeyPress;
    {@@ inherited from ancestors}
    property OnKeyUp;
    {@@ inherited from ancestors}
    property OnMouseDown;
    {@@ inherited from ancestors}
    property OnMouseMove;
    {@@ inherited from ancestors}
    property OnMouseUp;
    {@@ inherited from ancestors}
    property OnMouseWheel;
    {@@ inherited from ancestors}
    property OnMouseWheelDown;
    {@@ inherited from ancestors}
    property OnMouseWheelUp;
    {@@ inherited from ancestors}
    property OnPickListSelect;
    {@@ inherited from ancestors}
    property OnPrepareCanvas;
    {@@ inherited from ancestors}
    property OnResize;
    {@@ inherited from ancestors}
    property OnSelectEditor;
    {@@ inherited from ancestors}
    property OnSelection;
    {@@ inherited from ancestors}
    property OnSelectCell;
    {@@ inherited from ancestors}
    property OnSetEditText;
    {@@ inherited from ancestors}
    property OnShowHint;
    {@@ inherited from ancestors}
    property OnStartDock;
    {@@ inherited from ancestors}
    property OnStartDrag;
    {@@ inherited from ancestors}
    property OnTopLeftChanged;
    {@@ inherited from ancestors}
    property OnUTF8KeyPress;
    {@@ inherited from ancestors}
    property OnValidateEntry;
    {@@ inherited from ancestors}
    property OnContextPopup;
  end;

procedure Register;

implementation

uses
  Types, LCLType, LCLIntf, Math, fpCanvas, fpsUtils;

const
  {@@ Translation of the fpspreadsheet type of horizontal text alignment to that
      used in the graphics unit. }
  HOR_ALIGNMENTS: array[haLeft..haRight] of TAlignment = (
    taLeftJustify, taCenter, taRightJustify
  );
  {@@ Translation of the fpspreadsheet type of vertical text alignment to that
      used in the graphics unit. }
  VERT_ALIGNMENTS: array[TsVertAlignment] of TTextLayout = (
    tlBottom, tlTop, tlCenter, tlBottom
  );

var
  {@@ Auxiliary bitmap containing the fill pattern used by biff2 cell backgrounds. }
  FillPattern_BIFF2: TBitmap = nil;

{@@ ----------------------------------------------------------------------------
  Helper procedure which creates the fill pattern used by biff2 cell backgrounds.
-------------------------------------------------------------------------------}
procedure Create_FillPattern_BIFF2(ABkColor: TColor);
begin
  FreeAndNil(FillPattern_BIFF2);
  FillPattern_BIFF2 := TBitmap.Create;
  with FillPattern_BIFF2 do begin
    SetSize(4, 4);
    Canvas.Brush.Color := ABkColor;
    Canvas.FillRect(0, 0, Width, Height);
    Canvas.Pixels[0, 0] := clBlack;
    Canvas.Pixels[2, 2] := clBlack;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Helper procedure which draws a densely dotted horizontal line. In Excel
  this is called a "hair line".

  @param x1, x2   x coordinates of the end points of the line
  @param y        y coordinate of the horizontal line
-------------------------------------------------------------------------------}
procedure DrawHairLineHor(ACanvas: TCanvas; x1, x2, y: Integer);
var
  clr: TColor;
  x: Integer;
begin
  if odd(x1) then inc(x1);
  x := x1;
  clr := ACanvas.Pen.Color;
  while (x <= x2) do begin
    ACanvas.Pixels[x, y] := clr;
    inc(x, 2);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Helper procedure which draws a densely dotted vertical line. In Excel
  this is called a "hair line".

  @param x        x coordinate of the vertical line
  @param y1, y2   y coordinates of the end points of the line
-------------------------------------------------------------------------------}
procedure DrawHairLineVert(ACanvas: TCanvas; x, y1, y2: Integer);
var
  clr: TColor;
  y: Integer;
begin
  if odd(y1) then inc(y1);
  y := y1;
  clr := ACanvas.Pen.Color;
  while (y <= y2) do begin
    ACanvas.Pixels[x, y] := clr;
    inc(y, 2);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Wraps text by inserting line ending characters so that the lines are not
  longer than AMaxWidth.

  @param   ACanvas       Canvas on which the text will be drawn
  @param   AText         Text to be drawn
  @param   AMaxWidth     Maximimum line width (in pixels)
  @return  Text with inserted line endings such that the lines are shorter than
           AMaxWidth.

  @note    Based on ocde posted by user "taazz" in the Lazarus forum
           http://forum.lazarus.freepascal.org/index.php/topic,21305.msg124743.html#msg124743
-------------------------------------------------------------------------------}
function WrapText(ACanvas: TCanvas; const AText: string; AMaxWidth: integer): string;
var
  DC: HDC;
  textExtent: TSize;
  S, P, E: PChar;
  line: string;
  isFirstLine: boolean;
begin
  Result := '';
  DC := ACanvas.Handle;
  isFirstLine := True;
  P := PChar(AText);
  while P^ = ' ' do
    Inc(P);
  while P^ <> #0 do begin
    S := P;
    E := nil;
    while (P^ <> #0) and (P^ <> #13) and (P^ <> #10) do begin
      LCLIntf.GetTextExtentPoint(DC, S, P - S + 1, textExtent);
      if (textExtent.CX > AMaxWidth) and (E <> nil) then begin
        if (P^ <> ' ') and (P^ <> ^I) then begin
          while (E >= S) do
            case E^ of
              '.', ',', ';', '?', '!', '-', ':',
              ')', ']', '}', '>', '/', '\', ' ':
                break;
              else
                Dec(E);
            end;
          if E < S then
            E := P - 1;
        end;
        Break;
      end;
      E := P;
      Inc(P);
    end;
    if E <> nil then begin
      while (E >= S) and (E^ = ' ') do
        Dec(E);
    end;
    if E <> nil then
      SetString(Line, S, E - S + 1)
    else
      SetLength(Line, 0);
    if (P^ = #13) or (P^ = #10) then begin
      Inc(P);
      if (P^ <> (P - 1)^) and ((P^ = #13) or (P^ = #10)) then
        Inc(P);
      if P^ = #0 then
        line := line + LineEnding;
    end
    else if P^ <> ' ' then
      P := E + 1;
    while P^ = ' ' do
      Inc(P);
    if isFirstLine then begin
      Result := Line;
      isFirstLine := False;
    end else
      Result := Result + LineEnding + line;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Calculates a background color for selected cells. The procedures takes the
  original background color and dims or brightens it by adding the value ADelta
  to the RGB components.

  @param  c       Color to be modified
  @param  ADelta  Value to be added to the RGB components of the inpur color
  @result Modified color.
-------------------------------------------------------------------------------}
function CalcSelectionColor(c: TColor; ADelta: Byte) : TColor;
type
  TRGBA = record R,G,B,A: Byte end;
begin
  c := ColorToRGB(c);
  TRGBA(Result).A := 0;
  if TRGBA(c).R < 128
    then TRGBA(Result).R := TRGBA(c).R + ADelta
    else TRGBA(Result).R := TRGBA(c).R - ADelta;
  if TRGBA(c).G < 128
    then TRGBA(Result).G := TRGBA(c).G + ADelta
    else TRGBA(Result).G := TRGBA(c).G - ADelta;
  if TRGBA(c).B < 128
    then TRGBA(Result).B := TRGBA(c).B + ADelta
    else TRGBA(Result).B := TRGBA(c).B - ADelta;
end;

{@@ ----------------------------------------------------------------------------
  Registers the worksheet grid in the Lazarus component palette,
  page "Additional".
-------------------------------------------------------------------------------}
procedure Register;
begin
  RegisterComponents('Additional',[TsWorksheetGrid]);
end;


{*******************************************************************************
*                              TsCustomWorksheetGrid                           *
*******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Constructor of the grid. Activates the display of column and row headers
  and creates an internal "CellFont". Creates a pre-defined number of empty rows
  and columns.

  @param  AOwner   Owner of the grid
-------------------------------------------------------------------------------}
constructor TsCustomWorksheetGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  AutoAdvance := aaDown;
  ExtendedSelect := true;
  FHeaderCount := 1;
  FInitColCount := 26;
  FInitRowCount := 100;
  FCellFont := TFont.Create;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the grid: Destroys the workbook and the internal CellFont.
-------------------------------------------------------------------------------}
destructor TsCustomWorksheetGrid.Destroy;
begin
  FreeAndNil(FWorkbook);
  FreeAndNil(FCellFont);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Is called when goDblClickAutoSize is in the grid's options and a double click
  has occured at the border of a column header. Sets optimum column with.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.AutoAdjustColumn(ACol: Integer);
var
  gRow: Integer;  // row in grid coordinates
  c: Cardinal;
  r: Cardinal;
  lastRow: Cardinal;
  cell: PCell;
  w, maxw: Integer;
  fnt: TFont;
  txt: String;
begin
  if FWorksheet = nil then
    exit;

  c := GetWorksheetCol(ACol);
  lastRow := FWorksheet.GetLastOccupiedRowIndex;
  maxw := -1;
  for r := 0 to lastRow do
  begin
    gRow := GetGridRow(r);
    fnt := GetCellFont(ACol, gRow);
    txt := GetCellText(ACol, gRow);
    PrepareCanvas(ACol, gRow, []);
    w := Canvas.TextWidth(txt);
    if (txt <> '') and (w > maxw) then maxw := w;
  end;
  if maxw > -1 then
    maxw := maxw + 2*constCellPadding
  else
    maxw := DefaultColWidth;
  ColWidths[ACol] := maxW;
  HeaderSized(true, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Is called when goDblClickAutoSize is in the grid's options and a double click
  has occured at the border of a row header. Sets optimum row height.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.AutoAdjustRow(ARow: Integer);
begin
  if FWorksheet <> nil then
    RowHeights[ARow] := CalcAutoRowHeight(ARow)
  else
    RowHeights[ARow] := DefaultRowHeight;
  HeaderSized(false, ARow);
end;

{@@ ----------------------------------------------------------------------------
  The BeginUpdate/EndUpdate pair suppresses unnecessary painting of the grid.
  Call BeginUpdate to stop refreshing the grid, and call EndUpdate to release
  the lock and to repaint the grid again.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.BeginUpdate;
begin
  inc(FLockCount);
end;

{@@ ----------------------------------------------------------------------------
  Converts the column width, given in "characters" of the default font, to pixels.
  All chars are assumed to have the same width defined by the width of the
  "0" character. Therefore, this calculation is only approximate.

  @param   AWidth   Width of a column given as "character count".
  @return  Column width in pixels.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcColWidth(AWidth: Single): Integer;
var
  w0: Integer;
begin
  Convert_sFont_to_Font(FWorkbook.GetFont(0), Canvas.Font);
  w0 := Canvas.TextWidth('0');
  Result := Round(AWidth * w0);
end;

{@@ ----------------------------------------------------------------------------
  Finds the maximum cell height per row and uses this to define the RowHeights[].
  Returns DefaultRowHeight if the row does not contain any cells, or if the
  worksheet does not have a TRow record for this particular row.
  ARow is a grid row index.

  @param   ARow  Index of the row, in grid units
  @return  Row height
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcAutoRowHeight(ARow: Integer): Integer;
var
  c: Integer;
  h: Integer;
begin
  h := 0;
  for c := FHeaderCount to ColCount-1 do
    h := Max(h, GetCellHeight(c, ARow));
  if h = 0 then
    Result := DefaultRowHeight
  else
    Result := h;
end;

{@@ ----------------------------------------------------------------------------
  Converts the row height (from a worksheet row record), given in lines, to
  pixels as needed by the grid

  @param  AHeight  Row height expressed as default font line count from the
                   worksheet
  @result Row height in pixels.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcRowHeight(AHeight: Single): Integer;
var
  h_pts: Single;
begin
  h_pts := AHeight * (Workbook.GetFont(0).Size + ROW_HEIGHT_CORRECTION);
  Result := PtsToPX(h_pts, Screen.PixelsPerInch) + 4;
end;

{@@ ----------------------------------------------------------------------------
  Looks for overflowing cells: if the text of the given cell is longer than
  the cell width the function calculates the column indexes and the rectangle
  to show the complete text.
  Ony for non-wordwrapped label cells and for horizontal orientation.
  Function returns false if text overflow needs not to be considered.

  @param ACol, ARow   Column and row indexes (in grid coordinates) of the cell
                      to be drawn
  @param AState       GridDrawState of the cell (normal, fixed, selected etc)
  @param ACol1,ACol2  (output) Index of the first and last column covered by the
                      overflowing text
  @param ARect        (output) Pixel rectangle enclosing the cell and its neighbors
                      affected
  @return TRUE if text overflow into neighbor cells is to be considered,
          FALSE if not.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CellOverflow(ACol, ARow: Integer;
  AState: TGridDrawState; out ACol1, ACol2: Integer; var ARect: TRect): Boolean;
var
  txt: String;
  len: Integer;
  cell: PCell;
  txtalign: TsHorAlignment;
  r: Cardinal;
  w, w0: Integer;
begin
  Result := false;
  cell := FDrawingCell;

  // Nothing to do in these cases (like in Excel):
  if (cell = nil) or (cell^.ContentType <> cctUTF8String) then  // ... non-label cells
    exit;
  if (uffWordWrap in cell^.UsedFormattingFields) then           // ... word-wrap
    exit;
  if (uffTextRotation in cell^.UsedFormattingFields) and        // ... vertical text
     (cell^.TextRotation <> trHorizontal)
  then
    exit;

  txt := cell^.UTF8Stringvalue;
  if (uffHorAlign in cell^.UsedFormattingFields) then
    txtalign := cell^.HorAlignment
  else
    txtalign := haDefault;
  PrepareCanvas(ACol, ARow, AState);
  len := Canvas.TextWidth(txt) + 2*constCellPadding;
  ACol1 := ACol;
  ACol2 := ACol;
  r := GetWorksheetRow(ARow);
  case txtalign of
    haLeft, haDefault:
      // overflow to the right
      while (len > ARect.Right - ARect.Left) and (ACol2 < ColCount-1) do
      begin
        result := true;
        inc(ACol2);
        cell := FWorksheet.FindCell(r, GetWorksheetCol(ACol2));
        if (cell <> nil) and (cell^.ContentType <> cctEmpty) then
        begin
          dec(ACol2);
          break;
        end;
        ARect.Right := ARect.Right + ColWidths[ACol2];
      end;
    haRight:
      // overflow to the left
      while (len > ARect.Right - ARect.Left) and (ACol1 > FixedCols) do
      begin
        result := true;
        dec(ACol1);
        cell := FWorksheet.FindCell(r, GetWorksheetCol(ACol1));
        if (cell <> nil) and (cell^.ContentType <> cctEmpty) then
        begin
          inc(ACol1);
          break;
        end;
        ARect.Left := ARect.Left - ColWidths[ACol1];
      end;
    haCenter:
      begin
        len := len div 2;
        w0 := (ARect.Right - ARect.Left) div 2;
        w := w0;
        // right part
        while (len > w) and (ACol2 < ColCount-1) do
        begin
          Result := true;
          inc(ACol2);
          cell := FWorksheet.FindCell(r, GetWorksheetCol(ACol2));
          if (cell <> nil) and (cell^.ContentType <> cctEmpty) then
          begin
            dec(ACol2);
            break;
          end;
          ARect.Right := ARect.Right + ColWidths[ACol2];
          inc(w, ColWidths[ACol2]);
        end;
        // left part
        w := w0;
        while (len > w) and (ACol1 > FixedCols) do
        begin
          Result := true;
          dec(ACol1);
          cell := FWorksheet.FindCell(r, GetWorksheetCol(ACol1));
          if (cell <> nil) and (cell^.Contenttype <> cctEmpty) then
          begin
            inc(ACol1);
            break;
          end;
          ARect.Left := ARect.left - ColWidths[ACol1];
          inc(w, ColWidths[ACol1]);
        end;
      end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Handler for the event OnChangeCell fired by the worksheet when the contents
  or formatting of a cell have changed.
  As a consequence, the grid may have to update the cell.
  Row/Col coordinates are in worksheet units here!

  @param  ASender  Sender of the event OnChangeFont (the worksheet)
  @param  ARow     Row index of the changed cell, in worksheet units!
  @param  ACol     Column index of the changed cell, in worksheet units!
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.ChangedCellHandler(ASender: TObject; ARow, ACol:Cardinal);
begin
  Unused(ASender, ARow, ACol);
  if FLockCount = 0 then Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Handler for the event OnChangeFont fired by the worksheet when the font has
  changed in a cell.
  As a consequence, the grid may have to update the row height.
  Row/Col coordinates are in worksheet units here!

  @param  ASender  Sender of the event OnChangeFont (the worksheet)
  @param  ARow     Row index of the cell with the changed font, in worksheet units!
  @param  ACol     Column index of the cell with the changed font, in worksheet units!
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.ChangedFontHandler(ASender: TObject;
  ARow, ACol: Cardinal);
var
  lRow: PRow;
begin
  Unused(ASender, ACol);
  if (FWorksheet <> nil) then begin
    lRow := FWorksheet.FindRow(ARow);
    if lRow = nil then begin
      // There is no row record --> row height changes according to font height
      // Otherwise the row height would be fixed according to the value in the row record.
      ARow := ARow + FHeaderCount;  // convert row index to grid units
      RowHeights[ARow] := CalcAutoRowHeight(ARow);
    end;
    Invalidate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Converts a spreadsheet font to a font used for painting (TCanvas.Font).

  @param  sFont  Font as used by fpspreadsheet (input)
  @param  AFont  Font as used by TCanvas for painting (output)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Convert_sFont_to_Font(sFont: TsFont; AFont: TFont);
begin
  if Assigned(AFont) and Assigned(sFont) then begin
    AFont.Name := sFont.FontName;
    AFont.Size := round(sFont.Size);
    AFont.Style := [];
    if fssBold in sFont.Style then AFont.Style := AFont.Style + [fsBold];
    if fssItalic in sFont.Style then AFont.Style := AFont.Style + [fsItalic];
    if fssUnderline in sFont.Style then AFont.Style := AFont.Style + [fsUnderline];
    if fssStrikeout in sFont.Style then AFont.Style := AFont.Style + [fsStrikeout];
    AFont.Color := Workbook.GetPaletteColor(sFont.Color);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Converts a font used for painting (TCanvas.Font) to a spreadsheet font.

  @param  AFont  Font as used by TCanvas for painting (input)
  @param  sFont  Font as used by fpspreadsheet (output)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Convert_Font_to_sFont(AFont: TFont;
  sFont: TsFont);
begin
  if Assigned(AFont) and Assigned(sFont) then begin
    sFont.FontName := AFont.Name;
    sFont.Size := AFont.Size;
    sFont.Style := [];
    if fsBold in AFont.Style then Include(sFont.Style, fssBold);
    if fsItalic in AFont.Style then Include(sFont.Style, fssItalic);
    if fsUnderline in AFont.Style then Include(sFont.Style, fssUnderline);
    if fsStrikeout in AFont.Style then Include(sFont.Style, fssStrikeout);
    sFont.Color := FindNearestPaletteIndex(AFont.Color);
  end;
end;

{@@ ----------------------------------------------------------------------------
  This is one of the main painting methods inherited from TsCustomGrid. It is
  overridden here to achieve the feature of "frozen" cells which should be
  painted in the same style as normal cells.

  Internally, "frozen" cells are "fixed" cells of the grid. Therefore, it is
  not possible to select any cell within the frozen panes - in contrast to the
  standard spreadsheet applications.

  @param  ACol   Column index of the cell being drawn
  @param  ARow   Row index of the cell beging drawn
  @param  ARect  Rectangle, in grid pixels, covered by the cell
  @param  AState Grid drawing state, as defined by TsCustomGrid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DefaultDrawCell(aCol, aRow: Integer;
  var aRect: TRect; AState: TGridDrawState);
var
  wasFixed: Boolean;
begin
  wasFixed := false;
  if (gdFixed in AState) then
    if ShowHeaders then begin
      if ((ARow < FixedRows) and (ARow > 0) and (ACol > 0)) or
         ((ACol < FixedCols) and (ACol > 0) and (ARow > 0))
      then
        wasFixed := true;
    end else begin
      if (ARow < FixedRows) or (ACol < FixedCols) then
        wasFixed := true;
    end;

  if wasFixed then begin
    AState := AState - [gdFixed];
    Canvas.Brush.Color := clWindow;
    DoPrepareCanvas(ACol, ARow, AState);
  end;

  inherited DefaultDrawCell(ACol, ARow, ARect, AState);

  if wasFixed then begin
    DrawCellGrid(ACol, ARow, ARect, AState);
    AState := AState + [gdFixed];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the column specified.

  @param   AGridCol   Grid index of the column to be deleted
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DeleteCol(AGridCol: Integer);
begin
  if AGridCol < FHeaderCount then
    exit;

  FWorksheet.DeleteCol(GetWorksheetCol(AGridCol));
  UpdateColWidths(AGridCol);
end;

{@@ ----------------------------------------------------------------------------
  Deletes the row specified.

  @param  AGridRow   Grid index of the row to be deleted
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DeleteRow(AGridRow: Integer);
begin
  if AGridRow < FHeaderCount then
    exit;

  FWorksheet.DeleteRow(GetWorksheetRow(AGridRow));
  UpdateRowHeights(AGridRow);
end;


{@@ ----------------------------------------------------------------------------
  Creates a new empty workbook into which a file will be loaded. Destroys the
  previously used workbook.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.CreateNewWorkbook;
begin
  FreeAndNil(FWorkbook);
  FWorkbook := TsWorkbook.Create;
  if FReadFormulas then
    FWorkbook.Options := FWorkbook.Options + [boReadFormulas]
  else
    FWorkbook.Options := FWorkbook.Options - [boReadFormulas];
  SetAutoCalc(FAutoCalc);
end;

{@@ ----------------------------------------------------------------------------
  Is called when a Double-click occurs. Overrides the inherited method to
  react on double click on cell border in row headers to auto-adjust the
  row heights
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DblClick;
var
  oldHeight: Integer;
  gRow: Integer;
begin
  SelectActive := False;
  FGridState := gsNormal;
  if (goRowSizing in Options) and (Cursor = crVSplit) and (FHeaderCount > 0) then
  begin
    if (goDblClickAutoSize in Options) then
    begin
      gRow := GCache.MouseCell.y;
      if CellRect(0, gRow).Bottom - GCache.ClickMouse.y > 0 then dec(gRow);
      oldHeight := RowHeights[gRow];
      AutoAdjustRow(gRow);
      if oldHeight <> RowHeights[gRow] then
        Cursor := crDefault; //ChangeCursor;
    end
  end
  else
    inherited DblClick;
end;


{@@ ----------------------------------------------------------------------------
  Adjusts the grid's canvas before painting a given cell. Considers
  background color, horizontal alignment, vertical alignment, etc.

  @param  ACol    Column index of the cell being painted
  @param  ARow    Row index of the cell being painted
  @param  AState  Grid drawing state -- see TsCustomGrid.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DoPrepareCanvas(ACol, ARow: Integer;
  AState: TGridDrawState);
var
  ts: TTextStyle;
  lCell: PCell;
  r, c: Integer;
  fnt: TsFont;
  style: TFontStyles;
  isSelected: Boolean;
begin
  GetSelectedState(AState, isSelected);
  Canvas.Font.Assign(Font);
  Canvas.Brush.Bitmap := nil;
  Canvas.Brush.Color := Color;
  ts := Canvas.TextStyle;
  if ShowHeaders then
  begin
    // Formatting of row and column headers
    if ARow = 0 then
    begin
      ts.Alignment := taCenter;
      ts.Layout := tlCenter;
    end else
    if ACol = 0 then
    begin
      ts.Alignment := taRightJustify;
      ts.Layout := tlCenter;
    end;
    if ShowHeaders and ((ACol = 0) or (ARow = 0)) then
      Canvas.Brush.Color := FixedColor
  end;
  if (FWorksheet <> nil) and (ARow >= FHeaderCount) and (ACol >= FHeaderCount) then
  begin
    r := ARow - FHeaderCount;
    c := ACol - FHeaderCount;
    //lCell := FDrawingCell;
    lCell := FWorksheet.FindCell(r, c);
    if lCell <> nil then
    begin
      // Background color
      if (uffBackgroundColor in lCell^.UsedFormattingFields) then
      begin
        if FWorkbook.FileFormat = sfExcel2 then
        begin
          if (FillPattern_BIFF2 = nil) and (ComponentState = []) then
            Create_FillPattern_BIFF2(Color);
          Canvas.Brush.Style := bsImage;
          Canvas.Brush.Bitmap := FillPattern_BIFF2;
        end else
        begin
          Canvas.Brush.Style := bsSolid;
          if lCell^.BackgroundColor < FWorkbook.GetPaletteSize then
            Canvas.Brush.Color := FWorkbook.GetPaletteColor(lCell^.BackgroundColor)
          else
            Canvas.Brush.Color := Color;
        end;
      end else
      begin
        Canvas.Brush.Style := bsSolid;
        Canvas.Brush.Color := Color;
      end;
      // Font
      if (uffFont in lCell^.UsedFormattingFields) then
      begin
        fnt := FWorkbook.GetFont(lCell^.FontIndex);
        if fnt <> nil then
        begin
          Canvas.Font.Name := fnt.FontName;
          Canvas.Font.Color := FWorkbook.GetPaletteColor(fnt.Color);
          style := [];
          if fssBold in fnt.Style then Include(style, fsBold);
          if fssItalic in fnt.Style then Include(style, fsItalic);
          if fssUnderline in fnt.Style then Include(style, fsUnderline);
          if fssStrikeout in fnt.Style then Include(style, fsStrikeout);
          Canvas.Font.Style := style;
          Canvas.Font.Size := round(fnt.Size);
        end;
      end;
      if (lCell^.NumberFormat = nfCurrencyRed) and
         not IsNaN(lCell^.NumberValue) and (lCell^.NumberValue < 0)
      then
        Canvas.Font.Color := FWorkbook.GetPaletteColor(scRed);
      // Wordwrap, text alignment and text rotation are handled by "DrawTextInCell".
    end;
  end;

  if IsSelected then
    Canvas.Brush.Color := CalcSelectionColor(Canvas.Brush.Color, 16);

  Canvas.TextStyle := ts;

  inherited DoPrepareCanvas(ACol, ARow, AState);
end;

{@@ ----------------------------------------------------------------------------
  This method is inherited from TsCustomGrid, but is overridden here in order
  to paint the cell borders and the selection rectangle.
  Both features can extend into the neighboring cells and thus would be clipped
  at the cell borders by the standard painting mechanism. At the time when
  DrawAllRows is called, however, clipping at cell borders is no longer active.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawAllRows;
var
  cliprect: TRect;
  rgn: HRGN;
  tmp: Integer = 0;
begin
  inherited;

  Canvas.SaveHandleState;
  try
    // Avoid painting into the header cells
    cliprect := ClientRect;
    if FixedCols > 0 then
      ColRowToOffset(True, True, FixedCols-1, tmp, cliprect.Left);
    if FixedRows > 0 then
      ColRowToOffset(False, True, FixedRows-1, tmp, cliprect.Top);

    DrawFrozenPaneBorders(clipRect);

    rgn := CreateRectRgn(cliprect.Left, cliprect.top, cliprect.Right, cliprect.Bottom);
    SelectClipRgn(Canvas.Handle, Rgn);

    DrawCellBorders;
    DrawSelection;

    DeleteObject(rgn);
  finally
    Canvas.RestoreHandleState;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Draws the borders of all cells. Calls DrawCellBorder for each individual cell.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawCellBorders;
var
  cell: PCell;
  c, r: Integer;
  rect: TRect;
begin
  if FWorksheet = nil then
    exit;

  cell := FWorksheet.GetFirstCell;
  while cell <> nil do
  begin
    if (uffBorder in cell^.UsedFormattingFields) then
    begin
      c := cell^.Col + FHeaderCount;
      r := cell^.Row + FHeaderCount;
      rect := CellRect(c, r);
      DrawCellBorders(c, r, rect);
    end;
    cell := FWorksheet.GetNextCell;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Draws the border lines around a given cell. Note that when this procedure is
  called the output is clipped by the cell rectangle, but thick and double
  border styles extend into the neighboring cell. Therefore, these border lines
  are drawn in parts.

  @param  ACol   Column Index
  @param  ARow   Row index
  @param  ARect  Rectangle in pixels occupied by the cell.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawCellBorders(ACol, ARow: Integer; ARect: TRect);
const
  drawHor = 0;
  drawVert = 1;
  drawDiagUp = 2;
  drawDiagDown = 3;

  procedure DrawBorderLine(ACoord: Integer; ARect: TRect; ADrawDirection: Byte;
    ABorderStyle: TsCellBorderStyle);
  const
    // TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair);
    PEN_STYLES: array[TsLineStyle] of TPenStyle =
      (psSolid, psSolid, psDash, psDot, psSolid, psSolid, psSolid);
    PEN_WIDTHS: array[TsLineStyle] of Integer =
      (1, 2, 1, 1, 3, 1, 1);
  var
    width3: Boolean;     // line is 3 pixels wide
    deltax, deltay: Integer;
    angle: Double;
  begin
    Canvas.Pen.Style := PEN_STYLES[ABorderStyle.LineStyle];
    Canvas.Pen.Width := PEN_WIDTHS[ABorderStyle.LineStyle];
    Canvas.Pen.Color := FWorkbook.GetPaletteColor(ABorderStyle.Color);
    Canvas.Pen.EndCap := pecSquare;
    width3 := (ABorderStyle.LineStyle in [lsThick, lsDouble]);

    // Workaround until efficient drawing procedures for diagonal "hair" lines
    // is available
    if (ADrawDirection in [drawDiagUp, drawDiagDown]) and
       (ABorderStyle.LineStyle = lsHair)
    then
      ABorderStyle.LineStyle := lsDotted;

    // Tuning the rectangle to avoid issues at the grid borders and to get nice corners
    if (ABorderStyle.LineStyle in [lsMedium, lsThick, lsDouble]) then begin
      if ACol = ColCount-1 then
      begin
        if (ADrawDirection = drawVert) and (ACoord = ARect.Right-1) and width3
          then dec(ACoord);
        dec(ARect.Right);
      end;
      if ARow = RowCount-1 then
      begin
        if (ADrawDirection = drawHor) and (ACoord = ARect.Bottom-1) and width3
          then dec(ACoord);
        dec(ARect.Bottom);
      end;
    end;
    if ABorderStyle.LineStyle in [lsMedium, lsThick] then
    begin
      if (ADrawDirection = drawHor) then
        dec(ARect.Right, 1)
      else if (ADrawDirection = drawVert) then
        dec(ARect.Bottom, 1);
    end;

    // Painting
    case ABorderStyle.LineStyle of
      lsThin, lsMedium, lsThick, lsDotted, lsDashed:
        case ADrawDirection of
          drawHor     : Canvas.Line(ARect.Left, ACoord, ARect.Right, ACoord);
          drawVert    : Canvas.Line(ACoord, ARect.Top, ACoord, ARect.Bottom);
          drawDiagUp  : Canvas.Line(ARect.Left, ARect.Bottom, ARect.Right, ARect.Top);
          drawDiagDown: Canvas.Line(ARect.Left, ARect.Top, ARect.Right, ARect.Bottom);
        end;

      lsHair:
        case ADrawDirection of
          drawHor     : DrawHairLineHor(Canvas, ARect.Left, ARect.Right, ACoord);
          drawVert    : DrawHairLineVert(Canvas, ACoord, ARect.Top, ARect.Bottom);
          drawDiagUp  : ;
          drawDiagDown: ;
        end;

      lsDouble:
        case ADrawDirection of
          drawHor:
            begin
              Canvas.Line(ARect.Left, ACoord-1, ARect.Right, ACoord-1);
              Canvas.Line(ARect.Left, ACoord+1, ARect.Right, ACoord+1);
              Canvas.Pen.Color := Color;
              Canvas.Line(ARect.Left, ACoord, ARect.Right, ACoord);
            end;
          drawVert:
            begin
              Canvas.Line(ACoord-1, ARect.Top, ACoord-1, ARect.Bottom);
              Canvas.Line(ACoord+1, ARect.Top, ACoord+1, ARect.Bottom);
              Canvas.Pen.Color := Color;
              Canvas.Line(ACoord, ARect.Top, ACoord, ARect.Bottom);
            end;
          drawDiagUp:
            begin
              if ARect.Right = ARect.Left then
                angle := pi/2
              else
                angle := arctan((ARect.Bottom-ARect.Top) / (ARect.Right-ARect.Left));
              deltax := Max(1, round(1.0 / sin(angle)));
              deltay := Max(1, round(1.0 / cos(angle)));
              Canvas.Line(ARect.Left, ARect.Bottom-deltay-1, ARect.Right-deltax, ARect.Top-1);
              Canvas.Line(ARect.Left+deltax, ARect.Bottom-1, ARect.Right, ARect.Top+deltay-1);
            end;
          drawDiagDown:
            begin
              if ARect.Right = ARect.Left then
                angle := pi/2
              else
                angle := arctan((ARect.Bottom-ARect.Top) / (ARect.Right-ARect.Left));
              deltax := Max(1, round(1.0 / sin(angle)));
              deltay := Max(1, round(1.0 / cos(angle)));
              Canvas.Line(ARect.Left, ARect.Top+deltay-1, ARect.Right-deltax, ARect.Bottom-1);
              Canvas.Line(ARect.Left+deltax, ARect.Top-1, ARect.Right, ARect.Bottom-deltay-1);
            end;
        end;
    end;
  end;

var
  bs: TsCellBorderStyle;
  cell: PCell;
begin
  if Assigned(FWorksheet) then begin
    // Left border
    if GetBorderStyle(ACol, ARow, -1, 0, bs) then
      DrawBorderLine(ARect.Left-1, ARect, drawVert, bs);
    // Right border
    if GetBorderStyle(ACol, ARow, +1, 0, bs) then
      DrawBorderLine(ARect.Right-1, ARect, drawVert, bs);
    // Top border
    if GetBorderstyle(ACol, ARow, 0, -1, bs) then
      DrawBorderLine(ARect.Top-1, ARect, drawHor, bs);
    // Bottom border
    if GetBorderStyle(ACol, ARow, 0, +1, bs) then
      DrawBorderLine(ARect.Bottom-1, ARect, drawHor, bs);

    cell := FWorksheet.FindCell(ARow-FHeaderCount, ACol-FHeaderCount);
    if cell <> nil then begin
      // Diagonal up
      if cbDiagUp in cell^.Border then begin
        bs := cell^.Borderstyles[cbDiagUp];
        DrawBorderLine(0, ARect, drawDiagUp, bs);
      end;
      // Diagonal down
      if cbDiagDown in cell^.Border then begin
        bs := cell^.BorderStyles[cbDiagDown];
        DrawborderLine(0, ARect, drawDiagDown, bs);
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This procedure is responsible for painting the focus rectangle. We don't want
  the red dashed rectangle here, but prefer the thick Excel-like black border
  line. This new focus rectangle is drawn by the method DrawSelection.

  @param   ACol   Grid column index of the focused cell
  @param   ARow   Grid row index of the focused cell
  @param   ARect  Rectangle in pixels covered by the focused cell
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawFocusRect(aCol, aRow: Integer; ARect: TRect);
begin
  Unused(ACol, ARow, ARect);
  // Nothing do to
end;

{@@ ----------------------------------------------------------------------------
  Draws a solid line along the borders of frozen panes.

  @param  ARect  This rectangle indicates the area containing movable cells.
                 If the grid has frozen panes, a black line is drawn along the
                 upper and/or left edge of this rectangle (depending on the
                 value of FrozenRows and FrozenCols).
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawFrozenPaneBorders(ARect: TRect);
begin
  if FWorkSheet = nil then
    exit;
  if (soHasFrozenPanes in FWorksheet.Options) then begin
    Canvas.Pen.Style := psSolid;
    Canvas.Pen.Color := clBlack;
    Canvas.Pen.Width := 1;
    if FFrozenRows > 0 then
      Canvas.Line(ARect.Left, ARect.Top, ARect.Right, ARect.Top);
    if FFrozenCols > 0 then
      Canvas.Line(ARect.Left, ARect.Top, ARect.Left, ARect.Bottom);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Draws a complete row of cells. Is mostly duplicated from Grids.pas, but adds
  code for merged cells and overflow text, the section on drawing the default
  focus rectangle is removed.

  @param  ARow  Index of the row to be drawn (index in grid coordinates)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawRow(ARow: Integer);
var
  gds: TGridDrawState;
  sr, sc, sr1,sc1,sr2,sc2: Cardinal;                        // sheet row/column
  gr, gc, gcNext, gcLast, gc1, gc2, gcLastUsed: Integer;    // grid row/column
  i: Integer;
  rct, saved_rct, temp_rct: TRect;
  clipArea: Trect;
  cell: PCell;
  tmp: Integer;

  function IsPushCellActive: boolean;
  begin
    with GCache do
      result := (PushedCell.X <> -1) and (PushedCell.Y <> -1);
  end;

  function VerticalIntersect(const aRect,bRect: TRect): boolean;
  begin
    result := (aRect.Top < bRect.Bottom) and (aRect.Bottom > bRect.Top);
  end;

  function HorizontalIntersect(const aRect,bRect: TRect): boolean;
  begin
    result := (aRect.Left < bRect.Right) and (aRect.Right > bRect.Left);
  end;

  procedure DoDrawCell(_col, _row: Integer; _clipRect, _cellRect: TRect);
  var
    Rgn: HRGN;
  begin
    with GCache do begin
      if (_col = HotCell.x) and (_row = HotCell.y) and not IsPushCellActive() then begin
        Include(gds, gdHot);
        HotCellPainted := True;
      end;
      if ClickCellPushed and (_col = PushedCell.x) and (_row = PushedCell.y) then begin
        Include(gds, gdPushed);
      end;
    end;

    Canvas.SaveHandleState;
    try
      Rgn := CreateRectRgn(_clipRect.Left, _clipRect.Top, _clipRect.Right, _clipRect.Bottom);
      SelectClipRgn(Canvas.Handle, Rgn);
      DrawCell(_col, _row, _cellRect, gds);
      DeleteObject(Rgn);
    finally
      Canvas.RestoreHandleState;
    end;
  end;

begin
  // Upper and Lower bounds for this row
  rct := Rect(0, 0, 0, 0);
  ColRowToOffSet(False, True, ARow, rct.Top, rct.Bottom);
  saved_rct := rct;

  // is this row within the ClipRect?
  clipArea := Canvas.ClipRect;
  if (rct.Top >= rct.Bottom) or not VerticalIntersect(rct, clipArea) then begin
    {$IFDEF DbgVisualChange}
    DebugLn('Drawrow: Skipped row: ', IntToStr(aRow));
    {$ENDIF}
    exit;
  end;

  sr := GetWorksheetRow(ARow);

  // Draw columns in this row
  with GCache.VisibleGrid do
  begin
    gc := Left;

    // Because of possible cell overflow from cells left of the visible range
    // we have to seek to the left for the first occupied text cell
    // and start painting from here.
    if FTextOverflow and (sr <> Cardinal(-1)) and Assigned(FWorksheet) then
      while (gc > FixedCols) do
      begin
        dec(gc);
        cell := FWorksheet.FindCell(sr, GetWorksheetCol(gc));
        // Empty cell --> proceed with next cell to the left
        if (cell = nil) or (cell^.ContentType = cctEmpty) or
           ((cell^.ContentType = cctUTF8String) and (cell^.UTF8StringValue = ''))
        then
          Continue;
        // Overflow possible from non-merged, non-right-aligned, horizontal label cells
        if (cell^.MergeBase = nil) and (cell^.ContentType = cctUTF8String) and
           not (uffTextRotation in cell^.UsedFormattingFields) and
           (uffHorAlign in cell^.UsedFormattingFields) and (cell^.HorAlignment <> haRight)
        then
          Break;
        // All other cases --> no overflow --> return to initial left cell
        gc := Left;
        break;
      end;

    // Now find the last column. Again text can overflow into the visible area
    // from cells to the right.
    gcLast := Right;
    if FTextOverflow and (sr <> Cardinal(-1)) and Assigned(FWorksheet) then
    begin
      gcLastUsed := GetGridCol(FWorksheet.GetLastOccupiedColIndex);
      while (gcLast < ColCount-1) and (gcLast < gcLastUsed) do begin
        inc(gcLast);
        cell := FWorksheet.FindCell(sr, GetWorksheetCol(gcLast));
        // Empty cell --> proceed with next cell to the right
        if (cell = nil) or (cell^.ContentType = cctEmpty) or
           ((cell^.ContentType = cctUTF8String) and (cell^.UTF8StringValue = ''))
        then
          continue;
        // Overflow possible from non-merged, horizontal, non-left-aligned label cells
        if (cell^.MergeBase = nil) and (cell^.ContentType = cctUTF8String) and
           not (uffTextRotation in cell^.UsedFormattingFields) and
           (uffHorAlign in cell^.UsedFormattingFields) and (cell^.HorAlignment <> haLeft)
        then
          Break;
        // All other cases --> no overflow --> return to initial right column
        gcLast := Right;
        Break;
      end;
    end;

    // Here begins the drawing loop of all cells in the row
    while (gc <= gcLast) do begin
      gr := ARow;
      rct := saved_rct;
      // FDrawingCell is the cell which is currently being painted. We store
      // it to avoid excessive calls to "FindCell".
      FDrawingCell := nil;
      gcNext := gc + 1;
      if Assigned(FWorksheet) and (gr >= FixedRows) and (gc >= FixedCols) then
      begin
        cell := FWorksheet.FindCell(GetWorksheetRow(gr), GetWorksheetCol(gc));
        if (cell = nil) or (cell^.Mergebase = nil) then
        begin
          // single cell
          FDrawingCell := cell;
          // Special treatment of overflowing cells
          if FTextOverflow then
          begin
            gds := GetGridDrawState(gc, gr);
            ColRowToOffset(true, true, gc, rct.Left, rct.Right);
            if CellOverflow(gc, gr, gds, gc1, gc2, rct) then
            begin
              // Draw individual cells of the overflown range
              ColRowToOffset(true, true, gc1, rct.Left, tmp);    // rct is the clip rect
              ColRowToOffset(true, true, gc2, tmp, rct.Right);
              FDrawingCell := nil;
              temp_rct := rct;
              for i := gc1 to gc2 do begin
                ColRowToOffset(true, true, i, temp_rct.Left, temp_rct.Right);
                if HorizontalIntersect(temp_rct, clipArea) and (i <> gc) then
                begin
                  gds := GetGridDrawState(i, gr);
                  DoDrawCell(i, gr, rct, temp_rct);
                end;
              end;
              // Repaint the base cell text (it was partly overwritten before)
              FDrawingCell := cell;
              FTextOverflowing := true;
              ColRowToOffset(true, true, gc, temp_rct.Left, temp_rct.Right);
              if HorizontalIntersect(temp_rct, clipArea) then
              begin
                gds := GetGridDrawState(gc, gr);
                DoDrawCell(gc, gr, rct, temp_rct);
              end;
              FTextOverflowing := false;

              gcNext := gc2 + 1;
              gc := gcNext;
              continue;
            end;
          end;
        end
        else
        begin
          // merged cells
          FDrawingCell := FWorksheet.FindMergeBase(cell);
          FWorksheet.FindMergedRange(FDrawingCell, sr1, sc1, sr2, sc2);
          gr := GetGridRow(sr1);
          ColRowToOffSet(False, True, gr, rct.Top, tmp);
          ColRowToOffSet(False, True, gr + sr2 - sr1, tmp, rct.Bottom);
          gc := GetGridCol(sc1);
          gcNext := gc + (sc2 - sc1) + 1;
        end;
      end;

      ColRowToOffset(true, true, gc, rct.Left, tmp);
      ColRowToOffset(true, true, gcNext-1, tmp, rct.Right);

      if (rct.Left < rct.Right) and HorizontalIntersect(rct, clipArea) then
      begin
        gds := GetGridDrawState(gc, gr);
        DoDrawCell(gc, gr, rct, rct);
      end;

      gc := gcNext;
    end;
  end;    // with GCache.VisibleGrid ...

  // Draw Fixed Columns
  gr := ARow;
  for gc := 0 to FixedCols-1 do begin
    gds := [gdFixed];
    ColRowToOffset(True, True, gc, rct.Left, rct.Right);
    // is this column within the ClipRect?
    if (rct.Left < rct.Right) and HorizontalIntersect(rct, clipArea) then
    begin
      if Assigned(FWorksheet) then
        FDrawingCell := FWorksheet.FindCell(GetWorksheetRow(gr), GetWorksheetCol(gc))
      else
        FDrawingCell := nil;
      DoDrawCell(gc, gr, rct, rct);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Draws the selection rectangle around selected cells, 3 pixels wide as in Excel.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawSelection;
var
  P1, P2: TPoint;
begin
  // Cosmetics at the edges of the grid to avoid spurious rests
  P1 := CellRect(Selection.Left, Selection.Top).TopLeft;
  P2 := CellRect(Selection.Right, Selection.Bottom).BottomRight;
  if Selection.Top > TopRow then dec(P1.Y) else inc(P1.Y);
  if Selection.Left > LeftCol then dec(P1.X) else inc(P1.X);
  if Selection.Right = ColCount-1 then dec(P2.X);
  if Selection.Bottom = RowCount-1 then dec(P2.Y);
  // Set up the canvas
  Canvas.Pen.Style := psSolid;
  Canvas.Pen.Width := 3;
  Canvas.Pen.JoinStyle := pjsMiter;
  if UseXORFeatures then begin
    Canvas.Pen.Color := clWhite;
    Canvas.Pen.Mode := pmXOR;
  end else
    Canvas.Pen.Color := clBlack;
  Canvas.Brush.Style := bsClear;
  // Paint
  Canvas.Rectangle(P1.X, P1.Y, P2.X, P2.Y);
end;

{@@ ----------------------------------------------------------------------------
  Draws the cell text. Calls "GetCellText" to determine the text for the cell.
  Takes care of horizontal and vertical text alignment, text rotation and
  text wrapping.

  @param  ACol   Grid column index of the cell
  @param  ARow   Grid row index of the cell
  @param  ARect  Rectangle in pixels occupied by the cell
  @param  AState Drawing state of the grid -- see TCustomGrid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawTextInCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
var
  ts: TTextStyle;
  txt: String;
  c, r: Integer;
  wrapped: Boolean;
  horAlign: TsHorAlignment;
  vertAlign: TsVertAlignment;
  txtRot: TsTextRotation;
  lCell: PCell;
  txtLeft, txtRight: String;
  justif: Byte;
begin
  if (FWorksheet = nil) then
    exit;

  if (ACol < FHeaderCount) or (ARow < FHeaderCount) then
    lCell := nil
  else
    lCell := FDrawingCell;

  // Header
  if lCell = nil then
  begin
    if ShowHeaders and ((ACol = 0) or (ARow = 0)) then
    begin
      ts.Alignment := taCenter;
      ts.Layout := tlCenter;
      ts.Opaque := false;
      Canvas.TextStyle := ts;
    end;
    inherited DrawCellText(aCol, aRow, aRect, aState, GetCellText(ACol,ARow));
    exit;
  end;

  // Cells
  wrapped := (uffWordWrap in lCell^.UsedFormattingFields) or (lCell^.TextRotation = rtStacked);
  txtRot := lCell^.TextRotation;
  vertAlign := lCell^.VertAlignment;
  if vertAlign = vaDefault then vertAlign := vaBottom;
  if lCell^.HorAlignment <> haDefault then
    horAlign := lCell^.HorAlignment
  else
  begin
    if (lCell^.ContentType in [cctNumber, cctDateTime]) then
      horAlign := haRight
    else
      horAlign := haLeft;
  end;

  InflateRect(ARect, -constCellPadding, -constCellPadding);

//  txt := GetCellText(ACol, ARow);
  txt := GetCellText(GetGridRow(lCell^.Col), GetGridCol(lCell^.Row));
  if txt = '' then
    exit;

  case txtRot of
    trHorizontal:
      case horAlign of
        haLeft   : justif := 0;
        haCenter : justif := 1;
        haRight  : justif := 2;
      end;
    rtStacked,
    rt90DegreeClockwiseRotation:
      case vertAlign of
        vaTop   : justif := 0;
        vaCenter: justif := 1;
        vaBottom: justif := 2;
      end;
    rt90DegreeCounterClockwiseRotation:
      case vertAlign of
        vaTop   : justif := 2;
        vaCenter: justif := 1;
        vaBottom: justif := 0;
      end;
  end;
  InternalDrawTextInCell(txt, txt, ARect, justif, horAlign, vertAlign,
    txtRot, wrapped, false);
end;

{@@ ----------------------------------------------------------------------------
  This procedure is called when editing of a cell is completed. It determines
  the worksheet cell and writes the text into the worksheet. Tries to keep the
  format of the cell, but if it is a new cell, or the content type has changed,
  tries to figure out the content type (number, date/time, text).
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.EditingDone;
var
  oldText: String;
  cell: PCell;
begin
  if (not EditorShowing) and FEditing then
  begin
    oldText := GetCellText(Col, Row);
    if oldText <> FEditText then
    begin
      if FWorksheet = nil then
        FWorksheet := TsWorksheet.Create;
      cell := FWorksheet.GetCell(Row-FHeaderCount, Col-FHeaderCount);
      FWorksheet.WriteCellValueAsString(cell, FEditText);
      FEditText := '';
    end;
    inherited EditingDone;
  end;
  FEditing := false;
end;

{@@ ----------------------------------------------------------------------------
  The BeginUpdate/EndUpdate pair suppresses unnecessary painting of the grid.
  Call BeginUpdate to stop refreshing the grid, and call EndUpdate to release
  the lock and to repaint the grid again.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.EndUpdate;
begin
  dec(FLockCount);
  if FLockCount = 0 then Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Copies the borders of a cell to its neighbors. This avoids the nightmare of
  changing borders due to border conflicts of adjacent cells.

  @param  ACol  Grid column index of the cell
  @param  ARow  Grid row index of the cell
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.FixNeighborCellBorders(ACol, ARow: Integer);

  procedure SetNeighborBorder(NewRow, NewCol: Integer;
    ANewBorder: TsCellBorder; const ANewBorderStyle: TsCellBorderStyle;
    AInclude: Boolean);
  var
    neighbor: PCell;
    border: TsCellBorders;
  begin
    neighbor := FWorksheet.FindCell(NewRow, NewCol);
    if neighbor <> nil then
    begin
      border := neighbor^.Border;
      if AInclude then
      begin
        Include(border, ANewBorder);
        FWorksheet.WriteBorderStyle(NewRow, NewCol, ANewBorder, ANewBorderStyle);
      end else
        Exclude(border, ANewBorder);
      FWorksheet.WriteBorders(NewRow, NewCol, border);
    end;
  end;

var
  cell: PCell;
begin
  if FWorksheet = nil then
    exit;

  cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
  if (FWorksheet <> nil) and (cell <> nil) then
    with cell^ do
    begin
      SetNeighborBorder(Row, Col-1, cbEast, BorderStyles[cbWest], cbWest in Border);
      SetNeighborBorder(Row, Col+1, cbWest, BorderStyles[cbEast], cbEast in Border);
      SetNeighborBorder(Row-1, Col, cbSouth, BorderStyles[cbNorth], cbNorth in Border);
      SetNeighborBorder(Row+1, Col, cbNorth, BorderStyles[cbSouth], cbSouth in Border);
    end;
end;

{@@ ----------------------------------------------------------------------------
  The "colors" used by the spreadsheet are indexes into the workbook's color
  palette. If the user wants to set a color to a particular RGB value this is
  not possible in general. The method FindNearestPaletteIndex finds the bast
  matching color in the palette.

  @param  AColor  Color index into the workbook's palette
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.FindNearestPaletteIndex(AColor: TColor): TsColor;

  procedure ColorToHSL(RGB: TColor; out H, S, L : double);
  // Taken from https://code.google.com/p/thtmlviewer/source/browse/trunk/source/HSLUtils.pas?r=277
  // The procedure in GraphUtils crashes for some colors in Laz < 1.3
  var
    R, G, B, D, Cmax, Cmin: double;
  begin
    R := GetRValue(RGB) / 255;
    G := GetGValue(RGB) / 255;
    B := GetBValue(RGB) / 255;
    Cmax := Max(R, Max(G, B));
    Cmin := Min(R, Min(G, B));

    // calculate luminosity
    L := (Cmax + Cmin) / 2;

    if Cmax = Cmin then begin // it's grey
      H := 0; // it's actually undefined
      S := 0
    end else
    begin
      D := Cmax - Cmin;

      // calculate Saturation
      if L < 0.5 then
        S := D / (Cmax + Cmin)
      else
        S := D / (2 - Cmax - Cmin);

      // calculate Hue
      if R = Cmax then
        H := (G - B) / D
      else
      if G = Cmax then
        H := 2 + (B - R) /D
      else
        H := 4 + (R - G) / D;

      H := H / 6;
      if H < 0 then
        H := H + 1
    end
  end;

  function ColorDistance(color1, color2: TColor): Double;
  var
    H1,S1,L1, H2,S2,L2: Double;
  begin
    ColorToHSL(color1, H1, S1, L1);
    ColorToHSL(color2, H2, S2, L2);
    Result := sqr(H1-H2) + sqr(S1-S2) + sqr(L1-L2);
  end;

  {
  // To be activated when Lazarus 1.4 is available. (RgbToHLS bug in Laz < 1.3)

  function ColorDistance(color1, color2: TColor): Integer;
  type
    TRGBA = packed record R, G, B, A: Byte end;
  var
    H1,L1,S1, H2,L2,S2: Byte;
  begin
    ColorToHLS(color1, H1,L1,S1);
    ColorToHLS(color2, H2,L2,S2);
    result := sqr(Integer(H1)-H2) + sqr(Integer(L1)-L2) + sqr(Integer(S1)-S2);
  end;
  }

var
  i: Integer;
  dist, mindist: Double;
begin
  Result := 0;
  if Workbook <> nil then
  begin
    mindist := 1E308;
    for i:=0 to Workbook.GetPaletteSize-1 do
    begin
      dist := ColorDistance(AColor, TColor(Workbook.GetPaletteColor(i)));
      if dist < mindist then
      begin
        mindist := dist;
        Result := i;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the background color of a cell. The color is given as an index into
  the workbook's color palette.

  @param  ACol  Grid column index of the cell
  @param  ARow  Grid row index of the cell
  @result Color index of the cell's background color.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetBackgroundColor(ACol, ARow: Integer): TsColor;
var
  cell: PCell;
begin
  Result := scNotDefined;
  if Assigned(FWorksheet) then
  begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) and (uffBackgroundColor in cell^.UsedFormattingFields) then
      Result := cell^.BackgroundColor;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the background color of a cell range defined by a rectangle. The color
  is given as an index into the workbook's color palette. If the colors are
  different from cell to cell the value scUndefined is returned.

  @param  ARect  Cell range defined as a rectangle: Left/Top refers to the cell
                 in the left/top corner of the selection, Right/Bottom to the
                 right/bottom corner.
  @return Color index common to all cells within the selection. If the cells'
          background colors are different the value scUndefined is returned.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetBackgroundColors(ARect: TGridRect): TsColor;
var
  c, r: Integer;
  clr: TsColor;
begin
  Result := GetBackgroundColor(ARect.Left, ARect.Top);
  clr := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do
    begin
      Result := GetBackgroundColor(c, r);
      if Result <> clr then
      begin
        Result := scNotDefined;
        exit;
      end;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the cell borders which are drawn around a given cell.

  @param  ACol  Grid column index of the cell
  @param  ARow  Grid row index of the cell
  @return Set with flags indicating where borders are drawn (top/left/right/bottom)
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorder(ACol, ARow: Integer): TsCellBorders;
var
  cell: PCell;
begin
  Result := [];
  if Assigned(FWorksheet) then
  begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) and (uffBorder in cell^.UsedFormattingFields) then
      Result := cell^.Border;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the cell borders which are drawn around a given rectangular cell range.

  @param  ARect  Rectangle defining the range of cell.
  @return Set with flags indicating where borders are drawn (top/left/right/bottom)
          If the individual cells within the range have different borders an
          empty set is returned.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorders(ARect: TGridRect): TsCellBorders;
var
  c, r: Integer;
  b: TsCellBorders;
begin
  Result := GetCellBorder(ARect.Left, ARect.Top);
  b := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do
    begin
      Result := GetCellBorder(c, r);
      if Result <> b then
      begin
        Result := [];
        exit;
      end;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the style of the cell border line drawn along the edge specified
  by the parameter ABorder of a cell. The style is defined by line style and
  line color.

  @param   ACol     Grid column index of the cell
  @param   ARow     Grid row index of the cell
  @param   ABorder  Identifier of the border at which the line will be drawn
                    (see TsCellBorder)
  @return  CellBorderStyle record containing information on line style and
           line color.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorderStyle(ACol, ARow: Integer;
  ABorder: TsCellBorder): TsCellBorderStyle;
var
  cell: PCell;
begin
  Result := DEFAULT_BORDERSTYLES[ABorder];
  if Assigned(FWorksheet) then
  begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then
      Result := cell^.BorderStyles[ABorder];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the style of the cell border line drawn along the edge specified
  by the parameter ABorder of a range of cells defined by the rectangle of
  column and row indexes. The style is defined by linestyle and line color.

  @param   ARect    Rectangle whose edges define the limits of the grid row and
                    column indexes of the cells.
  @param   ABorder  Identifier of the border where the line will be drawn
                    (see TsCellBorder)
  @return  CellBorderStyle record containing information on line style and
           line color.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorderStyles(ARect: TGridRect;
  ABorder: TsCellBorder): TsCellBorderStyle;
var
  c, r: Integer;
  bs: TsCellBorderStyle;
begin
  Result := GetCellBorderStyle(ARect.Left, ARect.Top, ABorder);
  bs := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do
    begin
      Result := GetCellBorderStyle(c, r, ABorder);
      if (Result.LineStyle <> bs.LineStyle) or (Result.Color <> bs.Color) then
      begin
        Result := DEFAULT_BORDERSTYLES[ABorder];
        exit;
      end;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the font to be used when painting text in a cell.

  @param   ACol     Grid column index of the cell
  @param   ARow     Grid row index of the cell
  @return  Font usable when painting on a canvas.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellFont(ACol, ARow: Integer): TFont;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := nil;
  if (FWorkbook <> nil) and (FWorksheet <> nil) then
  begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then
    begin
      if (uffBold in cell^.UsedFormattingFields) then
        fnt := FWorkbook.GetFont(1)
      else
      if (uffFont in cell^.UsedFormattingFields) then
        fnt := FWorkbook.GetFont(cell^.FontIndex)
      else
        fnt := FWorkbook.GetDefaultFont;
//      fnt := FWorkbook.GetFont(cell^.FontIndex);
      Convert_sFont_to_Font(fnt, FCellFont);
      Result := FCellFont;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the font to be used when painting text in the cells defined by the
  rectangle of row/column indexes.

  @param   ARect    Rectangle whose edges define the limits of the grid row and
                    column indexes of the cells.
  @return  Font usable when painting on a canvas.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellFonts(ARect: TGridRect): TFont;
var
  c, r: Integer;
  sFont, sDefFont: TsFont;
  cell: PCell;
begin
  Result := GetCellFont(ARect.Left, ARect.Top);
  sDefFont := FWorkbook.GetFont(0);  // Default font
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do
    begin
      cell := FWorksheet.FindCell(GetWorksheetRow(r), GetWorksheetCol(c));
      if cell <> nil then
      begin
        sFont := FWorkbook.GetFont(cell^.FontIndex);
        if (sFont.FontName <> sDefFont.FontName) and (sFont.Size <> sDefFont.Size)
          and (sFont.Style <> sDefFont.Style) and (sFont.Color <> sDefFont.Color)
        then
        begin
          Convert_sFont_to_Font(sDefFont, FCellFont);
          Result := FCellFont;
          exit;
        end;
      end;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the height (in pixels) of the cell at ACol/ARow (of the grid).

  @param   ACol  Grid column index of the cell
  @param   ARow  Grid row index of the cell
  @result  Height of the cell in pixels. Wrapped text is handled correctly.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellHeight(ACol, ARow: Integer): Integer;
var
  lCell: PCell;
  s: String;
  wrapped: Boolean;
  txtR: TRect;
  cellR: TRect;
  flags: Cardinal;
  r1,c1,r2,c2: Cardinal;
begin
  Result := 0;
  if ShowHeaders and ((ACol = 0) or (ARow = 0)) then
    exit;
  if FWorksheet = nil then
    exit;

  lCell := FWorksheet.FindCell(ARow-FHeaderCount, ACol-FHeaderCount);
  if lCell <> nil then
  begin
    //if lCell^.MergedNeighbors <> [] then begin
    if (lCell^.Mergebase <> nil) then
    begin
      FWorksheet.FindMergedRange(lCell, r1, c1, r2, c2);
      if r1 <> r2 then
        // If the merged range encloses several rows we skip automatic row height
        // determination since only the height of the first row of the block
        // (containing the merge base cell) would change which is very confusing.
        exit;
    end;
    s := GetCellText(ACol, ARow);
    if s = '' then
      exit;
    DoPrepareCanvas(ACol, ARow, []);
    wrapped := (uffWordWrap in lCell^.UsedFormattingFields)
      or (lCell^.TextRotation = rtStacked);
    // *** multi-line text ***
    if wrapped then
    begin
      // horizontal
      if ( (uffTextRotation in lCell^.UsedFormattingFields) and
           (lCell^.TextRotation in [trHorizontal, rtStacked]))
         or not (uffTextRotation in lCell^.UsedFormattingFields)
      then
      begin
        cellR := CellRect(ACol, ARow);
        InflateRect(cellR, -constCellPadding, -constCellPadding);
        txtR := Bounds(cellR.Left, cellR.Top, cellR.Right-cellR.Left, cellR.Bottom-cellR.Top);
        flags := DT_WORDBREAK and not DT_SINGLELINE;
        LCLIntf.DrawText(Canvas.Handle, PChar(s), Length(s), txtR,
          DT_CALCRECT or flags);
        Result := txtR.Bottom - txtR.Top + 2*constCellPadding;
      end;
      // rotated wrapped text:
      // do not consider this because wrapping affects cell height.
    end else
    // *** single-line text ***
    begin
      // not rotated
      if ( not (uffTextRotation in lCell^.UsedFormattingFields) or
           (lCell^.TextRotation = trHorizontal) )
      then
        Result := Canvas.TextHeight(s) + 2*constCellPadding
      else
      // rotated by +/- 90°
      if (uffTextRotation in lCell^.UsedFormattingFields) and
         (lCell^.TextRotation in [rt90DegreeClockwiseRotation, rt90DegreeCounterClockwiseRotation])
      then
        Result := Canvas.TextWidth(s) + 2*constCellPadding;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This function returns the text to be shown in a grid cell. The text is looked
  up in the corresponding cell of the worksheet by calling its ReadAsUTF8Text
  method. In case of "stacked" text rotation, line endings are inserted after
  each character.

  @param   ACol   Grid column index of the cell
  @param   ARow   Grid row index of the cell
  @return  Text to be displayed in the cell.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellText(ACol, ARow: Integer): String;
var
  lCell: PCell;
  r, c, i: Integer;
  s: String;
begin
  Result := '';

  if ShowHeaders then
  begin
    // Headers
    if (ARow = 0) and (ACol = 0) then
      exit;
    if (ARow = 0) then
    begin
      Result := GetColString(ACol-FHeaderCount);
      exit;
    end
    else
    if (ACol = 0) then
    begin
      Result := IntToStr(ARow);
      exit;
    end;
  end;

  if FWorksheet <> nil then
  begin
    r := ARow - FHeaderCount;
    c := ACol - FHeaderCount;
    lCell := FWorksheet.FindCell(r, c);
    if lCell <> nil then
    begin
      Result := FWorksheet.ReadAsUTF8Text(lCell);
      if lCell^.TextRotation = rtStacked then
      begin
        s := Result;
        Result := '';
        for i:=1 to Length(s) do
        begin
          Result := Result + s[i];
          if i < Length(s) then Result := Result + LineEnding;
        end;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines the text to be passed to the cell editor. The text is determined
  from the underlying worksheet cell, but it is possible to intercept this by
  adding a handler for the OnGetEditText event.

  @param   ACol   Grid column index of the cell being edited
  @param   ARow   Grid row index of the grid cell being edited
  @return  Text to be passed to the cell editor.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetEditText(aCol, aRow: Integer): string;
begin
  Result := GetCellText(aCol, aRow);
  if Assigned(OnGetEditText) then OnGetEditText(Self, aCol, aRow, result);
end;

{@@ ----------------------------------------------------------------------------
  Determines the style of the border between a cell and its neighbor given by
  ADeltaCol and ADeltaRow (one of them must be 0, the other one can only be +/-1).
  ACol and ARow are in grid units.
  Result is FALSE if there is no border line.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer;
  out ABorderStyle: TsCellBorderStyle): Boolean;
var
  cell, neighborcell: PCell;
  border, neighborborder: TsCellBorder;
  r, c: Cardinal;
begin
  Result := true;
  if (ADeltaCol = -1) and (ADeltaRow = 0) then
  begin
    border := cbWest;
    neighborborder := cbEast;
  end else
  if (ADeltaCol = +1) and (ADeltaRow = 0) then
  begin
    border := cbEast;
    neighborborder := cbWest;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = -1) then
  begin
    border := cbNorth;
    neighborborder := cbSouth;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = +1) then
  begin
    border := cbSouth;
    neighborBorder := cbNorth;
  end else
    raise Exception.Create('[TsCustomWorksheetGrid] Incorrect col/row for GetBorderStyle.');

  r := GetWorksheetRow(ARow);
  c := GetWorksheetCol(ACol);
  cell := FWorksheet.FindCell(r, c);
  if (ARow - FHeaderCount + ADeltaRow < 0) or (ACol - FHeaderCount + ADeltaCol < 0) then
    neighborcell := nil
  else
    neighborcell := FWorksheet.FindCell(ARow - FHeaderCount + ADeltaRow, ACol - FHeaderCount + ADeltaCol);

  // Only cell has border, but neighbor has not
  if HasBorder(cell, border) and not HasBorder(neighborCell, neighborBorder) then
  begin
    if FWorksheet.IsMerged(cell) and FWorksheet.IsMerged(neighborcell) and
       (cell^.MergeBase = neighborcell^.Mergebase)
    then
      result := false
    else
      ABorderStyle := cell^.BorderStyles[border]
  end
  else
  // Only neighbor has border, cell has not
  if not HasBorder(cell, border) and HasBorder(neighborCell, neighborBorder) then
  begin
    if FWorksheet.IsMerged(cell) and FWorksheet.IsMerged(neighborcell) and
       (cell^.MergeBase = neighborcell^.Mergebase)
    then
      result := false
    else
      ABorderStyle := neighborcell^.BorderStyles[neighborborder]
  end
  else
  // Both cells have shared border -> use top or left border
  if HasBorder(cell, border) and HasBorder(neighborCell, neighborBorder) then
  begin
    if FWorksheet.IsMerged(cell) and FWorksheet.IsMerged(neighborcell) and
       (cell^.MergeBase = neighborcell^.Mergebase)
    then
      result := false
    else
    if (border in [cbNorth, cbWest]) then
      ABorderStyle := neighborcell^.BorderStyles[neighborborder]
    else
      ABorderStyle := cell^.BorderStyles[border];
  end else
    Result := false;
end;

{@@ ----------------------------------------------------------------------------
  Converts a column index of the worksheet to a column index usable in the grid.
  This is required because worksheet indexes always start at zero while
  grid indexes also have to account for the column/row headers.

  @param  ASheetCol   Worksheet column index
  @return Grid column index
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetGridCol(ASheetCol: Cardinal): Integer;
begin
  Result := Integer(ASheetCol) + FHeaderCount
end;

{@@ ----------------------------------------------------------------------------
  Converts a row index of the worksheet to a row index usable in the grid.
  This is required because worksheet indexes always start at zero while
  grid indexes also have to account for the column/row headers.

  @param  ASheetRow   Worksheet row index
  @return Grid row index
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetGridRow(ASheetRow: Cardinal): Integer;
begin
  Result := Integer(ASheetRow) + FHeaderCount;
end;

{@@ ----------------------------------------------------------------------------
  Returns a list of worksheets contained in the file. Useful for assigning to
  user controls like TabControl, Combobox etc. in order to select a sheet.

  @param  ASheets  List of strings containing the names of the worksheets of
                   the workbook
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.GetSheets(const ASheets: TStrings);
var
  i: Integer;
begin
  ASheets.Clear;
  if Assigned(FWorkbook) then
    for i:=0 to FWorkbook.GetWorksheetCount-1 do
      ASheets.Add(FWorkbook.GetWorksheetByIndex(i).Name);
end;

{@@ ----------------------------------------------------------------------------
  Calculates the index of the worksheet column that is displayed in the
  given column of the grid. If the sheet headers are turned on, both numbers
  differ by 1, otherwise they are equal. Saves an "if" in cases.

  @param   AGridCol   Index of a grid column
  @return  Index of a the corresponding worksheet column
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetWorksheetCol(AGridCol: Integer): cardinal;
begin
  if (FHeaderCount > 0) and (AGridCol = 0) then
    Result := Cardinal(-1)
  else
    Result := AGridCol - FHeaderCount;
end;

{@@ ----------------------------------------------------------------------------
  Calculates the index of the worksheet row that is displayed in the
  given row of the grid. If the sheet headers are turned on, both numbers
  differ by 1, otherwise they are equal. Saves an "if" in some cases.

  @param    AGridRow  Index of a grid row
  @resturn  Index of the corresponding worksheet row.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetWorksheetRow(AGridRow: Integer): Cardinal;
begin
  if (FHeaderCount > 0) and (AGridRow = 0) then
    Result := Cardinal(-1)
  else
    Result := AGridRow - FHeaderCount;
end;

{@@ ----------------------------------------------------------------------------
  Returns true if the cell has the given border.

  @param  ACell    Pointer to cell considered
  @param  ABorder  Indicator for border to be checked for visibility
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
begin
  Result := (ACell <> nil) and (uffBorder in ACell^.UsedFormattingfields) and
    (ABorder in ACell^.Border);
end;

{@@ ----------------------------------------------------------------------------
  Inherited from TCustomGrid. Is called when column widths or row heights
  have changed. Stores the new column width or row height in the worksheet.

  @param   IsColumn   Specifies whether the changed parameter is a column width
                      (true) or a row height (false)
  @param   Index      Index of the changed column or row
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.HeaderSized(IsColumn: Boolean; AIndex: Integer);
var
  w0: Integer;
  h, h_pts: Single;
begin
  if FWorksheet = nil then
    exit;

  Convert_sFont_to_Font(FWorkbook.GetDefaultFont, Canvas.Font);
  if IsColumn then begin
    // The grid's column width is in "pixels", the worksheet's column width is
    // in "characters".
    w0 := Canvas.TextWidth('0');
    FWorksheet.WriteColWidth(GetWorksheetCol(AIndex), ColWidths[AIndex] / w0);
  end else begin
    // The grid's row heights are in "pixels", the worksheet's row heights are
    // in "lines"
    h_pts := PxToPts(RowHeights[AIndex] - 4, Screen.PixelsPerInch);  // in points
    h := h_pts / (FWorkbook.GetFont(0).Size + ROW_HEIGHT_CORRECTION);
    FWorksheet.WriteRowHeight(GetWorksheetRow(AIndex), h);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Inserts an empty column before the column specified
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.InsertCol(AGridCol: Integer);
var
  c: Cardinal;
begin
  if AGridCol < FHeaderCount then
    exit;

  if FWorksheet.GetLastColIndex + 1 + FHeaderCount >= FInitColCount then
    ColCount := ColCount + 1;
  c := AGridCol - FHeaderCount;
  FWorksheet.InsertCol(c);

  UpdateColWidths(AGridCol);
end;

{@@ ----------------------------------------------------------------------------
  Inserts an empty row before the row specified
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.InsertRow(AGridRow: Integer);
var
  r: Cardinal;
begin
  if AGridRow < FHeaderCount then
    exit;

  if FWorksheet.GetlastRowIndex+1 + FHeaderCount >= FInitRowCount then
    RowCount := RowCount + 1;
  r := AGridRow - FHeaderCount;
  FWorksheet.InsertRow(r);

  UpdateRowHeights(AGridRow);
end;

{@@ ----------------------------------------------------------------------------
  Internal general text drawing method.

  @param  AText           Text to be drawn
  @param  AMeasureText    Text used for checking if the text fits into the
                          text rectangle. If too large and ReplaceTooLong = true,
                          a series of # is drawn.
  @param  ARect           Rectangle in which the text is drawn
  @param  AJustification  Determines whether the text is drawn at the "start" (0),
                          "center" (1) or "end" (2) of the drawing rectangle.
                          Start/center/end are seen along the text drawing
                          direction.
  @param ACellHorAlign    Is the HorAlignment property stored in the cell
  @param ACellVertAlign   Is the VertAlignment property stored in the cell
  @param ATextRot         Determines the rotation angle of the text.
  @param ATextWrap        Determines if the text can wrap into multiple lines
  @param ReplaceTooLang   If true too-long texts are replaced by a series of
                          # chars filling the cell.

  @Note The reason to separate AJustification from ACellHorAlign and ACelVertAlign is
  the output of nfAccounting formatted numbers where the numbers are always
  right-aligned, and the currency symbol is left-aligned.
  THIS FEATURE IS CURRENTLY NO LONGER SUPPORTED.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.InternalDrawTextInCell(AText, AMeasureText: String;
  ARect: TRect; AJustification: Byte; ACellHorAlign: TsHorAlignment;
  ACellVertAlign: TsVertAlignment; ATextRot: TsTextRotation;
  ATextWrap, ReplaceTooLong: Boolean);
var
  ts: TTextStyle;
  flags: Cardinal;
  txt: String;
  txtRect: TRect;
  P: TPoint;
  w, h, h0, hline: Integer;
  i: Integer;
  L: TStrings;
  wrapped: Boolean;
  pLeft, pRight: Integer;
begin
  wrapped := ATextWrap or (ATextRot = rtStacked);
  if AMeasureText = '' then txt := AText else txt := AMeasureText;
  flags := DT_WORDBREAK and not DT_SINGLELINE or DT_CALCRECT;
  txtRect := ARect;

  if (ATextRot in [trHorizontal, rtStacked]) then begin
    // HORIZONAL TEXT DRAWING DIRECTION
    Canvas.Font.Orientation := 0;
    ts := Canvas.TextStyle;
    ts.Opaque := false;
    ts.Clipping := not FTextOverflowing;
    if wrapped then begin
      ts.Wordbreak := true;
      ts.SingleLine := false;
      LCLIntf.DrawText(Canvas.Handle, PChar(txt), Length(txt), txtRect, flags);
      w := txtRect.Right - txtRect.Left;
      h := txtRect.Bottom - txtRect.Top;
    end else begin
      ts.WordBreak := false;
      ts.SingleLine := false;
      w := Canvas.TextWidth(AMeasureText);
      h := Canvas.TextHeight('Tg');
    end;

    if ATextRot = rtStacked then begin
      // Stacked
      ts.Alignment := HOR_ALIGNMENTS[ACellHorAlign];
      if h > ARect.Bottom - ARect.Top then begin
        if ReplaceTooLong then begin
          txt := '#';
          repeat
            txt := txt + '#';
            LCLIntf.DrawText(Canvas.Handle, PChar(txt), Length(txt), txtRect, flags);
          until txtRect.Bottom - txtRect.Top > ARect.Bottom - ARect.Top;
          AText := copy(txt, 1, Length(txt)-1);
        end;
        ts.Layout := tlTop;
      end else
        case AJustification of
          0: ts.Layout := tlTop;
          1: ts.Layout := tlCenter;
          2: ts.Layout := tlBottom;
        end;
      Canvas.TextStyle := ts;
      Canvas.TextRect(ARect, ARect.Left, ARect.Top, AText);
    end else begin
      // Horizontal
      if h > ARect.Bottom - ARect.Top then
        ts.Layout := tlTop
      else
        ts.Layout := VERT_ALIGNMENTS[ACellVertAlign];

      // too long text
      if w > ARect.Right - ARect.Left then
        if ReplaceTooLong then
        begin
          txt := '';
          repeat
            txt := txt + '#';
            LCLIntf.DrawText(Canvas.Handle, PChar(txt), Length(txt), txtRect, flags);
          until txtRect.Right - txtRect.Left > ARect.Right - ARect.Left;
          AText := Copy(txt, 1, Length(txt)-1);
          w := Canvas.TextWidth(AText);
        end;

      P := ARect.TopLeft;
      case AJustification of
        0: ts.Alignment := taLeftJustify;
        1: if (FDrawingCell <> nil) and (FDrawingCell^.MergeBase = nil) then //(FDrawingCell^.MergedNeighbors = []) then
           begin
             // Special treatment for overflowing cells: they must be centered
             // at their original column, not in the total enclosing rectangle.
             ColRowToOffset(true, true, FDrawingCell^.Col + FHeaderCount, pLeft, pRight);
             P.X := (pLeft + pRight - w) div 2;
             P.y := ARect.Top;
             ts.Alignment := taLeftJustify;
           end
           else
             ts.Alignment := taCenter;
        2: ts.Alignment := taRightJustify;
      end;
      Canvas.TextStyle := ts;
      Canvas.TextRect(ARect, P.X, P.Y, AText);
    end;
  end
  else
  begin
    // ROTATED TEXT DRAWING DIRECTION
    // Since there is no good API for multiline rotated text, we draw the text
    // line by line.
    L := TStringList.Create;
    try
      txtRect := Bounds(ARect.Left, ARect.Top, ARect.Bottom - ARect.Top, ARect.Right - ARect.Left);
      hline := Canvas.TextHeight('Tg');
      if wrapped then begin
        // Extract wrapped lines
        L.Text := WrapText(Canvas, txt, txtRect.Right - txtRect.Left);
        // Calculate size of wrapped text
        flags := DT_WORDBREAK and not DT_SINGLELINE or DT_CALCRECT;
        LCLIntf.DrawText(Canvas.Handle, PChar(L.Text), Length(L.Text), txtRect, flags);
        w := txtRect.Right - txtRect.Left;
        h := txtRect.Bottom - txtRect.Top;
        h0 := hline;
      end
      else begin
        L.Text := txt;
        w := Canvas.TextWidth(txt);
        h := hline;
        h0 := 0;
      end;
      // w and h are seen along the text direction, not x/y!

      if w > ARect.Bottom - ARect.Top then begin
        if ReplaceTooLong then begin
          txt := '#';
          repeat
            txt := txt + '#';
          until Canvas.TextWidth(txt) > ARect.Bottom - ARect.Top;
          L.Text := Copy(txt, 1, Length(txt)-1);
        end;
      end;

      ts := Canvas.TextStyle;
      ts.SingleLine := true;      // Draw text line by line
      ts.Clipping := false;
      ts.Layout := tlTop;
      ts.Alignment := taLeftJustify;
      ts.Opaque := false;

      if ATextRot = rt90DegreeClockwiseRotation then begin
        // Clockwise
        Canvas.Font.Orientation := -900;
        case ACellHorAlign of
          haLeft   : P.X := Min(ARect.Right-1, ARect.Left + h - h0);
          haCenter : P.X := Min(ARect.Right-1, (ARect.Left + ARect.Right + h) div 2);
          haRight  : P.X := ARect.Right - 1;
        end;
        for i:= 0 to L.Count-1 do begin
          w := Canvas.TextWidth(L[i]);
          case AJustification of
            0: P.Y := ARect.Top;                             // corresponds to "top"
            1: P.Y := Max(ARect.Top, (Arect.Top + ARect.Bottom - w) div 2);  // "center"
            2: P.Y := Max(ARect.Top, ARect.Bottom - w);      // "bottom"
          end;
          Canvas.TextRect(ARect, P.X, P.Y, L[i], ts);
          dec(P.X, hline);
        end
      end
      else begin
        // Counter-clockwise
        Canvas.Font.Orientation := +900;
        case ACellHorAlign of
          haLeft   : P.X := ARect.Left;
          haCenter : P.X := Max(ARect.Left, (ARect.Left + ARect.Right - h + h0) div 2);
          haRight  : P.X := MAx(ARect.Left, ARect.Right - h + h0);
        end;
        for i:= 0 to L.Count-1 do begin
          w := Canvas.TextWidth(L[i]);
          case AJustification of
            0: P.Y := ARect.Bottom;  // like "Bottom"
            1: P.Y := Min(ARect.Bottom, (ARect.Top + ARect.Bottom + w) div 2);  // "Center"
            2: P.Y := Min(ARect.Bottom, ARect.Top + w); // like "top"
          end;
          Canvas.TextRect(ARect, P.X, P.Y, L[i], ts);
          inc(P.X, hline);
        end;
      end;
    finally
      L.Free;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard key handling method inherited from TCustomGrid. Is overridden to
  catch the ESC key during editing in order to restore the old cell text

  @param  Key    Key which has been pressed
  @param  Shift  Additional shift keys which are pressed
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.KeyDown(var Key : Word; Shift : TShiftState);
begin
  if (Key = VK_ESCAPE) and FEditing then begin
    SetEditText(Col, Row, FOldEditText);
    EditorHide;
    exit;
  end;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Standard method inherited from TCustomGrid. Is overridden to create an
  empty workbook
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Loaded;
begin
  inherited;
  NewWorkbook(FInitColCount, FInitRowCount);
end;

{@@ ----------------------------------------------------------------------------
  Loads the worksheet into the grid and displays its contents.

  @param   AWorksheet   Worksheet to be displayed in the grid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadFromWorksheet(AWorksheet: TsWorksheet);
begin
  FWorksheet := AWorksheet;
  if FWorksheet <> nil then begin
    FWorksheet.OnChangeCell := @ChangedCellHandler;
    FWorksheet.OnChangeFont := @ChangedFontHandler;
    ShowHeaders := (soShowHeaders in FWorksheet.Options);
    ShowGridLines := (soShowGridLines in FWorksheet.Options);
    if (soHasFrozenPanes in FWorksheet.Options) then begin
      FrozenCols := FWorksheet.LeftPaneWidth;
      FrozenRows := FWorksheet.TopPaneHeight;
    end else begin
      FrozenCols := 0;
      FrozenRows := 0;
    end;
    Row := FrozenRows;
    Col := FrozenCols;
  end;
  Setup;
end;

{@@ ----------------------------------------------------------------------------
  Creates a new workbook and loads the given file into it. The file is assumed
  to have the given file format. Shows the sheet with the given sheet index.

  @param   AFileName        Name of the file to be loaded
  @param   AFormat          Spreadsheet file format assumed for the file
  @param   AWorksheetIndex  Index of the worksheet to be displayed in the grid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer);
begin
  BeginUpdate;
  try
    CreateNewWorkbook;
    FWorkbook.ReadFromFile(AFileName, AFormat);
    LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AWorksheetIndex));
  finally
    EndUpdate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a new workbook and loads the given file into it. The file format
  is determined automatically. Shows the sheet with the given sheet index.

  @param   AFileName        Name of the file to be loaded
  @param   AWorksheetIndex  Index of the worksheet to be shown in the grid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AWorksheetIndex: Integer);
begin
  BeginUpdate;
  try
    CreateNewWorkbook;
    FWorkbook.ReadFromFile(AFilename);
    LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AWorksheetIndex));
  finally
    EndUpdate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Merges the selected cells to a single large cell
  Only the upper left cell can have content and formatting (which is extended
  into the other cells).
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MergeCells;
begin
  FWorksheet.MergeCells(
    GetWorksheetRow(Selection.Top),
    GetWorksheetCol(Selection.Left),
    GetWorksheetRow(Selection.Bottom),
    GetWorksheetCol(Selection.Right)
  );
end;

{@@ ----------------------------------------------------------------------------
  Standard method inherited from TCustomGrid.
  Repaints the grid after moving selection to avoid spurious rests of the
  old thick selection border.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MoveSelection;
begin
  //Refresh;
  inherited;
  Refresh;
end;

{@@ ----------------------------------------------------------------------------
  Creates a new empty workbook with the specified number of columns and rows.

  @param   AColCount   Number of columns
  @param   ARowCount   Number of rows
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.NewWorkbook(AColCount, ARowCount: Integer);
begin
  BeginUpdate;
  try
    CreateNewWorkbook;
    FWorksheet := FWorkbook.AddWorksheet('Sheet1');
    FWorksheet.OnChangeCell := @ChangedCellHandler;
    FWorksheet.OnChangeFont := @ChangedFontHandler;
    FInitColCount := AColCount;
    FInitRowCount := ARowCount;
    Setup;
  finally
    EndUpdate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Splits a merged cell block into single cells
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UnmergeCells;
begin
  FWorksheet.UnmergeCells(
    GetWorksheetRow(Selection.Top),
    GetWorksheetCol(Selection.Left)
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the workbook represented by the grid to a spreadsheet file.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved.
  @param   AFormat            Spreadsheet file format in which the file is to be
                              saved.
  @param   AOverwriteExisting If the file already exists, it is overwritten in
                              the case of AOverwriteExisting = true, or an
                              exception is raised if AOverwriteExisting = false
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SaveToSpreadsheetFile(AFileName: String;
  AFormat: TsSpreadsheetFormat; AOverwriteExisting: Boolean = true);
begin
  if FWorkbook <> nil then
    FWorkbook.WriteToFile(AFileName, AFormat, AOverwriteExisting);
end;

{@@ ----------------------------------------------------------------------------
  Saves the workbook into a file with the specified file name. If this file
  name already exists the file is overwritten if AOverwriteExisting is true.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved
                              If the file format is not known is is written
                              as BIFF8/XLS.
  @param   AOverwriteExisting If this file already exists it is overwritten if
                              AOverwriteExisting = true, or an exception is
                              raised if AOverwriteExisting = false.
}
procedure TsCustomWorksheetGrid.SaveToSpreadsheetFile(AFileName: String;
  AOverwriteExisting: Boolean = true);
begin
  if FWorkbook <> nil then
    FWorkbook.WriteToFile(AFileName, AOverwriteExisting);
end;

{@@ ----------------------------------------------------------------------------
  Standard method inherited from TCustomGrid: Is called when editing starts.
  Is overridden here to store the old text just in case that the user presses
  ESC to cancel editing.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SelectEditor;
begin
  FOldEditText := GetCellText(Col, Row);
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Loads the workbook into the grid and selects the sheet with the given index.
  "Selected" means here that the sheet is loaded into the grid.

  @param   AIndex   Index of the worksheet to be shown in the grid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SelectSheetByIndex(AIndex: Integer);
begin
  if FWorkbook <> nil then
    LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AIndex));
end;

{@@ ----------------------------------------------------------------------------
  Standard method inherited from TCustomGrid. Fetches the text that is
  currently in the editor. It is not yet transferred to the worksheet because
  input will be checked only at the end of editing.

  @param  ACol    Grid column index of the cell being edited
  @param  ARow    Grid row index of the cell being edited
  @param  AValue  String which is currently in the cell editor
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SetEditText(ACol, ARow: Longint; const AValue: string);
begin
  FEditText := AValue;
  FEditing := true;
  inherited SetEditText(aCol, aRow, aValue);
end;

{@@ ----------------------------------------------------------------------------
  Helper method for setting up the rows and columns after a new workbook is
  loaded or created. Sets up the grid's column and row count, as well as the
  initial column widths and row heights.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Setup;
begin
  if (FWorksheet = nil) or (FWorksheet.GetCellCount = 0) then begin
    if ShowHeaders then begin
      ColCount := FInitColCount + 1; //2;
      RowCount := FInitRowCount + 1; //2;
      FixedCols := 1;
      FixedRows := 1;
      ColWidths[0] := Canvas.TextWidth(' 999999 ');
    end else begin
      FixedCols := 0;
      FixedRows := 0;
      ColCount := FInitColCount; //0;
      RowCount := FInitRowCount; //0;
    end;
  end else
  if FWorksheet <> nil then begin
    Convert_sFont_to_Font(FWorkbook.GetDefaultFont, Font);
    ColCount := Max(FWorksheet.GetLastColIndex + 1 + FHeaderCount, FInitColCount);
    RowCount := Max(FWorksheet.GetLastRowIndex + 1 + FHeaderCount, FInitRowCount);
    FixedCols := FFrozenCols + FHeaderCount;
    FixedRows := FFrozenRows + FHeaderCount;
    if ShowHeaders then begin
      ColWidths[0] := Canvas.TextWidth(' 999999 ');
      RowHeights[0] := DefaultRowHeight;
    end;
  end;
  UpdateColWidths;
  UpdateRowHeights;
  Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Updates column widths according to the data in the TCol records
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UpdateColWidths(AStartIndex: Integer = 0);
var
  i: Integer;
  lCol: PCol;
  w: Integer;
begin
  if AStartIndex = 0 then AStartIndex := FHeaderCount;
  for i := AStartIndex to ColCount-1 do begin
    w := DefaultColWidth;
    if FWorksheet <> nil then
    begin
      lCol := FWorksheet.FindCol(i - FHeaderCount);
      if lCol <> nil then
        w := CalcColWidth(lCol^.Width)
    end;
    ColWidths[i] := w;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Updates row heights by using the data from the TRow records or by auto-
  calculating the row height from the max of the cell heights
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UpdateRowHeights(AStartIndex: Integer = 0);
var
  i: Integer;
  lRow: PRow;
  h: Integer;
begin
  if AStartIndex <= 0 then AStartIndex := FHeaderCount;
  for i := AStartIndex to RowCount-1 do begin
    h := CalcAutoRowHeight(i);
    if FWorksheet <> nil then
    begin
      lRow := FWorksheet.FindRow(i - FHeaderCount);
      if (lRow <> nil) then
        h := CalcRowHeight(lRow^.Height);
    end;
    RowHeights[i] := h;
  end;
end;


{*******************************************************************************
*                      Setter / getter methods                                 *
*******************************************************************************}

function TsCustomWorksheetGrid.GetCellFontColor(ACol, ARow: Integer): TsColor;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := scNotDefined;
  if (FWorkbook <> nil) and (FWorksheet <> nil) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then begin
      fnt := FWorkbook.GetFont(cell^.FontIndex);
      if fnt <> nil then
        Result := fnt.Color;
    end;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontColors(ARect: TGridRect): TsColor;
var
  c, r: Integer;
  clr: TsColor;
begin
  Result := GetCellFontColor(ARect.Left, ARect.Top);
  clr := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetCellFontColor(c, r);
      if (Result <> clr) then begin
        Result := scNotDefined;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellFontName(ACol, ARow: Integer): String;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := '';
  if (FWorkbook <> nil) and (FWorksheet <> nil) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then begin
      fnt := FWorkbook.GetFont(cell^.FontIndex);
      if fnt <> nil then
        Result := fnt.FontName;
    end;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontNames(ARect: TGridRect): String;
var
  c, r: Integer;
  s: String;
begin
  Result := GetCellFontName(ARect.Left, ARect.Top);
  s := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetCellFontName(c, r);
      if (Result <> '') and (Result <> s) then begin
        Result := '';
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellFontSize(ACol, ARow: Integer): Single;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := -1.0;
  if (FWorkbook <> nil) and (FWorksheet <> nil) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then begin
      fnt := FWorkbook.GetFont(cell^.FontIndex);
      if fnt <> nil then
        Result := fnt.Size;
    end;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontSizes(ARect: TGridRect): Single;
var
  c, r: Integer;
  sz: Single;
begin
  Result := GetCellFontSize(ARect.Left, ARect.Top);
  sz := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetCellFontSize(c, r);
      if (Result <> -1) and not SameValue(Result, sz, 1E-3) then begin
        Result := -1.0;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellFontStyle(ACol, ARow: Integer): TsFontStyles;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := [];
  if (FWorkbook <> nil) and (FWorksheet <> nil) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then begin
      fnt := FWorkbook.GetFont(cell^.FontIndex);
      if fnt <> nil then
        Result := fnt.Style;
    end;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontStyles(ARect: TGridRect): TsFontStyles;
var
  c, r: Integer;
  style: TsFontStyles;
begin
  Result := GetCellFontStyle(ARect.Left, ARect.Top);
  style := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetCellFontStyle(c, r);
      if Result <> style then begin
        Result := [];
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetHorAlignment(ACol, ARow: Integer): TsHorAlignment;
var
  cell: PCell;
begin
  Result := haDefault;
  if Assigned(FWorksheet) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if cell <> nil then
      Result := cell^.HorAlignment;
  end;
end;

function TsCustomWorksheetGrid.GetHorAlignments(ARect: TGridRect): TsHorAlignment;
var
  c, r: Integer;
  horalign: TsHorAlignment;
begin
  Result := GetHorAlignment(ARect.Left, ARect.Top);
  horalign := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetHorAlignment(c, r);
      if Result <> horalign then begin
        Result := haDefault;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetShowGridLines: Boolean;
begin
  Result := (Options * [goHorzLine, goVertLine] <> []);
end;

function TsCustomWorksheetGrid.GetShowHeaders: Boolean;
begin
  Result := FHeaderCount <> 0;
end;

function TsCustomWorksheetGrid.GetTextRotation(ACol, ARow: Integer): TsTextRotation;
var
  cell: PCell;
begin
  Result := trHorizontal;
  if Assigned(FWorksheet) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then
      Result := cell^.TextRotation;
  end;
end;

function TsCustomWorksheetGrid.GetTextRotations(ARect: TGridRect): TsTextRotation;
var
  c, r: Integer;
  textrot: TsTextRotation;
begin
  Result := GetTextRotation(ARect.Left, ARect.Top);
  textrot := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetTextRotation(c, r);
      if Result <> textrot then begin
        Result := trHorizontal;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetVertAlignment(ACol, ARow: Integer): TsVertAlignment;
var
  cell: PCell;
begin
  Result := vaDefault;
  if Assigned(FWorksheet) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if cell <> nil then
      Result := cell^.VertAlignment;
  end;
end;

function TsCustomWorksheetGrid.GetVertAlignments(ARect: TGridRect): TsVertAlignment;
var
  c, r: Integer;
  vertalign: TsVertAlignment;
begin
  Result := GetVertalignment(ARect.Left, ARect.Top);
  vertalign := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetVertAlignment(c, r);
      if Result <> vertalign then begin
        Result := vaDefault;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetWordwrap(ACol, ARow: Integer): Boolean;
var
  cell: PCell;
begin
  Result := false;
  if Assigned(FWorksheet) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) and (uffWordwrap in cell^.UsedFormattingFields) then
      Result := true;
  end;
end;

function TsCustomWorksheetGrid.GetWordwraps(ARect: TGridRect): Boolean;
var
  c, r: Integer;
  wrapped: Boolean;
begin
  Result := GetWordwrap(ARect.Left, ARect.Top);
  wrapped := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetWordwrap(c, r);
      if Result <> wrapped then begin
        Result := false;
        exit;
      end;
    end;
end;

procedure TsCustomWorksheetGrid.SetAutoCalc(AValue: Boolean);
begin
  FAutoCalc := AValue;
  if Assigned(FWorkbook) then
  begin
    if FAutoCalc then
      FWorkbook.Options := FWorkbook.Options + [boAutoCalc]
    else
      FWorkbook.Options := FWorkbook.Options - [boAutoCalc];
  end;
end;

procedure TsCustomWorksheetGrid.SetBackgroundColor(ACol, ARow: Integer;
  AValue: TsColor);
var
  c, r: Cardinal;
begin
  if Assigned(FWorksheet) then begin
    BeginUpdate;
    try
      c := GetWorksheetCol(ACol);
      r := GetWorksheetRow(ARow);
      FWorksheet.WriteBackgroundColor(r, c, AValue);
    finally
      EndUpdate;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetBackgroundColors(ARect: TGridRect;
  AValue: TsColor);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetBackgroundColor(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBorder(ACol, ARow: Integer;
  AValue: TsCellBorders);
var
  c, r: Cardinal;
begin
  if Assigned(FWorksheet) then begin
    BeginUpdate;
    try
      c := GetWorksheetCol(ACol);
      r := GetWorksheetRow(ARow);
      FWorksheet.WriteBorders(r, c, AValue);
      FixNeighborCellBorders(ACol, ARow);
    finally
      EndUpdate;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBorders(ARect: TGridRect;
  AValue: TsCellBorders);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetCellBorder(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBorderStyle(ACol, ARow: Integer;
  ABorder: TsCellBorder; AValue: TsCellBorderStyle);
begin
  if Assigned(FWorksheet) then begin
    BeginUpdate;
    try
      FWorksheet.WriteBorderStyle(GetWorksheetRow(ARow), GetWorksheetCol(ACol), ABorder, AValue);
      FixNeighborCellBorders(ACol, ARow);
    finally
      EndUpdate;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBorderStyles(ARect: TGridRect;
  ABorder: TsCellBorder; AValue: TsCellBorderStyle);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetCellBorderStyle(c, r, ABorder, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFont(ACol, ARow: Integer; AValue: TFont);
var
  fnt: TsFont;
begin
  FCellFont.Assign(AValue);
  if Assigned(FWorksheet) then begin
    fnt := TsFont.Create;
    try
      Convert_Font_To_sFont(FCellFont, fnt);
      FWorksheet.WriteFont(GetWorksheetRow(ARow), GetWorksheetCol(ACol),
        fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
    finally
      fnt.Free;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFonts(ARect: TGridRect;
  AValue: TFont);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetCellFont(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontColor(ACol, ARow: Integer; AValue: TsColor);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteFontColor(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetCellFontColors(ARect: TGridRect; AValue: TsColor);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetCellFontColor(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontName(ACol, ARow: Integer; AValue: String);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteFontName(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetCellFontNames(ARect: TGridRect; AValue: String);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetCellFontName(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontSize(ACol, ARow: Integer;
  AValue: Single);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteFontSize(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetCellFontSizes(ARect: TGridRect;
  AValue: Single);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetCellFontSize(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontStyle(ACol, ARow: Integer;
  AValue: TsFontStyles);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteFontStyle(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetCellFontStyles(ARect: TGridRect;
  AValue: TsFontStyles);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetCellFontStyle(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetFrozenCols(AValue: Integer);
begin
  FFrozenCols := AValue;
  if FWorksheet <> nil then begin
    FWorksheet.LeftPaneWidth := FFrozenCols;
    if (FFrozenCols > 0) or (FFrozenRows > 0) then
      FWorksheet.Options := FWorksheet.Options + [soHasFrozenPanes]
    else
      FWorksheet.Options := FWorksheet.Options - [soHasFrozenPanes];
  end;
  Setup;
end;

procedure TsCustomWorksheetGrid.SetFrozenRows(AValue: Integer);
begin
  FFrozenRows := AValue;
  if FWorksheet <> nil then begin
    FWorksheet.TopPaneHeight := FFrozenRows;
    if (FFrozenCols > 0) or (FFrozenRows > 0) then
      FWorksheet.Options := FWorksheet.Options + [soHasFrozenPanes]
    else
      FWorksheet.Options := FWorksheet.Options - [soHasFrozenPanes];
  end;
  Setup;
end;

procedure TsCustomWorksheetGrid.SetHorAlignment(ACol, ARow: Integer;
  AValue: TsHorAlignment);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteHorAlignment(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetHorAlignments(ARect: TGridRect;
  AValue: TsHorAlignment);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetHorAlignment(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

{ Shows / hides the worksheet's grid lines }
procedure TsCustomWorksheetGrid.SetShowGridLines(AValue: Boolean);
begin
  if AValue = GetShowGridLines then
    Exit;

  if AValue then
    Options := Options + [goHorzLine, goVertLine]
  else
    Options := Options - [goHorzLine, goVertLine];

  if FWorksheet <> nil then
    if AValue then
      FWorksheet.Options := FWorksheet.Options + [soShowGridLines]
    else
      FWorksheet.Options := FWorksheet.Options - [soShowGridLines];
end;

{ Shows / hides the worksheet's row and column headers. }
procedure TsCustomWorksheetGrid.SetShowHeaders(AValue: Boolean);
begin
  if AValue = GetShowHeaders then Exit;

  FHeaderCount := ord(AValue);
  if FWorksheet <> nil then
    if AValue then
      FWorksheet.Options := FWorksheet.Options + [soShowHeaders]
    else
      FWorksheet.Options := FWorksheet.Options - [soShowHeaders];

  Setup;
end;

procedure TsCustomWorksheetGrid.SetTextRotation(ACol, ARow: Integer;
  AValue: TsTextRotation);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteTextRotation(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetTextRotations(ARect: TGridRect;
  AValue: TsTextRotation);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetTextRotation(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetVertAlignment(ACol, ARow: Integer;
  AValue: TsVertAlignment);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteVertAlignment(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetVertAlignments(ARect: TGridRect;
  AValue: TsVertAlignment);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetVertAlignment(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetWordwrap(ACol, ARow: Integer;
  AValue: Boolean);
begin
  if Assigned(FWorksheet) then
    FWorksheet.WriteWordwrap(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetWordwraps(ARect: TGridRect;
  AValue: Boolean);
var
  c,r: Integer;
begin
  BeginUpdate;
  try
    for c := ARect.Left to ARect.Right do
      for r := ARect.Top to ARect.Bottom do
        SetWordwrap(c, r, AValue);
  finally
    EndUpdate;
  end;
end;


initialization
  fpsutils.ScreenPixelsPerInch := Screen.PixelsPerInch;

finalization
  FreeAndNil(FillPattern_BIFF2);

end.
