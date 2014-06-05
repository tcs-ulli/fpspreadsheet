{
fpspreadsheetgrid.pas

Grid component which can load and write data from / to FPSpreadsheet documents

AUTHORS: Felipe Monteiro de Carvalho, Werner Pamler
}

{ To do:
 - When Lazarus 1.4 comes out remove the workaround for the RGB2HLS bug in
   FindNearestPaletteIndex.
}

unit fpspreadsheetgrid;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs, Grids,
  fpspreadsheet;

type

  { TsCustomWorksheetGrid }

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
    FReadFormulas: Boolean;
    function CalcAutoRowHeight(ARow: Integer): Integer;
    function CalcColWidth(AWidth: Single): Integer;
    function CalcRowHeight(AHeight: Single): Integer;
    procedure ChangedCellHandler(ASender: TObject; ARow, ACol: Cardinal);
    procedure ChangedFontHandler(ASender: TObject; ARow, ACol: Cardinal);
    procedure FixNeighborCellBorders(ACol, ARow: Integer);

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
    procedure DefaultDrawCell(ACol, ARow: Integer; var ARect: TRect; AState: TGridDrawState); override;
    procedure DoPrepareCanvas(ACol, ARow: Integer; AState: TGridDrawState); override;
    procedure DrawAllRows; override;
    procedure DrawCellBorders; overload;
    procedure DrawCellBorders(ACol, ARow: Integer; ARect: TRect); overload;
    procedure DrawFocusRect(aCol,aRow:Integer; ARect:TRect); override;
    procedure DrawSelection;
    procedure DrawTextInCell(ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState); override;
    function GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer;
      var ABorderStyle: TsCellBorderStyle): Boolean;
    function GetCellHeight(ACol, ARow: Integer): Integer;
    function GetCellText(ACol, ARow: Integer): String;
    function GetEditText(ACol, ARow: Integer): String; override;
    function HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
    procedure HeaderSized(IsColumn: Boolean; index: Integer); override;
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
    property DisplayFixedColRow: Boolean read GetShowHeaders write SetShowHeaders default true;
    property FrozenCols: Integer read FFrozenCols write SetFrozenCols;
    property FrozenRows: Integer read FFrozenRows write SetFrozenRows;
    property ReadFormulas: Boolean read FReadFormulas write FReadFormulas;
    property ShowGridLines: Boolean read GetShowGridLines write SetShowGridLines default true;
    property ShowHeaders: Boolean read GetShowHeaders write SetShowHeaders default true;

  public
    { public methods }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure BeginUpdate;
    procedure EditingDone; override;
    procedure EndUpdate;
    procedure GetSheets(const ASheets: TStrings);
    function GetGridCol(ASheetCol: Cardinal): Integer;
    function GetGridRow(ASheetRow: Cardinal): Integer;
    function GetWorksheetCol(AGridCol: Integer): Cardinal;
    function GetWorksheetRow(AGridRow: Integer): Cardinal;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AWorksheetIndex: Integer = 0); overload;
    procedure NewWorksheet(AColCount, ARowCount: Integer);
    procedure SaveToSpreadsheetFile(AFileName: string;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat;
      AOverwriteExisting: Boolean = true); overload;
    procedure SelectSheetByIndex(AIndex: Integer);

    { Utilities related to Workbooks }
    procedure Convert_sFont_to_Font(sFont: TsFont; AFont: TFont);
    procedure Convert_Font_to_sFont(AFont: TFont; sFont: TsFont);
    function FindNearestPaletteIndex(AColor: TColor): TsColor;

    { public properties }
    property Worksheet: TsWorksheet read FWorksheet;
    property Workbook: TsWorkbook read FWorkbook;
    property HeaderCount: Integer read FHeaderCount;

    { maybe these should become published ... }
    property BackgroundColor[ACol, ARow: Integer]: TsColor
        read GetBackgroundColor write SetBackgroundColor;
    property BackgroundColors[ARect: TGridRect]: TsColor
        read GetBackgroundColors write SetBackgroundColors;
    property CellBorder[ACol, ARow: Integer]: TsCellBorders
        read GetCellBorder write SetCellBorder;
    property CellBorders[ARect: TGridRect]: TsCellBorders
        read GetCellBorders write SetCellBorders;
    property CellBorderStyle[ACol, ARow: Integer; ABorder: TsCellBorder]: TsCellBorderStyle
        read GetCellBorderStyle write SetCellBorderStyle;
    property CellBorderStyles[ARect: TGridRect; ABorder: TsCellBorder]: TsCellBorderStyle
        read GetCellBorderStyles write SetCellBorderStyles;
    property CellFont[ACol, ARow: Integer]: TFont
        read GetCellFont write SetCellFont;
    property CellFonts[ARect: TGridRect]: TFont
        read GetCellFonts write SetCellFonts;
    property CellFontName[ACol, ARow: Integer]: String
        read GetCellFontName write SetCellFontName;
    property CellFontNames[ARect: TGridRect]: String
        read GetCellFontNames write SetCellFontNames;
    property CellFontStyle[ACol, ARow: Integer]: TsFontStyles
        read GetCellFontStyle write SetCellFontStyle;
    property CellFontStyles[ARect: TGridRect]: TsFontStyles
        read GetCellFontStyles write SetCellFontStyles;
    property CellFontSize[ACol, ARow: Integer]: Single
        read GetCellFontSize write SetCellFontSize;
    property CellFontSizes[ARect: TGridRect]: Single
        read GetCellFontSizes write SetCellFontSizes;
    property HorAlignment[ACol, ARow: Integer]: TsHorAlignment
        read GetHorAlignment write SetHorAlignment;
    property HorAlignments[ARect: TGridRect]: TsHorAlignment
        read GetHorAlignments write SetHorAlignments;
    property TextRotation[ACol, ARow: Integer]: TsTextRotation
        read GetTextRotation write SetTextRotation;
    property TextRotations[ARect: TGridRect]: TsTextRotation
        read GetTextRotations write SetTextRotations;
    property VertAlignment[ACol, ARow: Integer]: TsVertAlignment
        read GetVertAlignment write SetVertAlignment;
    property VertAlignments[ARect: TGridRect]: TsVertAlignment
        read GetVertAlignments write SetVertAlignments;
    property Wordwrap[ACol, ARow: Integer]: Boolean
        read GetWordwrap write SetWordwrap;
    property Wordwraps[ARect: TGridRect]: Boolean
        read GetWordwraps write SetWordwraps;
  end;

  { TsWorksheetGrid }

  TsWorksheetGrid = class(TsCustomWorksheetGrid)
  published
    // inherited from TsCustomWorksheetGrid
    property DisplayFixedColRow; deprecated 'Use ShowHeaders';
    property FrozenCols;
    property FrozenRows;
    property ReadFormulas;
    property ShowGridLines;
    property ShowHeaders;

    // inherited from other ancestors
    property Align;
    property AlternateColor;
    property Anchors;
    property AutoAdvance;
    property AutoEdit;
    property AutoFillColumns;
    //property BiDiMode;
    property BorderSpacing;
    property BorderStyle;
    property Color;
    property ColCount;
    //property Columns;
    property Constraints;
    property DefaultColWidth;
    property DefaultDrawing;
    property DefaultRowHeight;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property ExtendedSelect;
    property FixedColor;
    property Flat;
    property Font;
    property GridLineWidth;
    property HeaderHotZones;
    property HeaderPushZones;
    property MouseWheelOption;
    property Options;
    //property ParentBiDiMode;
    property ParentColor default false;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property RowCount;
    property ScrollBars;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property TitleFont;
    property TitleImageList;
    property TitleStyle;
    property UseXORFeatures;
    property Visible;
    property VisibleColCount;
    property VisibleRowCount;

    property OnBeforeSelection;
    property OnChangeBounds;
    property OnClick;
    property OnColRowDeleted;
    property OnColRowExchanged;
    property OnColRowInserted;
    property OnColRowMoved;
    property OnCompareCells;
    property OnDragDrop;
    property OnDragOver;
    property OnDblClick;
    property OnDrawCell;
    property OnEditButtonClick;
    property OnEditingDone;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnGetEditMask;
    property OnGetEditText;
    property OnHeaderClick;
    property OnHeaderSized;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseWheel;
    property OnMouseWheelDown;
    property OnMouseWheelUp;
    property OnPickListSelect;
    property OnPrepareCanvas;
    property OnResize;
    property OnSelectEditor;
    property OnSelection;
    property OnSelectCell;
    property OnSetEditText;
    property OnShowHint;
    property OnStartDock;
    property OnStartDrag;
    property OnTopLeftChanged;
    property OnUTF8KeyPress;
    property OnValidateEntry;
    property OnContextPopup;
  end;

procedure Register;

implementation

uses
  Types, LCLType, LCLIntf, Math, fpCanvas, GraphUtil, fpsUtils;

const
  HOR_ALIGNMENTS: array[haLeft..haRight] of TAlignment = (
    taLeftJustify, taCenter, taRightJustify
  );
  VERT_ALIGNMENTS: array[TsVertAlignment] of TTextLayout = (
    tlBottom, tlTop, tlCenter, tlBottom
  );

var
  FillPattern_BIFF2: TBitmap = nil;

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

function WrapText(ACanvas: TCanvas; const AText: string; AMaxWidth: integer): string;
// code posted by taazz in the Lazarus Forum:
// http://forum.lazarus.freepascal.org/index.php/topic,21305.msg124743.html#msg124743
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

procedure Register;
begin
  RegisterComponents('Additional',[TsWorksheetGrid]);
end;


{ TsCustomWorksheetGrid }

constructor TsCustomWorksheetGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FHeaderCount := 1;
  FInitColCount := 26;
  FInitRowCount := 100;
  FCellFont := TFont.Create;
end;

destructor TsCustomWorksheetGrid.Destroy;
begin
  FreeAndNil(FWorkbook);
  FreeAndNil(FCellFont);
  inherited Destroy;
end;

{ Suppresses unnecessary repaints. }
procedure TsCustomWorksheetGrid.BeginUpdate;
begin
  inc(FLockCount);
end;

{ Converts the column width, given in "characters" of the default font, to pixels
  All chars are assumed to have the same width defined by the "0".
  Therefore, this calculation is only approximate. }
function TsCustomWorksheetGrid.CalcColWidth(AWidth: Single): Integer;
var
  w0: Integer;
begin
  Convert_sFont_to_Font(FWorkbook.GetFont(0), Canvas.Font);
  w0 := Canvas.TextWidth('0');
  Result := Round(AWidth * w0);
end;

{ Finds the max cell height per row and uses this to define the RowHeights[].
  Returns DefaultRowHeight if the row does not contain any cells.
  ARow is a grid row index. }
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

{ Converts the row height (from a worksheet row), given in lines, to pixels }
function TsCustomWorksheetGrid.CalcRowHeight(AHeight: Single): Integer;
var
  h_pts: Single;
begin
  h_pts := AHeight * (Workbook.GetFont(0).Size + ROW_HEIGHT_CORRECTION);
  Result := PtsToPX(h_pts, Screen.PixelsPerInch) + 4;
end;

procedure TsCustomWorksheetGrid.ChangedCellHandler(ASender: TObject; ARow, ACol:Cardinal);
begin
  if FLockCount = 0 then Invalidate;
end;

{ Handler for the event that the font has changed in a given cell.
  As a consequence, the row height may have to be adapted.
  Row/Col coordinates are in worksheet units here! }
procedure TsCustomWorksheetGrid.ChangedFontHandler(ASender: TObject; ARow, ACol: Cardinal);
var
  h: Integer;
  lRow: PRow;
begin
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

{ Converts a spreadsheet font to a font used for painting (TCanvas.Font). }
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

{ Converts a font used for painting (TCanvas.Font) to a spreadsheet font }
procedure TsCustomWorksheetGrid.Convert_Font_to_sFont(AFont: TFont; sFont: TsFont);
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

{ Is overridden to show "frozen" cells in the same style as normal cells.
  "Frozen" cells are internally "fixed" cells of the grid. }
procedure TsCustomWorksheetGrid.DefaultDrawCell(aCol, aRow: Integer; var aRect: TRect;
  AState: TGridDrawState);
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
    wasFixed := true;                   // ?????
    AState := AState - [gdFixed];
    Canvas.Brush.Color := clWindow;
  end;

  inherited DefaultDrawCell(ACol, ARow, ARect, AState);

  if wasFixed then begin
    DrawCellGrid(ACol, ARow, ARect, AState);
    AState := AState + [gdFixed];
  end;
end;

{ Adjusts the grid's canvas before painting a given cell. Considers, e.g.
  background color, horizontal alignment, vertical alignment, etc. }
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
  if ShowHeaders then begin
    // Formatting of row and column headers
    if ARow = 0 then begin
      ts.Alignment := taCenter;
      ts.Layout := tlCenter;
    end else
    if ACol = 0 then begin
      ts.Alignment := taRightJustify;
      ts.Layout := tlCenter;
    end;
    if ShowHeaders and ((ACol = 0) or (ARow = 0)) then
      Canvas.Brush.Color := FixedColor
  end;
  if FWorksheet <> nil then begin
    r := ARow - FHeaderCount;
    c := ACol - FHeaderCount;
    lCell := FWorksheet.FindCell(r, c);
    if lCell <> nil then begin
      // Background color
      if (uffBackgroundColor in lCell^.UsedFormattingFields) then begin
        if FWorkbook.FileFormat = sfExcel2 then begin
          if (FillPattern_BIFF2 = nil) and (ComponentState = []) then
            Create_FillPattern_BIFF2(Color);
          Canvas.Brush.Style := bsImage;
          Canvas.Brush.Bitmap := FillPattern_BIFF2;
        end else begin
          Canvas.Brush.Style := bsSolid;
          if lCell^.BackgroundColor < FWorkbook.GetPaletteSize then
            Canvas.Brush.Color := FWorkbook.GetPaletteColor(lCell^.BackgroundColor)
          else
            Canvas.Brush.Color := Color;
        end;
      end else begin
        Canvas.Brush.Style := bsSolid;
        Canvas.Brush.Color := Color;
      end;
      // Font
      if (uffFont in lCell^.UsedFormattingFields) then begin
        fnt := FWorkbook.GetFont(lCell^.FontIndex);
        if fnt <> nil then begin
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
      if (lCell^.NumberFormat in [nfCurrencyRed, nfAccountingRed]) and
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

{ Is overridden in order to paint the cell borders and the selection rectangle.
  Both features can extend into the neighbor cells and thus are clipped at the
  cell borders by the standard painting mechanism. In DrawAllRows, clipping at
  cell borders is no longer active. }
procedure TsCustomWorksheetGrid.DrawAllRows;
var
  cliprect: TRect;
  rgn: HRGN;
  tmp: Integer;
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
    rgn := CreateRectRgn(cliprect.Left, cliprect.top, cliprect.Right, cliprect.Bottom);
    SelectClipRgn(Canvas.Handle, Rgn);

    DrawCellBorders;
    DrawSelection;

    DeleteObject(rgn);
  finally
    Canvas.RestoreHandleState;
  end;
end;

{ Draws the borders of all cells. }
procedure TsCustomWorksheetGrid.DrawCellBorders;
var
  cell: PCell;
  c, r: Integer;
  rect: TRect;
begin
  if FWorksheet = nil then exit;

  cell := FWorksheet.GetFirstCell;
  while cell <> nil do begin
    if (uffBorder in cell^.UsedFormattingFields) then begin
      c := cell^.Col + FHeaderCount;
      r := cell^.Row + FHeaderCount;
      rect := CellRect(c, r);
      DrawCellBorders(c, r, rect);
    end;
    cell := FWorksheet.GetNextCell;
  end;
end;

{ Draws the border lines around a given cell. Note that when this procedure is
  called the output is clipped by the cell rectangle, but thick and double
  border styles extend into the neighbor cell. Therefore, these border lines
  are drawn in parts. }
procedure TsCustomWorksheetGrid.DrawCellBorders(ACol, ARow: Integer; ARect: TRect);

  procedure DrawBorderLine(ACoord: Integer; ARect: TRect; IsHor: Boolean;
    ABorderStyle: TsCellBorderStyle);
  const
    // TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair);
    PEN_STYLES: array[TsLineStyle] of TPenStyle =
      (psSolid, psSolid, psDash, psDot, psSolid, psSolid, psSolid);
    PEN_WIDTHS: array[TsLineStyle] of Integer =
      (1, 2, 1, 1, 3, 1, 1);
  var
    width3: Boolean;     // line is 3 pixels wide
  begin
    Canvas.Pen.Style := PEN_STYLES[ABorderStyle.LineStyle];
    Canvas.Pen.Width := PEN_WIDTHS[ABorderStyle.LineStyle];
    Canvas.Pen.Color := FWorkbook.GetPaletteColor(ABorderStyle.Color);
    Canvas.Pen.EndCap := pecSquare;
    width3 := (ABorderStyle.LineStyle in [lsThick, lsDouble]);

    // Tuning the rectangle to avoid issues at the grid borders and to get nice corners
    if (ABorderStyle.LineStyle in [lsMedium, lsThick, lsDouble]) then begin
      if ACol = ColCount-1 then begin
        if not IsHor and (ACoord = ARect.Right-1) and width3 then dec(ACoord);
        dec(ARect.Right);
      end;
      if ARow = RowCount-1 then begin
        if IsHor and (ACoord = ARect.Bottom-1) and width3 then dec(ACoord);
        dec(ARect.Bottom);
      end;
    end;
    if ABorderStyle.LineStyle in [lsMedium, lsThick] then begin
      if IsHor then dec(ARect.Right, 1) else dec(ARect.Bottom, 1);
    end;

    // Painting
    case ABorderStyle.LineStyle of
      lsThin, lsMedium, lsThick, lsDotted, lsDashed:
        if IsHor then
          Canvas.Line(ARect.Left, ACoord, ARect.Right, ACoord)
        else
          Canvas.Line(ACoord, ARect.Top, ACoord, ARect.Bottom);

      lsHair:
        if IsHor then
          DrawHairLineHor(Canvas, ARect.Left, ARect.Right, ACoord)
        else
          DrawHairLineVert(Canvas, ACoord, ARect.Top, ARect.Bottom);

      lsDouble:
        if IsHor then begin
          Canvas.Line(ARect.Left, ACoord-1, ARect.Right, ACoord-1);
          Canvas.Line(ARect.Left, ACoord+1, ARect.Right, ACoord+1);
          Canvas.Pen.Color := Color;
          Canvas.Line(ARect.Left, ACoord, ARect.Right, ACoord);
        end else begin
          Canvas.Line(ACoord-1, ARect.Top, ACoord-1, ARect.Bottom);
          Canvas.Line(ACoord+1, ARect.Top, ACoord+1, ARect.Bottom);
          Canvas.Pen.Color := Color;
          Canvas.Line(ACoord, ARect.Top, ACoord, ARect.Bottom);
        end;
    end;
  end;

var
  bs: TsCellBorderStyle;
begin
  if Assigned(FWorksheet) then begin
    // Left border
    if GetBorderStyle(ACol, ARow, -1, 0, bs) then
      DrawBorderLine(ARect.Left-1, ARect, false, bs);
    // Right border
    if GetBorderStyle(ACol, ARow, +1, 0, bs) then
      DrawBorderLine(ARect.Right-1, ARect, false, bs);
    // Top border
    if GetBorderstyle(ACol, ARow, 0, -1, bs) then
      DrawBorderLine(ARect.Top-1, ARect, true, bs);
    // Bottom border
    if GetBorderStyle(ACol, ARow, 0, +1, bs) then
      DrawBorderLine(ARect.Bottom-1, ARect, true, bs);
  end;
end;

{ Is responsible for painting of the focus rectangle. We don't want the red
  dashed rectangle here, but the thick Excel-like rectangle. }
procedure TsCustomWorksheetGrid.DrawFocusRect(aCol, aRow: Integer; ARect: TRect);
begin
  // Nothing do to
end;

{ Draws the selection rectangle, 3 pixels wide as in Excel. }
procedure TsCustomWorksheetGrid.DrawSelection;
var
  P1, P2: TPoint;
  selrect: TRect;
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

{ Draws the cell text. Calls "GetCellText" to determine the text in the cell.
  Takes care of horizontal and vertical text alignment, text rotation and
  text wrapping }
procedure TsCustomWorksheetGrid.DrawTextInCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
var
  ts: TTextStyle;
  flags: Cardinal;
  txt: String;
  txtL, txtR: String;
  txtRect: TRect;
  P: TPoint;
  w, h, h0, hline: Integer;
  i: Integer;
  L: TStrings;
  c, r: Integer;
  wrapped: Boolean;
  horAlign: TsHorAlignment;
  vertAlign: TsVertAlignment;
  txtRot: TsTextRotation;
  lCell: PCell;
  txtLeft, txtRight: String;
  justif: Byte;
begin
  if FWorksheet = nil then
    exit;

  c := ACol - FHeaderCount;
  r := ARow - FHeaderCount;
  lCell := FWorksheet.FindCell(r, c);

  // Header
  if lCell = nil then begin
    if ShowHeaders and ((ACol = 0) or (ARow = 0)) then begin
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
  else begin
    if (lCell^.ContentType in [cctNumber, cctDateTime]) then
      horAlign := haRight
    else
      horAlign := haLeft;
    {
    if txtRot = rt90DegreeCounterClockwiseRotation then begin
      if horAlign = haRight then horAlign := haLeft else horAlign := haRight;
    end;
    }
  end;

  InflateRect(ARect, -constCellPadding, -constCellPadding);

  if (lCell^.NumberFormat in [nfAccounting, nfAccountingRed]) and not IsNaN(lCell^.Numbervalue)
  then begin
    case SplitAccountingFormatString(lCell^.NumberFormatStr, Sign(lCell^.NumberValue),
                                     txtLeft, txtRight) of
      1: begin
           txtLeft := FormatFloat(txtLeft, lCell^.NumberValue);
           if txtLeft = '' then exit;
           txt := txtLeft + ' ' + txtRight;
         end;
      2: begin
           txtRight := FormatFloat(txtRight, lCell^.NumberValue);
           if txtRight = '' then exit;
           txt := txtLeft + ' ' + txtRight;
         end;
    end;
    InternalDrawTextInCell(txtLeft, txt, ARect, 0, horAlign, vertAlign,
      txtRot, wrapped, true);
    InternalDrawTextInCell(txtRight, txt, ARect, 2, horAlign, vertAlign,
      txtRot, wrapped, true);
  end else begin
    txt := GetCellText(ACol, ARow);
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
end;

(*





  procedure InternalDrawTextInCell(AText, AMeasureText: String; ARect: TRect;
    AJustification: Byte; ACellHorAlign: TsHorAlignment;
    ACellVertAlign: TsVertAlignment; ATextRot: TsTextRotation;
    ATextWrap, ReplaceTooLong: Boolean);




  if (lCell^.TextRotation in [trHorizontal, rtStacked]) or
     (not (uffTextRotation in lCell^.UsedFormattingFields))
  then begin
    // HORIZONAL TEXT DRAWING DIRECTION
    ts := Canvas.TextStyle;
    if wrapped then begin
      ts.Wordbreak := true;
      ts.SingleLine := false;
      flags := DT_WORDBREAK and not DT_SINGLELINE;
      LCLIntf.DrawText(Canvas.Handle, PChar(txt), Length(txt), txtRect,
        DT_CALCRECT or flags);
      w := txtRect.Right - txtRect.Left;
      h := txtRect.Bottom - txtRect.Top;
    end else begin
      ts.WordBreak := false;
      ts.SingleLine := false;
      w := Canvas.TextWidth(txt);
      h := Canvas.TextHeight('Tg');
    end;

    Canvas.Font.Orientation := 0;
    ts.Alignment := HOR_ALIGNMENTS[horAlign];
    ts.Opaque := false;
    if h > ARect.Bottom - ARect.Top then
      ts.Layout := tlTop
    else
      ts.Layout := VERT_ALIGNMENTS[vertAlign];

    Canvas.TextStyle := ts;
    Canvas.TextRect(ARect, ARect.Left, ARect.Top, txt);
  end
  else
  begin
    // ROTATED TEXT DRAWING DIRECTION
    L := TStringList.Create;
    try
      txtRect := Bounds(ARect.Left, ARect.Top, ARect.Bottom - ARect.Top, ARect.Right - ARect.Left);
      hline := Canvas.TextHeight('Tg');
      if wrapped then begin
        L.Text := WrapText(Canvas, txt, txtRect.Right - txtRect.Left);
        flags := DT_WORDBREAK and not DT_SINGLELINE;
        LCLIntf.DrawText(Canvas.Handle, PChar(L.Text), Length(L.Text), txtRect,
          DT_CALCRECT or flags);
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

      ts := Canvas.TextStyle;
      ts.SingleLine := true;      // Draw text line by line
      ts.Clipping := false;
      ts.Layout := tlTop;
      ts.Alignment := taLeftJustify;
      ts.Opaque := false;

      if lCell^.TextRotation = rt90DegreeClockwiseRotation then begin
        // Clockwise
        Canvas.Font.Orientation := -900;
        case horAlign of
          haLeft   : P.X := Min(ARect.Right-1, ARect.Left + h - h0);
          haCenter : P.X := Min(ARect.Right-1, (ARect.Left + ARect.Right + h) div 2);
          haRight  : P.X := ARect.Right - 1;
        end;
        for i:= 0 to L.Count-1 do begin
          w := Canvas.TextWidth(L[i]);
          case vertAlign of
            vaTop    : P.Y := ARect.Top;
            vaCenter : P.Y := Max(ARect.Top, (ARect.Top + ARect.Bottom - w) div 2);
            vaBottom : P.Y := Max(ARect.Top, ARect.Bottom - w);
          end;
          Canvas.TextRect(ARect, P.X, P.Y, L[i], ts);
          dec(P.X, hline);
        end
      end
      else begin
        // Counter-clockwise
        Canvas.Font.Orientation := +900;
        case horAlign of
          haLeft   : P.X := ARect.Left;
          haCenter : P.X := Max(ARect.Left, (ARect.Left + ARect.Right - h + h0) div 2);
          haRight  : P.X := MAx(ARect.Left, ARect.Right - h + h0);
        end;
        for i:= 0 to L.Count-1 do begin
          w := Canvas.TextWidth(L[i]);
          case vertAlign of
            vaTop    : P.Y := Min(ARect.Bottom, ARect.Top + w);
            vaCenter : P.Y := Min(ARect.Bottom, (ARect.Top + ARect.Bottom + w) div 2);
            vaBottom : P.Y := ARect.Bottom;
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
*)

procedure TsCustomWorksheetGrid.EditingDone;
var
  oldText: String;
  cell: PCell;
begin
  if (not EditorShowing) and FEditing then begin
    oldText := GetCellText(Col, Row);
    if oldText <> FEditText then begin
      if FWorksheet = nil then
        FWorksheet := TsWorksheet.Create;
      cell := FWorksheet.GetCell(Row-FHeaderCount, Col-FHeaderCount);
      if FEditText = '' then
        cell^.ContentType := cctEmpty
      else
      if TryStrToFloat(FEditText, cell^.NumberValue) then
        cell^.ContentType := cctNumber
      else
      if TryStrToDateTime(FEditText, cell^.DateTimeValue) then begin
        cell^.ContentType := cctDateTime;
        if cell^.DateTimeValue < 1.0 then begin      // this is a TTime
          if not (cell^.NumberFormat in [nfShortDateTime, nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM])
            then cell^.NumberFormat := nfLongTime;
        end else
        if frac(cell^.DateTimeValue) = 0 then begin  // this is a TDate
          if not (cell^.NumberFormat in [nfShortDateTime, nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM])
            then cell^.NumberFormat := nfShortDate
        end else
          cell^.NumberFormat := nfShortDateTime;
      end else begin
        cell^.UTF8StringValue := FEditText;
        cell^.ContentType := cctUTF8String;
      end;
      FEditText := '';
    end;
    inherited EditingDone;
  end;
  FEditing := false;
end;

procedure TsCustomWorksheetGrid.EndUpdate;
begin
  dec(FLockCount);
  if FLockCount = 0 then Invalidate;
end;

{ Copies the borders of a cell to its neighbors. This avoids the nightmare of
  changing borders due to border conflicts of adjacent cells. }
procedure TsCustomWorksheetGrid.FixNeighborCellBorders(ACol, ARow: Integer);

  procedure SetNeighborBorder(NewRow, NewCol: Integer;
    ANewBorder: TsCellBorder; const ANewBorderStyle: TsCellBorderStyle;
    AInclude: Boolean);
  var
    neighbor: PCell;
    border: TsCellBorders;
  begin
    neighbor := FWorksheet.FindCell(NewRow, NewCol);
    if neighbor <> nil then begin
      border := neighbor^.Border;
      if AInclude then begin
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
  if FWorksheet = nil then exit;
  cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
  if (FWorksheet <> nil) and (cell <> nil) then
    with cell^ do begin
      SetNeighborBorder(Row, Col-1, cbEast, BorderStyles[cbWest], cbWest in Border);
      SetNeighborBorder(Row, Col+1, cbWest, BorderStyles[cbEast], cbEast in Border);
      SetNeighborBorder(Row-1, Col, cbSouth, BorderStyles[cbNorth], cbNorth in Border);
      SetNeighborBorder(Row+1, Col, cbNorth, BorderStyles[cbSouth], cbSouth in Border);
    end;
end;

{ The "colors" used by the spreadsheet are indexes into the workbook's color
  palette. If the user wants to set a color to a particular rgb value this is
  not possible in general. The method FindNearestPaletteIndex finds the bast
  matching color in the palette. }
function TsCustomWorksheetGrid.FindNearestPaletteIndex(AColor: TColor): TsColor;

  procedure ColorToHSL(RGB: TColor; var H, S, L : double);
  // Taken from https://code.google.com/p/thtmlviewer/source/browse/trunk/source/HSLUtils.pas?r=277
  // The procedure in GraphUtils is crashing for some colors in Laz < 1.3
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
    end else begin
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
  type
    TRGBA = packed record R,G,B,A: Byte end;
  var
    H1,S1,L1, H2,S2,L2: Double;
  begin
    ColorToHSL(color1, H1, S1, L1);
    ColorToHSL(color2, H2, S2, L2);
    Result := sqr(H1-H2) + sqr(S1-S2) + sqr(L1-L2);
  end;

  {
  // will be activated when Lazarus 1.4 is available. (RgbToHLS bug in Laz < 1.3)

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
  if Workbook <> nil then begin
    mindist := 1E308;
    for i:=0 to Workbook.GetPaletteSize-1 do begin
      dist := ColorDistance(AColor, TColor(Workbook.GetPaletteColor(i)));
      if dist < mindist then begin
        mindist := dist;
        Result := i;
      end;
    end;
  end;
end;

function TsCustomWorksheetGrid.GetBackgroundColor(ACol, ARow: Integer): TsColor;
var
  cell: PCell;
begin
  Result := scNotDefined;
  if Assigned(FWorksheet) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) and (uffBackgroundColor in cell^.UsedFormattingFields) then
      Result := cell^.BackgroundColor;
  end;
end;

function TsCustomWorksheetGrid.GetBackgroundColors(ARect: TGridRect): TsColor;
var
  c, r: Integer;
  clr: TsColor;
begin
  Result := GetBackgroundColor(ARect.Left, ARect.Top);
  clr := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetBackgroundColor(c, r);
      if Result <> clr then begin
        Result := scNotDefined;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellBorder(ACol, ARow: Integer): TsCellBorders;
var
  cell: PCell;
begin
  Result := [];
  if Assigned(FWorksheet) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) and (uffBorder in cell^.UsedFormattingFields) then
      Result := cell^.Border;
  end;
end;

function TsCustomWorksheetGrid.GetCellBorders(ARect: TGridRect): TsCellBorders;
var
  c, r: Integer;
  b: TsCellBorders;
begin
  Result := GetCellBorder(ARect.Left, ARect.Top);
  b := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetCellBorder(c, r);
      if Result <> b then begin
        Result := [];
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellBorderStyle(ACol, ARow: Integer;
  ABorder: TsCellBorder): TsCellBorderStyle;
var
  cell: PCell;
begin
  Result := DEFAULT_BORDERSTYLES[ABorder];
  if Assigned(FWorksheet) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then
      Result := cell^.BorderStyles[ABorder];
  end;
end;

function TsCustomWorksheetGrid.GetCellBorderStyles(ARect: TGridRect;
  ABorder: TsCellBorder): TsCellBorderStyle;
var
  c, r: Integer;
  bs: TsCellBorderStyle;
begin
  Result := GetCellBorderStyle(ARect.Left, ARect.Top, ABorder);
  bs := Result;
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      Result := GetCellBorderStyle(c, r, ABorder);
      if (Result.LineStyle <> bs.LineStyle) or (Result.Color <> bs.Color) then begin
        Result := DEFAULT_BORDERSTYLES[ABorder];
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellFont(ACol, ARow: Integer): TFont;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := nil;
  if (FWorkbook <> nil) and (FWorksheet <> nil) then begin
    cell := FWorksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if (cell <> nil) then begin
      fnt := FWorkbook.GetFont(cell^.FontIndex);
      Convert_sFont_to_Font(fnt, FCellFont);
      Result := FCellFont;
    end;
  end;
end;

function TsCustomWorksheetGrid.GetCellFonts(ARect: TGridRect): TFont;
var
  c, r: Integer;
  sFont, sDefFont: TsFont;
  cell: PCell;
begin
  Result := GetCellFont(ARect.Left, ARect.Top);
  sDefFont := FWorkbook.GetFont(0);  // Default font
  for c := ARect.Left to ARect.Right do
    for r := ARect.Top to ARect.Bottom do begin
      cell := FWorksheet.FindCell(GetWorksheetRow(r), GetWorksheetCol(c));
      if cell <> nil then begin
        sFont := FWorkbook.GetFont(cell^.FontIndex);
        if (sFont.FontName <> sDefFont.FontName) and (sFont.Size <> sDefFont.Size)
          and (sFont.Style <> sDefFont.Style) and (sFont.Color <> sDefFont.Color)
        then begin
          Convert_sFont_to_Font(sDefFont, FCellFont);
          Result := FCellFont;
          exit;
        end;
      end;
    end;
end;

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

{ Returns the height (in pixels) of the cell at ACol/ARow (of the grid). }
function TsCustomWorksheetGrid.GetCellHeight(ACol, ARow: Integer): Integer;
var
  lCell: PCell;
  s: String;
  wrapped: Boolean;
  txtR: TRect;
  cellR: TRect;
  flags: Cardinal;
begin
  Result := 0;
  if ShowHeaders and ((ACol = 0) or (ARow = 0)) then
    exit;
  if FWorksheet = nil then
    exit;

  lCell := FWorksheet.FindCell(ARow-FHeaderCount, ACol-FHeaderCount);
  if lCell <> nil then begin
    s := GetCellText(ACol, ARow);
    if s = '' then
      exit;
    DoPrepareCanvas(ACol, ARow, []);
    wrapped := (uffWordWrap in lCell^.UsedFormattingFields)
      or (lCell^.TextRotation = rtStacked);
    // *** multi-line text ***
    if wrapped then begin
      // horizontal
      if ( (uffTextRotation in lCell^.UsedFormattingFields) and
           (lCell^.TextRotation in [trHorizontal, rtStacked]))
         or not (uffTextRotation in lCell^.UsedFormattingFields)
      then begin
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

{ GetCellText function returns the text to be written in the cell }
function TsCustomWorksheetGrid.GetCellText(ACol, ARow: Integer): String;
var
  lCell: PCell;
  r, c, i: Integer;
  s: String;
begin
  Result := '';

  if ShowHeaders then begin
    // Headers
    if (ARow = 0) and (ACol = 0) then
      exit;
    if (ARow = 0) then begin
      Result := GetColString(ACol-FHeaderCount);
      exit;
    end
    else
    if (ACol = 0) then begin
      Result := IntToStr(ARow);
      exit;
    end;
  end;

  if FWorksheet <> nil then begin
    r := ARow - FHeaderCount;
    c := ACol - FHeaderCount;
    lCell := FWorksheet.FindCell(r, c);
    if lCell <> nil then begin
      Result := FWorksheet.ReadAsUTF8Text(r, c);
      if lCell^.TextRotation = rtStacked then begin
        s := Result;
        Result := '';
        for i:=1 to Length(s) do begin
          Result := Result + s[i];
          if i < Length(s) then Result := Result + LineEnding;
        end;
      end;
    end;
  end;
end;

{ Determines the text to be passed to the cell editor. }
function TsCustomWorksheetGrid.GetEditText(aCol, aRow: Integer): string;
begin
  Result := GetCellText(aCol, aRow);
  if Assigned(OnGetEditText) then OnGetEditText(Self, aCol, aRow, result);
end;

{ Determines the style of the border between a cell and its neighbor given by
  ADeltaCol and ADeltaRow (one of them must be 0, the other one can only be +/-1).
  ACol and ARow are in grid units. }
function TsCustomWorksheetGrid.GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer;
  var ABorderStyle: TsCellBorderStyle): Boolean;
var
  cell, neighborcell: PCell;
  border, neighborborder: TsCellBorder;
  r, c: Cardinal;
begin
  Result := true;
  if (ADeltaCol = -1) and (ADeltaRow = 0) then begin
    border := cbWest;
    neighborborder := cbEast;
  end else
  if (ADeltaCol = +1) and (ADeltaRow = 0) then begin
    border := cbEast;
    neighborborder := cbWest;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = -1) then begin
    border := cbNorth;
    neighborborder := cbSouth;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = +1) then begin
    border := cbSouth;
    neighborBorder := cbNorth;
  end else
    raise Exception.Create('TsCustomWorksheetGrid: incorrect col/row for GetBorderStyle.');
  r := GetWorksheetRow(ARow);
  c := GetWorksheetCol(ACol);
  cell := FWorksheet.FindCell(r, c);
  neighborcell := FWorksheet.FindCell(r+ADeltaRow, c+ADeltaCol);
  // Only cell has border, but neighbor has not
  if ((cell <> nil) and (border in cell^.Border)) and
     ((neighborcell = nil) or (neighborborder in neighborcell^.Border))
  then
    ABorderStyle := cell^.BorderStyles[border]
  else
  // Only neighbor has border, cell has not
  if ((cell = nil) or not (border in cell^.Border)) and
     (neighborcell <> nil) and (neighborborder in neighborcell^.Border)
  then
    ABorderStyle := neighborcell^.BorderStyles[neighborborder]
  else
  // Both cells have shared border -> use top or left border
  if (cell <> nil) and (border in cell^.Border) and
     (neighborcell <> nil) and (neighborborder in neighborcell^.Border)
  then begin
    if (border in [cbNorth, cbWest]) then
      ABorderStyle := neighborcell^.BorderStyles[neighborborder]
    else
      ABorderStyle := cell^.BorderStyles[border];
  end else
    Result := false;
end;

function TsCustomWorksheetGrid.GetGridCol(ASheetCol: Cardinal): Integer;
begin
  Result := ASheetCol + FHeaderCount
end;

function TsCustomWorksheetGrid.GetGridRow(ASheetRow: Cardinal): Integer;
begin
  Result := ASheetRow + FHeaderCount;
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

{ Returns a list of worksheets contained in the file. Useful for assigning to
  user controls like TabControl, Combobox etc. in order to select a sheet. }
procedure TsCustomWorksheetGrid.GetSheets(const ASheets: TStrings);
var
  i: Integer;
begin
  ASheets.Clear;
  if Assigned(FWorkbook) then
    for i:=0 to FWorkbook.GetWorksheetCount-1 do
      ASheets.Add(FWorkbook.GetWorksheetByIndex(i).Name);
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

{ Calculates the index of the worksheet column that is displayed in the
  given column of the grid. If the sheet headers are turned on, both numbers
  differ by 1, otherwise they are equal. Saves an "if" in cases. }
function TsCustomWorksheetGrid.GetWorksheetCol(AGridCol: Integer): cardinal;
begin
  Result := AGridCol - FHeaderCount;
end;

{ Calculates the index of the worksheet row that is displayed in the
  given row of the grid. If the sheet headers are turned on, both numbers
  differ by 1, otherwise they are equal. Save an "if" in cases. }
function TsCustomWorksheetGrid.GetWorksheetRow(AGridRow: Integer): Cardinal;
begin
  Result := AGridRow - FHeaderCount;
end;

{ Returns if the cell has the given border }
function TsCustomWorksheetGrid.HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
begin
  Result := (ACell <> nil) and (uffBorder in ACell^.UsedFormattingfields) and
    (ABorder in ACell^.Border);
end;

{ Column width or row heights have changed. Stores the new number in the
  worksheet. }
procedure TsCustomWorksheetGrid.HeaderSized(IsColumn: Boolean; index: Integer);
var
  w0: Integer;
  h, h_pts: Single;
begin
  if FWorksheet = nil then
    exit;

  Convert_sFont_to_Font(FWorkbook.GetFont(0), Canvas.Font);
  if IsColumn then begin
    // The grid's column width is in "pixels", the worksheet's column width is
    // in "characters".
    w0 := Canvas.TextWidth('0');
    FWorksheet.WriteColWidth(GetWorksheetCol(Index), ColWidths[Index] div w0);
  end else begin
    // The grid's row heights are in "pixels", the worksheet's row heights are
    // in "lines"
    h_pts := PxToPts(RowHeights[Index] - 4, Screen.PixelsPerInch);  // in points
    h := h_pts / (FWorkbook.GetFont(0).Size + ROW_HEIGHT_CORRECTION);
    FWorksheet.WriteRowHeight(GetWorksheetRow(Index), h);
  end;
end;


{ Internal general text drawing method.
  - AText: text to be drawn
  - AMeasureText: text used for checking if the text fits into the text rectangle.
    If too large and ReplaceTooLong = true, a series of # is drawn.
  - ARect: Rectangle in which the text is drawn
  - AJustification: determines whether the text is drawn at the "start" (0),
    "center" (1) or "end" (2) of the drawing rectangle. Start/center/end are
    seen along the text drawing direction.
  - ACellHorAlign: Is the HorAlignment property stored in the cell
  - ACellVertAlign: Is the VertAlignment property stored in the cell
  - ATextRot: determines the rotation angle of the text.
  - ATextWrap: determines if the text can wrap into multiple lines
  - ReplaceTooLang: if true too-long texts are replaced by a series of # chars
    filling the cell.
  The reason to separate AJustification from ACellHorAlign and ACelVertAlign is
  the output of nfAccounting formatted numbers where the numbers are always
  right-aligned, and the currency symbol is left-aligned. }
procedure TsCustomWorksheetGrid.InternalDrawTextInCell(AText, AMeasureText: String;
  ARect: TRect; AJustification: Byte; ACellHorAlign: TsHorAlignment;
  ACellVertAlign: TsVertAlignment; ATextRot: TsTextRotation;
  ATextWrap, ReplaceTooLong: Boolean);
var
  ts: TTextStyle;
  flags: Cardinal;
  txt: String;
  txtL, txtR: String;
  txtRect: TRect;
  P: TPoint;
  w, h, h0, hline: Integer;
  i: Integer;
  L: TStrings;
  c, r: Integer;
  wrapped: Boolean;
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
      if w > ARect.Right - ARect.Left then begin
        if ReplaceTooLong then begin
          txt := '';
          repeat
            txt := txt + '#';
            LCLIntf.DrawText(Canvas.Handle, PChar(txt), Length(txt), txtRect, flags);
          until txtRect.Right - txtRect.Left > ARect.Right - ARect.Left;
          AText := Copy(txt, 1, Length(txt)-1);
        end;
        ts.Alignment := taLeftJustify;
      end else
        case AJustification of
          0: ts.Alignment := taLeftJustify;
          1: ts.Alignment := taCenter;
          2: ts.Alignment := taRightJustify;
        end;
      Canvas.TextStyle := ts;
      Canvas.TextRect(ARect,ARect.Left, ARect.Top, AText);
    end;
  end
  else
  begin
    // ROTATED TEXT DRAWING DIRECTION
    // Since there is not good API for multiline rotated text, we draw the text
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
          end;                                                                          {
          case vertAlign of
            vaTop    : P.Y := ARect.Top;
            vaCenter : P.Y := Max(ARect.Top, (ARect.Top + ARect.Bottom - w) div 2);
            vaBottom : P.Y := Max(ARect.Top, ARect.Bottom - w);
          end;}
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
          end;                                                                      {
          case vertAlign of
            vaTop    : P.Y := Min(ARect.Bottom, ARect.Top + w);
            vaCenter : P.Y := Min(ARect.Bottom, (ARect.Top + ARect.Bottom + w) div 2);
            vaBottom : P.Y := ARect.Bottom;
          end;}
          Canvas.TextRect(ARect, P.X, P.Y, L[i], ts);
          inc(P.X, hline);
        end;
      end;
    finally
      L.Free;
    end;
  end;
end;



{ Catches the ESC key during editing in order to restore the old cell text }
procedure TsCustomWorksheetGrid.KeyDown(var Key : Word; Shift : TShiftState);
begin
  if (Key = VK_ESCAPE) and FEditing then begin
    SetEditText(Col, Row, FOldEditText);
    EditorHide;
  end;
end;

procedure TsCustomWorksheetGrid.Loaded;
begin
  inherited;
  NewWorksheet(FInitColCount, FInitRowCount);
end;

{ Repaints after moving selection to avoid spurious rests of the old thick
  selection border. }
procedure TsCustomWorksheetGrid.MoveSelection;
begin
  //Refresh;
  inherited;
  Refresh;
end;

{ Is called when editing starts. Stores the old text just for the case that
  the user presses ESC to cancel editing. }
procedure TsCustomWorksheetGrid.SelectEditor;
begin
  FOldEditText := GetCellText(Col, Row);
  inherited;
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

{ Fetches the text that is currently in the editor. It is not yet transferred
  to the Worksheet because input is checked only at the end of editing. }
procedure TsCustomWorksheetGrid.SetEditText(ACol, ARow: Longint; const AValue: string);
begin
  FEditText := AValue;
  FEditing := true;
  inherited SetEditText(aCol, aRow, aValue);
end;

procedure TsCustomWorksheetGrid.SetFrozenCols(AValue: Integer);
begin
  FFrozenCols := AValue;
  Setup;
end;

procedure TsCustomWorksheetGrid.SetFrozenRows(AValue: Integer);
begin
  FFrozenRows := AValue;
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

procedure TsCustomWorksheetGrid.Setup;
var
  i: Integer;
  lCol: PCol;
  lRow: PRow;
  fc, fr: Integer;
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
    ColCount := FWorksheet.GetLastColIndex + 1 + FHeaderCount;
    RowCount := FWorksheet.GetLastRowIndex + 1 + FHeaderCount;
    FixedCols := FFrozenCols + FHeaderCount;
    FixedRows := FFrozenRows + FHeaderCount;
    if ShowHeaders then begin
      ColWidths[0] := Canvas.TextWidth(' 999999 ');
      RowHeights[0] := DefaultRowHeight;
    end;
    for i := FHeaderCount to ColCount-1 do begin
      lCol := FWorksheet.FindCol(i - FHeaderCount);
      if (lCol <> nil) then
        ColWidths[i] := CalcColWidth(lCol^.Width)
      else
        ColWidths[i] := DefaultColWidth;
    end;
    for i := FHeaderCount to RowCount-1 do begin
      lRow := FWorksheet.FindRow(i - FHeaderCount);
      if (lRow = nil) then
        RowHeights[i] := CalcAutoRowHeight(i)
      else
        RowHeights[i] := CalcRowHeight(lRow^.Height);
    end;
  end;
  Invalidate;
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

procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer);
begin
  BeginUpdate;
  try
    FreeAndNil(FWorkbook);
    FWorkbook := TsWorkbook.Create;
    FWorkbook.ReadFormulas := FReadFormulas;
    FWorkbook.ReadFromFile(AFileName, AFormat);
    LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AWorksheetIndex));
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AWorksheetIndex: Integer);
begin
  BeginUpdate;
  try
    FreeAndNil(FWorkbook);
    FWorkbook := TsWorkbook.Create;
    FWorkbook.ReadFormulas := FReadFormulas;
    FWorkbook.ReadFromFile(AFilename);
    LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AWorksheetIndex));
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.NewWorksheet(AColCount, ARowCount: Integer);
begin
  BeginUpdate;
  try
    FreeAndNil(FWorkbook);
    FWorkbook := TsWorkbook.Create;
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

{ Writes the workbook behind the grid to a spreadsheet file. }
procedure TsCustomWorksheetGrid.SaveToSpreadsheetFile(AFileName: String;
  AFormat: TsSpreadsheetFormat; AOverwriteExisting: Boolean = true);
begin
  if FWorkbook <> nil then
    FWorkbook.WriteToFile(AFileName, AFormat, AOverwriteExisting);
end;

procedure TsCustomWorksheetGrid.SaveToSpreadsheetFile(AFileName: String;
  AOverwriteExisting: Boolean = true);
begin
  if FWorkbook <> nil then
    FWorkbook.WriteToFile(AFileName, AOverwriteExisting);
end;

procedure TsCustomWorksheetGrid.SelectSheetByIndex(AIndex: Integer);
begin
  if FWorkbook <> nil then
    LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AIndex));
end;


initialization
  fpsutils.ScreenPixelsPerInch := Screen.PixelsPerInch;

finalization
  FreeAndNil(FillPattern_BIFF2);

end.
