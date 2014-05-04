{
fpspreadsheetgrid.pas

Grid component which can load and write data from / to FPSpreadsheet documents

AUTHORS: Felipe Monteiro de Carvalho, Werner Pamler
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
    FFrozenCols: Integer;
    FFrozenRows: Integer;
    function CalcAutoRowHeight(ARow: Integer): Integer;
    function CalcColWidth(AWidth: Single): Integer;
    function CalcRowHeight(AHeight: Single): Integer;
    function GetShowGridLines: Boolean;
    function GetShowHeaders: Boolean;
    procedure SetFrozenCols(AValue: Integer);
    procedure SetFrozenRows(AValue: Integer);
    procedure SetShowGridLines(AValue: Boolean);
    procedure SetShowHeaders(AValue: Boolean);
  protected
    { Protected declarations }
    procedure DefaultDrawCell(ACol, ARow: Integer; var ARect: TRect; AState: TGridDrawState); override;
    procedure DoPrepareCanvas(ACol, ARow: Integer; AState: TGridDrawState); override;
    procedure DrawAllRows; override;
    procedure DrawTextInCell(ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState); override;
    function GetCellHeight(ACol, ARow: Integer): Integer;
    function GetCellText(ACol, ARow: Integer): String;
    procedure Loaded; override;
    procedure Setup;
    property DisplayFixedColRow: Boolean read GetShowHeaders write SetShowHeaders default true;
    property FrozenCols: Integer read FFrozenCols write SetFrozenCols;
    property FrozenRows: Integer read FFrozenRows write SetFrozenRows;
    property ShowGridLines: Boolean read GetShowGridLines write SetShowGridLines default true;
    property ShowHeaders: Boolean read GetShowHeaders write SetShowHeaders default true;
  public
    { public methods }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure GetSheets(const ASheets: TStrings);
    procedure LoadFromWorksheet(AWorksheet: TsWorksheet);
    procedure LoadFromSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string; AWorksheetIndex: Integer = 0); overload;
    procedure SaveToWorksheet(AWorksheet: TsWorksheet);
    procedure SelectSheetByIndex(AIndex: Integer);
    { public properties }
    property Worksheet: TsWorksheet read FWorksheet;
    property Workbook: TsWorkbook read FWorkbook;
  end;

  { TsWorksheetGrid }

  TsWorksheetGrid = class(TsCustomWorksheetGrid)
  published
    // inherited from TsCustomWorksheetGrid
    property DisplayFixedColRow; deprecated 'Use ShowHeaders';
    property FrozenCols;
    property FrozenRows;
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
  Types, LCLType, LCLIntf, Math, fpCanvas, fpsUtils;

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

procedure Register;
begin
  RegisterComponents('Additional',[TsWorksheetGrid]);
end;


{ TsCustomWorksheetGrid }

constructor TsCustomWorksheetGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FHeaderCount := 1;
end;

destructor TsCustomWorksheetGrid.Destroy;
begin
  FreeAndNil(FWorkbook);
  inherited Destroy;
end;

// Converts the column width, given in "characters", to pixels
// All chars are assumed to have the same width defined by the "0".
// Therefore, this calculation is only approximate.
function TsCustomWorksheetGrid.CalcColWidth(AWidth: Single): Integer;
var
  w0: Integer;
begin
  w0 := Canvas.TextWidth('0');
  Result := Round(AWidth * w0);
end;

{ Finds the max cell height per row and uses this to define the RowHeights[].
  Returns DefaultRowHeight if the row does not contain any cells. }
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

{ Converts the row height, given in mm, to pixels }
function TsCustomWorksheetGrid.CalcRowHeight(AHeight: Single): Integer;
begin
  Result := round(AHeight / 25.4 * Screen.PixelsPerInch) + 4;
end;

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
    wasFixed := true;
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
begin
  Canvas.Font.Assign(Font);
  Canvas.Brush.Bitmap := nil;
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
      // Wordwrap, text alignment and text rotation are handled by "DrawTextInCell".
    end;
  end;
  Canvas.TextStyle := ts;

  inherited DoPrepareCanvas(ACol, ARow, AState);
end;

{ Paints the cell borders. This cannot be done in DrawCellGrid because the
  lower border line is overwritten when painting the next row. }
procedure TsCustomWorksheetGrid.DrawAllRows;
var
  cell: PCell;
  c, r, tmp: Integer;
  rect, cliprect: TRect;
  rgn: HRGN;

  procedure DrawBorderLine(ACell: PCell; ARect: TRect; ABorder: TsCellBorder;
    ALineStyle: TsLineStyle);
  const
    // TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble);
    PEN_WIDTHS: array[TsLineStyle] of Byte =
      (1, 2, 1, 1, 3, 1);
    PEN_STYLES: array[TsLineStyle] of TPenStyle =
      (psSolid, psSolid, psDash, psDot, psSolid, psSolid);
//      (psSolid, psSolid, psPattern, psPattern, psSolid, psSolid);
    PEN_PATTERNS: array[TsLineStyle] of TPenPattern =
      ($FFFFFFFF, $FFFFFFFF, $07070707, $AAAAAAAA, $FFFFFFFF, $FFFFFFFF);
  var
    w: Integer;
  begin
    if ALineStyle = lsDouble then
      case ABorder of
        cbEast, cbWest:
          begin
            InflateRect(ARect, -1, 0);
            DrawBorderLine(ACell, ARect, ABorder, lsThin);
            InflateRect(ARect, +2, 0);
            DrawBorderLine(ACell, ARect, ABorder, lsThin);
          end;
        cbNorth, cbSouth:
          begin
            InflateRect(ARect, 0, -1);
            DrawBorderLine(ACell, ARect, ABorder, lsThin);
            InflateRect(ARect, 0, +2);
            DrawBorderLine(ACell, ARect, ABorder, lsThin)
          end;
      end
    else begin
      w := PEN_WIDTHS[ACell^.BorderStyles[ABorder].LineStyle] div 2;
      Canvas.Pen.Style := PEN_STYLES[ACell^.BorderStyles[ABorder].LineStyle];
      Canvas.Pen.Width := PEN_WIDTHS[ACell^.BorderStyles[ABorder].LineStyle];
      Canvas.Pen.Color := FWorkBook.GetPaletteColor(ACell^.BorderStyles[ABorder].Color);
      //Canvas.Pen.Pattern := PEN_PATTERNS[ACell^.BorderStyles[ABorder].LineStyle];
      //Canvas.Pen.EndCap := pecSquare;

      case ABorder of
        cbEast : Canvas.Line(ARect.Right-1, ARect.Top, ARect.Right-1, ARect.Bottom-w);
        cbSouth: Canvas.Line(ARect.Left-1, ARect.Bottom-1, ARect.Right-w, ARect.Bottom-1);
        cbWest : Canvas.Line(ARect.Left-1, ARect.Top, ARect.Left-1, ARect.Bottom-w);
        cbNorth: Canvas.Line(ARect.Left-1, ARect.Top-1, ARect.Right-w, ARect.Top-1);
      end;
    end;
  end;

begin
  inherited;
  if FWorksheet = nil then exit;

  Canvas.SaveHandleState;
  try
    // Avoid painting into the fixed cells
    cliprect := ClientRect;
    if FixedCols > 0 then
      ColRowToOffset(True, True, FixedCols-1, tmp, cliprect.Left);
    if FixedRows > 0 then
      ColRowToOffset(False, True, FixedRows-1, tmp, cliprect.Top);
    rgn := CreateRectRgn(cliprect.Left, cliprect.top, cliprect.Right, cliprect.Bottom);
    SelectClipRgn(Canvas.Handle, Rgn);

    cell := FWorksheet.GetFirstCell;
    while cell <> nil do begin
      if (uffBorder in cell^.UsedFormattingFields) then begin
        c := cell^.Col + FHeaderCount;
        r := cell^.Row + FHeaderCount;
        rect := CellRect(c, r);
        if (cbNorth in cell^.Border) then
          DrawBorderLine(cell, rect, cbNorth, cell^.BorderStyles[cbNorth].LineStyle);
        if cbEast in cell^.Border then
          DrawBorderLine(cell, rect, cbEast, cell^.BorderStyles[cbEast].LineStyle);
        if cbSouth in cell^.Border then
          DrawBorderLine(cell, rect, cbSouth, cell^.BorderStyles[cbSouth].LineStyle);
        if cbWest in cell^.Border then
          DrawBorderLine(cell, rect, cbWest, cell^.BorderStyles[cbWest].LineStyle);
      end;
      cell := FWorksheet.GetNextCell;
    end;
    DeleteObject(rgn);
  finally
    Canvas.RestoreHandleState;
  end;
end;

{ Draws the cell text. Calls "GetCellText" to determine the text in the cell.
  Takes care of horizontal and vertical text alignment, text rotation and
  text wrapping }
procedure TsCustomWorksheetGrid.DrawTextInCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
const
  HOR_ALIGNMENTS: array[haLeft..haRight] of TAlignment = (
    taLeftJustify, taCenter, taRightJustify
  );
  VERT_ALIGNMENTS: array[TsVertAlignment] of TTextLayout = (
    tlBottom, tlTop, tlCenter, tlBottom
  );
var
  ts: TTextStyle;
  flags: Cardinal;
  txt: String;
  txtRect: TRect;
  P: TPoint;
  w, h, h0, hline: Integer;
  i: Integer;
  L: TStrings;
  c, r: Integer;
  wordwrap: Boolean;
  horAlign: TsHorAlignment;
  vertAlign: TsVertAlignment;
  lCell: PCell;
begin
  if FWorksheet = nil then
    exit;

  c := ACol - FHeaderCount;
  r := ARow - FHeaderCount;
  lCell := FWorksheet.FindCell(r, c);
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

  txt := GetCellText(ACol, ARow);
  if txt = '' then
    exit;

  if lCell^.HorAlignment <> haDefault then
    horAlign := lCell^.HorAlignment
  else begin
    if lCell^.ContentType = cctNumber then
      horAlign := haRight
    else
      horAlign := haLeft;
    if lCell^.TextRotation = rt90DegreeCounterClockwiseRotation then begin
      if horAlign = haRight then horAlign := haLeft else horAlign := haRight;
    end;
  end;
  vertAlign := lCell^.VertAlignment;
  wordwrap := (uffWordWrap in lCell^.UsedFormattingFields)
    or (lCell^.TextRotation = rtStacked);

  InflateRect(ARect, -constCellPadding, -constCellPadding);

  if (lCell^.TextRotation in [trHorizontal, rtStacked]) or
     (not (uffTextRotation in lCell^.UsedFormattingFields))
  then begin
    // HORIZONAL TEXT DRAWING DIRECTION
    ts := Canvas.TextStyle;
    if wordwrap then begin
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
      if wordwrap then begin
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

{ Returns the height (in pixels) of the cell at ACol/ARow. }
function TsCustomWorksheetGrid.GetCellHeight(ACol, ARow: Integer): Integer;
var
  lCell: PCell;
  s: String;
  wordwrap: Boolean;
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
    wordwrap := (uffWordWrap in lCell^.UsedFormattingFields)
      or (lCell^.TextRotation = rtStacked);
    // *** multi-line text ***
    if wordwrap then begin
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
        for i:=1 to Length(s) do
          Result := Result + s[i] + LineEnding;
//        Result := Result + s[Length(s)];
      end;
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

procedure TsCustomWorksheetGrid.Loaded;
begin
  inherited;
  Setup;
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

procedure TsCustomWorksheetGrid.SetShowGridLines(AValue: Boolean);
begin
  if AValue = GetShowGridLines then Exit;
  if AValue then
    Options := Options + [goHorzLine, goVertLine]
  else
    Options := Options - [goHorzLine, goVertLine];
end;

procedure TsCustomWorksheetGrid.SetShowHeaders(AValue: Boolean);
begin
  if AValue = GetShowHeaders then Exit;
  FHeaderCount := ord(AValue);
  Setup;
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
      ColCount := 2;
      RowCount := 2;
      FixedCols := 1;
      FixedRows := 1;
      ColWidths[0] := Canvas.TextWidth(' 999999 ');
    end else begin
      FixedCols := 0;
      FixedRows := 0;
      ColCount := 0;
      RowCount := 0;
    end;
  end else
  if FWorksheet <> nil then begin
    ColCount := FWorksheet.GetLastColNumber + 1 + FHeaderCount;
    RowCount := FWorksheet.GetLastRowNumber + 1 + FHeaderCount;
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
      if (lRow = nil) or lRow^.AutoHeight then
        RowHeights[i] := CalcAutoRowHeight(i)
      else
        RowHeights[i] := CalcRowHeight(lRow^.Height);
    end;
  end;
end;

procedure TsCustomWorksheetGrid.LoadFromWorksheet(AWorksheet: TsWorksheet);
begin
  FWorksheet := AWorksheet;
  if FWorksheet <> nil then begin
    ShowHeaders := (soShowHeaders in FWorksheet.Options);
    ShowGridLines := (soShowGridLines in FWorksheet.Options);
    if (soHasFrozenPanes in FWorksheet.Options) then begin
      FrozenCols := FWorksheet.LeftPaneWidth;
      FrozenRows := FWorksheet.TopPaneHeight;
    end else begin
      FrozenCols := 0;
      FrozenRows := 0;
    end;
  end;
  Setup;
end;

procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer);
begin
  FreeAndNil(FWorkbook);
  FWorkbook := TsWorkbook.Create;
  FWorkbook.ReadFromFile(AFileName, AFormat);
  LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AWorksheetIndex));
end;

procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AWorksheetIndex: Integer);
begin
  FreeAndNil(FWorkbook);
  FWorkbook := TsWorkbook.Create;
  FWorkbook.ReadFromFile(AFilename);
  LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AWorksheetIndex));
end;

procedure TsCustomWorksheetGrid.SaveToWorksheet(AWorksheet: TsWorksheet);
var
  x, y: Integer;
  Str: string;
begin
  if AWorksheet = nil then Exit;

  { Copy the contents }
  for x := 0 to ColCount - 1 do
    for y := 0 to RowCount - 1 do
    begin
      Str := GetCells(x, y);
      if Str <> '' then AWorksheet.WriteUTF8Text(y, x, Str);
    end;
end;

procedure TsCustomWorksheetGrid.SelectSheetByIndex(AIndex: Integer);
begin
  LoadFromWorksheet(FWorkbook.GetWorksheetByIndex(AIndex));
end;

initialization

finalization
  FreeAndNil(FillPattern_BIFF2);

end.