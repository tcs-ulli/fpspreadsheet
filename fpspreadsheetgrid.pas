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
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FDisplayFixedColRow: Boolean;
    function CalcColWidth(AWidth: Single): Integer;
    function CalcRowHeight(AHeight: Single): Integer;
    procedure SetDisplayFixedColRow(const AValue: Boolean);
    { Private declarations }
  protected
    { Protected declarations }
    procedure DoPrepareCanvas(ACol, ARow: Integer; AState: TGridDrawState); override;
    procedure DrawAllRows; override;
    procedure DrawTextInCell(ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState); override;
    function GetCellText(ACol, ARow: Integer): String;
    procedure Loaded; override;
    procedure Setup;
  public
    { methods }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure GetSheets(const ASheets: TStrings);
    procedure LoadFromWorksheet(AWorksheet: TsWorksheet);
    procedure LoadFromSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string; AWorksheetIndex: Integer = 0); overload;
    procedure SaveToWorksheet(AWorksheet: TsWorksheet);
    procedure SelectSheetByIndex(AIndex: Integer);
    property DisplayFixedColRow: Boolean read FDisplayFixedColRow write SetDisplayFixedColRow;
    property Worksheet: TsWorksheet read FWorksheet;
    property Workbook: TsWorkbook read FWorkbook;
  end;

  { TsWorksheetGrid }

  TsWorksheetGrid = class(TsCustomWorksheetGrid)
  published
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
  fpCanvas, fpsUtils;

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

procedure Register;
begin
  RegisterComponents('Additional',[TsWorksheetGrid]);
end;


{ TsCustomWorksheetGrid }

constructor TsCustomWorksheetGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDisplayFixedColRow := true;
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

// Converts the row height, given in mm, to pixels
function TsCustomWorksheetGrid.CalcRowHeight(AHeight: Single): Integer;
begin
  Result := round(AHeight / 25.4 * Screen.PixelsPerInch) + 4;
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
  Canvas.Brush.Bitmap := nil;
  ts := Canvas.TextStyle;
  if FDisplayFixedColRow then begin
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
    r := ARow - FixedRows;
    c := ACol - FixedCols;
    lCell := FWorksheet.FindCell(r, c);
    if lCell <> nil then begin
      // Horizontal alignment
      case lCell^.HorAlignment of
        haDefault: if lCell^.ContentType = cctNumber then
                     ts.Alignment := taRightJustify
                   else
                     ts.Alignment := taLeftJustify;
        haLeft   : ts.Alignment := taLeftJustify;
        haCenter : ts.Alignment := taCenter;
        haRight  : ts.Alignment := taRightJustify;
      end;
      // Vertical alignment
      case lCell^.VertAlignment of
        vaDefault: ts.Layout := tlBottom;
        vaTop    : ts.Layout := tlTop;
        vaCenter : ts.Layout := tlCenter;
        vaBottom : ts.layout := tlBottom;
      end;
      // Word wrap
      if (uffWordWrap in lCell^.UsedFormattingFields) then begin
        ts.Wordbreak := true;
        ts.SingleLine := false;
      end;
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
  c, r: Integer;
  rect: TRect;
begin
  inherited;
  if FWorksheet = nil then exit;

  cell := FWorksheet.GetFirstCell;
  while cell <> nil do begin
    if (uffBorder in cell^.UsedFormattingFields) then begin
      c := cell^.Col + FixedCols;
      r := cell^.Row + FixedRows;
      rect := CellRect(c, r);
      Canvas.Pen.Style := psSolid;
      Canvas.Pen.Color := clBlack;
      if (cbNorth in cell^.Border) then
        Canvas.Line(rect.Left, rect.Top-1, rect.Right, rect.Top-1);
      if (cbWest in cell^.Border) then
        Canvas.Line(rect.Left-1, rect.Top, rect.Left-1, rect.Bottom);
      if (cbEast in cell^.Border) then
        Canvas.Line(rect.Right-1, rect.Top, rect.Right-1, rect.Bottom);
      if (cbSouth in cell^.Border) then
        Canvas.Line(rect.Left, rect.Bottom-1, rect.Right, rect.Bottom-1);
    end;
    cell := FWorksheet.GetNextCell;
  end;
end;

{ Draws the cell text. Calls "GetCellText" to determine the text in the cell. }
procedure TsCustomWorksheetGrid.DrawTextInCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
begin
  DrawCellText(aCol, aRow, aRect, aState, GetCellText(ACol,ARow));
end;

{ This function returns the text to be written in the cell }
function TsCustomWorksheetGrid.GetCellText(ACol, ARow: Integer): String;
var
  lCell: PCell;
  r, c: Integer;
begin
  Result := '';

  if FDisplayFixedColRow then begin
    // Titles
    if (ARow = 0) and (ACol = 0) then
      exit;
    if (ARow = 0) then begin
      Result := GetColString(ACol-FixedCols);
      exit;
    end
    else
    if (ACol = 0) then begin
      Result := IntToStr(ARow);
      exit;
    end;
  end;

  if FWorksheet <> nil then begin
    r := ARow - FixedRows;
    c := ACol - FixedCols;
    lCell := FWorksheet.FindCell(r, c);
    if lCell <> nil then
      Result := FWorksheet.ReadAsUTF8Text(r, c);
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

procedure TsCustomWorksheetGrid.Loaded;
begin
  inherited;
  Setup;
end;

procedure TsCustomWorksheetGrid.SetDisplayFixedColRow(const AValue: Boolean);
begin
  if AValue = FDisplayFixedColRow then Exit;

  FDisplayFixedColRow := AValue;
  Setup;
end;

procedure TsCustomWorksheetGrid.Setup;
var
  i: Integer;
  lCol: PCol;
  lRow: PRow;
begin
  if (FWorksheet = nil) or (FWorksheet.GetCellCount = 0) then begin
    if FDisplayFixedColRow then begin
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
  if FDisplayFixedColRow then begin
    ColCount := FWorksheet.GetLastColNumber + 2;
    RowCount := FWorksheet.GetLastRowNumber + 2;
    FixedCols := 1;
    FixedRows := 1;
    ColWidths[0] := Canvas.TextWidth(' 999999 ');
    // Setup column widths
    for i := FixedCols to ColCount-1 do begin
      lCol := FWorksheet.FindCol(i - FixedCols);
      if (lCol <> nil) then
        ColWidths[i] := CalcColWidth(lCol^.Width)
      else
        ColWidths[i] := DefaultColWidth;
    end;
  end else begin
    ColCount := FWorksheet.GetLastColNumber + 1;
    RowCount := FWorksheet.GetLastRowNumber + 1;
    FixedCols := 0;
    FixedRows := 0;
    for i := 0 to ColCount-1 do begin
      lCol := FWorksheet.FindCol(i);
      if (lCol <> nil) then
        ColWidths[i] := CalcColWidth(lCol^.Width)
      else
        ColWidths[i] := DefaultColWidth;
    end;
  end;
  if FWorksheet <> nil then begin
    RowHeights[0] := DefaultRowHeight;
    for i := FixedRows to RowCount-1 do begin
      lRow := FWorksheet.FindRow(i - FixedRows);
      if (lRow <> nil) then
        RowHeights[i] := CalcRowHeight(lRow^.Height)
      else
        RowHeights[i] := DefaultRowHeight;
    end
  end
  else
    for i:=0 to RowCount-1 do begin
      RowHeights[i] := DefaultRowHeight;
    end;
  Invalidate;
end;

procedure TsCustomWorksheetGrid.LoadFromWorksheet(AWorksheet: TsWorksheet);
begin
  FWorksheet := AWorksheet;
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
