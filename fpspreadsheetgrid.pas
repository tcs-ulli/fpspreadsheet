{
fpspreadsheetgrid.pas

Grid component which can load and write data from / to FPSpreadsheet documents

AUTHORS: Felipe Monteiro de Carvalho
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
    procedure DrawCellGrid(aCol,aRow: Integer; aRect: TRect; aState: TGridDrawState); override;
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
  fpsUtils;

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
  Result := round(AHeight / 25.4 * Screen.PixelsPerInch);
end;

procedure TsCustomWorksheetGrid.DoPrepareCanvas(ACol, ARow: Integer;
  AState: TGridDrawState);
var
  ts: TTextStyle;
  lCell: PCell;
  r, c: Integer;
begin
  ts := Canvas.TextStyle;
  if FDisplayFixedColRow then begin
    // Formatting of row and column headers
    if ARow = 0 then
      ts.Alignment := taCenter
    else
    if ACol = 0 then
      ts.Alignment := taRightJustify;
  end;
  if FWorksheet <> nil then begin
    r := ARow - FixedRows;
    c := ACol - FixedCols;
    lCell := FWorksheet.FindCell(r, c);
    if lCell <> nil then begin
      // Default alignment of number is right-justify
      if lCell^.ContentType = cctNumber then
        ts.Alignment := taRightJustify;
      // Word wrap?
      if (uffWordWrap in lCell^.UsedFormattingFields) then begin
        ts.Wordbreak := true;
        ts.SingleLine := false;
      end;
    end;
  end;
  Canvas.TextStyle := ts;

  inherited DoPrepareCanvas(ACol, ARow, AState);
end;

procedure TsCustomWorksheetGrid.DrawCellGrid(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
var
  lCell: PCell;
  r, c: Integer;
begin
  inherited;

  if FWorksheet = nil then exit;

  r := ARow - FixedRows;
  c := ACol - FixedCols;
  lCell := FWorksheet.FindCell(r, c);
  if (lCell <> nil) and (uffBorder in lCell^.UsedFormattingFields) then begin
    Canvas.Pen.Style := psSolid;
    Canvas.Pen.Color := clBlack;
    if (cbNorth in lCell^.Border) then
      Canvas.Line(ARect.Left, ARect.Top, ARect.Right, ARect.Top)
    else
    if (cbWest in lCell^.Border) then
      Canvas.Line(ARect.Left, ARect.Top, ARect.Left, ARect.Bottom)
    else
    if (cbEast in lCell^.Border) then
      Canvas.Line(ARect.Right-1, ARect.Top, ARect.Right-1, ARect.Bottom)
    else
    if (cbSouth in lCell^.Border) then
      Canvas.Line(ARect.Left, ARect.Bottom-1, ARect.Right, ARect.Bottom-1)
  end;
end;

procedure TsCustomWorksheetGrid.DrawTextInCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
begin
  DrawCellText(aCol, aRow, aRect, aState, GetCellText(ACol,ARow));
end;

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

end.
