unit fpspreadsheetgrid;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs, Grids,
  fpspreadsheet;

type

  { TsWorksheetGrid }

  { TsCustomWorksheetGrid }

  TsCustomWorksheetGrid = class(TCustomStringGrid)
  private
    FDisplayFixedColRow: Boolean;
    FWorksheet: TsWorksheet;
    procedure SetDisplayFixedColRow(const AValue: Boolean);
    { Private declarations }
  protected
    { Protected declarations }
  public
    { methods }
    constructor Create(AOwner: TComponent); override;
    procedure LoadFromWorksheet(AWorksheet: TsWorksheet);
    property DisplayFixedColRow: Boolean read FDisplayFixedColRow write SetDisplayFixedColRow;
  end;

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
    property Columns;
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

procedure Register;
begin
  RegisterComponents('Additional',[TsWorksheetGrid]);
end;

{ TsCustomWorksheetGrid }

procedure TsCustomWorksheetGrid.SetDisplayFixedColRow(const AValue: Boolean);
var
  x: Integer;
begin
  if AValue = FDisplayFixedColRow then Exit;

  FDisplayFixedColRow := AValue;

  if AValue then
  begin
    for x := 1 to ColCount - 1 do
      SetCells(x, 0, 'A');
    for x := 1 to RowCount - 1 do
      SetCells(0, x, IntToStr(x));
  end;
end;

constructor TsCustomWorksheetGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FDisplayFixedColRow := False;
  FixedCols := 0;
  FixedRows := 0;
end;

procedure TsCustomWorksheetGrid.LoadFromWorksheet(AWorksheet: TsWorksheet);
var
  x, lRow, lCol: Integer;
  lStr: string;
  lCell: PCell;
begin
  FWorksheet := AWorksheet;

  { First get the size of the table }
  if FWorksheet.GetCellCount = 0 then
  begin
    ColCount := 0;
    RowCount := 0;
  end
  else
  begin
    if DisplayFixedColRow then
    begin
      ColCount := FWorksheet.GetLastColNumber() + 2;
      RowCount := FWorksheet.GetLastRowNumber() + 2;
    end
    else
    begin
      ColCount := FWorksheet.GetLastColNumber() + 1;
      RowCount := FWorksheet.GetLastRowNumber() + 1;
    end;
  end;

  { Now copy the contents }

  lCell := FWorksheet.GetFirstCell();
  for x := 0 to FWorksheet.GetCellCount() - 1 do
  begin
    lCol := lCell^.Col;
    lRow := lCell^.Row;
    lStr := lCell^.UTF8StringValue;

    if DisplayFixedColRow then
      SetCells(lCol + 1, lRow + 1, lStr)
    else
      SetCells(lCol, lRow, lStr);

    lCell := FWorksheet.GetNextCell();
  end;
end;

end.
