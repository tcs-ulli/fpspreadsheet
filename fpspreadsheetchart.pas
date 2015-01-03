{
fpspreadsheetgrid.pas

Chart data source designed to work together with TChart from Lazarus to display the data
and with TsWorksheetGrid from FPSpreadsheet to load data from a grid.

AUTHORS: Felipe Monteiro de Carvalho
}
unit fpspreadsheetchart;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs,
  // TChart
  {tasources,} TACustomSource,
  // FPSpreadsheet Visual
  fpspreadsheetctrls, fpspreadsheetgrid,
  // FPSpreadsheet
  fpspreadsheet, fpsutils;

type

  {@@ Chart data source designed to work together with TChart from Lazarus
    to display the data.

    The data can be loaded from a TsWorksheetGrid Grid component or
    directly from a TsWorksheet FPSpreadsheet Worksheet }

  { TsWorksheetChartSource }

  { DEPRECTATED - use TsWorkbookChartSource instead! }

  TsWorksheetChartSource = class(TCustomChartSource)
  private
    FInternalWorksheet: TsWorksheet;
    FPointsNumber: Integer;
    FXSelectionDirection: TsSelectionDirection;
    FYSelectionDirection: TsSelectionDirection;
//    FWorksheetGrid: TsWorksheetGrid;
    FXFirstCellCol: Cardinal;
    FXFirstCellRow: Cardinal;
    FYFirstCellCol: Cardinal;
    FYFirstCellRow: Cardinal;
    procedure SetPointsNumber(const AValue: Integer);
    procedure SetXSelectionDirection(const AValue: TsSelectionDirection);
    procedure SetYSelectionDirection(const AValue: TsSelectionDirection);
    procedure SetXFirstCellCol(const AValue: Cardinal);
    procedure SetXFirstCellRow(const AValue: Cardinal);
    procedure SetYFirstCellCol(const AValue: Cardinal);
    procedure SetYFirstCellRow(const AValue: Cardinal);
  protected
    FDataWorksheet: TsWorksheet;
    FCurItem: TChartDataItem;
    function GetCount: Integer; override;
    function GetItem(AIndex: Integer): PChartDataItem; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure LoadPropertiesFromStrings(AXInterval, AYInterval, AXTitle, AYTitle, ATitle: string);
  public
    procedure LoadFromWorksheetGrid(const AValue: TsWorksheetGrid);
  published
//    property WorksheetGrid: TsWorksheetGrid read FWorksheetGrid write SetWorksheetGrid;
    property PointsNumber: Integer read FPointsNumber write SetPointsNumber default 0;
    property XFirstCellCol: Cardinal read FXFirstCellCol write SetXFirstCellCol default 0;
    property XFirstCellRow: Cardinal read FXFirstCellRow write SetXFirstCellRow default 0;
    property YFirstCellCol: Cardinal read FYFirstCellCol write SetYFirstCellCol default 0;
    property YFirstCellRow: Cardinal read FYFirstCellRow write SetYFirstCellRow default 0;
    property XSelectionDirection: TsSelectionDirection read FXSelectionDirection write SetXSelectionDirection;
    property YSelectionDirection: TsSelectionDirection read FYSelectionDirection write SetYSelectionDirection;
  end;


  { TsWorkbookChartSource }

  TsXYRange = (rngX, rngY);

  TsWorkbookChartSource = class(TCustomChartSource)
  private
    FWorkbookSource: TsWorkbookSource;
    FWorkbook: TsWorkbook;
    FWorksheets: array[TsXYRange] of TsWorksheet;
    FRanges: array[TsXYRange] of TsCellRangeArray;
    FDirections: array[TsXYRange] of TsSelectionDirection;
    FPointsNumber: Cardinal;
    function GetRange(AIndex: TsXYRange): String;
    function GetWorkbook: TsWorkbook;
    procedure GetXYItem(XOrY:TsXYRange; APointIndex: Integer;
      out ANumber: Double; out AText: String);
    procedure SetRange(AIndex: TsXYRange; const AValue: String);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    FCurItem: TChartDataItem;
    function CountValues(AIndex: TsXYRange): Integer;
    function GetCount: Integer; override;
    function GetItem(AIndex: Integer): PChartDataItem; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure SetYCount(AValue: Cardinal); override;
  public
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems; AData: Pointer = nil);
    procedure Reset;
    property PointsNumber: Cardinal read FPointsNumber;
    property Workbook: TsWorkbook read GetWorkbook;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    property XRange: String index rngX read GetRange write SetRange;
    property YRange: String index rngY read GetRange write SetRange;
  end;



procedure Register;

implementation

uses
  Math;


procedure Register;
begin
  RegisterComponents('Chart', [TsWorksheetChartSource, TsWorkbookChartSource]);
end;


{ TsWorksheetChartSource }

procedure TsWorksheetChartSource.SetPointsNumber(const AValue: Integer);
begin
  if FPointsNumber = AValue then exit;
  FPointsNumber := AValue;
  InvalidateCaches;
  Notify;
end;

procedure TsWorksheetChartSource.SetXSelectionDirection(
  const AValue: TsSelectionDirection);
begin
  if FXSelectionDirection=AValue then exit;
  FXSelectionDirection:=AValue;
  InvalidateCaches;
  Notify;
end;

procedure TsWorksheetChartSource.SetYSelectionDirection(
  const AValue: TsSelectionDirection);
begin
  if FYSelectionDirection=AValue then exit;
  FYSelectionDirection:=AValue;
  InvalidateCaches;
  Notify;
end;

procedure TsWorksheetChartSource.SetXFirstCellCol(const AValue: Cardinal);
begin
  if FXFirstCellCol=AValue then exit;
  FXFirstCellCol:=AValue;
  InvalidateCaches;
  Notify;
end;

procedure TsWorksheetChartSource.SetXFirstCellRow(const AValue: Cardinal);
begin
  if FXFirstCellRow=AValue then exit;
  FXFirstCellRow:=AValue;
  InvalidateCaches;
  Notify;
end;

procedure TsWorksheetChartSource.SetYFirstCellCol(const AValue: Cardinal);
begin
  if FYFirstCellCol=AValue then exit;
  FYFirstCellCol:=AValue;
  InvalidateCaches;
  Notify;
end;

procedure TsWorksheetChartSource.SetYFirstCellRow(const AValue: Cardinal);
begin
  if FYFirstCellRow=AValue then exit;
  FYFirstCellRow:=AValue;
  InvalidateCaches;
  Notify;
end;

function TsWorksheetChartSource.GetCount: Integer;
begin
  Result := FPointsNumber;
end;

function TsWorksheetChartSource.GetItem(AIndex: Integer): PChartDataItem;
var
  XRow, XCol, YRow, YCol: Integer;
begin
  // First calculate the cell position
  if XSelectionDirection = fpsVerticalSelection then
  begin
    XRow := Integer(FXFirstCellRow) + AIndex;
    XCol := FXFirstCellCol;
  end
  else
  begin
    XRow := FXFirstCellRow;
    XCol := Integer(FXFirstCellCol) + AIndex;
  end;

  if YSelectionDirection = fpsVerticalSelection then
  begin
    YRow := Integer(FYFirstCellRow) + AIndex;
    YCol := FYFirstCellCol;
  end
  else
  begin
    YRow := FYFirstCellRow;
    YCol := Integer(FYFirstCellCol) + AIndex;
  end;

  // Check the corresponding cell, if it is empty, use zero
  // If not, then get a number value

  FCurItem.X := FDataWorksheet.ReadAsNumber(XRow, XCol);
  FCurItem.Y := FDataWorksheet.ReadAsNumber(YRow, YCol);

  Result := @FCurItem;
end;

constructor TsWorksheetChartSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FInternalWorksheet := TsWorksheet.Create;
  FDataWorksheet := FInternalWorksheet;
end;

destructor TsWorksheetChartSource.Destroy;
begin
  if FInternalWorksheet <> nil then FInternalWorksheet.Free;
  inherited Destroy;
end;

procedure TsWorksheetChartSource.LoadFromWorksheetGrid(const AValue: TsWorksheetGrid);
begin
  if AValue = nil then Exit;

  FDataWorksheet := AValue.Worksheet;
//  AValue.SaveToWorksheet(FDataWorksheet);

  InvalidateCaches;
  Notify;
end;

procedure TsWorksheetChartSource.LoadPropertiesFromStrings(AXInterval,
  AYInterval, AXTitle, AYTitle, ATitle: string);
var
  lXCount, lYCount: Cardinal;
begin
  Unused(AXTitle, AYTitle, ATitle);
  ParseIntervalString(AXInterval, FXFirstCellRow, FXFirstCellCol, lXCount, FXSelectionDirection);
  ParseIntervalString(AYInterval, FYFirstCellRow, FYFirstCellCol, lYCount, FYSelectionDirection);
  if lXCount <> lYCount then raise Exception.Create(
    'TsWorksheetChartSource.LoadPropertiesFromStrings: Interval sizes don''t match');
  FPointsNumber := lXCount;
end;


{------------------------------------------------------------------------------}
{                             TsWorkbookChartSource                            }
{------------------------------------------------------------------------------}

destructor TsWorkbookChartSource.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Counts the number of x or y values contained in the x/y ranges
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.CountValues(AIndex: TsXYRange): Integer;
var
  ir: Integer;
begin
  Result := 0;
  case FDirections[AIndex] of
    fpsVerticalSelection:
      for ir:=0 to High(FRanges[AIndex]) do
        inc(Result, FRanges[AIndex, ir].Row2 - FRanges[AIndex, ir].Row1 + 1);
    fpsHorizontalSelection:
      for ir:=0 to High(FRanges[AIndex]) do
        inc(Result, FRanges[AIndex, ir].Col2 - FRanges[AIndex, ir].Col1 + 1);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Inherited ChartSource method telling the series how many data points are
  available
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.GetCount: Integer;
begin
  Result := FPointsNumber;
end;

{@@ ----------------------------------------------------------------------------
  Main ChartSource method called from the series requiring data for plotting.
  Retrieves the data from the workbook.
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.GetItem(AIndex: Integer): PChartDataItem;
var
  dummy: String;
begin
  GetXYItem(rngX, AIndex, FCurItem.X, FCurItem.Text);
  GetXYItem(rngY, AIndex, FCurItem.Y, dummy);
  Result := @FCurItem;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the cell range used for x or y coordinates (or x labels)

  @param   AIndex   Determines whether the methods deals with x or y values
  @return  An Excel string containing workbookname and cell block(s) in A1
           notation. Multiple blocks are separated by the ListSeparator defined
           by the workbook's FormatSettings.
-------------------------------------------------------------------------------}
function TsWorkbookChartsource.GetRange(AIndex: TsXYRange): String;
var
  L: TStrings;
  ir: Integer;
begin
  if FWorksheets[AIndex] = nil then
  begin
    Result := '';
    exit;
  end;

  L := TStringList.Create;
  try
    L.Delimiter := Workbook.FormatSettings.ListSeparator;
    for ir:=0 to High(FRanges[AIndex]) do
      L.Add(GetCellRangeString(FRanges[AIndex, ir], rfAllRel, true));
    Result := FWorksheets[AIndex].Name + SHEETSEPARATOR + L.DelimitedText;
  finally
    L.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the linked workbook
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := WorkbookSource.Workbook
  else
    Result := nil;
  FWorkbook := Result;
end;

procedure TsWorkbookChartSource.GetXYItem(XOrY:TsXYRange; APointIndex: Integer;
  out ANumber: Double; out AText: String);
var
  range: TsCellRange;
  i, j: Integer;
  len: Integer;
  row, col: Cardinal;
  cell: PCell;
begin
  cell := nil;
  i := 0;
  case FDirections[XOrY] of
    fpsVerticalSelection:
      for j:=0 to High(FRanges[XOrY]) do begin
        range := FRanges[XOrY, j];
        len := range.Row2 - range.Row1 + 1;
        if (APointIndex >= i) and (APointIndex < i + len) then begin
          row := range.Row1 + APointIndex - i;
          col := range.Col1;
          cell := FWorksheets[XOrY].FindCell(row, col);
          break;
        end;
        inc(i, len);
      end;

    fpsHorizontalSelection:
      for j:=0 to High(FRanges[XOrY]) do begin
        range := FRanges[XOrY, j];
        len := range.Col2 - range.Col1 + 1;
        if (APointIndex >= i) and (APointIndex < i + len) then begin
          row := range.Row1;
          col := range.Col1 + APointIndex - i;
          cell := FWorksheets[XOrY].FindCell(row, col);
          break;
        end;
        inc(i, len);
      end;
  end;
  if cell = nil then begin
    ANumber := NaN;
    AText := '';
  end else
  if cell^.ContentType = cctUTF8String then begin
    ANumber := APointIndex;
    AText := FWorksheets[rngX].ReadAsUTF8Text(cell);
  end else
  begin
    ANumber := FWorksheets[rngX].ReadAsNumber(cell);
    AText := '';
  end;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which
  spreadsheet item has changed.
  Responds to workbook changes by reading the worksheet names into the tabs,
  and to worksheet changes by selecting the tab corresponding to the selected
  worksheet.

  @param  AChangedItems  Set with elements identifying whether workbook, worksheet
                         cell content or cell formatting has changed
  @param  AData          Additional data, not used here
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  ir: Integer;
  cell: PCell;
  ResetDone: Boolean;
  xy: TsXYRange;
begin
  Unused(AData);

  // Worksheet changes
  if (lniWorksheet in AChangedItems) and (Workbook <> nil) then
    Reset;

  // Cell changes: Enforce recalculation of axes if modified cell is within the
  // x or y range(s).
  if (lniCell in AChangedItems) and (Workbook <> nil) then
  begin
    cell := PCell(AData);
    if (cell <> nil) then begin
      ResetDone := false;
      for xy in TsXYrange do
        for ir:=0 to High(FRanges[xy]) do
        begin
          if FWorksheets[xy].CellInRange(cell^.Row, cell^.Col, FRanges[xy, ir]) then
          begin
            Reset;
            ResetDone := true;
            break;
          end;
        if ResetDone then break;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification: The ChartSource is notified that the
  WorkbookSource is being removed.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Resets internal buffers and notfies chart elements of the changes,
  in particular, enforces recalculation of axis limits
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Reset;
begin
  InvalidateCaches;
  Notify;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the cell range used for x or y data (or labels) in the chart
  If it does not contain the worksheet name the currently active worksheet of
  the WorkbookSource is assumed.

  @param   AIndex     Distinguishes whether the method deals with x or y ranges
  @param   AValue     String in Excel syntax containing the cell range to be
                      used for x or y (depending on AIndex). Can contain multiple
                      cell blocks which must be separator by the ListSeparator
                      character defined in the Workbook's FormatSettings.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetRange(AIndex: TsXYRange; const AValue: String);
var
  s: String;
  p, i: Integer;
  L: TStrings;
  sd: TsSelectionDirection;
  sd0: TsSelectionDirection;
begin
  if (FWorkbook = nil) then
    exit;

  p := pos(SHEETSEPARATOR, AValue);
  if p = 0 then
  begin
    FWorksheets[AIndex] := FWorkbook.ActiveWorksheet;
    s := AValue;
  end else
  begin
    s := Copy(AValue, 1, p-1);
    FWorksheets[AIndex] := FWorkbook.GetWorksheetByName(s);
    if FWorksheets[AIndex] = nil then
      raise Exception.CreateFmt('%s cell range "%s" is in a non-existing '+
        'worksheet.', [''+char(ord('x')+ord(AIndex)), AValue]);
    s := Copy(AValue, p+1, Length(AValue));
  end;
  L := TStringList.Create;
  try
    L.Delimiter := FWorkbook.FormatSettings.ListSeparator;
    L.DelimitedText := s;
    if L.Count = 0 then
      raise Exception.CreateFmt('No %s cell range contained in "%s".',
        [''+char(ord('x')+ord(AIndex)), AValue]
      );
    sd := fpsVerticalSelection;
    SetLength(FRanges[AIndex], L.Count);
    for i:=0 to L.Count-1 do
      if ParseCellRangeString(L[i], FRanges[AIndex, i]) then begin
        if FRanges[AIndex, i].Col1 = FRanges[AIndex, i].Col2 then
          sd := fpsVerticalSelection
        else
        if FRanges[AIndex, i].Row1 = FRanges[AIndex, i].Row2 then
          sd := fpsHorizontalSelection
        else
          raise Exception.Create('Selection can only be 1 column wide or 1 row high');
      end else
        raise Exception.CreateFmt('No valid %s cell range in "%s".',
          [''+char(ord('x')+ord(AIndex)), L[i]]
        );
    FPointsNumber := Max(CountValues(rngX), CountValues(rngY));
    // If x and y ranges are of different size empty data points will be plotted.
    Reset;
  finally
    L.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  FWorkbook := GetWorkbook;
  ListenerNotification([lniWorkbook, lniWorksheet]);
end;

{@@ ----------------------------------------------------------------------------
  Inherited ChartSource method telling the series how many y values are used.
  Currently we support only single valued data
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetYCount(AValue: Cardinal);
begin
  FYCount := AValue;
  // currently not used
end;


end.
