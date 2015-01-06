{ fpspreadsheetchart.pas }

{@@ ----------------------------------------------------------------------------
Chart data source designed to work together with TChart from Lazarus
to display the data and with FPSpreadsheet to load data.

AUTHORS: Felipe Monteiro de Carvalho, Werner Pamler

LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
         distribution, for details about the license.
-------------------------------------------------------------------------------}

unit fpspreadsheetchart;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs,
  // TChart
  TACustomSource,
  // FPSpreadsheet
  fpspreadsheet, fpsutils,
  // FPSpreadsheet Visual
  fpspreadsheetctrls, fpspreadsheetgrid;

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

  TsWorkbookChartSource = class(TCustomChartSource, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FWorkbook: TsWorkbook;
    FWorksheets: array[TsXYRange] of TsWorksheet;
    FRangeStr: array[TsXYRange] of String;
    FRanges: array[TsXYRange] of TsCellRangeArray;
    FPointsNumber: Cardinal;
    function GetRange(AIndex: TsXYRange): String;
    function GetWorkbook: TsWorkbook;
    procedure GetXYItem(XOrY:TsXYRange; APointIndex: Integer;
      out ANumber: Double; out AText: String);
    procedure SetRange(AIndex: TsXYRange; const AValue: String);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    FCurItem: TChartDataItem;
    function BuildRangeStr(AIndex: TsXYRange; AListSeparator: char = #0): String;
    function CountValues(AIndex: TsXYRange): Integer;
    function GetCount: Integer; override;
    function GetItem(AIndex: Integer): PChartDataItem; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure Prepare; overload;
    procedure Prepare(AIndex: TsXYRange); overload;
    procedure SetYCount(AValue: Cardinal); override;
  public
    destructor Destroy; override;
    procedure Reset;
    property PointsNumber: Cardinal read FPointsNumber;
    property Workbook: TsWorkbook read GetWorkbook;
  public
    // Interface to TsWorkbookSource
    procedure ListenerNotification(AChangedItems: TsNotificationItems; AData: Pointer = nil);
    procedure RemoveWorkbookSource;
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

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookChartSource.
  Removes itself from the WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsWorkbookChartSource.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Constructs the range string from the stored internal information. Is needed
  to have the worksheet name in the range string in order to make the range
  string unique.
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.BuildRangeStr(AIndex: TsXYRange;
  AListSeparator: Char = #0): String;
var
  L: TStrings;
  range: TsCellRange;
begin
  if (FWorkbook = nil) or (FWorksheets[AIndex] = nil) or (Length(FRanges) = 0) then
    exit('');

  L := TStringList.Create;
  try
    if AListSeparator = #0 then
      L.Delimiter := FWorkbook.FormatSettings.ListSeparator
    else
      L.Delimiter := AListSeparator;
    L.StrictDelimiter := true;
    for range in FRanges[AIndex] do
      L.Add(GetCellRangeString(range, rfAllRel, true));
    Result := FWorksheets[AIndex].Name + SHEETSEPARATOR + L.DelimitedText;
  finally
    L.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Counts the number of x or y values contained in the x/y ranges

  @param   AIndex   Identifies whether values in the x or y ranges are counted.
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.CountValues(AIndex: TsXYRange): Integer;
var
  ir: Integer;
  range: TsCellRange;
begin
  Result := 0;
  for range in FRanges[AIndex] do
  begin
    if range.Col1 = range.Col2 then
      inc(Result, range.Row2 - range.Row1 + 1)
    else
    if range.Row1 = range.Row2 then
      inc(Result, range.Col2 - range.Col1 + 1)
    else
      raise Exception.Create('x/y ranges can only be 1 column wide or 1 row high.');
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

  @param   AIndex   Index of the data point in the series.
  @return  Pointer to a TChartDataItem record containing the x and y coordinates,
           the data point mark text, and the individual data point color.
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
begin
  Result := FRangeStr[AIndex];
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

{@@ ----------------------------------------------------------------------------
  Helper method the prepare the information required for the series data point.

  @param  XOrY         Identifies whether the method retrieves the x or y
                       coordinate.
  @param  APointIndex  Index of the data point for which the data are required
  @param  ANumber      (output) x or y coordinate of the data point
  @param  AText        Data point marks label text
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.GetXYItem(XOrY:TsXYRange; APointIndex: Integer;
  out ANumber: Double; out AText: String);
var
  range: TsCellRange;
  idx, ir: Integer;
  len: Integer;
  row, col: Cardinal;
  cell: PCell;
begin
  cell := nil;
  idx := 0;
  if FRanges[XOrY] = nil then
    exit;

  for range in FRanges[XOrY] do
  begin
    if (range.Col1 = range.Col2) then  // vertical range
    begin
      len := range.Row2 - range.Row1 + 1;
      if (APointIndex >= idx) and (APointIndex < idx + len) then
      begin
        row := range.Row1 + APointIndex - idx;
        col := range.Col1;
        break;
      end;
      inc(idx, len);
    end else  // horizontal range
    if (range.Row1 = range.Row2) then
    begin
      len := range.Col2 - range.Col1 + 1;
      if (APointIndex >= idx) and (APointIndex < idx + len) then
      begin
        row := range.Row1;
        col := range.Col1 + APointIndex - idx;
        break;
      end;
    end else
      raise Exception.Create('x/y ranges can only be 1 column wide or 1 row high');
  end;

  cell := FWorksheets[XOrY].FindCell(row, col);

  if cell = nil then
  begin
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

  @param  AChangedItems  Set with elements identifying whether workbook,
                         worksheet, cell content or cell formatting has changed
  @param  AData          Additional data, not used here

  @see    TsNotificationItem
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

  // Workbook has been successfully loaded, all sheets are ready
  if (lniWorkbook in AChangedItems) then
    Prepare;

  // Used worksheet has been renamed?
  if (lniWorksheetRename in AChangedItems) then
    for xy in TsXYRange do
      if TsWorksheet(AData) = FWorksheets[xy] then begin
        FRangeStr[xy] := BuildRangeStr(xy);
        Prepare(xy);
      end;

  // Used worksheet will be deleted?
  if (lniWorksheetRemoving in AChangedItems) then
    for xy in TsXYRange do
      if TsWorksheet(AData) = FWorksheets[xy] then begin
        FWorksheets[xy] := nil;
        FRangeStr[xy] := BuildRangeStr(xy);
        Prepare(xy);
      end;

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
  Parses the x and y cell range strings and extracts internal information
  (worksheet used, cell range coordinates)
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Prepare;
begin
  Prepare(rngX);
  Prepare(rngY);
end;

{@@ ----------------------------------------------------------------------------
  Parses the range string of the data specified by AIndex and extracts internal
  information (worksheet used, cell range coordinates)

  @param  AIndex   Identifies whether x or y cell ranges are analyzed
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Prepare(AIndex: TsXYRange);
const
  XY: array[TsXYRange] of string = ('x', 'y');
var
  range: TsCellRange;
begin
  if (FWorkbook = nil) or (FRangeStr[AIndex] = '') then begin
    FWorksheets[AIndex] := nil;
    SetLength(FRanges[AIndex], 0);
    FPointsNumber := 0;
    Reset;
    exit;
  end;

  if FWorkbook.TryStrToCellRanges(FRangeStr[AIndex], FWorksheets[AIndex], FRanges[AIndex])
  then begin
    for range in FRanges[AIndex] do
      if (range.Col1 <> range.Col2) and (range.Row1 <> range.Row2) then
        raise Exception.Create('x/y ranges can only be 1 column wide or 1 row high');
    FPointsNumber := Max(CountValues(rngX), CountValues(rngY));
    // If x and y ranges are of different size empty data points will be plotted.
    Reset;
    // Make sure to include worksheet name in RangeString.
    FRangeStr[AIndex] := BuildRangeStr(AIndex);
  end else
  if (FWorkbook.GetWorksheetCount > 0) then begin
    if FWorksheets[AIndex] = nil then
      raise Exception.CreateFmt('Worksheet of %s cell range "%s" does not exist.',
        [XY[AIndex], FRangeStr[AIndex]])
    else
      raise Exception.CreateFmt('No valid %s cell range in "%s".',
        [XY[AIndex], FRangeStr[AIndex]]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the ChartSource to the WorkbookSource.
  Required before destruction.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.RemoveWorkbookSource;
begin
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
begin
  FRangeStr[AIndex] := AValue;
  Prepare;
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
  Currently we support only single valued data (YCount = 1).
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetYCount(AValue: Cardinal);
begin
  FYCount := AValue;
end;


end.
