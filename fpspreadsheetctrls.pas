unit fpspreadsheetctrls;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Controls, StdCtrls, ComCtrls, ValEdit, ActnList,
  fpspreadsheet, fpsAllFormats;

type
  TsWorkbookSourceErrorEvent = procedure (Sender: TObject;
    const AMsg: String) of object;

  TsNotificationItem = (lniWorkbook, lniWorksheet, lniCell, lniSelection);
  TsNotificationItems = set of TsNotificationItem;


  { TsWorkbookSource }

  TsWorkbookSource = class(TComponent)
  private
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FListeners: TFPList;
    FAutoDetectFormat: Boolean;
    FFileName: TFileName;
    FFileFormat: TsSpreadsheetFormat;
    FOptions: TsWorkbookOptions;
    FOnError: TsWorkbookSourceErrorEvent;
    procedure CellChangedHandler(Sender: TObject; ARow, ACol: Cardinal);
    procedure CellSelectedHandler(Sender: TObject; ARow, ACol: Cardinal);
    procedure InternalCreateNewWorkbook;
    procedure InternalLoadFromFile(AFileName: string; AAutoDetect: Boolean;
      AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0);
    procedure SetFileName(const AFileName: TFileName);
    procedure SetOptions(AValue: TsWorkbookOptions);
    procedure WorksheetAddedHandler(Sender: TObject; ASheet: TsWorksheet);
    procedure WorksheetChangedHandler(Sender: TObject; ASheet: TsWorksheet);
    procedure WorksheetRemovedHandler(Sender: TObject; ASheetIndex: Integer);

  protected
    procedure DoShowError(const AErrorMsg: String);
    procedure Loaded; override;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

  public
    procedure AddListener(AListener: TComponent);
    procedure RemoveListener(AListener: TComponent);
    procedure NotifyListeners(AChangedItems: TsNotificationItems; AData: Pointer = nil);

  public
    procedure CreateNewWorkbook;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AWorksheetIndex: Integer = 0); overload;
    procedure SaveToSpreadsheetFile(AFileName: string;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat;
      AOverwriteExisting: Boolean = true); overload;
    procedure SelectCell(ASheetRow, ASheetCol: Cardinal);
    procedure SelectWorksheet(AWorkSheet: TsWorksheet);

  public
    property Workbook: TsWorkbook read FWorkbook;
    property SelectedWorksheet: TsWorksheet read FWorksheet;

  published
    property AutoDetectFormat: Boolean read FAutoDetectFormat write FAutoDetectFormat;
    property FileFormat: TsSpreadsheetFormat read FFileFormat write FFileFormat default sfExcel8;
    property FileName: TFileName read FFileName write SetFileName;  // using this property loads the file at design-time!
    property Options: TsWorkbookOptions read FOptions write SetOptions;
    property OnError: TsWorkbookSourceErrorEvent read FOnError write FOnError;
  end;


  { TsWorkbookTabControl }

  TsWorkbookTabControl = class(TTabControl)
  private
    FWorkbookSource: TsWorkbookSource;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Change; override;
    procedure GetSheetList(AList: TStrings);
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    property Workbook: TsWorkbook read GetWorkbook;
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
  end;


  { TsCellEdit }

  TsCellEdit = class(TMemo)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetSelectedCell: PCell;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure ShowCell(ACell: PCell); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure EditingDone; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    property SelectedCell: PCell read GetSelectedCell;
    property Workbook: TsWorkbook read GetWorkbook;
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
  end;


  { TsCellIndicator }

  TsCellIndicator = class(TEdit)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure EditingDone; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    property Workbook: TsWorkbook read GetWorkbook;
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    property Alignment default taCenter;
  end;


  { TsSpreadsheetInspector }
  TsInspectorMode = (imWorkbook, imWorksheet, imCellValue, imCellProperties);

  TsSpreadsheetInspector = class(TValueListEditor)
  private
    FWorkbookSource: TsWorkbookSource;
    FMode: TsInspectorMode;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetMode(AValue: TsInspectorMode);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure DoUpdate; virtual;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure UpdateCellValue(ACell: PCell); virtual;
    procedure UpdateCellProperties(ACell: PCell); virtual;
    procedure UpdateWorkbook(AWorkbook: TsWorkbook); virtual;
    procedure UpdateWorksheet(ASheet: TsWorksheet); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    property Workbook: TsWorkbook read GetWorkbook;
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    property Mode: TsInspectorMode read FMode write SetMode;
    property DisplayOptions default [doColumnTitles, doAutoColResize];
    property FixedCols default 0;
  end;


procedure Register;


implementation

uses
  Dialogs, TypInfo,
  fpsStrings, fpsUtils, fpSpreadsheetGrid;


{@@ ----------------------------------------------------------------------------
  Registers the spreadsheet components in the Lazarus component palette,
  page "FPSpreadsheet".
-------------------------------------------------------------------------------}
procedure Register;
begin
  RegisterComponents('FPSpreadsheet', [TsWorkbookSource, TsWorkbookTabControl,
    TsCellEdit, TsCellIndicator, TsSpreadsheetInspector]);
end;


{------------------------------------------------------------------------------}
{                            TsWorkbookSource                                  }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the workbook source class. Creates the internal list for the
  notified ("listening") components, and creates an empty workbook.
-------------------------------------------------------------------------------}
constructor TsWorkbookSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FListeners := TFPList.Create;
  FFileFormat := sfExcel8;
  CreateNewWorkbook;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the workbook source. Cleans up the of listening component list
  and destroys the linked workbook.
-------------------------------------------------------------------------------}
destructor TsWorkbookSource.Destroy;
var
  i: Integer;
begin
  // Tell listeners that the workbook source will no longer exist
  for i:= FListeners.Count-1 downto 0 do
    RemoveListener(TComponent(FListeners[i]));
  // Destroy listener list
  FListeners.Free;
  // Destroy the instance of the workbook
  FWorkbook.Free;
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Adds a component to the list of listeners. All these components are
  notified of changes in the workbook.

  @param  AListener  Component to be added to the listener list notified for
                     changes
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.AddListener(AListener: TComponent);
begin
  if FListeners.IndexOf(AListener) = -1 then  // Avoid duplicates
    FListeners.Add(AListener);
end;

{@@ ----------------------------------------------------------------------------
  Event handler for the OnChange event of TsWorksheet which is fired whenver
  cell content or formatting changes.

  @param   Sender   Pointer to the worksheet
  @param   ARow     Row index (in sheet notation) of the cell changed
  @param   ACol     Column index (in sheet notation) of the cell changed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CellChangedHandler(Sender: TObject;
  ARow, ACol: Cardinal);
begin
  if FWorksheet <> nil then
    NotifyListeners([lniCell], FWorksheet.FindCell(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Event handler for the OnSelectCell event of TsWorksheet which is fired
  whenever another cell is selected in the worksheet. Notifies the listeners
  of the changed selection.

  @param  Sender   Pointer to the worksheet
  @param  ARow     Row index (in sheet notation) of the cell selected
  @param  ACol     Column index (in sheet notation) of the cell selected
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CellSelectedHandler(Sender: TObject;
  ARow, ACol: Cardinal);
begin
  Unused(ARow, ACol);
  NotifyListeners([lniSelection]);
end;

{@@ ----------------------------------------------------------------------------
  Creates a new empty workbook and adds a single worksheet
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CreateNewWorkbook;
begin
  InternalCreateNewWorkbook;
  FWorksheet := FWorkbook.AddWorksheet('Sheet1');
  SelectWorksheet(FWorksheet);

  // notify dependent controls
  NotifyListeners([lniWorkbook, lniWorksheet, lniSelection]);
end;

{ An error has occured during loading of the workbook. Shows a message box by
  default. But a different behavior can be obtained by means of the OnError
  event. }
procedure TsWorkbookSource.DoShowError(const AErrorMsg: String);
begin
  if Assigned(FOnError) then
    FOnError(self, AErrorMsg)
  else
    MessageDlg(AErrorMsg, mtError, [mbOK], 0);
end;

{@@ ----------------------------------------------------------------------------
  Helper method which creates a new workbook without sheets
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.InternalCreateNewWorkbook;
begin
  FreeAndNil(FWorkbook);
  FWorksheet := nil;
  FWorkbook := TsWorkbook.Create;
  FWorkbook.OnAddWorksheet := @WorksheetAddedHandler;
  FWorkbook.OnChangeWorksheet := @WorksheetChangedHandler;
  FWorkbook.OnRemoveWorksheet := @WorksheetRemovedHandler;
  // Pass options to workbook
  SetOptions(FOptions);
end;

{@@ ----------------------------------------------------------------------------
  Internal loader for the spreadsheet file. Is called with various combinations
  of arguments from several procedures.

  @param  AFilename        Name of the spreadsheet file to be loaded
  @param  AAutoDetect      Instructs the loader to automatically detect the
                           file format from the extension or by temporarily
                           opening the file in all available formats. Note that
                           an exception is raised in the IDE when an incorrect
                           format is tested.
  @param  AFormat          Spreadsheet file format assumed
  @param  AWorksheetIndex  Index of the worksheet to be loaded from the file
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.InternalLoadFromFile(AFileName: string;
  AAutoDetect: Boolean; AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0);
begin
  Unused(AWorksheetIndex);

  // Create a new empty workbook
  InternalCreateNewWorkbook;

  // Read workbook from file and get worksheet
  if AAutoDetect then
    FWorkbook.ReadFromFile(AFileName)
  else
    FWorkbook.ReadFromFile(AFileName, AFormat);

  SelectWorksheet(FWorkbook.GetWorkSheetByIndex(AWorksheetIndex));

  // If required, display loading error message
  if FWorkbook.ErrorMsg <> '' then
    DoShowError(FWorkbook.ErrorMsg);
end;

{@@ ----------------------------------------------------------------------------
  Inherited method which is called after loading from the lfm file.
  Is overridden here to open a spreadsheet file specified at design-time.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.Loaded;
begin
  inherited;
  if (FFileName <> '') then
    SetFileName(FFilename);
end;

{@@ ----------------------------------------------------------------------------
  Public spreadsheet loader to be used if file format is known.

  @param  AFilename        Name of the spreadsheet file to be loaded
  @param  AFormat          Spreadsheet file format assumed for the file
  @param  AWorksheetIndex  Index of the worksheet to be loaded from the file
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromSpreadsheetFile(AFileName: string;
  AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = 0);
begin
  InternalLoadFromFile(AFileName, false, AFormat, AWorksheetIndex);
end;

{ ------------------------------------------------------------------------------
  Public spreadsheet loader to be used if file format is not known. The file
  format is determined from the file extension, or - if this is valid for
  several formats (such as .xls) - by assuming a format. Note that exceptions
  are raised in the IDE if in incorrect format is tested. This does not occur
  outside the IDE:

  @param  AFilename        Name of the spreadsheet file to be loaded
  @param  AWorksheetIndex  Index of the worksheet to be loaded from the file
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromSpreadsheetFile(AFileName: string;
  AWorksheetIndex: Integer = 0);
const
  sfNotNeeded = sfExcel8;
  // The parameter AFormat if InternalLoadFromFile is not needed here,
  // but the compiler wants a value...
begin
  InternalLoadFromFile(AFileName, true, sfNotNeeded, AWorksheetIndex);
end;

{@@ ----------------------------------------------------------------------------
  Notifies listeners of workbook, worksheet, cell, or selection changes.
  The changed item is identified by the parameter AChangedItems.

  @param  AChangedItems  A set containing elements lniWorkbook, lniWorksheet,
                         lniCell, lniSelection which indicate which item has
                         changed.
  @param  AData          Additional data on the change. Is used only for
                         lniCell and points to the cell with changed value or
                         formatting.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.NotifyListeners(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
var
  i: Integer;
begin
  for i:=0 to FListeners.Count-1 do
    if TObject(FListeners[i]) is TsCellIndicator then
      TsCellIndicator(FListeners[i]).ListenerNotification(AChangedItems, AData)
    else
    if TObject(FListeners[i]) is TsCellEdit then
      TsCellEdit(FListeners[i]).ListenerNotification(AChangedItems, AData)
    else
    if TObject(FListeners[i]) is TsWorkbookTabControl then
      TsWorkbookTabControl(FListeners[i]).ListenerNotification(AChangedItems, AData)
    else
    if TObject(FListeners[i]) is TsWorksheetGrid then
      TsWorksheetGrid(FListeners[i]).ListenerNotification(AChangedItems, AData)
    else
    if TObject(FListeners[i]) is TsSpreadsheetInspector then
      TsSpreadsheetInspector(FListeners[i]).ListenerNotification(AChangedItems, AData)
    else                                    {
    if TObject(FListeners[i]) is TsSpreadsheetAction then
      TsSpreadsheetAction(FListeners[i]).ListenerNotifiation(AChangedItems, AData)
    else                                     }
      raise Exception.CreateFmt('Class %s is not prepared to be a spreadsheet listener.',
        [TObject(FListeners[i]).ClassName]);
end;

{@@ ----------------------------------------------------------------------------
  Removes a component from the listener list. The component is no longer
  notified of changes in workbook, worksheet or cells

  @param  AComponent  Component to be removed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.RemoveListener(AListener: TComponent);
var
  i: Integer;
begin
  for i:= FListeners.Count-1 downto 0 do
    if TComponent(FListeners[i]) = AListener then
    begin
      FListeners.Delete(i);
      if (AListener is TsCellIndicator) then
        TsCellIndicator(AListener).WorkbookSource := nil
      else
      if (AListener is TsCellEdit) then
        TsCellEdit(AListener).WorkbookSource := nil
      else
      if (AListener is TsWorkbookTabControl) then
        TsWorkbookTabControl(AListener).WorkbookSource := nil
      else
      if (AListener is TsWorksheetGrid) then
        TsWorksheetGrid(AListener).WorkbookSource := nil
      else
      if (AListener is TsSpreadsheetInspector) then
        TsSpreadsheetInspector(AListener).WorkbookSource := nil
      else                         {
      if (AListener is TsSpreadsheetAction) then
        TsSpreadsheetAction(AListener).WorksheetLink := nil
      else                          }
        raise Exception.CreateFmt('Class %s not prepared for listening.',[AListener.ClassName]);
      exit;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the workbook loaded into the WorkbookSource component to a
  spreadsheet file.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved.
  @param   AFormat            Spreadsheet file format in which the file is to be
                              saved.
  @param   AOverwriteExisting If the file already exists, it is overwritten in
                              the case of AOverwriteExisting = true, or an
                              exception is raised otherwise.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SaveToSpreadsheetFile(AFileName: String;
  AFormat: TsSpreadsheetFormat; AOverwriteExisting: Boolean = true);
begin
  if Workbook <> nil then
    Workbook.WriteToFile(AFileName, AFormat, AOverwriteExisting);
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
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SaveToSpreadsheetFile(AFileName: String;
  AOverwriteExisting: Boolean = true);
begin
  if Workbook <> nil then
    Workbook.WriteToFile(AFileName, AOverwriteExisting);
end;

{@@ ----------------------------------------------------------------------------
  Usually called by code or from the spreadsheet grid component. The
  method identifies a cell as "selected". Stores its coordinates in the
  worksheet and notifies the controls
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SelectCell(ASheetRow, ASheetCol: Cardinal);
begin
  if SelectedWorksheet <> nil then
    FWorksheet.SelectCell(ASheetRow, ASheetCol);
  NotifyListeners([lniSelection]);
end;

{@@ ----------------------------------------------------------------------------
  Selects a worksheet and notifies the controls. This method is usually called
  by code or from the worksheet tabcontrol.

  @param  AWorksheet  Instsance of the newly selected worksheet.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SelectWorksheet(AWorkSheet: TsWorksheet);
begin
  if AWorksheet = nil then
    exit;
  FWorksheet := AWorkSheet;
  FWorksheet.OnChangeCell := @CellChangedHandler;
  FWorksheet.OnSelectCell := @CellSelectedHandler;
  NotifyListeners([lniWorksheet]);
  SelectCell(FWorksheet.ActiveCellRow, FWorksheet.ActiveCellCol);
end;

{@@ ----------------------------------------------------------------------------
  Setter for the file name property. Loads the spreadsheet file and uses the
  values of the properties AutoDetectFormat and FileFormat.
  Useful if the spreadsheet is to be loaded at design time.
  But note that an exception can be raised if the file format cannot be
  determined from the file extension alone.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SetFileName(const AFileName: TFileName);
begin
  if AFileName = '' then
  begin
    CreateNewWorkbook;
    FFileName := '';
    exit;
  end;

  if FileExists(AFileName) then
  begin
    if FAutoDetectFormat then
      LoadFromSpreadsheetFile(AFileName)
    else
      LoadFromSpreadsheetFile(AFileName, FFileFormat);
    FFileName := AFileName;
  end else
    raise Exception.CreateFmt(rsFileNotFound, [AFileName]);
end;

{@@ ----------------------------------------------------------------------------
  Setter for the property Options. Copies the options of the WorkbookSource
  to the workbook
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SetOptions(AValue: TsWorkbookOptions);
begin
  FOptions := AValue;
  if Workbook <> nil then
    Workbook.Options := FOptions;
end;

{@@ ----------------------------------------------------------------------------
  Event handler called whenever a new worksheet is added to the workbook
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetAddedHandler(Sender: TObject;
  ASheet: TsWorksheet);
begin
  NotifyListeners([lniWorkbook]);
  SelectWorksheet(ASheet);
end;

{@@ ----------------------------------------------------------------------------
  Event handler canned whenever worksheet properties changed. Currently only
  for changing the workbook name.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetChangedHandler(Sender: TObject;
  ASheet: TsWorksheet);
begin
  Unused(ASheet);
  NotifyListeners([lniWorkbook, lniWorksheet]);
end;

{@@ ----------------------------------------------------------------------------
  Event handler called AFTER a worksheet has been removed (deleted) from
  the workbook

  @param  ASheetIndex  Index of the sheet that was deleted. The sheet itself
                       does not exist any more.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetRemovedHandler(Sender: TObject;
  ASheetIndex: Integer);
var
  i, sheetCount: Integer;
  sheet: TsWorksheet;
begin
  // It is very possible that the currently selected worksheet has been deleted.
  // Look for the selected worksheet in the workbook. Does it still exist? ...
  i := Workbook.GetWorksheetIndex(FWorksheet);
  if i = -1 then
  begin
    // ... no - it must have been the sheet deleted.
    // We have to select another worksheet.
    sheetCount := Workbook.GetWorksheetCount;
    if (ASheetIndex >= sheetCount) then
      sheet := Workbook.GetWorksheetByIndex(sheetCount-1)
    else
      sheet := Workbook.GetWorksheetbyIndex(ASheetIndex);
  end else
    sheet := FWorksheet;
  NotifyListeners([lniWorkbook]);
  SelectWorksheet(sheet);
end;


{------------------------------------------------------------------------------}
{                            TsWorkbookTabControl                              }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookTabControl. Removes itself from the
  WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsWorkbookTabControl.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  The currently active tab has been changed. The WorkbookSource must activate
  the corresponding worksheet and notify its listening components of the change.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.Change;
begin
  if FWorkbookSource <> nil then
    FWorkbookSource.SelectWorksheet(Workbook.GetWorksheetByIndex(TabIndex));
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Creates a (string) list containing the names of the workbook's sheet names
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.GetSheetList(AList: TStrings);
var
  i: Integer;
  oldTabIndex: Integer;
begin
  oldTabIndex := TabIndex;
  AList.BeginUpdate;
  try
    AList.Clear;
    if Workbook <> nil then
      for i:=0 to Workbook.GetWorksheetCount-1 do
        AList.Add(Workbook.GetWorksheetByIndex(i).Name);
  finally
    AList.EndUpdate;
    if oldtabIndex < AList.Count then
      TabIndex := oldTabIndex;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for property "Workbook"
-------------------------------------------------------------------------------}
function TsWorkbookTabControl.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for property "Worksheet"
-------------------------------------------------------------------------------}
function TsWorkbookTabControl.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.SelectedWorksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which
  spreadsheet item has changed.
  Responds to workbook changes by reading the worksheet names into the tabs,
  and to worksheet changes by selecting the tab corresponding to the selected
  worksheet.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  i: Integer;
begin
  Unused(AData);

  // Workbook changed
  if (lniWorkbook in AChangedItems) then
    GetSheetList(Tabs);

  // Worksheet changed
  if (lniWorksheet in AChangedItems) and (Worksheet <> nil) then
  begin
    i := Tabs.IndexOf(Worksheet.Name);
    if i <> TabIndex then
      TabIndex := i;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification. Must clean up the WorkbookSource field
  when the workbook source is going to be deleted.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  ListenerNotification([lniWorkbook, lniWorksheet]);
end;


{------------------------------------------------------------------------------}
{                               TsCellEdit                                     }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the spreadsheet edit control. Disables RETURN and TAB keys.
  RETURN characters can still be entered into the edited text by pressing
  CTRL+RETURN
-------------------------------------------------------------------------------}
constructor TsCellEdit.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  WantReturns := false;
  WantTabs := false;
  AutoSize := true;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsCellEdit. Removes itself from the WorkbookSource's
  listener list.
-------------------------------------------------------------------------------}
destructor TsCellEdit.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  EditingDone is called when the user presses the RETURN key to finish editing,
  or the TAB key which removes focus from the control, or clicks somewhere else
  The edited text is written to the worksheet which tries to figure out the
  data type. In particular, if the text begins with an "=" sign then the text
  is written as a formula.
-------------------------------------------------------------------------------}
procedure TsCellEdit.EditingDone;
var
  r, c: Cardinal;
  s: String;
begin
  if Worksheet = nil then
    exit;
  r := Worksheet.ActiveCellRow;
  c := Worksheet.ActiveCellCol;
  s := Lines.Text;
  if (s <> '') and (s[1] = '=') then
    Worksheet.WriteFormula(r, c, Copy(s, 2, Length(s)), true)
  else
    Worksheet.WriteCellValueAsString(r, c, s);
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property SelectedCell which points to the currently
  selected cell in the selected worksheet
-------------------------------------------------------------------------------}
function TsCellEdit.GetSelectedCell: PCell;
begin
  if (Worksheet <> nil) then
    with Worksheet do
      Result := FindCell(ActiveCellRow, ActiveCellCol)
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Workbook which is currently loaded into the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellEdit.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Worksheet which is currently "selected" in the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellEdit.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.SelectedWorksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which item of the
  spreadsheet has changed.
  Responds to selection and cell changes by updating the cell content.
-------------------------------------------------------------------------------}
procedure TsCellEdit.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
begin
  if (FWorkbookSource = nil) then
    exit;

  if  (lniSelection in AChangedItems) or
     ((lniCell in AChangedItems) and (PCell(AData) = SelectedCell))
  then
    ShowCell(SelectedCell);
end;

{ Standard component notification when the workbook link is deleted. }
procedure TsCellEdit.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsCellEdit.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  Text := '';
  ListenerNotification([lniSelection]);
end;

{@@ ----------------------------------------------------------------------------
  Loads the contents of a cell into the editor.
  Shows the formula if available. Numbers are displayed in full precision.
  Date and time values are shown in the long formats.
-------------------------------------------------------------------------------}
procedure TsCellEdit.ShowCell(ACell: PCell);
var
  s: String;
begin
  if (FWorkbookSource <> nil) and (ACell <> nil) then
  begin
    s := Worksheet.ReadFormulaAsString(ACell, true);
    if s <> '' then begin
      if s[1] <> '=' then s := '=' + s;
      Lines.Text := s;
    end else
      case ACell^.ContentType of
        cctNumber:
          Lines.Text := FloatToStr(ACell^.NumberValue);
        cctDateTime:
          if ACell^.DateTimeValue < 1.0 then
            Lines.Text := FormatDateTime('tt', ACell^.DateTimeValue)
          else
            Lines.Text := FormatDateTime('c', ACell^.DateTimeValue);
        else
          Lines.Text := Worksheet.ReadAsUTF8Text(ACell);
      end;
  end else
    Clear;
end;

{------------------------------------------------------------------------------}
{                              TsCellIndicator                                 }
{------------------------------------------------------------------------------}

constructor TsCellIndicator.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Alignment := taCenter;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the cell indicator. Removes itself from the WorkbookSource's
  listener list.
-------------------------------------------------------------------------------}
destructor TsCellIndicator.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  EditingDone is called when the user presses the RETURN key to finish editing,
  or the TAB key which removes focus from the control, or clicks somewhere else
  The edited text is interpreted as a cell address. The corresponding cell is
  selected.
-------------------------------------------------------------------------------}
procedure TsCellIndicator.EditingDone;
var
  r, c: Cardinal;
begin
  if (WorkbookSource <> nil) and ParseCellString(Text, r, c) then
    WorkbookSource.SelectCell(r, c);
end;

function TsCellIndicator.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

function TsCellIndicator.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.SelectedWorksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  The cell indicator responds to notification that the selection has changed
  and displays the address of the selected cell as editable text.
-------------------------------------------------------------------------------}
procedure TsCellIndicator.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
begin
  Unused(AData);
  if (lniSelection in AChangedItems) and (Worksheet <> nil) then
    Text := GetCellString(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol)
  else
    Text := '';
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification called when the WorkbookSource is deleted.
-------------------------------------------------------------------------------}
procedure TsCellIndicator.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbooksource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsCellIndicator.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  Text := '';
  ListenerNotification([lniSelection]);
end;


{------------------------------------------------------------------------------}
{                          TsSpreadsheetInspector                              }
{------------------------------------------------------------------------------}

constructor TsSpreadsheetInspector.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DisplayOptions := DisplayOptions - [doKeyColFixed];
  FixedCols := 0;
  TitleCaptions.Add('Properties');
  TitleCaptions.Add('Values');
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the spreadsheet inspector. Removes itself from the
  WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsSpreadsheetInspector.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Updates the data shown by the inspector grid. Display depends on the FMode
  setting (workbook, worksheet, cell values, cell properties).
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.DoUpdate;
var
  cell: PCell;
  sheet: TsWorksheet;
  book: TsWorkbook;
begin
  Strings.Clear;

  cell := nil;
  sheet := nil;
  book := nil;
  if FWorkbookSource <> nil then
  begin
    book := FWorkbookSource.Workbook;
    sheet := FWorkbookSource.SelectedWorksheet;
    if sheet <> nil then
      cell := sheet.FindCell(sheet.ActiveCellRow, sheet.ActiveCellCol);
  end;

  case FMode of
    imCellValue      : UpdateCellValue(cell);
    imCellProperties : UpdateCellProperties(cell);
    imWorksheet      : UpdateWorksheet(sheet);
    imWorkbook       : UpdateWorkbook(book);
  end;
end;

function TsSpreadsheetInspector.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

function TsSpreadsheetInspector.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.SelectedWorksheet
  else
    Result := nil;
end;

procedure TsSpreadsheetInspector.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
begin
  Unused(AData);
  case FMode of
    imWorkbook:
      if lniWorkbook in AChangedItems then DoUpdate;
    imWorksheet:
      if lniWorksheet in AChangedItems then DoUpdate;
    imCellValue,
    imCellProperties:
      if ([lniCell, lniSelection]*AChangedItems <> []) then DoUpdate;
  end;
end;

procedure TsSpreadsheetInspector.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

procedure TsSpreadsheetInspector.SetMode(AValue: TsInspectorMode);
begin
  if AValue = FMode then
    exit;
  FMode := AValue;
  DoUpdate;
end;

procedure TsSpreadsheetInspector.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  ListenerNotification([lniWorkbook, lniWorksheet, lniSelection]);
end;

procedure TsSpreadsheetInspector.UpdateCellProperties(ACell: PCell);
var
  s: String;
  cb: TsCellBorder;
  r1, r2, c1, c2: Cardinal;
begin
  if (ACell = nil) or not (uffFont in ACell^.UsedFormattingFields)
    then Strings.Add('FontIndex=')
    else Strings.Add(Format('FontIndex=%d (%s)', [
           ACell^.FontIndex,
           Workbook.GetFontAsString(ACell^.FontIndex)
         ]));

  if (ACell=nil) or not (uffTextRotation in ACell^.UsedFormattingFields)
    then Strings.Add('TextRotation=')
    else Strings.Add(Format('TextRotation=%s', [
           GetEnumName(TypeInfo(TsTextRotation), ord(ACell^.TextRotation))
         ]));

  if (ACell=nil) or not (uffHorAlign in ACell^.UsedFormattingFields)
    then Strings.Add('HorAlignment=')
    else Strings.Add(Format('HorAlignment=%s', [
           GetEnumName(TypeInfo(TsHorAlignment), ord(ACell^.HorAlignment))
         ]));

  if (ACell=nil) or not (uffVertAlign in ACell^.UsedFormattingFields)
    then Strings.Add('VertAlignment=')
    else Strings.Add(Format('VertAlignment=%s', [
           GetEnumName(TypeInfo(TsVertAlignment), ord(ACell^.VertAlignment))
         ]));

  if (ACell=nil) or not (uffBorder in ACell^.UsedFormattingFields) then
    Strings.Add('Borders=')
  else
  begin
    s := '';
    for cb in TsCellBorder do
      if cb in ACell^.Border then
        s := s + ', ' + GetEnumName(TypeInfo(TsCellBorder), ord(cb));
    if s <> '' then Delete(s, 1, 2);
    Strings.Add('Borders='+s);
  end;

  for cb in TsCellBorder do
    if ACell = nil then
      Strings.Add(Format('BorderStyles[%s]=', [
        GetEnumName(TypeInfo(TsCellBorder), ord(cb))]))
    else
      Strings.Add(Format('BorderStyles[%s]=%s, %s', [
        GetEnumName(TypeInfo(TsCellBorder), ord(cb)),
        GetEnumName(TypeInfo(TsLineStyle), ord(ACell^.BorderStyles[cbEast].LineStyle)),
        Workbook.GetColorName(ACell^.BorderStyles[cbEast].Color)]));

  if (ACell = nil) or not (uffBackgroundColor in ACell^.UsedformattingFields)
    then Strings.Add('BackgroundColor=')
    else Strings.Add(Format('BackgroundColor=%d (%s)', [
           ACell^.BackgroundColor,
           Workbook.GetColorName(ACell^.BackgroundColor)]));

  if (ACell = nil) or not (uffNumberFormat in ACell^.UsedFormattingFields) then
  begin
    Strings.Add('NumberFormat=');
    Strings.Add('NumberFormatStr=');
  end else
  begin
    Strings.Add(Format('NumberFormat=%s', [
      GetEnumName(TypeInfo(TsNumberFormat), ord(ACell^.NumberFormat))]));
    Strings.Add('NumberFormatStr=' + ACell^.NumberFormatStr);
  end;

  if (Worksheet = nil) or not Worksheet.IsMerged(ACell) then
    Strings.Add('Merged range=')
  else
  begin
    Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    Strings.Add('Merged range=' + GetCellRangeString(r1, c1, r2, c2));
  end;
end;

procedure TsSpreadsheetInspector.UpdateCellValue(ACell: PCell);
begin
  if ACell = nil then
  begin
    if Worksheet <> nil then
    begin
      Strings.Add(Format('Row=%d', [Worksheet.ActiveCellRow]));
      Strings.Add(Format('Col=%d', [Worksheet.ActiveCellCol]));
    end else
    begin
      Strings.Add('Row=');
      Strings.Add('Col=');
    end;
    Strings.Add('ContentType=(none)');
  end else
  begin
    Strings.Add(Format('Row=%d', [ACell^.Row]));
    Strings.Add(Format('Col=%d', [ACell^.Col]));
    Strings.Add(Format('ContentType=%s', [
      GetEnumName(TypeInfo(TCellContentType), ord(ACell^.ContentType))
    ]));
    Strings.Add(Format('NumberValue=%g', [ACell^.NumberValue]));
    Strings.Add(Format('DateTimeValue=%g', [ACell^.DateTimeValue]));
    Strings.Add(Format('UTF8StringValue=%s', [ACell^.UTF8StringValue]));
    Strings.Add(Format('BoolValue=%s', [BoolToStr(ACell^.BoolValue)]));
    Strings.Add(Format('ErrorValue=%s', [
      GetEnumName(TypeInfo(TsErrorValue), ord(ACell^.ErrorValue))
    ]));
    Strings.Add(Format('FormulaValue=%s', [Worksheet.ReadFormulaAsString(ACell, true)])); //^.FormulaValue]));
    if ACell^.SharedFormulaBase = nil then
      Strings.Add('SharedFormulaBase=')
    else
      Strings.Add(Format('SharedFormulaBase=%s', [GetCellString(
        ACell^.SharedFormulaBase^.Row, ACell^.SharedFormulaBase^.Col)
      ]));
  end;
end;

procedure TsSpreadsheetInspector.UpdateWorkbook(AWorkbook: TsWorkbook);
var
  bo: TsWorkbookOption;
  s: String;
  i: Integer;
begin
  if AWorkbook = nil then
  begin
    Strings.Add('FileName=');
    Strings.Add('FileFormat=');
    Strings.Add('Options=');
    Strings.Add('FormatSettings=');
  end else
  begin
    Strings.Add(Format('FileName=%s', [AWorkbook.FileName]));
    Strings.Add(Format('FileFormat=%s', [
      GetEnumName(TypeInfo(TsSpreadsheetFormat), ord(AWorkbook.FileFormat))
    ]));

    s := '';
    for bo in TsWorkbookOption do
      if bo in AWorkbook.Options then
        s := s + ', ' + GetEnumName(TypeInfo(TsWorkbookOption), ord(bo));
    if s <> '' then Delete(s, 1, 2);
    Strings.Add('Options='+s);

    Strings.Add('FormatSettings=');
    Strings.Add('  ThousandSeparator='+AWorkbook.FormatSettings.ThousandSeparator);
    Strings.Add('  DecimalSeparator='+AWorkbook.FormatSettings.DecimalSeparator);
    Strings.Add('  ListSeparator='+AWorkbook.FormatSettings.ListSeparator);
    Strings.Add('  DateSeparator='+AWorkbook.FormatSettings.DateSeparator);
    Strings.Add('  TimeSeparator='+AWorkbook.FormatSettings.TimeSeparator);
    Strings.Add('  ShortDateFormat='+AWorkbook.FormatSettings.ShortDateFormat);
    Strings.Add('  LongDateFormat='+AWorkbook.FormatSettings.LongDateFormat);
    Strings.Add('  ShortTimeFormat='+AWorkbook.FormatSettings.ShortTimeFormat);
    Strings.Add('  LongTimeFormat='+AWorkbook.FormatSettings.LongTimeFormat);
    Strings.Add('  TimeAMString='+AWorkbook.FormatSettings.TimeAMString);
    Strings.Add('  TimePMString='+AWorkbook.FormatSettings.TimePMString);
    s := AWorkbook.FormatSettings.ShortMonthNames[1];
    for i:=2 to 12 do
      s := s + ', ' + AWorkbook.FormatSettings.ShortMonthNames[i];
    Strings.Add('  ShortMonthNames='+s);
    s := AWorkbook.FormatSettings.LongMonthnames[1];
    for i:=2 to 12 do
      s := s +', ' + AWorkbook.FormatSettings.LongMonthNames[i];
    Strings.Add('  LongMontNames='+s);
    s := AWorkbook.FormatSettings.ShortDayNames[1];
    for i:=2 to 7 do
      s := s + ', ' + AWorkbook.FormatSettings.ShortDayNames[i];
    Strings.Add('  ShortMonthNames='+s);
    s := AWorkbook.FormatSettings.LongDayNames[1];
    for i:=2 to 7 do
      s := s +', ' + AWorkbook.FormatSettings.LongDayNames[i];
    Strings.Add('  LongMontNames='+s);
    Strings.Add('  CurrencyString='+AWorkbook.FormatSettings.CurrencyString);
    Strings.Add('  PosCurrencyFormat='+IntToStr(AWorkbook.FormatSettings.CurrencyFormat));
    Strings.Add('  NegCurrencyFormat='+IntToStr(AWorkbook.FormatSettings.NegCurrFormat));
    Strings.Add('  TwoDigitYearCenturyWindow='+IntToStr(AWorkbook.FormatSettings.TwoDigitYearCenturyWindow));
  end;
end;

procedure TsSpreadsheetInspector.UpdateWorksheet(ASheet: TsWorksheet);
begin
  if ASheet = nil then
  begin
    Strings.Add('First row=');
    Strings.Add('Last row=');
    Strings.Add('First column=');
    Strings.Add('Last column=');
  end else
  begin
    Strings.Add(Format('First row=%d', [Integer(ASheet.GetFirstRowIndex)]));
    Strings.Add(Format('Last row=%d', [ASheet.GetLastRowIndex]));
    Strings.Add(Format('First column=%d', [Integer(ASheet.GetFirstColIndex)]));
    Strings.Add(Format('Last column=%d', [ASheet.GetLastColIndex]));
  end;
end;


end.
