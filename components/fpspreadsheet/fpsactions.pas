unit fpsActions;

interface

uses
  SysUtils, Classes, Controls, ActnList,
  fpspreadsheet, fpspreadsheetctrls;

type
  TsSpreadsheetAction = class(TAction)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    property Worksheet: TsWorksheet read GetWorksheet;
  public
    function HandlesTarget(Target: TObject): Boolean; override;
    procedure UpdateTarget(Target: TObject); override;
    property Workbook: TsWorkbook read GetWorkbook;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write FWorkbookSource;
  end;

  { --- Actions related to worksheets --- }

  TsWorksheetAction = class(TsSpreadsheetAction)
  private
  public
    function HandlesTarget(Target: TObject): Boolean; override;
    procedure UpdateTarget(Target: TObject); override;
    property Worksheet;
  end;

  { Action for adding a worksheet }
  TsWorksheetAddAction = class(TsWorksheetAction)
  private
    FNameMask: String;
    procedure SetNameMask(const AValue: String);
  protected
    function GetUniqueSheetName: String;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  published
    property NameMask: String read FNameMask write SetNameMask;
  end;

  { Action for deleting selected worksheet }
  TsWorksheetDeleteAction = class(TsWorksheetAction)
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  end;

  { Action for renaming selected worksheet }
  TsWorksheetRenameAction = class(TsWorksheetAction)
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  end;

  procedure Register;

implementation

uses
  Dialogs;

procedure Register;
begin
  RegisterActions('FPSpreadsheet', [
    TsWorksheetAddAction, TsWorksheetDeleteAction, TsWorksheetRenameAction
  ], nil);
end;


{ TsSpreadsheetAction }

function TsSpreadsheetAction.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

function TsSpreadsheetAction.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

function TsSpreadsheetAction.HandlesTarget(Target: TObject): Boolean;
begin
  Result := (Target <> nil) and (Target = FWorkbookSource);
end;

procedure TsSpreadsheetAction.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    FWorkbookSource := nil;
end;

procedure TsSpreadsheetAction.UpdateTarget(Target: TObject);
begin
  Enabled := HandlesTarget(Target);
end;


{ TsWorksheetAction }

function TsWorksheetAction.HandlesTarget(Target: TObject): Boolean;
begin
  Result := inherited HandlesTarget(Target) and (Worksheet <> nil);
end;

procedure TsWorksheetAction.UpdateTarget(Target: TObject);
begin
  Enabled := inherited Enabled and (Worksheet <> nil);
end;


{ TsWorksheetAddAction }

constructor TsWOrksheetAddAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Add';
  Hint := 'Add empty worksheet';
  FNameMask := 'Sheet%d';
end;

function TsWorksheetAddAction.GetUniqueSheetName: String;
var
  i: Integer;
begin
  Result := '';
  if Workbook = nil then
    exit;

  i := 0;
  repeat
    inc(i);
    Result := Format(FNameMask, [i]);
  until Workbook.GetWorksheetByName(Result) = nil
end;

procedure TsWorksheetAddAction.ExecuteTarget(Target: TObject);
var
  sheetName: String;
begin
  if HandlesTarget(Target) then
  begin
    sheetName := GetUniqueSheetName;
    Workbook.AddWorksheet(sheetName);
  end;
end;

procedure TsWorksheetAddAction.SetNameMask(const AValue: String);
begin
  if AValue = FNameMask then
    exit;

  if pos('%d', AValue) = 0 then
    raise Exception.Create('Worksheet name mask must contain a %d place-holder.');

  FNameMask := AValue;
end;


{ TsWorksheetDeleteAction }

constructor TsWorksheetDeleteAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Delete';
  Hint := 'Delete worksheet';
end;

procedure TsWorksheetDeleteAction.ExecuteTarget(Target: TObject);
begin
  if HandlesTarget(Target) then
  begin
    if Workbook.GetWorksheetCount = 1 then
    begin
      MessageDlg('The workbook must contain at least 1 worksheet', mtError, [mbOK], 0);
      exit;
    end;

    if MessageDlg(
      Format('Do you really want to delete worksheet "%s"?', [Worksheet.Name]),
      mtConfirmation, [mbYes, mbNo], 0) <> mrYes
    then
      exit;

    Workbook.RemoveWorksheet(Worksheet);
    // The workbooksource takes care of selecting the next worksheet
  end;
end;


{ TsWorksheetRenameAction }

constructor TsWorksheetRenameAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Rename...';
  Hint := 'Rename worksheet';
end;

procedure TsWorksheetRenameAction.ExecuteTarget(Target: TObject);
var
  s: String;
begin
  if HandlesTarget(Target) then
  begin
    s := Worksheet.Name;
    if InputQuery('Rename worksheet', 'New worksheet name', s) then
      Worksheet.Name := s;
  end;
end;


                        (*
  { TsSpreadsheetAction }

  TsSpreadsheetAction = class(TAction)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookLink(AValue: TsWorkbookSource);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure UpdateCell; virtual;
    procedure UpdateWorkbook; virtual;
    procedure UpdateWorksheet; virtual;
  public
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    property Workbook: TsWorkbook read GetWorkbook;
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property WorkbookLink: TsWorkbookSource read FWorkbookSource write SetWorkbookLink;
  end;

  {TsWorksheetNavigateAction}
  TsWorksheetNavigateAction = class(TsSpreadsheetAction)
  public
    function Update: Boolean; override;
  end;

  {TsNextWorksheetAction}
  TsNextWorksheetAction = class(TAction)
  public
    function Execute: Boolean; override;
  end;

  {TsPreviosWorksheetAction}
  TsPreviousWorksheetAction = class(TAction)
  public
    function Execute: Boolean; override;
  end;                *)


  end.
