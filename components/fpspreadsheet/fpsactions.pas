unit fpsActions;

interface

uses
  SysUtils, Classes, Controls, ActnList,
  fpspreadsheet, fpspreadsheetctrls;

type
  TsSpreadsheetAction = class(TCustomAction)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetSelection: TsCellRangeArray;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    property Selection: TsCellRangeArray read GetSelection;
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
  published
    property Caption;
    property Enabled;
    property HelpContext;
    property HelpKeyword;
    property HelpType;
    property Hint;
    property ImageIndex;
    property OnExecute;
    property OnHint;
    property OnUpdate;
    property SecondaryShortCuts;
    property ShortCut;
    property Visible;
  end;

  TsWorksheetNameEvent = procedure (Sender: TObject; AWorksheet: TsWorksheet;
    var ASheetName: String) of object;

  { Action for adding a worksheet }
  TsWorksheetAddAction = class(TsWorksheetAction)
  private
    FNameMask: String;
    FOnGetWorksheetName: TsWorksheetNameEvent;
    procedure SetNameMask(const AValue: String);
  protected
    function GetUniqueSheetName: String;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  published
    property NameMask: String read FNameMask write SetNameMask;
    property OnGetWorksheetName: TsWorksheetNameEvent
      read FOnGetWorksheetName write FOnGetWorksheetName;
  end;

  { Action for deleting selected worksheet }
  TsWorksheetDeleteAction = class(TsWorksheetAction)
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  end;

  { Action for renaming selected worksheet }
  TsWorksheetRenameAction = class(TsWorksheetAction)
  private
    FOnGetWorksheetName: TsWorksheetNameEvent;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  published
    property OnGetWorksheetName: TsWorksheetNameEvent
      read FOnGetWorksheetName write FOnGetWorksheetName;
  end;


  { --- Actions related to cell and cell selection formatting--- }

  TsCellFormatAction = class(TsSpreadsheetAction)
  private
    //
  protected
    procedure ApplyFormatToCell(ACell: PCell); virtual;
    procedure ExtractFromCell(ACell: PCell); virtual;
  public
    procedure ExecuteTarget(Target: TObject); override;
    function HandlesTarget(Target: TObject): Boolean; override;
    procedure UpdateTarget(Target: TObject); override;
    property Selection;
    property Worksheet;
  published
    property Caption;
    property Enabled;
    property HelpContext;
    property HelpKeyword;
    property HelpType;
    property Hint;
    property ImageIndex;
    property OnExecute;
    property OnHint;
    property OnUpdate;
    property SecondaryShortCuts;
    property ShortCut;
    property Visible;
  end;


  { TsFontStyleAction }

  TsFontStyleAction = class(TsCellFormatAction)
  private
    FFontStyle: TsFontStyle;
    procedure SetFontStyle(AValue: TsFontStyle);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property FontStyle: TsFontStyle
      read FFontStyle write SetFontStyle;
  end;


  { TsHorAlignmentAction }

  TsHorAlignmentAction = class(TsCellFormatAction)
  private
    FHorAlign: TsHorAlignment;
    procedure SetHorAlign(AValue: TsHorAlignment);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property HorAlignment: TsHorAlignment
      read FHorAlign write SetHorAlign default haDefault;
  end;


  { TsVertAlignmentAction }

  TsVertAlignmentAction = class(TsCellFormatAction)
  private
    FVertAlign: TsVertAlignment;
    procedure SetVertAlign(AValue: TsVertAlignment);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property VertAlignment: TsVertAlignment
      read FVertAlign write SetVertAlign default vaDefault;
  end;


  { TsTextRotationAction }

  TsTextRotationAction = class(TsCellFormatAction)
  private
    FTextRotation: TsTextRotation;
    procedure SetTextRotation(AValue: TsTextRotation);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property TextRotation: TsTextRotation
      read FTextRotation write SetTextRotation default trHorizontal;
  end;


  { TsWordwrapAction }

  TsWordwrapAction = class(TsCellFormatAction)
  private
    function GetWordwrap: Boolean;
    procedure SetWordwrap(AValue: Boolean);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Wordwrap: boolean
      read GetWordwrap write SetWordwrap default false;
  end;


  { TsNumberFormatAction }

  TsNumberFormatAction = class(TsCellFormatAction)
  private
    FNumberFormat: TsNumberFormat;
    FNumberFormatStr: string;
    procedure SetNumberFormat(AValue: TsNumberFormat);
    procedure SetNumberFormatStr(AValue: String);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property NumberFormat: TsNumberFormat
      read FNumberFormat write SetNumberFormat default nfGeneral;
    property NumberFormatString: string
      read FNumberFormatStr write SetNumberFormatStr;
  end;


  { TsDecimalsAction }
  TsDecimalsAction = class(TsCellFormatAction)
  private
    FDecimals: Integer;
    FDelta: Integer;
    procedure SetDecimals(AValue: Integer);
    procedure SetDelta(AValue: Integer);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
    property Decimals: Integer
      read FDecimals write SetDecimals default 2;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Delta: Integer
      read FDelta write SetDelta default +1;
  end;


procedure Register;


implementation

uses
  Dialogs,
  fpsutils;

procedure Register;
begin
  RegisterActions('FPSpreadsheet', [
    // Worksheet-releated actions
    TsWorksheetAddAction, TsWorksheetDeleteAction, TsWorksheetRenameAction,
    // Cell or cell range formatting actions
    TsFontStyleAction,
    TsHorAlignmentAction, TsVertAlignmentAction,
    TsTextRotationAction, TsWordWrapAction,
    TsNumberFormatAction, TsDecimalsAction
  ], nil);
end;


{ TsSpreadsheetAction }

function TsSpreadsheetAction.GetSelection: TsCellRangeArray;
begin
  Result := Worksheet.GetSelection;
end;

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

constructor TsWorksheetAddAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Add';
  Hint := 'Add empty worksheet';
  FNameMask := 'Sheet%d';
end;

{ Helper procedure which creates a default worksheetname by counting a number
  up until it provides in the NameMask a unique worksheet name. }
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
    // Get default name of the new worksheet
    sheetName := GetUniqueSheetName;
    // If available use own procedure to specify new worksheet name
    if Assigned(FOnGetWorksheetName) then
      FOnGetWorksheetName(self, Worksheet, sheetName);
    // Check validity of worksheet name
    if not Workbook.ValidWorksheetName(sheetName) then
    begin
      MessageDlg(Format('"5s" is not a valid worksheet name.', [sheetName]), mtError, [mbOK], 0);
      exit;
    end;
    // Add new worksheet using the worksheet name.
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
  Caption := 'Delete...';
  Hint := 'Delete worksheet';
end;

procedure TsWorksheetDeleteAction.ExecuteTarget(Target: TObject);
begin
  if HandlesTarget(Target) then
  begin
    // Make sure that the last worksheet is not deleted - there must always be
    // at least 1 worksheet.
    if Workbook.GetWorksheetCount = 1 then
    begin
      MessageDlg('The workbook must contain at least 1 worksheet', mtError, [mbOK], 0);
      exit;
    end;

    // Confirmation dialog
    if MessageDlg(
      Format('Do you really want to delete worksheet "%s"?', [Worksheet.Name]),
      mtConfirmation, [mbYes, mbNo], 0) <> mrYes
    then
      exit;

    // Remove the worksheet; the workbookSource takes care of selecting the
    // next worksheet after deletion.
    Workbook.RemoveWorksheet(Worksheet);
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
    // If requested, override input box by own input
    if Assigned(FOnGetWorksheetName) then
      FOnGetWorksheetName(self, Worksheet, s)
    else
      s := InputBox('Rename worksheet', 'New worksheet name', s);
    // No change
    if s = Worksheet.Name then
      exit;
    // Check validity of new worksheet name
    if Workbook.ValidWorksheetName(s) then
      Worksheet.Name := s
    else
      MessageDlg(Format('"%s" is not a valid worksheet name.', [s]), mtError, [mbOK], 0);
  end;
end;


{ TsCellFormatAction }

{ Copies the format item for which the action is responsible to the
  specified cell. Must be overridden by descendants. }
procedure TsCellFormatAction.ApplyFormatToCell(ACell: PCell);
begin
end;

procedure TsCellFormatAction.ExecuteTarget(Target: TObject);
var
  range: Integer;
  r,c: Cardinal;
  sel: TsCellRangeArray;
  cell: PCell;
begin
  if not HandlesTarget(Target) then
    exit;
  sel := GetSelection;
  for range := 0 to High(sel) do
    for r := sel[range].Row1 to sel[range].Row2 do
      for c := sel[range].Col1 to sel[range].Col2 do
      begin
        cell := Worksheet.GetCell(r, c); // Use "GetCell", empty cells will be formatted!
        if cell <> nil then
          ApplyFormatToCell(cell);
      end;
end;

{ Extracts the format item for which the action is responsible from the
  specified cell. Must be overridden by descendants. }
procedure TsCellFormatAction.ExtractFromCell(ACell: PCell);
begin
end;

function TsCellFormatAction.HandlesTarget(Target: TObject): Boolean;
begin
  Result := inherited HandlesTarget(Target) and (Worksheet <> nil) and (Length(GetSelection) > 0);
end;

procedure TsCellFormatAction.UpdateTarget(Target: TObject);
var
  cell: PCell;
begin
  Enabled := inherited Enabled and (Worksheet <> nil) and (Length(GetSelection) > 0);
  if not Enabled then
    exit;

  cell := Worksheet.FindCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  ExtractFromCell(cell);
end;


{ TsFontStyleAction }

constructor TsFontStyleAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  AutoCheck := true;
end;

procedure TsFontStyleAction.ApplyFormatToCell(ACell: PCell);
var
  fnt: TsFont;
  fs: TsFontStyles;
begin
  fnt := Workbook.GetFont(ACell^.FontIndex);
  fs := fnt.Style;
  if Checked then
    Include(fs, FFontStyle)
  else
    Exclude(fs, FFontStyle);
  Worksheet.WriteFontStyle(ACell, fs);
end;

procedure TsFontStyleAction.ExtractFromCell(ACell: PCell);
var
  fnt: TsFont;
  fs: TsFontStyles;
begin
  if (ACell = nil) then
    Checked := false
  else
  if (uffBold in ACell^.UsedFormattingFields) then
    Checked := (FFontStyle = fssBold)
  else
  if (uffFont in ACell^.UsedFormattingFields) then
  begin
    fnt := Workbook.GetFont(ACell^.FontIndex);
    Checked := (FFontStyle in fnt.Style);
  end else
    Checked := false;
end;

procedure TsFontStyleAction.SetFontStyle(AValue: TsFontStyle);
begin
  FFontStyle := AValue;
  case AValue of
    fssBold: begin Caption := 'Bold'; Hint := 'Bold font'; end;
    fssItalic: begin Caption := 'Italic'; Hint := 'Italic font'; end;
    fssUnderline: begin Caption := 'Underline'; Hint := 'Underlines font'; end;
    fssStrikeout: begin Caption := 'Strikeout'; Hint := 'Strike-out font'; end;
  end;
end;


{ TsHorAlignmentAction }

constructor TsHorAlignmentAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411122312;    // Date/time when this was written
  AutoCheck := true;
end;

procedure TsHorAlignmentAction.ApplyFormatToCell(ACell: PCell);
begin
  if Checked then
    Worksheet.WriteHorAlignment(ACell, FHorAlign)
  else
    Worksheet.WriteHorAlignment(ACell, haDefault);
end;

procedure TsHorAlignmentAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffHorAlign in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := ACell^.HorAlignment = FHorAlign;
end;

procedure TsHorAlignmentAction.SetHorAlign(AValue: TsHorAlignment);
begin
  FHorAlign := AValue;
  case FHorAlign of
    haLeft   : begin Caption := 'Left'; Hint := 'Left-aligned text'; end;
    haCenter : begin Caption := 'Center'; Hint := 'Centered text'; end;
    haRight  : begin Caption := 'Right'; Hint := 'Right-aligned text'; end;
    haDefault: begin Caption := 'Default'; Hint := 'Default horizontal text alignment'; end;
  end;
end;


{ TsVertAlignmentAction }

constructor TsVertAlignmentAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411122322;    // Date/time when this was written
  AutoCheck := true;
end;

procedure TsVertAlignmentAction.ApplyFormatToCell(ACell: PCell);
begin
  if Checked then
    Worksheet.WriteVertAlignment(ACell, FVertAlign)
  else
    Worksheet.WriteVertAlignment(ACell, vaDefault);
end;

procedure TsVertAlignmentAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffVertAlign in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := ACell^.VertAlignment = FVertAlign;
end;

procedure TsVertAlignmentAction.SetVertAlign(AValue: TsVertAlignment);
begin
  FVertAlign := AValue;
  case FVertAlign of
    vaTop    : begin Caption := 'Top'; Hint := 'Top-aligned text'; end;
    vaCenter : begin Caption := 'Center'; Hint := 'Vertically centered text'; end;
    vaBottom : begin Caption := 'Bottom'; Hint := 'Bottom-aligned text'; end;
    vaDefault: begin Caption := 'Default'; Hint := 'Default vertical text alignment'; end;
  end;
end;


{ TsTextRotationAction }

constructor TsTextRotationAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411141108;    // Date/time when this was written
  AutoCheck := true;
end;

procedure TsTextRotationAction.ApplyFormatToCell(ACell: PCell);
begin
  if Checked then
    Worksheet.WriteTextRotation(ACell, FTextRotation)
  else
    Worksheet.WriteTextRotation(ACell, trHorizontal);
end;

procedure TsTextRotationAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffTextRotation in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := ACell^.TextRotation = FTextRotation;
end;

procedure TsTextRotationAction.SetTextRotation(AValue: TsTextRotation);
begin
  FTextRotation := AValue;
  case FTextRotation of
    trHorizontal:
      begin Caption := 'Horizontal'; Hint := 'Horizontal text'; end;
    rt90DegreeClockwiseRotation:
      begin Caption := '90° clockwise'; Hint := '90° clockwise rotated text'; end;
    rt90DegreeCounterClockwiseRotation:
      begin Caption := '90° counter-clockwise'; Hint := '90° counter-clockwise rotated text'; end;
    rtStacked:
      begin Caption := 'Stacked'; Hint := 'Vertically stacked horizontal letters'; end;
  end;
end;


{ TsWordwrapAction }

constructor TsWordwrapAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  AutoCheck := true;
  Caption := 'Word-wrap';
  Hint := 'Word-wrapped text';
end;

procedure TsWordwrapAction.ApplyFormatToCell(ACell: PCell);
begin
  Worksheet.WriteWordwrap(ACell, Checked);
end;

procedure TsWordwrapAction.ExtractFromCell(ACell: PCell);
begin
  Checked := (ACell <> nil) and (uffWordwrap in ACell^.UsedFormattingFields);
end;

function TsWordwrapAction.GetWordwrap: Boolean;
begin
  Result := Checked;
end;

procedure TsWordwrapAction.SetWordwrap(AValue: Boolean);
begin
  Checked := AValue;
end;


{ TsNumberFormatAction }

constructor TsNumberFormatAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411141258;    // Date/time when this was written
  AutoCheck := true;
  Caption := 'Number format';
  Hint := 'Number format';
end;

procedure TsNumberFormatAction.ApplyFormatToCell(ACell: PCell);
var
  nf: TsNumberFormat;
  nfstr: String;
begin
  if Checked then
  begin
    nf := FNumberFormat;
    nfstr := FNumberFormatStr;
  end else
  begin
    nf := nfGeneral;
    nfstr := '';
  end;
  if IsDateTimeFormat(nf) then
    Worksheet.WriteDateTimeFormat(ACell, nf, nfstr)
  else
    Worksheet.WriteNumberFormat(ACell, nf, nfstr);
end;

procedure TsNumberFormatAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffNumberFormat in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := (ACell^.NumberFormat = FNumberFormat)
      and (ACell^.NumberFormatStr = FNumberFormatStr);
end;

procedure TsNumberFormatAction.SetNumberFormat(AValue: TsNumberFormat);
begin
  FNumberFormat := AValue;
  case FNumberFormat of
    nfGeneral:
      begin Caption := 'General'; Hint := 'General format'; end;
    nfFixed:
      begin Caption := 'Fixed'; Hint := 'Fixed decimals format'; end;
    nfFixedTh:
      begin Caption := 'Fixed w/thousand separator'; Hint := 'Fixed decimal count with thousand separator'; end;
    nfExp:
      begin Caption := 'Exponential'; Hint := 'Exponential format'; end;
    nfPercentage:
      begin Caption := 'Percent'; Hint := 'Percent format'; end;
    nfCurrency:
      begin Caption := 'Currency'; Hint := 'Currency format'; end;
    nfCurrencyRed:
      begin Caption := 'Currency (red)'; Hint := 'Currency format (negative values in red)'; end;
    nfShortDateTime:
      begin Caption := 'Date/time'; Hint := 'Date and time'; end;
    nfShortDate:
      begin Caption := 'Short date'; Hint := 'Short date format'; end;
    nfLongDate:
      begin Caption := 'Long date'; Hint := 'Long date format'; end;
    nfShortTime:
      begin Caption := 'Short time'; Hint := 'Short time format'; end;
    nfLongTime:
      begin Caption := 'Long time'; Hint := 'Long time foramt'; end;
    nfShortTimeAM:
      begin Caption := 'Short time AM/PM'; Hint := 'Short 12-hour time format'; end;
    nfLongTimeAM:
      begin Caption := 'Long time AM/PM'; Hint := 'Long 12-hour time format'; end;
    nfTimeInterval:
      begin Caption := 'Time interval'; Hint := 'Time interval format'; end;
    nfCustom:
      begin Caption := 'Custom'; Hint := 'User-defined custom format'; end;
  end;
end;

procedure TsNumberFormatAction.SetNumberFormatStr(AValue: String);
begin
  FNumberFormatStr := AValue;
end;


{ TsDecimalsAction }

constructor TsDecimalsAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Decimals';
  Hint := 'Decimal places';
  FDelta := +1;
end;

procedure TsDecimalsAction.ApplyFormatToCell(ACell: PCell);
var
  decs: Integer;
  currSym: String;
begin
  if not (uffNumberFormat in ACell^.UsedFormattingFields) or
         (ACell^.NumberFormat = nfGeneral)
  then
    Worksheet.WriteNumberFormat(ACell, nfFixed, '0')
  else
  if IsDateTimeFormat(ACell^.NumberFormat) then
    exit
  else
  begin
    decs := Decimals + FDelta;
    if decs < 0 then decs := 0;
    Worksheet.WriteDecimals(ACell, decs);
  end;
end;

procedure TsDecimalsAction.ExtractFromCell(ACell: PCell);
var
  csym: String;
  decs: Byte;
begin
  decs := 2;
  if (ACell <> nil) and (uffNumberFormat in ACell^.UsedFormattingFields) then
    Worksheet.GetNumberFormatAttributes(ACell, decs, csym);
  Decimals := decs
end;

procedure TsDecimalsAction.SetDecimals(AValue: Integer);
begin
  FDecimals := AValue;
  if FDecimals < 0 then FDecimals := 0;
end;

procedure TsDecimalsAction.SetDelta(AValue: Integer);
begin
  FDelta := AValue;
  if FDelta > 0 then
    Hint := 'More decimal places'
  else
    Hint := 'Less decimal places';
end;

end.
