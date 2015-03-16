unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls, Grids, Buttons, Menus, ActnList, StdActns,
  fpsTypes, fpSpreadsheet, fpSpreadsheetCtrls, fpSpreadsheetGrid, fpsActions;

type

  { TForm1 }

  TForm1 = class(TForm)
    ActionList: TActionList;
    BtnLoad: TButton;
    CbLoader: TComboBox;
    AcFileExit: TFileExit;
    ImageList1: TImageList;
    Label1: TLabel;
    MainMenu: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MnuNumFormatTimeInterval: TMenuItem;
    MnuNumFormatLongTimeAM: TMenuItem;
    MnuNumFormatShortTimeAM: TMenuItem;
    MnuNumFormatLongTime: TMenuItem;
    MnuNumFormatShortTime: TMenuItem;
    MnuNumFormatShortDateTime: TMenuItem;
    MenuItem2: TMenuItem;
    MnuNumFormatLongDate: TMenuItem;
    MnuNumFormatShortDate: TMenuItem;
    MnuNumFormatCurrency: TMenuItem;
    MnuNumFormatCurrencyRed: TMenuItem;
    MenuItem4: TMenuItem;
    MnuNumberFormat: TMenuItem;
    MnuNumFormatGeneral: TMenuItem;
    MenuItem3: TMenuItem;
    MnuNumFormatFixed: TMenuItem;
    MnuNumFormatFixedTh: TMenuItem;
    MnuNumFormatExp: TMenuItem;
    MnuNumFormatPercentage: TMenuItem;
    MenuItem8: TMenuItem;
    MnuTextRotHor: TMenuItem;
    MnuTextRot90CW: TMenuItem;
    MnuTextRot90CCW: TMenuItem;
    MnuTextRotStacked: TMenuItem;
    MnuTextRotation: TMenuItem;
    MnuWordwrap: TMenuItem;
    MnuVertAlignTop: TMenuItem;
    MnuVertAlignCenter: TMenuItem;
    MnuVertAlignBottom: TMenuItem;
    MnuVertAlignment: TMenuItem;
    MnuBOld: TMenuItem;
    MnuItalic: TMenuItem;
    MnuUnderline: TMenuItem;
    MnuStrikeout: TMenuItem;
    MnuFontStyle: TMenuItem;
    MnuHorAlignRight: TMenuItem;
    MnuHorAlignCenter: TMenuItem;
    MnuHorAlignLeft: TMenuItem;
    MnuCells: TMenuItem;
    MnuHorAlignment: TMenuItem;
    MnuFileExit: TMenuItem;
    MnuRenameWorksheet: TMenuItem;
    MnuDeleteWorksheet: TMenuItem;
    MnuAddWorksheet: TMenuItem;
    MnuWorksheets: TMenuItem;
    MnuEdit: TMenuItem;
    MnuFile: TMenuItem;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Splitter1: TSplitter;
    InspectorTabControl: TTabControl;
    ToolBar1: TToolBar;
    TbBold: TToolButton;
    TbItalic: TToolButton;
    TbUnderline: TToolButton;
    TbStrikeout: TToolButton;
    ToolButton1: TToolButton;
    TbHorAlignLeft: TToolButton;
    TbHorAlignCenter: TToolButton;
    TbHorAlignRight: TToolButton;
    ToolButton2: TToolButton;
    TbVertAlignTop: TToolButton;
    TbVertAlignCenter: TToolButton;
    TbVertAlignBottom: TToolButton;
    ToolButton6: TToolButton;
    procedure BtnLoadClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure InspectorTabControlChange(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
  private
    { private declarations }
    WorkbookSource: TsWorkbookSource;
    WorkbookTabControl: TsWorkbookTabControl;
    WorksheetGrid: TsWorksheetGrid;
    CellIndicator: TsCellIndicator;
    CellEdit: TsCellEdit;
    Inspector: TsSpreadsheetInspector;
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.BtnLoadClick(Sender: TObject);
var
  fn: ansistring;
begin
  if OpenDialog.Execute then begin
    WorkbookSource.AutodetectFormat := false;
    case OpenDialog.FilterIndex of
      1: WorkbookSource.AutoDetectFormat := true;      // All spreadsheet files
      2: WorkbookSource.AutoDetectFormat := true;      // All Excel files
      3: WorkbookSource.FileFormat := sfOOXML;         // Excel 2007+
      4: WorkbookSource.FileFormat := sfExcel8;        // Excel 97-2003
      5: WorkbookSource.FileFormat := sfExcel5;        // Excel 5.0
      6: WorkbookSource.FileFormat := sfExcel2;        // Excel 2.1
      7: WorkbookSource.FileFormat := sfOpenDocument;  // Open/LibreOffice
      8: WOrkbookSource.FileFormat := sfCSV;           // Text files
      9: WorkbookSource.FileFormat := sfWikiTable_WikiMedia;  // wiki tables
    end;

    // There are 3 possibilities to open a file:
    fn := UTF8ToAnsi(OpenDialog.Filename);
    case CbLoader.ItemIndex of
      0:  if WorkbookSource.AutodetectFormat then
            WorkbookSource.Workbook.ReadFromFile(fn)
          else
            WorkbookSource.Workbook.ReadFromFile(fn, WorkbookSource.FileFormat);
      1:  WorkbookSource.FileName := fn;    // this loads the file
      2:  if WorkbookSource.AutodetectFormat then
            WorksheetGrid.LoadFromSpreadsheetFile(fn)
          else
            WorksheetGrid.LoadFromSpreadsheetFile(fn, WorkbookSource.FileFormat);
    end;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  actn: TCustomAction;
begin
  WorkbookSource := TsWorkbookSource.Create(self);
  with WorkbookSource do
  begin
    Options := [boReadFormulas, boAutoCalc];
  end;

  WorkbookTabControl := TsWorkbookTabControl.Create(self);
  with WorkbookTabControl do
  begin
    Parent := self;
    Align := alClient;
    WorkbookSource := Self.WorkbookSource;
  end;

  WorksheetGrid := TsWorksheetGrid.Create(self);
  with WorksheetGrid do
  begin
    Parent := WorkbookTabControl;
    Align := alClient;
    Options := Options + [goEditing, goRowSizing, goColSizing];
    TextOverflow := true;
    WorkbookSource := Self.WorkbookSource;
  end;

  CellIndicator := TsCellIndicator.Create(self);
  with CellIndicator do
  begin
    Parent := Panel1;
    Left := BtnLoad.Left + BtnLoad.Width + 24;
    Top := BtnLoad.Top;
    WorkbookSource := Self.WorkbookSource;
  end;

  CellEdit := TsCellEdit.Create(self);
  with CellEdit do
  begin
    Parent := Panel1;
    Left := CellIndicator.Left + CellIndicator.Width + 24;
    Top := CellIndicator.Top;
    WorkbookSource := Self.WorkbookSource;
  end;

  Inspector := TsSpreadsheetInspector.Create(self);
  with Inspector do
  begin
    Parent := InspectorTabControl;
    Align := alClient;
    WorkbookSource := Self.WorkbookSource;
  end;

  actn := TsWorksheetAddAction.Create(self);
  with TsWorksheetAddAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Add';
    Hint := 'Add worksheet';
    WorkbookSource := Self.WorkbookSource;
  end;
  MnuAddWorksheet.Action := actn;

  actn := TsWorksheetDeleteAction.Create(self);
  with TsWorksheetDeleteAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Delete...';
    Hint := 'Delete worksheet';
    WorkbookSource := Self.WorkbookSource;
  end;
  MnuDeleteWorksheet.Action := actn;

  actn := TsWorksheetRenameAction.Create(self);
  with TsWorksheetRenameAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Rename...';
    Hint := 'Rename worksheet';
    WorkbookSource := Self.WorkbookSource;
  end;
  MnuRenameWorksheet.Action := actn;

  { Font names }
  with TsCellCombobox.Create(self) do
  begin
//  with TsFontnameCombobox.Create(self) do begin
    Parent := Toolbar1;
    WorkbookSource := Self.WorkbookSource;
    Width := 160;
    CellFormatItem := cfiFontName;
  end;

  { Font size }
  with TsCellCombobox.Create(self) do
  begin
    Parent := Toolbar1;
    WorkbookSource := Self.WorkbookSource;
    Width := 50;
    CellFormatItem := cfiFontSize;
  end;

  { Font color }
  with TsCellCombobox.Create(self) do
  begin
    Parent := Toolbar1;
    WorkbookSource := Self.WorkbookSource;
    Width := 120;
    CellFormatItem := cfiFontColor;
  end;

  { Font styles }
  actn := TsFontStyleAction.Create(self);
  with TsFontStyleAction(actn) do
  begin
    ActionList := Self.ActionList;
    Caption := 'Bold';
    Hint := 'Bold';
    ImageIndex := 0;
    WorkbookSource := Self.WorkbookSource;
    FontStyle := fssBold;
  end;
  MnuBold.Action := actn;
  tbBold.Action := actn;

  actn := TsFontStyleAction.Create(self);
  with TsFontStyleAction(actn) do
  begin
    ActionList := Self.ActionList;
    Caption := 'Italic';
    Hint := 'Italic';
    ImageIndex := 1;
    WorkbookSource := Self.WorkbookSource;
    FontStyle := fssItalic;
  end;
  MnuItalic.Action := actn;
  TbItalic.Action := actn;

  actn := TsFontStyleAction.Create(self);
  with TsFontStyleAction(actn) do
  begin
    ActionList := Self.ActionList;
    Caption := 'Underline';
    Hint := 'Underline';
    ImageIndex := 2;
    WorkbookSource := Self.WorkbookSource;
    FontStyle := fssUnderline;
  end;
  MnuUnderline.Action := actn;
  TbUnderline.Action := actn;

  actn := TsFontStyleAction.Create(self);
  with TsFontStyleAction(actn) do
  begin
    ActionList := Self.ActionList;
    Caption := 'Strikeout';
    Hint := 'Strikeout';
    ImageIndex := 3;
    WorkbookSource := Self.WorkbookSource;
    FontStyle := fssStrikeout;
  end;
  MnuStrikeout.Action := actn;
  TbStrikeout.Action := actn;

  { Horizontal alignments }
  actn := TsHorAlignmentAction.Create(self);
  with TsHorAlignmentAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Left';
    Hint := 'Left-aligned';
    ImageIndex := 4;
    WorkbookSource := Self.WorkbookSource;
    HorAlignment := haLeft;
  end;
  MnuHorAlignLeft.Action := actn;
  TbHorAlignLeft.Action := actn;

  actn := TsHorAlignmentAction.Create(self);
  with TsHorAlignmentAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Center';
    Hint := 'Centered';
    ImageIndex := 5;
    WorkbookSource := Self.WorkbookSource;
    HorAlignment := haCenter;
  end;
  MnuHorAlignCenter.Action := actn;
  TbHorAlignCenter.Action := actn;

  actn := TsHorAlignmentAction.Create(self);
  with TsHorAlignmentAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Right';
    Hint := 'Right-aligned';
    ImageIndex := 6;
    WorkbookSource := Self.WorkbookSource;
    HorAlignment := haRight;
  end;
  MnuHorAlignRight.Action := actn;
  TbHorAlignRight.Action := Actn;

  { Vertical alignments }
  actn := TsVertAlignmentAction.Create(self);
  with TsVertAlignmentAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Top';
    Hint := 'Top-aligned';
    ImageIndex := 7;
    WorkbookSource := Self.WorkbookSource;
    VertAlignment := vaTop;
  end;
  MnuVertAlignTop.Action := actn;
  TbVertAlignTop.Action := actn;

  actn := TsVertAlignmentAction.Create(self);
  with TsVertAlignmentAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Middle';
    Hint := 'Middle';
    ImageIndex := 8;
    WorkbookSource := Self.WorkbookSource;
    VertAlignment := vaCenter;
  end;
  MnuVertAlignCenter.Action := actn;
  TbVertAlignCenter.Action := actn;

  actn := TsVertAlignmentAction.Create(self);
  with TsVertAlignmentAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Bottom';
    Hint := 'Bottom-aligned';
    ImageIndex := 9;
    WorkbookSource := Self.WorkbookSource;
    VertAlignment := vaBottom;
  end;
  MnuVertAlignBottom.Action := actn;
  TbVertAlignBottom.Action := Actn;

  { Text rotation }
  actn := TsTextRotationAction.Create(self);
  with TsTextRotationAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Horizontal';
    Hint := 'Horizontal text';
    WorkbookSource := Self.WorkbookSource;
    TextRotation := trHorizontal;
  end;
  MnuTextRotHor.Action := actn;

  actn := TsTextRotationAction.Create(self);
  with TsTextRotationAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := '90° clockwise rotation';
    Hint := '90° clockwise rotated text';
    WorkbookSource := Self.WorkbookSource;
    TextRotation := rt90DegreeClockwiseRotation;
  end;
  MnuTextRot90CW.Action := actn;

  actn := TsTextRotationAction.Create(self);
  with TsTextRotationAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := '90° couner-clockwise rotation';
    Hint := '90° counter-clockwise rotated text';
    WorkbookSource := Self.WorkbookSource;
    TextRotation := rt90DegreeCounterClockwiseRotation;
  end;
  MnuTextRot90CCW.Action := actn;

  actn := TsTextRotationAction.Create(self);
  with TsTextRotationAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Stacked';
    Hint := 'Stacked text';
    WorkbookSource := Self.WorkbookSource;
    TextRotation := rtStacked;
  end;
  MnuTextRotStacked.Action := actn;

  { Word wrap }
  actn := TsWordwrapAction.Create(self);
  with TsWordwrapAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Word-wrap';
    Hint := 'Word-wrapped text';
    WorkbookSource := Self.WorkbookSource;
    Wordwrap := false;
  end;
  MnuWordwrap.Action := actn;

  { Number format }
  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'General';
    Hint := 'General';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfGeneral;
  end;
  MnuNumFormatGeneral.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Fixed';
    Hint := 'Fixed decimals';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfFixed;
  end;
  MnuNumFormatFixed.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Fixed w/Thousand separator';
    Hint := 'Fixed decimals with thousand separator';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfFixedTh;
  end;
  MnuNumFormatFixedTh.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Exponential';
    Hint := 'Exponential';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfExp;
  end;
  MnuNumFormatExp.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Percentage';
    Hint := 'Percentage';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfPercentage;
  end;
  MnuNumFormatPercentage.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Currency';
    Hint := 'Currency';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfCurrency;
  end;
  MnuNumFormatCurrency.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Currency (red)';
    Hint := 'Currency (negative values in red)';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfCurrencyRed;
  end;
  MnuNumFormatCurrencyRed.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Date and time';
    Hint := 'Date and time';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfShortDateTime;
  end;
  MnuNumFormatShortDateTime.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Long date';
    Hint := 'Long date';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfLongDate;
  end;
  MnuNumFormatLongDate.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    WorkbookSource := Self.WorkbookSource;
    Caption := 'Short date';
    Hint := 'Short date';
    NumberFormat := nfShortDate;
  end;
  MnuNumFormatShortDate.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Long time';
    Hint := 'Long time';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfLongTime;
  end;
  MnuNumFormatLongTime.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Short time';
    Hint := 'Short time';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfShortTime;
  end;
  MnuNumFormatShortTime.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    Caption := 'Long time AM/PM';
    Hint := 'Long time with AM/PM';
    WorkbookSource := Self.WorkbookSource;
    NumberFormat := nfLongTimeAM;
  end;
  MnuNumFormatLongTimeAM.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    WorkbookSource := Self.WorkbookSource;
    Caption := 'Short time AM/PM';
    Hint := 'Short time with AM/PM';
    NumberFormat := nfShortTimeAM;
  end;
  MnuNumFormatShortTimeAM.Action := actn;

  actn := TsNumberFormatAction.Create(self);
  with TsNumberFormatAction(actn) do
  begin
    ActionList := self.ActionList;
    WorkbookSource := Self.WorkbookSource;
    Caption := 'Time interval';
    Hint := 'Time interval';
    NumberFormat := nfTimeInterval;
  end;
  MnuNumFormatTimeInterval.Action := actn;
end;

procedure TForm1.InspectorTabControlChange(Sender: TObject);
begin
  Inspector.Mode := TsInspectorMode(InspectorTabControl.TabIndex);
end;

procedure TForm1.SpeedButton1Click(Sender: TObject);
// The same effect is obtained by using the built-in TsWorksheetAddAction.
var
  sheetname: String;
  i: Integer;
begin
  i := WorkbookSource.Workbook.GetWorksheetCount;
  repeat
    inc(i);
    sheetName := Format('Sheet %d', [i]);
  until (WorkbookSource.Workbook.GetWorksheetByName(sheetname) = nil);
  WorkbookSource.Workbook.AddWorksheet(sheetName);
end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
// The same effect is obtained by using the built-in TsWorksheetDeleteAction.
begin
  if WorkbookSource.Workbook.GetWorksheetCount = 1 then
    MessageDlg('There must be a least 1 worksheet.', mtError, [mbOK], 0)
  else
  if MessageDlg('Do you really want to delete this worksheet?', mtConfirmation,
    [mbYes, mbNo], 0) = mrYes
  then
    WorkbookSource.Workbook.RemoveWorksheet(WorkbookSource.Worksheet);
end;

procedure TForm1.SpeedButton3Click(Sender: TObject);
// The same effect can be obtained by using the built-in TsWorksheetRenameAction
var
  s: String;
begin
  s := WorkbookSource.Worksheet.Name;
  if InputQuery('Edit worksheet name', 'New name', s) then
  begin
    if WorkbookSource.Workbook.ValidWorksheetName(s) then
      WorkbookSource.Worksheet.Name := s
    else
      MessageDlg('Invalid worksheet name.', mtError, [mbOK], 0);
  end;
end;

end.

