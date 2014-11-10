unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls, Grids, Buttons, Menus, ActnList,
  fpspreadsheet, fpspreadsheetctrls, fpSpreadsheetGrid, fpsActions;

type

  { TForm1 }

  TForm1 = class(TForm)
    ActionList: TActionList;
    BtnLoad: TButton;
    CbLoader: TComboBox;
    Label1: TLabel;
    MainMenu: TMainMenu;
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
    case CbLoader.ItemIndex of
      0:  if WorkbookSource.AutodetectFormat then
            WorkbookSource.Workbook.ReadFromFile(OpenDialog.FileName)
          else
            WorkbookSource.Workbook.ReadFromFile(OpenDialog.Filename, WorkbookSource.FileFormat);
      1:  WorkbookSource.FileName := OpenDialog.FileName;    // this loads the file
      2:  if WorkbookSource.AutodetectFormat then
            WorksheetGrid.LoadFromSpreadsheetFile(OpenDialog.FileName)
          else
            WorksheetGrid.LoadFromSpreadsheetFile(OpenDialog.FileName, WorkbookSource.FileFormat);
    end;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  actn: TAction;
begin
  WorkbookSource := TsWorkbookSource.Create(self);
  with WorkbookSource do begin
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
    Options := Options + [goEditing];
    TextOverflow := true;
    WorkbookSource := Self.WorkbookSource;
  end;

  CellIndicator := TsCellIndicator.Create(self);
  with CellIndicator do begin
    Parent := Panel1;
    Left := BtnLoad.Left + BtnLoad.Width + 24;
    Top := BtnLoad.Top;
    WorkbookSource := Self.WorkbookSource;
  end;

  CellEdit := TsCellEdit.Create(self);
  with CellEdit do begin
    Parent := Panel1;
    Left := CellIndicator.Left + CellIndicator.Width + 24;
    Top := CellIndicator.Top;
    WorkbookSource := Self.WorkbookSource;
  end;

  Inspector := TsSpreadsheetInspector.Create(self);
  with Inspector do begin
    Parent := InspectorTabControl;
    Align := alClient;
    WorkbookSource := Self.WorkbookSource;
  end;

  actn := TsWorksheetAddAction.Create(self);
  with TsWorksheetAddAction(actn) do begin
    ActionList := self.ActionList;
    WorkbookSource := Self.WorkbookSource;
  end;
  MnuAddWorksheet.Action := actn;

  actn := TsWorksheetDeleteAction.Create(self);
  with TsWorksheetDeleteAction(actn) do begin
    ActionList := self.ActionList;
    WOrkbookSource := Self.WorkbookSource;
  end;
  MnuDeleteWorksheet.Action := actn;

end;

procedure TForm1.InspectorTabControlChange(Sender: TObject);
begin
  Inspector.Mode := TsInspectorMode(InspectorTabControl.TabIndex);
end;

procedure TForm1.SpeedButton1Click(Sender: TObject);
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

