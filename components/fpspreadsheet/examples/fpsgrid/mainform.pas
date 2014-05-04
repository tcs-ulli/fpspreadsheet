unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Spin,
  fpspreadsheetgrid, fpspreadsheet, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    AcOpen: TAction;
    AcSaveAs: TAction;
    AcQuit: TAction;
    ActionList1: TActionList;
    btnPopulateGrid: TButton;
    CbShowHeaders: TCheckBox;
    CbShowGridLines: TCheckBox;
    EdFrozenRows: TSpinEdit;
    ImageList1: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    mnuFile: TMenuItem;
    mnuOpen: TMenuItem;
    mnuQuit: TMenuItem;
    mnuSaveAs: TMenuItem;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    Panel1: TPanel;
    SaveDialog1: TSaveDialog;
    EdFrozenCols: TSpinEdit;
    sWorksheetGrid1: TsWorksheetGrid;
    TabSheet1: TTabSheet;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton5: TToolButton;
    procedure btnPopulateGridClick(Sender: TObject);
    procedure CbShowHeadersClick(Sender: TObject);
    procedure CbShowGridLinesClick(Sender: TObject);
    procedure acOpenExecute(Sender: TObject);
    procedure acQuitExecute(Sender: TObject);
    procedure acSaveAsExecute(Sender: TObject);
    procedure EdFrozenColsChange(Sender: TObject);
    procedure EdFrozenRowsChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
  private
    { private declarations }
    procedure LoadFile(const AFileName: String);
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

uses
  Grids, fpcanvas;

{ TForm1 }

procedure TForm1.btnPopulateGridClick(Sender: TObject);
// Populate grid with some demo data
var
  lCell: PCell;
begin
  // create a cell (2,2) if not yet available
  lCell := sWorksheetGrid1.Worksheet.GetCell(2, 2);
  sWorksheetGrid1.Worksheet.WriteUTF8Text(2, 2, 'Algo');
  sWorksheetGrid1.Invalidate;
end;

procedure TForm1.CbShowHeadersClick(Sender: TObject);
begin
  sWorksheetGrid1.ShowHeaders := CbShowHeaders.Checked;
end;

procedure TForm1.CbShowGridLinesClick(Sender: TObject);
begin
  sWorksheetGrid1.ShowGridLines := CbShowGridLines.Checked;
end;

procedure TForm1.acOpenExecute(Sender: TObject);
begin
  if OpenDialog1.Execute then
    LoadFile(OpenDialog1.FileName);
end;

procedure TForm1.acQuitExecute(Sender: TObject);
begin
  Close;
end;

procedure TForm1.acSaveAsExecute(Sender: TObject);
// Saves sheet in grid to file, overwriting existing file
var
  lWorkBook: TsWorkbook;
  lWorkSheet:TsWorksheet;
begin
  ShowMessage('Not implemented...');
  exit;

  if SaveDialog1.Execute then
  begin
    lWorkBook := TsWorkBook.Create;
    lWorkSheet := lWorkBook.AddWorksheet('Sheet1');
    try
      sWorksheetGrid1.SaveToWorksheet(lWorkSheet);
      lWorkBook.WriteToFile(SaveDialog1.FileName,true);
    finally
      lWorkBook.Free;
    end;
  end;
end;

procedure TForm1.EdFrozenColsChange(Sender: TObject);
begin
  sWorksheetGrid1.FrozenCols := EdFrozenCols.Value;
end;

procedure TForm1.EdFrozenRowsChange(Sender: TObject);
begin
  sWorksheetGrid1.FrozenRows := EdFrozenRows.Value;
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
  if ParamCount > 0 then
    LoadFile(ParamStr(1));
end;

procedure TForm1.LoadFile(const AFileName: String);
// Loads first worksheet from file into grid
var
  pages: TStrings;
  i: Integer;
begin
  // Load file
  sWorksheetGrid1.LoadFromSpreadsheetFile(AFileName);

  // Update user interface
  Caption := Format('fpsGrid - %s (%s)', [
    AFilename,
    GetFileFormatName(sWorksheetGrid1.Workbook.FileFormat)
  ]);
  CbShowGridLines.Checked := (soShowGridLines in sWorksheetGrid1.Worksheet.Options);
  CbShowHeaders.Checked := (soShowHeaders in sWorksheetGrid1.Worksheet.Options);
  EdFrozenCols.Value := sWorksheetGrid1.FrozenCols;
  EdFrozenRows.Value := sWorksheetGrid1.FrozenRows;

  // Create a tab in the pagecontrol for each worksheet contained in the workbook
  // This would be easer with a TTabControl. This has display issues, though.
  pages := TStringList.Create;
  try
    sWorksheetGrid1.GetSheets(pages);
    sWorksheetGrid1.Parent := PageControl1.Pages[0];
    while PageControl1.PageCount > pages.Count do PageControl1.Pages[1].Free;
    while PageControl1.PageCount < pages.Count do PageControl1.AddTabSheet;
    for i:=0 to PageControl1.PageCount-1 do
      PageControl1.Pages[i].Caption := pages[i];
  finally
    pages.Free;
  end;
end;

procedure TForm1.PageControl1Change(Sender: TObject);
begin
  sWorksheetGrid1.Parent := PageControl1.Pages[PageControl1.ActivePageIndex];
  sWorksheetGrid1.SelectSheetByIndex(PageControl1.ActivePageIndex);
end;

initialization
  {$I mainform.lrs}

end.

