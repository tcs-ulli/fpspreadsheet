unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList,
  fpspreadsheetgrid, fpspreadsheet, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    AcOpen: TAction;
    AcSaveAs: TAction;
    AcQuit: TAction;
    ActionList1: TActionList;
    btnPopulateGrid: TButton;
    CbDisplayFixedColRow: TCheckBox;
    CbDisplayGrid: TCheckBox;
    ImageList1: TImageList;
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
    sWorksheetGrid1: TsWorksheetGrid;
    TabSheet1: TTabSheet;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton5: TToolButton;
    procedure btnPopulateGridClick(Sender: TObject);
    procedure CbDisplayFixedColRowClick(Sender: TObject);
    procedure CbDisplayGridClick(Sender: TObject);
    procedure acOpenExecute(Sender: TObject);
    procedure acQuitExecute(Sender: TObject);
    procedure acSaveAsExecute(Sender: TObject);
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

procedure TForm1.CbDisplayFixedColRowClick(Sender: TObject);
begin
  sWorksheetGrid1.DisplayFixedColRow := CbDisplayFixedColRow.Checked;
end;

procedure TForm1.CbDisplayGridClick(Sender: TObject);
begin
  if CbDisplayGrid.Checked then
    sWorksheetGrid1.Options := sWorksheetGrid1.Options + [goHorzLine, goVertLine]
  else
    sWorksheetGrid1.Options := sWorksheetGrid1.Options - [goHorzLine, goVertLine];
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
  sWorksheetGrid1.LoadFromSpreadsheetFile(AFileName);
  Caption := Format('fpsGrid - %s (%s)', [
    AFilename,
    GetFileFormatName(sWorksheetGrid1.Workbook.FileFormat)
  ]);
  CbDisplayGrid.Checked := sWorksheetGrid1.Worksheet.ShowGridLines;

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

