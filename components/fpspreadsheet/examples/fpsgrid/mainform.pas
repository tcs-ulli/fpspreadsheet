unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus,
  fpspreadsheetgrid, fpspreadsheet, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    buttonPopulateGrid: TButton;
    MainMenu1: TMainMenu;
    mnuFile: TMenuItem;
    mnuOpen: TMenuItem;
    mnuQuit: TMenuItem;
    mnuSaveAs: TMenuItem;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    sWorksheetGrid1: TsWorksheetGrid;
    procedure buttonPopulateGridClick(Sender: TObject);
    procedure mnuOpenClick(Sender: TObject);
    procedure mnuQuitClick(Sender: TObject);
    procedure mnuSaveAsClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

{ TForm1 }

procedure TForm1.buttonPopulateGridClick(Sender: TObject);
// Populate grid with some demo data
var
  lWorksheet: TsWorksheet;
begin
  lWorksheet := TsWorksheet.Create;
  try
    lWorksheet.WriteUTF8Text(2, 2, 'Algo');
    sWorksheetGrid1.LoadFromWorksheet(lWorksheet);
  finally
    lWorksheet.Free;
  end;
end;

procedure TForm1.mnuOpenClick(Sender: TObject);
// Loads first worksheet from file into grid
begin
  if OpenDialog1.Execute then
  begin
    sWorksheetGrid1.LoadFromSpreadsheetFile(OpenDialog1.FileName);
  end;
end;

procedure TForm1.mnuQuitClick(Sender: TObject);
begin
  Close;
end;

procedure TForm1.mnuSaveAsClick(Sender: TObject);
// Saves sheet in grid to file, overwriting existing file
var
  lWorkBook: TsWorkbook;
  lWorkSheet:TsWorksheet;
begin
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


initialization
  {$I mainform.lrs}

end.

