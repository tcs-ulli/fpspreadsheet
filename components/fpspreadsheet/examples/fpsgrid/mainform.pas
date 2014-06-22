unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Spin, Grids, ColorBox, Buttons,
  ButtonPanel, fpspreadsheetgrid, fpspreadsheet, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    BtnOpen: TButton;
    BtnSave: TButton;
    BtnNew: TButton;
    SheetsCombo: TComboBox;
    Label1: TLabel;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    SaveDialog: TSaveDialog;
    WorksheetGrid: TsWorksheetGrid;
    procedure BtnNewClick(Sender: TObject);
    procedure BtnOpenClick(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure SheetsComboSelect(Sender: TObject);
  private
    { private declarations }
    procedure LoadFile(const AFileName: String);
    procedure SetupBackgroundColorBox;
    procedure UpdateBackgroundColorIndex;
    procedure UpdateFontNameIndex;
    procedure UpdateFontSizeIndex;
    procedure UpdateFontStyleActions;
    procedure UpdateHorAlignmentActions;
    procedure UpdateNumFormatActions;
    procedure UpdateTextRotationActions;
    procedure UpdateVertAlignmentActions;
    procedure UpdateWordwraps;
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

uses
  fpcanvas, fpsutils;


{ TForm1 }

procedure TForm1.BtnNewClick(Sender: TObject);
var
  dlg: TForm;
  edCols, edRows: TSpinEdit;
  x: Integer;
begin
  dlg := TForm.Create(nil);
  try
    dlg.Width := 220;
    dlg.Height := 128;
    dlg.Position := poMainFormCenter;
    dlg.Caption := 'New workbook';
    edCols := TSpinEdit.Create(dlg);
    with edCols do begin
      Parent := dlg;
      Left := dlg.ClientWidth - Width - 24;
      Top := 16;
      Value := WorksheetGrid.ColCount - ord(WorksheetGrid.DisplayFixedColRow);
    end;
    with TLabel.Create(dlg) do begin
      Parent := dlg;
      Left := 24;
      Top := edCols.Top + 3;
      Caption := 'Columns:';
      FocusControl := edCols;
    end;
    edRows := TSpinEdit.Create(dlg);
    with edRows do begin
      Parent := dlg;
      Left := edCols.Left;
      Top := edCols.Top + edCols.Height + 8;
      Value := WorksheetGrid.RowCount - ord(WorksheetGrid.DisplayFixedColRow);
    end;
    with TLabel.Create(dlg) do begin
      Parent := dlg;
      Left := 24;
      Top := edRows.Top + 3;
      Caption := 'Rows:';
      FocusControl := edRows;
    end;
    with TButtonPanel.Create(dlg) do begin
      Parent := dlg;
      Align := alBottom;
      ShowButtons := [pbCancel, pbOK];
    end;
    if dlg.ShowModal = mrOK then begin
      WorksheetGrid.NewWorksheet(edCols.Value, edRows.Value);
      SheetsCombo.Items.Clear;
      SheetsCombo.Items.Add('Sheet 1');
      SheetsCombo.ItemIndex := 0;
    end;
  finally
    dlg.Free;
  end;
end;

procedure TForm1.BtnOpenClick(Sender: TObject);
begin
  if OpenDialog.Execute then
    LoadFile(OpenDialog.FileName);
end;

// Saves sheet in grid to file, overwriting existing file
procedure TForm1.BtnSaveClick(Sender: TObject);
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  if SaveDialog.Execute then
  begin
    Screen.Cursor := crHourglass;
    try
      WorksheetGrid.SaveToSpreadsheetFile(SaveDialog.FileName);
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TForm1.SheetsComboSelect(Sender: TObject);
begin
  WorksheetGrid.SelectSheetByIndex(SheetsCombo.ItemIndex);
end;

// Loads first worksheet from file into grid
procedure TForm1.LoadFile(const AFileName: String);
var
  i: Integer;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    WorksheetGrid.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));

    // Update user interface
    Caption := Format('fpsGrid - %s (%s)', [
      AFilename,
      GetFileFormatName(WorksheetGrid.Workbook.FileFormat)
    ]);

    // Collect the sheet names in the combobox for switching sheets.
    WorksheetGrid.GetSheets(SheetsCombo.Items);
    SheetsCombo.ItemIndex := 0;

//    WorksheetGridSelection(nil, WorksheetGrid.Col, WorksheetGrid.Row);
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.SetupBackgroundColorBox;
begin

end;

procedure TForm1.UpdateBackgroundColorIndex;
begin

end;

procedure TForm1.UpdateFontNameIndex;
begin

end;

procedure TForm1.UpdateFontSizeIndex;
begin

end;

procedure TForm1.UpdateFontStyleActions;
begin

end;

procedure TForm1.UpdateHorAlignmentActions;
begin

end;

procedure TForm1.UpdateNumFormatActions;
begin

end;

procedure TForm1.UpdateTextRotationActions;
begin

end;

procedure TForm1.UpdateVertAlignmentActions;
begin

end;

procedure TForm1.UpdateWordwraps;
begin

end;

initialization
  {$I mainform.lrs}

end.

