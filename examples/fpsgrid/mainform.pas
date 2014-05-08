unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Spin, ColorBox,graphutil,
  fpspreadsheetgrid, fpspreadsheet, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    AcOpen: TAction;
    AcSaveAs: TAction;
    AcQuit: TAction;
    AcEdit: TAction;
    AcLeftAlign: TAction;
    AcHorCenterAlign: TAction;
    AcRightAlign: TAction;
    AcHorDefaultAlign: TAction;
    AcFontBold: TAction;
    AcFontItalic: TAction;
    AcFontStrikeout: TAction;
    AcFontUnderline: TAction;
    AcFont: TAction;
    ActionList1: TActionList;
    CbShowHeaders: TCheckBox;
    CbShowGridLines: TCheckBox;
    FontComboBox: TComboBox;
    EdFrozenRows: TSpinEdit;
    FontDialog1: TFontDialog;
    FontSizeComboBox: TComboBox;
    ImageList1: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MnuHorDefault: TMenuItem;
    MnuHorAlignment: TMenuItem;
    mnuFormat: TMenuItem;
    mnuEdit: TMenuItem;
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
    FormatToolBar: TToolBar;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    procedure AcEditExecute(Sender: TObject);
    procedure AcFontExecute(Sender: TObject);
    procedure AcFontStyleExecute(Sender: TObject);
    procedure AcHorAlignmentExecute(Sender: TObject);
    procedure CbShowHeadersClick(Sender: TObject);
    procedure CbShowGridLinesClick(Sender: TObject);
    procedure acOpenExecute(Sender: TObject);
    procedure acQuitExecute(Sender: TObject);
    procedure acSaveAsExecute(Sender: TObject);
    procedure EdFrozenColsChange(Sender: TObject);
    procedure EdFrozenRowsChange(Sender: TObject);
    procedure FontComboBoxSelect(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure sWorksheetGrid1SelectCell(Sender: TObject; aCol, aRow: Integer;
      var CanSelect: Boolean);
  private
    { private declarations }
    procedure LoadFile(const AFileName: String);
    procedure UpdateHorAlignment(AValue: TsHorAlignment);
    procedure UpdateFont(AFont: TsFont);
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

uses
  fpcanvas, Grids;

const
  HORALIGN_TAG = 100;


{ TForm1 }

procedure TForm1.AcEditExecute(Sender: TObject);
begin
  if AcEdit.Checked then
    sWorksheetGrid1.Options := sWorksheetGrid1.Options + [goEditing]
  else
    sWorksheetGrid1.Options := sWorksheetGrid1.Options - [goEditing];
end;

{ Changes the font of the selected cell by calling a standard font dialog.
  Note that the worksheet's and grid's fonts are implemented differently.
  In particular, the worksheet's font color is an index into the workbook's
  palette while the grid's font color is an rgb value. }
procedure TForm1.AcFontExecute(Sender: TObject);
var
  r,c: Cardinal;
  f: Integer;
  style: TsFontStyles;
  lFont: TsFont;
begin
  with sWorksheetGrid1 do begin
    if Worksheet <> nil then begin
      c := GetWorksheetCol(Col);
      r := GetWorksheetRow(Row);
      f := Worksheet.GetCell(r, c)^.FontIndex;
      Convert_sFont_to_Font(Workbook.GetFont(f), FontDialog1.Font);
      if FontDialog1.Execute then begin
        lFont := TsFont.Create;
        try
          Convert_Font_to_sFont(FontDialog1.Font, lFont);
          WorkSheet.WriteFont(r, c, lFont.FontName, lFont.Size, lFont.Style, lFont.Color);
        finally
          lFont.Free;
        end;
      end;
    end;
  end;
end;

procedure TForm1.AcFontStyleExecute(Sender: TObject);
var
  style: TsFontstyles;
  f: Integer;
  r,c: Cardinal;
  lFont: TsFont;
begin
  with sWorksheetGrid1 do begin
    c := GetWorksheetCol(Col);
    r := GetWorksheetRow(Row);
    if Worksheet <> nil then begin
      f := Worksheet.GetCell(r, c)^.FontIndex;
      lFont := Workbook.GetFont(f);
      style := lFont.Style;
      if TAction(Sender).Checked then
        Include(style, TsFontStyle(TAction(Sender).Tag))
      else
        Exclude(style, TsFontStyle(TAction(Sender).Tag));
      Worksheet.WriteFontStyle(r, c, style);
    end;
  end;
end;

procedure TForm1.AcHorAlignmentExecute(Sender: TObject);
var
  horalign: TsHorAlignment;
  c, r: Cardinal;
begin
  horalign := TsHorAlignment(TAction(Sender).Tag - HORALIGN_TAG);
  if TAction(Sender).Checked then
    horalign := haDefault;
  UpdateHorAlignment(horalign);
  with sWorksheetGrid1 do begin
    c := GetWorksheetCol(Col);
    r := GetWorksheetRow(Row);
    if Worksheet <> nil then
      Worksheet.WriteHorAlignment(r, c, horalign);
  end;
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
begin
  if sWorksheetGrid1.Workbook = nil then
    exit;

  if SaveDialog1.Execute then
    sWorksheetGrid1.SaveToSpreadsheetFile(SaveDialog1.FileName);
end;

procedure TForm1.EdFrozenColsChange(Sender: TObject);
begin
  sWorksheetGrid1.FrozenCols := EdFrozenCols.Value;
end;

procedure TForm1.EdFrozenRowsChange(Sender: TObject);
begin
  sWorksheetGrid1.FrozenRows := EdFrozenRows.Value;
end;

procedure TForm1.FontComboBoxSelect(Sender: TObject);
var
  c, r: Cardinal;
  f: Integer;
  lFont: TsFont;
  h: Integer;
  s: String;
begin
  if sWorksheetGrid1.Workbook = nil then
    exit;

  with sWorksheetGrid1 do begin
    c := GetWorksheetCol(Col);
    r := GetWorksheetRow(Row);
    f := Worksheet.GetCell(r, c)^.FontIndex;
    lFont := Workbook.GetFont(f);

    if FontCombobox.ItemIndex = -1 then
      s := lFont.FontName
    else
      s := FontCombobox.Items[FontCombobox.ItemIndex];

    if FontSizeCombobox.ItemIndex = -1 then
      h := round(lFont.Size)
    else
      h := StrToInt(FontSizeCombobox.Items[FontSizeCombobox.ItemIndex]);

    Worksheet.WriteFont(r, c, s, h, lFont.Style, lFont.Color);
  end;
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
  if ParamCount > 0 then
    LoadFile(ParamStr(1));
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  // Populate font combobox
  FontCombobox.Items.Assign(Screen.Fonts);
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

procedure TForm1.sWorksheetGrid1SelectCell(Sender: TObject;
  aCol, aRow: Integer; var CanSelect: Boolean);
var
  cell: PCell;
  c, r: Cardinal;
  lFont: TsFont;
begin
  with sWorksheetGrid1 do begin
    if Worksheet = nil then exit;
    c := GetWorksheetCol(ACol);
    r := GetWorksheetRow(ARow);
    cell := Worksheet.FindCell(r, c);
  end;
  if cell = nil then
    exit;
  UpdateHorAlignment(cell^.HorAlignment);
  lFont := sWorksheetGrid1.Workbook.GetFont(cell^.FontIndex);
  UpdateFont(lFont);
end;

procedure TForm1.UpdateHorAlignment(AValue: TsHorAlignment);
var
  i: Integer;
  ac: TAction;
begin
  for i:=0 to ActionList1.ActionCount-1 do begin
    ac := TAction(ActionList1.Actions[i]);
    if (ac.Tag >= HORALIGN_TAG) and (ac.Tag < HORALIGN_TAG+10) then
      ac.Checked := ((ac.Tag - HORALIGN_TAG) = ord(AValue));
  end;
end;

procedure TForm1.UpdateFont(AFont: TsFont);
begin
  FontCombobox.ItemIndex := FontCombobox.Items.IndexOf(AFont.FontName);
  FontsizeCombobox.ItemIndex := FontSizeCombobox.Items.IndexOf(IntToStr(Round(AFont.Size)));
  AcFontBold.Checked := fssBold in AFont.Style;
  AcFontItalic.Checked := fssItalic in AFont.Style;
  AcFontUnderline.Checked := fssUnderline in AFont.Style;
  AcFontStrikeout.Checked := fssStrikeOut in AFont.Style;
end;

initialization
  {$I mainform.lrs}

end.

