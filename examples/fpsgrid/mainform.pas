unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Spin, Grids, graphutil,
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
    AcBorderTop: TAction;
    AcBorderBottom: TAction;
    AcBorderBottomDbl: TAction;
    AcBorderBottomMedium: TAction;
    AcBorderLeft: TAction;
    AcBorderRight: TAction;
    AcVAlignDefault: TAction;
    AcVAlignTop: TAction;
    AcVAlignCenter: TAction;
    AcVAlignBottom: TAction;
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
    MenuItem7: TMenuItem;
    MnuVertBottom: TMenuItem;
    MnuVertCentered: TMenuItem;
    MnuVertTop: TMenuItem;
    MnuVertDefault: TMenuItem;
    MnuVertAlignment: TMenuItem;
    MnuFOnt: TMenuItem;
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
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    ToolButton2: TToolButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    procedure AcBorderExecute(Sender: TObject);
    procedure AcEditExecute(Sender: TObject);
    procedure AcFontExecute(Sender: TObject);
    procedure AcFontStyleExecute(Sender: TObject);
    procedure AcHorAlignmentExecute(Sender: TObject);
    procedure AcVertAlignmentExecute(Sender: TObject);
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
    procedure sWorksheetGrid1Selection(Sender: TObject; aCol, aRow: Integer);
  private
    { private declarations }
    procedure LoadFile(const AFileName: String);
    procedure UpdateBorderActions(ACell: PCell);
    procedure UpdateHorAlignmentActions;
    procedure UpdateFontActions(AFont: TsFont);
    procedure UpdateVertAlignmentActions;
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

uses
  fpcanvas;

const
  HORALIGN_TAG = 100;
  VERTALIGN_TAG = 110;


{ TForm1 }

procedure TForm1.AcEditExecute(Sender: TObject);
begin
  if AcEdit.Checked then
    sWorksheetGrid1.Options := sWorksheetGrid1.Options + [goEditing]
  else
    sWorksheetGrid1.Options := sWorksheetGrid1.Options - [goEditing];
end;

procedure TForm1.AcBorderExecute(Sender: TObject);
var
  r,c: Cardinal;
  borders: TsCellBorders;
  lCell: PCell;
begin
  with sWorksheetGrid1 do begin
    if Worksheet <> nil then begin
      c := GetWorksheetCol(Col);
      r := GetWorksheetRow(Row);
      borders := [];
      if AcBorderTop.Checked then borders := borders + [cbNorth];
      if AcBorderLeft.Checked then borders := borders + [cbWest];
      if AcBorderRight.Checked then borders := borders + [cbEast];
      if AcBorderBottom.Checked or AcBorderBottomDbl.Checked or AcBorderBottomMedium.Checked then
        borders := borders + [cbSouth];
      Worksheet.WriteBorders(r, c, borders);
      if AcBorderBottom.Checked then
        Worksheet.WriteBorderLineStyle(r, c, cbSouth, lsThin);
      if AcBorderBottomMedium.Checked then
        Worksheet.WriteBorderLineStyle(r, c, cbSouth, lsMedium);
      if AcBorderBottomDbl.Checked then
        Worksheet.WriteBorderLineStyle(r, c, cbSouth, lsDouble);
    end;
  end;
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
  hor_align: TsHorAlignment;
begin
  if TAction(Sender).Checked then
    hor_align := TsHorAlignment(TAction(Sender).Tag - HORALIGN_TAG)
  else
    hor_align := haDefault;
  with sWorksheetGrid1 do HorAlignments[Selection] := hor_align;
  UpdateHorAlignmentActions;
end;

procedure TForm1.AcVertAlignmentExecute(Sender: TObject);
var
  vert_align: TsVertAlignment;
begin
  if TAction(Sender).Checked then
    vert_align := TsVertAlignment(TAction(Sender).Tag - VERTALIGN_TAG)
  else
    vert_align := vaDefault;
  with sWorksheetGrid1 do VertAlignments[Selection] := vert_align;
  UpdateVertAlignmentActions;
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
  sWorksheetGrid1.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));

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

procedure TForm1.sWorksheetGrid1Selection(Sender: TObject; aCol, aRow: Integer);
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
  UpdateHorAlignmentActions;
  UpdateVertAlignmentActions;
  UpdateBorderactions(cell);
  if cell = nil then
    exit;
  lFont := sWorksheetGrid1.Workbook.GetFont(cell^.FontIndex);
  UpdateFontActions(lFont);
end;

procedure TForm1.UpdateBorderActions(ACell: PCell);
begin
  AcBorderTop.Checked := (ACell <> nil) and (cbNorth in ACell^.Border);
  AcBorderLeft.Checked := (ACell <> nil) and (cbWest in ACell^.Border);
  AcBorderRight.Checked := (ACell <> nil) and (cbEast in ACell^.Border);
  AcBorderBottom.Checked := (ACell <> nil) and (cbSouth in ACell^.Border) and
   (ACell^.BorderStyles[cbSouth].LineStyle = lsThin);
  AcBorderBottomDbl.Checked := (ACell <> nil) and (cbSouth in ACell^.Border) and
    (ACell^.BorderStyles[cbSouth].LineStyle = lsDouble);
  AcBorderBottomMedium.Checked := (ACell <> nil) and (cbSouth in ACell^.Border) and
    (ACell^.BorderStyles[cbSouth].LineStyle = lsMedium);
end;

procedure TForm1.UpdateHorAlignmentActions;
var
  i: Integer;
  ac: TAction;
  hor_align: TsHorAlignment;
begin
  with sWorksheetGrid1 do hor_align := HorAlignments[Selection];
  for i:=0 to ActionList1.ActionCount-1 do begin
    ac := TAction(ActionList1.Actions[i]);
    if (ac.Tag >= HORALIGN_TAG) and (ac.Tag < HORALIGN_TAG+10) then
      ac.Checked := ((ac.Tag - HORALIGN_TAG) = ord(hor_align));
  end;
end;

procedure TForm1.UpdateFontActions(AFont: TsFont);
begin
  FontCombobox.ItemIndex := FontCombobox.Items.IndexOf(AFont.FontName);
  FontsizeCombobox.ItemIndex := FontSizeCombobox.Items.IndexOf(IntToStr(Round(AFont.Size)));
  AcFontBold.Checked := fssBold in AFont.Style;
  AcFontItalic.Checked := fssItalic in AFont.Style;
  AcFontUnderline.Checked := fssUnderline in AFont.Style;
  AcFontStrikeout.Checked := fssStrikeOut in AFont.Style;
end;

procedure TForm1.UpdateVertAlignmentActions;
var
  i: Integer;
  ac: TAction;
  vert_align: TsVertAlignment;
  t: Integer;
begin
  with sWorksheetGrid1 do vert_align := VertAlignments[Selection];
  for i:=0 to ActionList1.ActionCount-1 do begin
    ac := TAction(ActionList1.Actions[i]);
    t := ac.tag;
    if (ac.Tag >= VERTALIGN_TAG) and (ac.Tag < VERTALIGN_TAG+10) then
      ac.Checked := ((ac.Tag - VERTALIGN_TAG) = ord(vert_align));
  end;
end;


initialization
  {$I mainform.lrs}

end.

