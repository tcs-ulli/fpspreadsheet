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
    AcBorderNone: TAction;
    AcBorderHCenter: TAction;
    AcBorderVCenter: TAction;
    AcBorderTopBottom: TAction;
    AcBorderTopBottomThick: TAction;
    AcBorderInner: TAction;
    AcBorderAll: TAction;
    AcBorderOuter: TAction;
    AcBorderOuterMedium: TAction;
    AcWordwrap: TAction;
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
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    MnuWordwrap: TMenuItem;
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
    BordersPopupMenu: TPopupMenu;
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
    TbBorders: TToolButton;
    ToolButton20: TToolButton;
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
    procedure AcWordwrapExecute(Sender: TObject);
    procedure CbShowHeadersClick(Sender: TObject);
    procedure CbShowGridLinesClick(Sender: TObject);
    procedure AcOpenExecute(Sender: TObject);
    procedure AcQuitExecute(Sender: TObject);
    procedure AcSaveAsExecute(Sender: TObject);
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
    procedure UpdateFontActions(AFont: TsFont);
    procedure UpdateHorAlignmentActions;
    procedure UpdateVertAlignmentActions;
    procedure UpdateWordwraps;
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

  LEFT_BORDER_THIN       = $0001;
  LEFT_BORDER_THICK      = $0002;
  LR_INNER_BORDER_THIN   = $0008;
  RIGHT_BORDER_THIN      = $0010;
  RIGHT_BORDER_THICK     = $0020;
  TOP_BORDER_THIN        = $0100;
  TOP_BORDER_THICK       = $0200;
  TB_INNER_BORDER_THIN   = $0800;
  BOTTOM_BORDER_THIN     = $1000;
  BOTTOM_BORDER_THICK    = $2000;
  BOTTOM_BORDER_DOUBLE   = $3000;
  LEFT_BORDER_MASK       = $0007;
  RIGHT_BORDER_MASK      = $0070;
  TOP_BORDER_MASK        = $0700;
  BOTTOM_BORDER_MASK     = $7000;
  LR_INNER_BORDER        = $0008;
  TB_INNER_BORDER        = $0800;
  // Use a combination of these bits for the "Tag" of the Border actions.

{ TForm1 }

procedure TForm1.AcEditExecute(Sender: TObject);
begin
  if AcEdit.Checked then
    sWorksheetGrid1.Options := sWorksheetGrid1.Options + [goEditing]
  else
    sWorksheetGrid1.Options := sWorksheetGrid1.Options - [goEditing];
end;

procedure TForm1.AcBorderExecute(Sender: TObject);
const
  LINESTYLES: Array[1..3] of TsLinestyle = (lsThin, lsMedium, lsDouble);
var
  r,c: Integer;
  ls: integer;
  bs: TsCellBorderStyle;
begin
  bs.Color := scBlack;

  with sWorksheetGrid1 do begin
    TbBorders.Action := TAction(Sender);

    BeginUpdate;
    try
      if TAction(Sender).Tag = 0 then begin
        CellBorders[Selection] := [];
        exit;
      end;
      // Top and bottom edges
      for c := Selection.Left to Selection.Right do begin
        ls := (TAction(Sender).Tag and TOP_BORDER_MASK) shr 8;
        if (ls <> 0) then begin
          CellBorder[c, Selection.Top] := CellBorder[c, Selection.Top] + [cbNorth];
          bs.LineStyle := LINESTYLES[ls];
          CellBorderStyle[c, Selection.Top, cbNorth] := bs;
        end;
        ls := (TAction(Sender).Tag and BOTTOM_BORDER_MASK) shr 12;
        if ls <> 0 then begin
          CellBorder[c, Selection.Bottom] := CellBorder[c, Selection.Bottom] + [cbSouth];
          bs.LineStyle := LINESTYLES[ls];
          CellBorderStyle[c, Selection.Bottom, cbSouth] := bs;
        end;
      end;
      // Left and right edges
      for r := Selection.Top to Selection.Bottom do begin
        ls := (TAction(Sender).Tag and LEFT_BORDER_MASK);
        if ls <> 0 then begin
          CellBorder[Selection.Left, r] := CellBorder[Selection.Left, r] + [cbWest];
          bs.LineStyle := LINESTYLES[ls];
          CellBorderStyle[Selection.Left, r, cbWest] := bs;
        end;
        ls := (TAction(Sender).Tag and RIGHT_BORDER_MASK) shr 4;
        if ls <> 0 then begin
          CellBorder[Selection.Right, r] := CellBorder[Selection.Right, r] + [cbEast];
          bs.LineStyle := LINESTYLES[ls];
          CellBorderStyle[Selection.Right, r, cbEast] := bs;
        end;
      end;
      // Inner edges along row (vertical border lines) - we assume only thin lines.
      bs.LineStyle := lsThin;
      if (TAction(Sender).Tag and LR_INNER_BORDER <> 0) and (Selection.Right > Selection.Left)
      then
        for r := Selection.Top to Selection.Bottom do begin
          CellBorder[Selection.Left, r] := CellBorder[Selection.Left, r] + [cbEast];
          CellBorderStyle[Selection.Left, r, cbEast] := bs;
          for c := Selection.Left+1 to Selection.Right-1 do begin
            CellBorder[c,r] := CellBorder[c, r] + [cbEast, cbWest];
            CellBorderStyle[c, r, cbEast] := bs;
            CellBorderStyle[c, r, cbWest] := bs;
          end;
          CellBorder[Selection.Right, r] := CellBorder[Selection.Right, r] + [cbWest];
          CellBorderStyle[Selection.Right, r, cbWest] := bs;
        end;
      // Inner edges along column (horizontal border lines)
      if (TAction(Sender).Tag and TB_INNER_BORDER <> 0) and (Selection.Bottom > Selection.Top)
      then
        for c := Selection.Left to Selection.Right do begin
          CellBorder[c, Selection.Top] := CellBorder[c, Selection.Top] + [cbSouth];
          CellBorderStyle[c, Selection.Top, cbSouth] := bs;
          for r := Selection.Top+1 to Selection.Bottom-1 do begin
            CellBorder[c, r] := CellBorder[c, r] + [cbNorth, cbSouth];
            CellBorderStyle[c, r, cbNorth] := bs;
            CellBorderStyle[c, r, cbSouth] := bs;
          end;
          CellBorder[c, Selection.Bottom] := CellBorder[c, Selection.Bottom] + [cbNorth];
          CellBorderStyle[c, Selection.Bottom, cbNorth] := bs;
        end;
    finally
      EndUpdate;
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

procedure TForm1.AcWordwrapExecute(Sender: TObject);
begin
  with sWorksheetGrid1 do Wordwraps[Selection] := TAction(Sender).Checked;
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
  // Adjust format toolbar height, looks strange at 120 dpi
  FormatToolbar.Height := FontCombobox.Height + 2*FontCombobox.Top;
  FormatToolbar.ButtonHeight := FormatToolbar.Height - 4;

  // Populate font combobox
  FontCombobox.Items.Assign(Screen.Fonts);

  // Set the Tags of the Border actions
  AcBorderNone.Tag := 0;
  AcBorderLeft.Tag := LEFT_BORDER_THIN;
  AcBorderHCenter.Tag := LR_INNER_BORDER_THIN;
  AcBorderRight.Tag := RIGHT_BORDER_THIN;
  AcBorderTop.Tag := TOP_BORDER_THIN;
  AcBorderVCenter.Tag := TB_INNER_BORDER_THIN;
  AcBorderBottom.Tag := BOTTOM_BORDER_THIN;
  AcBorderBottomDbl.Tag := BOTTOM_BORDER_DOUBLE;
  AcBorderBottomMedium.Tag := BOTTOM_BORDER_THICK;
  AcBorderTopBottom.Tag := TOP_BORDER_THIN + BOTTOM_BORDER_THIN;
  AcBorderTopBottomThick.Tag := TOP_BORDER_THIN + BOTTOM_BORDER_THICK;
  AcBorderInner.Tag := LR_INNER_BORDER_THIN + TB_INNER_BORDER_THIN;
  AcBorderOuter.Tag := LEFT_BORDER_THIN + RIGHT_BORDER_THIN + TOP_BORDER_THIN + BOTTOM_BORDER_THIN;
  AcBorderOuterMedium.Tag := LEFT_BORDER_THICK + RIGHT_BORDER_THICK + TOP_BORDER_THICK + BOTTOM_BORDER_THICK;
  AcBorderAll.Tag := AcBorderOuter.Tag + AcBorderInner.Tag;
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
  UpdateWordwraps;
  if cell = nil then
    exit;
  lFont := sWorksheetGrid1.Workbook.GetFont(cell^.FontIndex);
  UpdateFontActions(lFont);
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

procedure TForm1.UpdateWordwraps;
var
  wrapped: Boolean;
begin
  with sWorksheetGrid1 do wrapped := Wordwraps[Selection];
  AcWordwrap.Checked := wrapped;
end;

initialization
  {$I mainform.lrs}

end.

