unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Spin, Grids, graphutil,
  ColorBox, fpspreadsheetgrid, fpspreadsheet, fpsallformats;

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
    AcTextHoriz: TAction;
    AcTextVertCW: TAction;
    AcTextVertCCW: TAction;
    AcTextStacked: TAction;
    AcNFFixed: TAction;
    AcNFFixedTh: TAction;
    AcNFPercentage: TAction;
    AcIncDecimals: TAction;
    AcDecDecimals: TAction;
    AcNFGeneral: TAction;
    AcNFExp: TAction;
    AcNFSci: TAction;
    AcCopyFormat: TAction;
    AcWordwrap: TAction;
    AcVAlignDefault: TAction;
    AcVAlignTop: TAction;
    AcVAlignCenter: TAction;
    AcVAlignBottom: TAction;
    ActionList1: TActionList;
    CbShowHeaders: TCheckBox;
    CbShowGridLines: TCheckBox;
    CbBackgroundColor: TColorBox;
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
    MenuItem29: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem34: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem36: TMenuItem;
    MenuItem37: TMenuItem;
    MenuItem38: TMenuItem;
    MnuNumberFormat: TMenuItem;
    MnuNFFixed: TMenuItem;
    MnuNFFixedTh: TMenuItem;
    MnuNFPercentage: TMenuItem;
    MnuNFExp: TMenuItem;
    MnuNFSci: TMenuItem;
    MnuNFGeneral: TMenuItem;
    MnuTextRotation: TMenuItem;
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
    NumFormatPopupMenu: TPopupMenu;
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
    TbNumFormats: TToolButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    ToolButton26: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    procedure AcBorderExecute(Sender: TObject);
    procedure AcCopyFormatExecute(Sender: TObject);
    procedure AcEditExecute(Sender: TObject);
    procedure AcFontExecute(Sender: TObject);
    procedure AcFontStyleExecute(Sender: TObject);
    procedure AcHorAlignmentExecute(Sender: TObject);
    procedure AcIncDecDecimalsExecute(Sender: TObject);
    procedure AcNumFormatExecute(Sender: TObject);
    procedure AcOpenExecute(Sender: TObject);
    procedure AcQuitExecute(Sender: TObject);
    procedure AcSaveAsExecute(Sender: TObject);
    procedure AcTextRotationExecute(Sender: TObject);
    procedure AcVertAlignmentExecute(Sender: TObject);
    procedure AcWordwrapExecute(Sender: TObject);
    procedure CbBackgroundColorSelect(Sender: TObject);
    procedure CbShowHeadersClick(Sender: TObject);
    procedure CbShowGridLinesClick(Sender: TObject);
    procedure CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
    procedure EdFrozenColsChange(Sender: TObject);
    procedure EdFrozenRowsChange(Sender: TObject);
    procedure FontComboBoxSelect(Sender: TObject);
    procedure FontSizeComboBoxSelect(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure sWorksheetGrid1Selection(Sender: TObject; aCol, aRow: Integer);
  private
    { private declarations }
    FCopiedFormat: TCell;
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
  fpcanvas;

const
  HORALIGN_TAG = 100;
  VERTALIGN_TAG = 110;
  TEXTROT_TAG = 130;
  NUMFMT_TAG = 150;  // needs 20

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
  // Use a combination of these bits for the "Tag" of the Border actions - see FormCreate.

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

procedure TForm1.AcCopyFormatExecute(Sender: TObject);
var
  cell: PCell;
  r, c: Cardinal;
begin
  with sWorksheetGrid1 do begin
    if Workbook = nil then
      exit;

    if AcCopyFormat.Checked then begin
      r := GetWorksheetRow(Row);
      c := GetWorksheetCol(Col);
      cell := Worksheet.FindCell(r, c);
      if cell <> nil then
        FCopiedFormat := cell^;
    end;
  end;
end;

{ Changes the font of the selected cell by calling a standard font dialog. }
procedure TForm1.AcFontExecute(Sender: TObject);
begin
  with sWorksheetGrid1 do begin
    if Workbook = nil then
      exit;
    FontDialog1.Font := CellFonts[Selection];
    if FontDialog1.Execute then
      CellFonts[Selection] := FontDialog1.Font;
  end;
end;

procedure TForm1.AcFontStyleExecute(Sender: TObject);
var
  style: TsFontstyles;
begin
  with sWorksheetGrid1 do begin
    if Workbook = nil then
      exit;
    style := [];
    if AcFontBold.Checked then Include(style, fssBold);
    if AcFontItalic.Checked then Include(style, fssItalic);
    if AcFontStrikeout.Checked then Include(style, fssStrikeout);
    if AcFontUnderline.Checked then Include(style, fssUnderline);
    CellFontStyles[Selection] := style;
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

procedure TForm1.AcIncDecDecimalsExecute(Sender: TObject);
var
  cell: PCell;
  decs: Byte;
begin
  with sWorksheetGrid1 do begin
    if Workbook = nil then
      exit;
    cell := Worksheet.FindCell(GetWorksheetRow(Row), GetWorksheetCol(Col));
    if (cell <> nil) then begin
      decs := cell^.NumberDecimals;
      if (Sender = AcIncDecimals) then
        Worksheet.WriteDecimals(cell, decs+1);
      if (Sender = AcDecDecimals) and (decs > 0) then
        Worksheet.WriteDecimals(cell, decs-1);
    end;
  end;
end;

procedure TForm1.AcNumFormatExecute(Sender: TObject);
var
  nf: TsNumberFormat;
  c, r: Cardinal;
begin
  if sWorksheetGrid1.Worksheet = nil then
    exit;

  if TAction(Sender).Checked then
    nf := TsNumberFormat(TAction(Sender).Tag - NUMFMT_TAG)
  else
    nf := nfGeneral;

  with sWorksheetGrid1 do begin
    c := GetWorksheetCol(Col);
    r := GetWorksheetRow(Row);
    Worksheet.WriteNumberFormat(r, c, nf);
  end;

  UpdateNumFormatActions;
end;

procedure TForm1.AcTextRotationExecute(Sender: TObject);
var
  text_rot: TsTextRotation;
begin
  if TAction(Sender).Checked then
    text_rot := TsTextRotation(TAction(Sender).Tag - TEXTROT_TAG)
  else
    text_rot := trHorizontal;
  with sWorksheetGrid1 do TextRotations[Selection] := text_rot;
  UpdateTextRotationActions;
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

procedure TForm1.CbBackgroundColorSelect(Sender: TObject);
begin
  with sWorksheetGrid1 do BackgroundColors[Selection] := CbBackgroundColor.ItemIndex;
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

procedure TForm1.CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
type
  TRGB = packed record R,G,B: byte end;
var
  clr: TColor;
  rgb: TRGB absolute clr;
  i: Integer;
begin
  if sWorksheetGrid1.Workbook <> nil then begin
    Items.Clear;
    for i:=0 to sWorksheetGrid1.Workbook.GetPaletteSize-1 do begin
      clr := sWorksheetGrid1.Workbook.GetPaletteColor(i);
      Items.AddObject(Format('Color %d: %.2x%.2x%.2x', [i, rgb.R, rgb.G, rgb.B]),
        TObject(PtrInt(clr)));
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

procedure TForm1.FontComboBoxSelect(Sender: TObject);
var
  fname: String;
begin
  fname := FontCombobox.Items[FontCombobox.ItemIndex];
  if fname <> '' then
    with sWorksheetGrid1 do CellFontNames[Selection] := fName;
end;

procedure TForm1.FontSizeComboBoxSelect(Sender: TObject);
var
  sz: Integer;
begin
  sz := StrToInt(FontSizeCombobox.Items[FontSizeCombobox.ItemIndex]);
  if sz > 0 then
    with sWorksheetGrid1 do CellFontSizes[Selection] := sz;
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
  SetupBackgroundColorBox;

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

  sWorksheetGrid1Selection(nil, sWorksheetGrid1.Col, sWorksheetGrid1.Row);
end;

procedure TForm1.PageControl1Change(Sender: TObject);
begin
  sWorksheetGrid1.Parent := PageControl1.Pages[PageControl1.ActivePageIndex];
  sWorksheetGrid1.SelectSheetByIndex(PageControl1.ActivePageIndex);
end;

procedure TForm1.SetupBackgroundColorBox;
begin
  // This change triggers re-reading of the workbooks palette by the OnGetColors
  // event of the ColorBox.
  CbBackgroundColor.Style := CbBackgroundColor.Style - [cbCustomColors];
  CbBackgroundColor.Style := CbBackgroundColor.Style + [cbCustomColors];
end;

procedure TForm1.sWorksheetGrid1Selection(Sender: TObject; aCol, aRow: Integer);
var
  r, c: Cardinal;
begin
  if sWorksheetGrid1.Workbook = nil then
    exit;

  if AcCopyFormat.Checked then begin
    r := sWorksheetGrid1.GetWorksheetRow(ARow);
    c := sWorksheetGrid1.GetWorksheetCol(ACol);
    sWorksheetGrid1.Worksheet.CopyFormat(@FCopiedFormat, r, c);
    AcCopyFormat.Checked := false;
  end;

  UpdateHorAlignmentActions;
  UpdateVertAlignmentActions;
  UpdateWordwraps;
  UpdateBackgroundColorIndex;
//  UpdateFontActions;
  UpdateFontNameIndex;
  UpdateFontSizeIndex;
  UpdateFontStyleActions;
  UpdateNumFormatActions;
end;

procedure TForm1.UpdateBackgroundColorIndex;
var
  sClr: TsColor;
begin
  with sWorksheetGrid1 do sClr := BackgroundColors[Selection];
  if sClr = scNotDefined then
    CbBackgroundColor.ItemIndex := -1
  else
    CbBackgroundColor.ItemIndex := sClr;
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

procedure TForm1.UpdateFontNameIndex;
var
  fname: String;
begin
  with sWorksheetGrid1 do fname := CellFontNames[Selection];
  if fname = '' then
    FontCombobox.ItemIndex := -1
  else
    FontCombobox.ItemIndex := FontCombobox.Items.IndexOf(fname);
end;

procedure TForm1.UpdateFontSizeIndex;
var
  sz: Single;
begin
  with sWorksheetGrid1 do sz := CellFontSizes[Selection];
  if sz < 0 then
    FontSizeCombobox.ItemIndex := -1
  else
    FontSizeCombobox.ItemIndex := FontSizeCombobox.Items.IndexOf(IntToStr(Round(sz)));
end;

procedure TForm1.UpdateFontStyleActions;
var
  style: TsFontStyles;
begin
  with sWorksheetGrid1 do style := CellFontStyles[Selection];
  AcFontBold.Checked := fssBold in style;
  AcFontItalic.Checked := fssItalic in style;
  AcFontUnderline.Checked := fssUnderline in style;
  AcFontStrikeout.Checked := fssStrikeOut in style;
end;

procedure TForm1.UpdateNumFormatActions;
var
  i: Integer;
  ac: TAction;
  nf: TsNumberFormat;
  cell: PCell;
  r,c: Cardinal;
begin
  with sWorksheetGrid1 do begin
    r := GetWorksheetRow(Row);
    c := GetWorksheetCol(Col);
    cell := Worksheet.FindCell(r, c);
    if (cell = nil) or (cell^.ContentType <> cctNumber) then
      nf := nfGeneral
    else
      nf := cell^.NumberFormat;
    for i:=0 to ActionList1.ActionCount-1 do begin
      ac := TAction(ActionList1.Actions[i]);
      if (ac.Tag >= NUMFMT_TAG) and (ac.Tag < NUMFMT_TAG + 20) then
        ac.Checked := ((ac.Tag - NUMFMT_TAG) = ord(nf));
    end;
  end;
end;

procedure TForm1.UpdateTextRotationActions;
var
  i: Integer;
  ac: TAction;
  text_rot: TsTextRotation;
begin
  with sWorksheetGrid1 do text_rot := TextRotations[Selection];
  for i:=0 to ActionList1.ActionCount-1 do begin
    ac := TAction(ActionList1.Actions[i]);
    if (ac.Tag >= TEXTROT_TAG) and (ac.Tag < TEXTROT_TAG+10) then
      ac.Checked := ((ac.Tag - TEXTROT_TAG) = ord(text_rot));
  end;
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

