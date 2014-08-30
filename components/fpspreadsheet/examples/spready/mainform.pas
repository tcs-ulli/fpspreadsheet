unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Spin, Grids,
  ColorBox, ValEdit, fpspreadsheetgrid, fpspreadsheet, fpsallformats;

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
    AcCopyFormat: TAction;
    AcNFCurrency: TAction;
    AcNFCurrencyRed: TAction;
    AcNFShortDateTime: TAction;
    AcNFShortDate: TAction;
    AcNFLongDate: TAction;
    AcNFShortTime: TAction;
    AcNFLongTime: TAction;
    AcNFShortTimeAM: TAction;
    AcNFLongTimeAM: TAction;
    AcNFTimeInterval: TAction;
    AcNFCustomDM: TAction;
    AcNFCustomMY: TAction;
    AcNFCusstomMS: TAction;
    AcNFCustomMSZ: TAction;
    AcNew: TAction;
    AcAddColumn: TAction;
    AcAddRow: TAction;
    AcViewInspector: TAction;
    AcWordwrap: TAction;
    AcVAlignDefault: TAction;
    AcVAlignTop: TAction;
    AcVAlignCenter: TAction;
    AcVAlignBottom: TAction;
    ActionList: TActionList;
    CbShowHeaders: TCheckBox;
    CbShowGridLines: TCheckBox;
    CbBackgroundColor: TColorBox;
    CbReadFormulas: TCheckBox;
    CbHeaderStyle: TComboBox;
    CbAutoCalcFormulas: TCheckBox;
    EdFormula: TEdit;
    EdCellAddress: TEdit;
    FontComboBox: TComboBox;
    EdFrozenRows: TSpinEdit;
    FontDialog: TFontDialog;
    FontSizeComboBox: TComboBox;
    ImageList: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    MainMenu: TMainMenu;
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
    MenuItem39: TMenuItem;
    MenuItem40: TMenuItem;
    MenuItem41: TMenuItem;
    MenuItem42: TMenuItem;
    MenuItem43: TMenuItem;
    MenuItem44: TMenuItem;
    MenuItem45: TMenuItem;
    MenuItem46: TMenuItem;
    MenuItem47: TMenuItem;
    MenuItem48: TMenuItem;
    MenuItem49: TMenuItem;
    MenuItem50: TMenuItem;
    MenuItem51: TMenuItem;
    MenuItem52: TMenuItem;
    MenuItem53: TMenuItem;
    MenuItem54: TMenuItem;
    MenuItem55: TMenuItem;
    MenuItem56: TMenuItem;
    MenuItem57: TMenuItem;
    MenuItem58: TMenuItem;
    MenuItem59: TMenuItem;
    MenuItem60: TMenuItem;
    MenuItem61: TMenuItem;
    MenuItem62: TMenuItem;
    MenuItem63: TMenuItem;
    MenuItem64: TMenuItem;
    MenuItem65: TMenuItem;
    MenuItem66: TMenuItem;
    mnuInspector: TMenuItem;
    mnuView: TMenuItem;
    MnuFmtDateTimeMSZ: TMenuItem;
    MnuTimeInterval: TMenuItem;
    MnuShortTimeAM: TMenuItem;
    MnuLongTimeAM: TMenuItem;
    MnuFmtDateTimeMY: TMenuItem;
    MnuFmtDateTimeDM: TMenuItem;
    MnuShortTime: TMenuItem;
    MnuShortDate: TMenuItem;
    MnuLongTime: TMenuItem;
    MnuLongDate: TMenuItem;
    MnuShortDateTime: TMenuItem;
    MnuCurrencyRed: TMenuItem;
    MnuCurrency: TMenuItem;
    MnuNumberFormat: TMenuItem;
    MnuNFFixed: TMenuItem;
    MnuNFFixedTh: TMenuItem;
    MnuNFPercentage: TMenuItem;
    MnuNFExp: TMenuItem;
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
    OpenDialog: TOpenDialog;
    PageControl1: TPageControl;
    InspectorPageControl: TPageControl;
    Panel1: TPanel;
    BordersPopupMenu: TPopupMenu;
    NumFormatPopupMenu: TPopupMenu;
    SaveDialog: TSaveDialog;
    EdFrozenCols: TSpinEdit;
    FormulaToolBar: TToolBar;
    FormulaToolbarSplitter: TSplitter;
    InspectorSplitter: TSplitter;
    PgCellValue: TTabSheet;
    PgProperties: TTabSheet;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    ToolButton27: TToolButton;
    CellInspector: TValueListEditor;
    WorksheetGrid: TsWorksheetGrid;
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
    procedure AcAddColumnExecute(Sender: TObject);
    procedure AcAddRowExecute(Sender: TObject);
    procedure AcBorderExecute(Sender: TObject);
    procedure AcCopyFormatExecute(Sender: TObject);
    procedure AcEditExecute(Sender: TObject);
    procedure AcFontExecute(Sender: TObject);
    procedure AcFontStyleExecute(Sender: TObject);
    procedure AcHorAlignmentExecute(Sender: TObject);
    procedure AcIncDecDecimalsExecute(Sender: TObject);
    procedure AcNewExecute(Sender: TObject);
    procedure AcNumFormatExecute(Sender: TObject);
    procedure AcOpenExecute(Sender: TObject);
    procedure AcQuitExecute(Sender: TObject);
    procedure AcSaveAsExecute(Sender: TObject);
    procedure AcTextRotationExecute(Sender: TObject);
    procedure AcVertAlignmentExecute(Sender: TObject);
    procedure AcViewInspectorExecute(Sender: TObject);
    procedure AcWordwrapExecute(Sender: TObject);
    procedure CbAutoCalcFormulasChange(Sender: TObject);
    procedure CbBackgroundColorSelect(Sender: TObject);
    procedure CbHeaderStyleChange(Sender: TObject);
    procedure CbReadFormulasChange(Sender: TObject);
    procedure CbShowHeadersClick(Sender: TObject);
    procedure CbShowGridLinesClick(Sender: TObject);
    procedure CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
    procedure EdCellAddressEditingDone(Sender: TObject);
    procedure EdFormulaEditingDone(Sender: TObject);
    procedure EdFrozenColsChange(Sender: TObject);
    procedure EdFrozenRowsChange(Sender: TObject);
    procedure FontComboBoxSelect(Sender: TObject);
    procedure FontSizeComboBoxSelect(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure InspectorPageControlChange(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure WorksheetGridSelection(Sender: TObject; aCol, aRow: Integer);

  private
    { private declarations }
    FCopiedFormat: TCell;
    procedure LoadFile(const AFileName: String);
    procedure SetupBackgroundColorBox;
    procedure UpdateBackgroundColorIndex;
    procedure UpdateCellInfo(ACell: PCell);
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
  StrUtils, TypInfo,
  fpcanvas, fpsutils, fpsnumformatparser;

const
  DROPDOWN_COUNT = 24;

  HORALIGN_TAG = 100;
  VERTALIGN_TAG = 110;
  TEXTROT_TAG = 130;
  NUMFMT_TAG = 1000;  // differnce 10 per format item

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
    WorksheetGrid.Options := WorksheetGrid.Options + [goEditing]
  else
    WorksheetGrid.Options := WorksheetGrid.Options - [goEditing];
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

  with WorksheetGrid do begin
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

procedure TForm1.AcAddColumnExecute(Sender: TObject);
begin
  WorksheetGrid.InsertCol(WorksheetGrid.Col);
  WorksheetGrid.Col := WorksheetGrid.Col + 1;
end;

procedure TForm1.AcAddRowExecute(Sender: TObject);
begin
  WorksheetGrid.InsertRow(WorksheetGrid.Row);
  WorksheetGrid.Row := WorksheetGrid.Row + 1;
end;

procedure TForm1.AcCopyFormatExecute(Sender: TObject);
var
  cell: PCell;
  r, c: Cardinal;
begin
  with WorksheetGrid do begin
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
  with WorksheetGrid do begin
    if Workbook = nil then
      exit;
    FontDialog.Font := CellFonts[Selection];
    if FontDialog.Execute then
      CellFonts[Selection] := FontDialog.Font;
  end;
end;

procedure TForm1.AcFontStyleExecute(Sender: TObject);
var
  style: TsFontstyles;
begin
  with WorksheetGrid do begin
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
  with WorksheetGrid do HorAlignments[Selection] := hor_align;
  UpdateHorAlignmentActions;
end;

procedure TForm1.AcIncDecDecimalsExecute(Sender: TObject);
var
  cell: PCell;
  decs: Byte;
  currsym: String;
begin
  currsym := Sender.ClassName;
  with WorksheetGrid do begin
    if Workbook = nil then
      exit;
    cell := Worksheet.FindCell(GetWorksheetRow(Row), GetWorksheetCol(Col));
    if (cell <> nil) then begin
      if cell^.NumberFormat = nfGeneral then begin
        Worksheet.WriteNumberFormat(cell, nfFixed, '0.00');
        exit;
      end;
      Worksheet.GetNumberFormatAttributes(cell, decs, currSym);
      if (Sender = AcIncDecimals) then
        Worksheet.WriteDecimals(cell, decs+1)
      else
      if (Sender = AcDecDecimals) and (decs > 0) then
        Worksheet.WriteDecimals(cell, decs-1);
    end;
  end;
end;

procedure TForm1.AcNewExecute(Sender: TObject);
begin
  WorksheetGrid.NewWorkbook(26, 100);

  WorksheetGrid.BeginUpdate;
  try
    WorksheetGrid.Col := WorksheetGrid.FixedCols;
    WorksheetGrid.Row := WorksheetGrid.FixedRows;
  finally
    WorksheetGrid.EndUpdate;
  end;
end;

procedure TForm1.AcNumFormatExecute(Sender: TObject);
const
  DATETIME_CUSTOM: array[0..4] of string = ('', 'dd/mmm', 'mmm/yy', 'nn:ss', 'nn:ss.zzz');
var
  nf: TsNumberFormat;
  c, r: Cardinal;
  cell: PCell;
  fmt: String;
  decs: Byte;
  cs: String;
  isDateTimeFmt: Boolean;
begin
  if TAction(Sender).Checked then
    nf := TsNumberFormat((TAction(Sender).Tag - NUMFMT_TAG) div 10)
  else
    nf := nfGeneral;

  fmt := '';
  isDateTimeFmt := IsDateTimeFormat(nf);
  if nf = nfCustom then begin
    fmt := DATETIME_CUSTOM[TAction(Sender).Tag mod 10];
    isDateTimeFmt := true;
  end;

  with WorksheetGrid do begin
    c := GetWorksheetCol(Col);
    r := GetWorksheetRow(Row);
    cell := Worksheet.GetCell(r, c);
    Worksheet.GetNumberFormatAttributes(cell, decs, cs);
    if cs = '' then cs := '?';
    case cell^.ContentType of
      cctNumber, cctDateTime:
        if isDateTimeFmt then begin
          if IsDateTimeFormat(cell^.NumberFormat) then
            Worksheet.WriteDateTime(cell, cell^.DateTimeValue, nf, fmt)
          else
            Worksheet.WriteDateTime(cell, cell^.NumberValue, nf, fmt);
        end else
        if IsCurrencyFormat(nf) then begin
          if IsDateTimeFormat(cell^.NumberFormat) then
            Worksheet.WriteCurrency(cell, cell^.DateTimeValue, nf, decs, cs)
          else
            Worksheet.WriteCurrency(cell, cell^.Numbervalue, nf, decs, cs);
        end else begin
          if IsDateTimeFormat(cell^.NumberFormat) then
            Worksheet.WriteNumber(cell, cell^.DateTimeValue, nf, decs)
          else
            Worksheet.WriteNumber(cell, cell^.NumberValue, nf, decs)
        end;
      else
        Worksheet.WriteNumberformat(cell, nf, fmt);
    end;
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
  with WorksheetGrid do TextRotations[Selection] := text_rot;
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
  with WorksheetGrid do VertAlignments[Selection] := vert_align;
  UpdateVertAlignmentActions;
end;

procedure TForm1.AcViewInspectorExecute(Sender: TObject);
begin
  InspectorPageControl.Visible := AcViewInspector.Checked;
  InspectorSplitter.Visible := AcViewInspector.Checked;
  InspectorSplitter.Left := 0;
end;

procedure TForm1.AcWordwrapExecute(Sender: TObject);
begin
  with WorksheetGrid do Wordwraps[Selection] := TAction(Sender).Checked;
end;

procedure TForm1.CbAutoCalcFormulasChange(Sender: TObject);
begin
  WorksheetGrid.AutoCalc := CbAutoCalcFormulas.Checked;;
end;

procedure TForm1.CbBackgroundColorSelect(Sender: TObject);
begin
  with WorksheetGrid do BackgroundColors[Selection] := CbBackgroundColor.ItemIndex;
end;

procedure TForm1.CbHeaderStyleChange(Sender: TObject);
begin
  WorksheetGrid.TitleStyle := TTitleStyle(CbHeaderStyle.ItemIndex);
end;

procedure TForm1.CbReadFormulasChange(Sender: TObject);
begin
  WorksheetGrid.ReadFormulas := CbReadFormulas.Checked;
end;

procedure TForm1.CbShowHeadersClick(Sender: TObject);
begin
  WorksheetGrid.ShowHeaders := CbShowHeaders.Checked;
end;

procedure TForm1.CbShowGridLinesClick(Sender: TObject);
begin
  WorksheetGrid.ShowGridLines := CbShowGridLines.Checked;
end;

procedure TForm1.acOpenExecute(Sender: TObject);
begin
  if OpenDialog.Execute then
    LoadFile(OpenDialog.FileName);
end;

procedure TForm1.acQuitExecute(Sender: TObject);
begin
  Close;
end;

procedure TForm1.acSaveAsExecute(Sender: TObject);
// Saves sheet in grid to file, overwriting existing file
var
  err: String = '';
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
      err := WorksheetGrid.Workbook.ErrorMsg;
      if err <> '' then
        MessageDlg(err, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TForm1.CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
type
  TRGB = packed record R,G,B: byte end;
var
  clr: TColor;
  rgb: TRGB absolute clr;
  i: Integer;
begin
  if WorksheetGrid.Workbook <> nil then begin
    Items.Clear;
    for i:=0 to WorksheetGrid.Workbook.GetPaletteSize-1 do begin
      clr := WorksheetGrid.Workbook.GetPaletteColor(i);
      Items.AddObject(Format('Color %d: %.2x%.2x%.2x', [i, rgb.R, rgb.G, rgb.B]),
        TObject(PtrInt(clr)));
    end;
  end;
end;

procedure TForm1.EdCellAddressEditingDone(Sender: TObject);
var
  c, r: cardinal;
begin
  if ParseCellString(EdCellAddress.Text, r, c) then begin
    WorksheetGrid.Row := WorksheetGrid.GetGridRow(r);
    WorksheetGrid.Col := WorksheetGrid.GetGridCol(c);
  end;
end;

procedure TForm1.EdFormulaEditingDone(Sender: TObject);
var
  r, c: Cardinal;
  s: String;
begin
  r := WorksheetGrid.GetWorksheetRow(WorksheetGrid.Row);
  c := WorksheetGrid.GetWorksheetCol(WorksheetGrid.Col);
  s := EdFormula.Text;
  if (s <> '') and (s[1] = '=') then
    WorksheetGrid.Worksheet.WriteFormula(r, c, Copy(s, 2, Length(s)))
  else
    WorksheetGrid.Worksheet.WriteCellValueAsString(r, c, EdFormula.Text);
end;

procedure TForm1.EdFrozenColsChange(Sender: TObject);
begin
  WorksheetGrid.FrozenCols := EdFrozenCols.Value;
end;

procedure TForm1.EdFrozenRowsChange(Sender: TObject);
begin
  WorksheetGrid.FrozenRows := EdFrozenRows.Value;
end;

procedure TForm1.FontComboBoxSelect(Sender: TObject);
var
  fname: String;
begin
  fname := FontCombobox.Items[FontCombobox.ItemIndex];
  if fname <> '' then
    with WorksheetGrid do CellFontNames[Selection] := fName;
end;

procedure TForm1.FontSizeComboBoxSelect(Sender: TObject);
var
  sz: Integer;
begin
  sz := StrToInt(FontSizeCombobox.Items[FontSizeCombobox.ItemIndex]);
  if sz > 0 then
    with WorksheetGrid do CellFontSizes[Selection] := sz;
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

  CbBackgroundColor.ItemHeight := FontCombobox.ItemHeight;

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

  FontCombobox.DropDownCount := DROPDOWN_COUNT;
  FontSizeCombobox.DropDownCount := DROPDOWN_COUNT;
  CbBackgroundColor.DropDownCount := DROPDOWN_COUNT;
end;

procedure TForm1.InspectorPageControlChange(Sender: TObject);
var
  r,c: Cardinal;
  cell: PCell;
begin
  CellInspector.Parent := InspectorPageControl.ActivePage;

  r := WorksheetGrid.GetWorksheetRow(WorksheetGrid.Row);
  c := WorksheetGrid.GetWorksheetCol(WorksheetGrid.Col);
  cell := WorksheetGrid.Worksheet.FindCell(r, c);
  UpdateCellInfo(cell);
end;

procedure TForm1.LoadFile(const AFileName: String);
// Loads first worksheet from file into grid
var
  pages: TStrings;
  i: Integer;
  err: String;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    WorksheetGrid.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));

    // Update user interface
    Caption := Format('spready - %s (%s)', [
      AFilename,
      GetFileFormatName(WorksheetGrid.Workbook.FileFormat)
    ]);
    CbShowGridLines.Checked := (soShowGridLines in WorksheetGrid.Worksheet.Options);
    CbShowHeaders.Checked := (soShowHeaders in WorksheetGrid.Worksheet.Options);
    EdFrozenCols.Value := WorksheetGrid.FrozenCols;
    EdFrozenRows.Value := WorksheetGrid.FrozenRows;
    SetupBackgroundColorBox;

    // Create a tab in the pagecontrol for each worksheet contained in the workbook
    // This would be easier with a TTabControl. This has display issues, though.
    pages := TStringList.Create;
    try
      WorksheetGrid.GetSheets(pages);
      WorksheetGrid.Parent := PageControl1.Pages[0];
      while PageControl1.PageCount > pages.Count do PageControl1.Pages[1].Free;
      while PageControl1.PageCount < pages.Count do PageControl1.AddTabSheet;
      for i:=0 to PageControl1.PageCount-1 do
        PageControl1.Pages[i].Caption := pages[i];
    finally
      pages.Free;
    end;

    WorksheetGridSelection(nil, WorksheetGrid.Col, WorksheetGrid.Row);

  finally
    Screen.Cursor := crDefault;

    err := WorksheetGrid.Workbook.ErrorMsg;
    if err <> '' then
      MessageDlg(err, mtError, [mbOK], 0);
  end;
end;

procedure TForm1.PageControl1Change(Sender: TObject);
begin
  WorksheetGrid.Parent := PageControl1.Pages[PageControl1.ActivePageIndex];
  WorksheetGrid.SelectSheetByIndex(PageControl1.ActivePageIndex);
end;

procedure TForm1.SetupBackgroundColorBox;
begin
  // This change triggers re-reading of the workbooks palette by the OnGetColors
  // event of the ColorBox.
  CbBackgroundColor.Style := CbBackgroundColor.Style - [cbCustomColors];
  CbBackgroundColor.Style := CbBackgroundColor.Style + [cbCustomColors];
end;

procedure TForm1.WorksheetGridSelection(Sender: TObject; aCol, aRow: Integer);
var
  r, c: Cardinal;
  cell: PCell;
  s: String;
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  r := WorksheetGrid.GetWorksheetRow(ARow);
  c := WorksheetGrid.GetWorksheetCol(ACol);

  if AcCopyFormat.Checked then begin
    WorksheetGrid.Worksheet.CopyFormat(@FCopiedFormat, r, c);
    AcCopyFormat.Checked := false;
  end;

  cell := WorksheetGrid.Worksheet.FindCell(r, c);
  if cell <> nil then begin
    s := WorksheetGrid.Worksheet.ReadFormulaAsString(cell);
    if s <> '' then begin
      if s[1] <> '=' then s := '=' + s;
      EdFormula.Text := s;
    end
    else
      case cell^.ContentType of
        cctNumber:
          EdFormula.Text := FloatToStr(cell^.NumberValue);
        cctDateTime:
          if cell^.DateTimeValue < 1.0 then
            EdFormula.Text := FormatDateTime('tt', cell^.DateTimeValue)
          else
            EdFormula.Text := FormatDateTime('c', cell^.DateTimeValue);
        cctUTF8String:
          EdFormula.Text := cell^.UTF8StringValue;
        else
          EdFormula.Text := WorksheetGrid.Worksheet.ReadAsUTF8Text(cell);
      end;
  end else
    EdFormula.Text := '';

  EdCellAddress.Text := GetCellString(r, c, [rfRelRow, rfRelCol]);

  UpdateHorAlignmentActions;
  UpdateVertAlignmentActions;
  UpdateWordwraps;
  UpdateBackgroundColorIndex;
//  UpdateFontActions;
  UpdateFontNameIndex;
  UpdateFontSizeIndex;
  UpdateFontStyleActions;
  UpdateTextRotationActions;
  UpdateNumFormatActions;

  UpdateCellInfo(cell);

end;

procedure TForm1.UpdateBackgroundColorIndex;
var
  sClr: TsColor;
begin
  with WorksheetGrid do sClr := BackgroundColors[Selection];
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
  with WorksheetGrid do hor_align := HorAlignments[Selection];
  for i:=0 to ActionList.ActionCount-1 do begin
    ac := TAction(ActionList.Actions[i]);
    if (ac.Tag >= HORALIGN_TAG) and (ac.Tag < HORALIGN_TAG+10) then
      ac.Checked := ((ac.Tag - HORALIGN_TAG) = ord(hor_align));
  end;
end;

procedure TForm1.UpdateCellInfo(ACell: PCell);
var
  i: Integer;
  s: String;
  cb: TsCellBorder;
begin
  with CellInspector do begin
    TitleCaptions[0] := 'Properties';
    TitleCaptions[1] := 'Values';
    Strings.Clear;
    if InspectorPageControl.ActivePage = PgCellValue then
    begin
      if ACell=nil
        then Strings.Add('Row=')
        else Strings.Add(Format('Row=%d', [ACell^.Row]));
      if ACell=nil
        then Strings.Add('Column=')
        else Strings.Add(Format('Column=%d', [ACell^.Col]));
      if ACell=nil
        then Strings.Add('ContentType=')
        else Strings.Add(Format('ContentType=%s', [GetEnumName(TypeInfo(TCellContentType), ord(ACell^.ContentType))]));
      if ACell=nil
        then Strings.Add('NumberValue=')
        else Strings.Add(Format('NumberValue=%g', [ACell^.NumberValue]));
      if ACell=nil
        then Strings.Add('DateTimeValue=')
        else Strings.Add(Format('DateTimeValue=%g', [ACell^.DateTimeValue]));
      if ACell=nil
        then Strings.Add('UTF8StringValue=')
        else Strings.Add(Format('UTF8StringValue=%s', [ACell^.UTF8StringValue]));
      if ACell=nil
        then Strings.Add('BoolValue=')
        else Strings.Add(Format('BoolValue=%s', [BoolToStr(ACell^.BoolValue)]));
      if ACell=nil
        then Strings.Add('ErrorValue=')
        else Strings.Add(Format('ErrorValue=%s', [
               GetEnumName(TypeInfo(TsErrorValue), ord(ACell^.ErrorValue)) ]));
      if (ACell=nil) or (Length(ACell^.FormulaValue)=0)
        then Strings.Add('FormulaValue=')
        else Strings.Add(Format('FormulaValue="%s"', [ACell^.FormulaValue]));
      if (ACell=nil) or (ACell^.SharedFormulaBase=nil)
        then Strings.Add('SharedFormulaBase=')
        else Strings.Add(Format('SharedFormulaBase=%s', [GetCellString(
               ACell^.SharedFormulaBase^.Row, ACell^.SharedFormulaBase^.Col)]));
    end
    else
    begin
      if (ACell=nil) or not (uffFont in ACell^.UsedFormattingFields)
        then Strings.Add('FontIndex=')
        else Strings.Add(Format('FontIndex=%d (%s(', [
               ACell^.FontIndex,
               WorksheetGrid.Workbook.GetFontAsString(ACell^.FontIndex)]));
      if (ACell=nil) or not (uffTextRotation in ACell^.UsedFormattingFields)
        then Strings.Add('TextRotation=')
        else Strings.Add(Format('TextRotation=%s', [GetEnumName(TypeInfo(TsTextRotation), ord(ACell^.TextRotation))]));
      if (ACell=nil) or not (uffHorAlign in ACell^.UsedFormattingFields)
        then Strings.Add('HorAlignment=')
        else Strings.Add(Format('HorAlignment=%s', [GetEnumName(TypeInfo(TsHorAlignment), ord(ACell^.HorAlignment))]));
      if (ACell=nil) or not (uffVertAlign in ACell^.UsedFormattingFields)
        then Strings.Add('VertAlignment=')
        else Strings.Add(Format('VertAlignment=%s', [GetEnumName(TypeInfo(TsVertAlignment), ord(ACell^.VertAlignment))]));
      if (ACell=nil) or not (uffBorder in ACell^.UsedFormattingFields) then
        Strings.Add('Borders=')
      else begin
        s := '';
        if cbNorth in ACell^.Border then s := s + ', cbNorth';
        if cbSouth in ACell^.Border then s := s + ', cbSouth';
        if cbEast in ACell^.Border then s := s + ', cbEast';
        if cbWest in ACell^.Border then s := s + ', cbWest';
        if cbDiagUp in ACell^.Border then s := s + ', cbDiagUp';
        if cbDiagDown in ACell^.Border then s := s + ', cbDiagDown';
        if s <> '' then Delete(s, 1, 2);
        Strings.Add('Borders='+s);
      end;
      for cb in TsCellBorder do
        if ACell=nil then
          Strings.Add(Format('BorderStyles[%s]=', [
            GetEnumName(TypeInfo(TsCellBorder), ord(cb))
          ]))
        else
          Strings.Add(Format('BorderStyles[%s]=%s, %s', [
            GetEnumName(TypeInfo(TsCellBorder), ord(cb)),
            GetEnumName(TypeInfo(TsLineStyle), ord(ACell^.BorderStyles[cbEast].LineStyle)),
            WorksheetGrid.Workbook.GetColorName(ACell^.BorderStyles[cbEast].Color)
          ]));
      if (ACell=nil) or not (uffBackgroundColor in ACell^.UsedformattingFields)
        then Strings.Add('BackgroundColor=')
        else Strings.Add(Format('BackgroundColor=%d (%s)', [
               ACell^.BackgroundColor,
               WorksheetGrid.Workbook.GetColorName(Acell^.BackgroundColor)
             ]));
      if (ACell=nil) or not (uffNumberFormat in ACell^.UsedFormattingFields)
        then Strings.Add('NumberFormat=')
        else Strings.Add(Format('NumberFormat=%s', [GetEnumName(TypeInfo(TsNumberFormat), ord(ACell^.NumberFormat))]));
      if (ACell=nil) or not (uffNumberFormat in ACell^.UsedFormattingFields)
        then Strings.Add('NumberFormatStr=')
        else Strings.Add('NumberFormatStr=' + ACell^.NumberFormatStr);
    end;
  end;
end;

procedure TForm1.UpdateFontNameIndex;
var
  fname: String;
begin
  with WorksheetGrid do fname := CellFontNames[Selection];
  if fname = '' then
    FontCombobox.ItemIndex := -1
  else
    FontCombobox.ItemIndex := FontCombobox.Items.IndexOf(fname);
end;

procedure TForm1.UpdateFontSizeIndex;
var
  sz: Single;
begin
  with WorksheetGrid do sz := CellFontSizes[Selection];
  if sz < 0 then
    FontSizeCombobox.ItemIndex := -1
  else
    FontSizeCombobox.ItemIndex := FontSizeCombobox.Items.IndexOf(IntToStr(Round(sz)));
end;

procedure TForm1.UpdateFontStyleActions;
var
  style: TsFontStyles;
begin
  with WorksheetGrid do style := CellFontStyles[Selection];
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
  found: Boolean;
begin
  with WorksheetGrid do begin
    r := GetWorksheetRow(Row);
    c := GetWorksheetCol(Col);
    cell := Worksheet.FindCell(r, c);
    if (cell = nil) or not (cell^.ContentType in [cctNumber, cctDateTime]) then
      nf := nfGeneral
    else
      nf := cell^.NumberFormat;
    for i:=0 to ActionList.ActionCount-1 do begin
      ac := TAction(ActionList.Actions[i]);
      if (ac.Tag >= NUMFMT_TAG) and (ac.Tag < NUMFMT_TAG + 200) then begin
        found := ((ac.Tag - NUMFMT_TAG) div 10 = ord(nf));
        if nf = nfCustom then
          case (ac.Tag - NUMFMT_TAG) mod 10 of
            1: found := cell^.NumberFormatStr = 'dd/mmm';
            2: found := cell^.NumberFormatStr = 'mmm/yy';
            3: found := cell^.NumberFormatStr = 'nn:ss';
            4: found := cell^.NumberFormatStr = 'nn:ss.z';
          end;
        ac.Checked := found;
      end;
    end;
    Invalidate;
  end;
end;

procedure TForm1.UpdateTextRotationActions;
var
  i: Integer;
  ac: TAction;
  text_rot: TsTextRotation;
begin
  with WorksheetGrid do text_rot := TextRotations[Selection];
  for i:=0 to ActionList.ActionCount-1 do begin
    ac := TAction(ActionList.Actions[i]);
    if (ac.Tag >= TEXTROT_TAG) and (ac.Tag < TEXTROT_TAG+10) then
      ac.Checked := ((ac.Tag - TEXTROT_TAG) = ord(text_rot));
  end;
end;

procedure TForm1.UpdateVertAlignmentActions;
var
  i: Integer;
  ac: TAction;
  vert_align: TsVertAlignment;
begin
  with WorksheetGrid do vert_align := VertAlignments[Selection];
  for i:=0 to ActionList.ActionCount-1 do begin
    ac := TAction(ActionList.Actions[i]);
    if (ac.Tag >= VERTALIGN_TAG) and (ac.Tag < VERTALIGN_TAG+10) then
      ac.Checked := ((ac.Tag - VERTALIGN_TAG) = ord(vert_align));
  end;
end;

procedure TForm1.UpdateWordwraps;
var
  wrapped: Boolean;
begin
  with WorksheetGrid do wrapped := Wordwraps[Selection];
  AcWordwrap.Checked := wrapped;
end;

initialization
  {$I mainform.lrs}

end.

