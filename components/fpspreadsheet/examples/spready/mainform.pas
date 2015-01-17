unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Spin, Grids,
  ColorBox, ValEdit,
  fpstypes, fpspreadsheetgrid, fpspreadsheet, {%H-}fpsallformats;

type

  { TMainFrm }

  TMainFrm = class(TForm)
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
    AcMergeCells: TAction;
    AcShowHeaders: TAction;
    AcShowGridlines: TAction;
    AcDeleteColumn: TAction;
    AcDeleteRow: TAction;
    AcCSVParams: TAction;
    AcFormatSettings: TAction;
    AcSortColAsc: TAction;
    AcSort: TAction;
    AcCurrencySymbols: TAction;
    AcViewInspector: TAction;
    AcWordwrap: TAction;
    AcVAlignDefault: TAction;
    AcVAlignTop: TAction;
    AcVAlignCenter: TAction;
    AcVAlignBottom: TAction;
    ActionList: TActionList;
    CbBackgroundColor: TColorBox;
    CbReadFormulas: TCheckBox;
    CbHeaderStyle: TComboBox;
    CbAutoCalcFormulas: TCheckBox;
    CbTextOverflow: TCheckBox;
    EdCellAddress: TEdit;
    FontComboBox: TComboBox;
    EdFrozenRows: TSpinEdit;
    FontDialog: TFontDialog;
    FontSizeComboBox: TComboBox;
    ImageList: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    MainMenu: TMainMenu;
    FormulaMemo: TMemo;
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
    MenuItem67: TMenuItem;
    MenuItem68: TMenuItem;
    MenuItem69: TMenuItem;
    MenuItem70: TMenuItem;
    MenuItem71: TMenuItem;
    MenuItem72: TMenuItem;
    MenuItem73: TMenuItem;
    MenuItem74: TMenuItem;
    MnuNumberFormatSettings: TMenuItem;
    MenuItem76: TMenuItem;
    MenuItem77: TMenuItem;
    MenuItem78: TMenuItem;
    MenuItem79: TMenuItem;
    MnuCurrencySymbol: TMenuItem;
    MnuCSVParams: TMenuItem;
    MnuSettings: TMenuItem;
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
    InspectorPageControl: TPageControl;
    Panel1: TPanel;
    BordersPopupMenu: TPopupMenu;
    NumFormatPopupMenu: TPopupMenu;
    AddressPanel: TPanel;
    SaveDialog: TSaveDialog;
    EdFrozenCols: TSpinEdit;
    FormulaToolBar: TToolBar;
    FormulaToolbarSplitter: TSplitter;
    InspectorSplitter: TSplitter;
    PgCellValue: TTabSheet;
    PgProperties: TTabSheet;
    Splitter1: TSplitter;
    TabControl: TTabControl;
    PgSheet: TTabSheet;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    ToolButton27: TToolButton;
    CellInspector: TValueListEditor;
    ToolButton28: TToolButton;
    ToolButton29: TToolButton;
    ToolButton30: TToolButton;
    ToolButton31: TToolButton;
    WorksheetGrid: TsWorksheetGrid;
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
    procedure AcCSVParamsExecute(Sender: TObject);
    procedure AcCurrencySymbolsExecute(Sender: TObject);
    procedure AcDeleteColumnExecute(Sender: TObject);
    procedure AcDeleteRowExecute(Sender: TObject);
    procedure AcEditExecute(Sender: TObject);
    procedure AcFontExecute(Sender: TObject);
    procedure AcFontStyleExecute(Sender: TObject);
    procedure AcFormatSettingsExecute(Sender: TObject);
    procedure AcHorAlignmentExecute(Sender: TObject);
    procedure AcIncDecDecimalsExecute(Sender: TObject);
    procedure AcMergeCellsExecute(Sender: TObject);
    procedure AcNewExecute(Sender: TObject);
    procedure AcNumFormatExecute(Sender: TObject);
    procedure AcOpenExecute(Sender: TObject);
    procedure AcQuitExecute(Sender: TObject);
    procedure AcSaveAsExecute(Sender: TObject);
    procedure AcShowGridlinesExecute(Sender: TObject);
    procedure AcShowHeadersExecute(Sender: TObject);
    procedure AcSortColAscExecute(Sender: TObject);
    procedure AcSortExecute(Sender: TObject);
    procedure AcTextRotationExecute(Sender: TObject);
    procedure AcVertAlignmentExecute(Sender: TObject);
    procedure AcViewInspectorExecute(Sender: TObject);
    procedure AcWordwrapExecute(Sender: TObject);
    procedure CbAutoCalcFormulasChange(Sender: TObject);
    procedure CbBackgroundColorSelect(Sender: TObject);
    procedure CbHeaderStyleChange(Sender: TObject);
    procedure CbReadFormulasChange(Sender: TObject);
    procedure CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
    procedure CbTextOverflowChange(Sender: TObject);
    procedure EdCellAddressEditingDone(Sender: TObject);
    procedure EdFrozenColsChange(Sender: TObject);
    procedure EdFrozenRowsChange(Sender: TObject);
    procedure FontComboBoxSelect(Sender: TObject);
    procedure FontSizeComboBoxSelect(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure InspectorPageControlChange(Sender: TObject);
    procedure MemoFormulaEditingDone(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
    procedure WorksheetGridHeaderClick(Sender: TObject; IsColumn: Boolean;
      Index: Integer);
    procedure WorksheetGridSelection(Sender: TObject; aCol, aRow: Integer);
  private
    FCopiedFormat: TCell;
    procedure LoadFile(const AFileName: String);
    procedure SetupBackgroundColorBox;
    procedure UpdateBackgroundColorIndex;
    procedure UpdateCellInfo(ACell: PCell);
    procedure UpdateFontNameIndex;
    procedure UpdateFontSizeIndex;
    procedure UpdateFontStyleActions;
    procedure UpdateHorAlignmentActions;
    procedure UpdateInspector;
    procedure UpdateNumFormatActions;
    procedure UpdateTextRotationActions;
    procedure UpdateVertAlignmentActions;
    procedure UpdateWordwraps;

  public
    procedure BeforeRun;

  end;

//  Excel 97-2003 spreadsheet (*.xls)|*.xls|Excel 5.0 spreadsheet (*.xls)|*.xls|Excel 2.1 spreadsheet (*.xls)|*.xls|Excel XML spreadsheet (*.xlsx)|*.xlsx|LibreOffice/OpenOffice spreadsheet (*.ods)|*.ods|Comma-delimited file (*.csv)|*.csv|Wikitable (wikimedia) (.wikitable_wikimedia)|*.wikitable_wikimedia
var
  MainFrm: TMainFrm;

implementation

uses
  TypInfo, LCLIntf, LCLType, LCLVersion, fpcanvas,
  fpsutils, fpscsv, fpsNumFormatParser,
  sFormatSettingsForm, sCSVParamsForm, sSortParamsForm, sfCurrencyForm;

const
  DROPDOWN_COUNT = 24;

  HORALIGN_TAG   = 100;
  VERTALIGN_TAG  = 110;
  TEXTROT_TAG    = 130;
  NUMFMT_TAG     = 1000;  // difference 10 per format item

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


{ TMainFrm }

procedure TMainFrm.AcEditExecute(Sender: TObject);
begin
  if AcEdit.Checked then
    WorksheetGrid.Options := WorksheetGrid.Options + [goEditing]
  else
    WorksheetGrid.Options := WorksheetGrid.Options - [goEditing];
end;

procedure TMainFrm.AcBorderExecute(Sender: TObject);
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

procedure TMainFrm.AcAddColumnExecute(Sender: TObject);
begin
  WorksheetGrid.InsertCol(WorksheetGrid.Col);
  WorksheetGrid.Col := WorksheetGrid.Col + 1;
end;

procedure TMainFrm.AcAddRowExecute(Sender: TObject);
begin
  WorksheetGrid.InsertRow(WorksheetGrid.Row);
  WorksheetGrid.Row := WorksheetGrid.Row + 1;
end;

procedure TMainFrm.AcCopyFormatExecute(Sender: TObject);
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

procedure TMainFrm.AcCSVParamsExecute(Sender: TObject);
var
  F: TCSVParamsForm;
begin
  F := TCSVParamsForm.Create(nil);
  try
    F.SetParams(fpscsv.CSVParams);
    if F.ShowModal = mrOK then
      F.GetParams(fpscsv.CSVParams);
  finally
    F.Free;
  end;
end;

procedure TMainFrm.AcCurrencySymbolsExecute(Sender: TObject);
var
  F: TCurrencyForm;
begin
  F := TCurrencyForm.Create(nil);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

procedure TMainFrm.AcDeleteColumnExecute(Sender: TObject);
var
  c: Integer;
begin
  c := WorksheetGrid.Col;
  WorksheetGrid.DeleteCol(c);
  WorksheetGrid.Col := c;
end;

procedure TMainFrm.AcDeleteRowExecute(Sender: TObject);
var
  r: Integer;
begin
  r := WorksheetGrid.Row;
  WorksheetGrid.DeleteRow(r);
  WorksheetGrid.Row := r;
end;

{ Changes the font of the selected cell by calling a standard font dialog. }
procedure TMainFrm.AcFontExecute(Sender: TObject);
begin
  with WorksheetGrid do begin
    if Workbook = nil then
      exit;
    FontDialog.Font := CellFonts[Selection];
    if FontDialog.Execute then
      CellFonts[Selection] := FontDialog.Font;
  end;
end;

procedure TMainFrm.AcFontStyleExecute(Sender: TObject);
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

procedure TMainFrm.AcFormatSettingsExecute(Sender: TObject);
var
  F: TFormatSettingsForm;
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  F := TFormatSettingsForm.Create(nil);
  try
    F.FormatSettings := WorksheetGrid.Workbook.FormatSettings;
    if F.ShowModal = mrOK then
    begin
      WorksheetGrid.Workbook.FormatSettings := F.FormatSettings;
      WorksheetGrid.Invalidate;
    end;
  finally
    F.Free;
  end;
end;

procedure TMainFrm.AcHorAlignmentExecute(Sender: TObject);
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

procedure TMainFrm.AcIncDecDecimalsExecute(Sender: TObject);
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

procedure TMainFrm.AcMergeCellsExecute(Sender: TObject);
begin
  AcMergeCells.Checked := not AcMergeCells.Checked;
  if AcMergeCells.Checked then
    WorksheetGrid.MergeCells
  else
    WorksheetGrid.UnmergeCells;
  WorksheetGridSelection(nil, WorksheetGrid.Col, WorksheetGrid.Row);
end;

procedure TMainFrm.AcNewExecute(Sender: TObject);
begin
  WorksheetGrid.NewWorkbook(26, 100);

  WorksheetGrid.BeginUpdate;
  try
    WorksheetGrid.Col := WorksheetGrid.FixedCols;
    WorksheetGrid.Row := WorksheetGrid.FixedRows;
    SetupBackgroundColorBox;
    WorksheetGridSelection(nil, WorksheetGrid.Col, WorksheetGrid.Row);
  finally
    WorksheetGrid.EndUpdate;
  end;
end;

procedure TMainFrm.AcNumFormatExecute(Sender: TObject);
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


procedure TMainFrm.acOpenExecute(Sender: TObject);
begin
  if OpenDialog.Execute then
    LoadFile(OpenDialog.FileName);
end;

procedure TMainFrm.acQuitExecute(Sender: TObject);
begin
  Close;
end;

procedure TMainFrm.acSaveAsExecute(Sender: TObject);
// Saves sheet in grid to file, overwriting existing file
var
  err: String = '';
  fmt: TsSpreadsheetFormat;
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  if SaveDialog.Execute then
  begin
    Screen.Cursor := crHourglass;
    case SaveDialog.FilterIndex of
      1: fmt := sfExcel8;
      2: fmt := sfExcel5;
      3: fmt := sfExcel2;
      4: fmt := sfOOXML;
      5: fmt := sfOpenDocument;
      6: fmt := sfCSV;
      7: fmt := sfWikiTable_wikimedia;
    end;
    try
      WorksheetGrid.SaveToSpreadsheetFile(SaveDialog.FileName, fmt);
    finally
      Screen.Cursor := crDefault;
      err := WorksheetGrid.Workbook.ErrorMsg;
      if err <> '' then
        MessageDlg(err, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TMainFrm.AcShowGridlinesExecute(Sender: TObject);
begin
  WorksheetGrid.ShowGridLines := AcShowGridLines.Checked;
end;

procedure TMainFrm.AcShowHeadersExecute(Sender: TObject);
begin
  WorksheetGrid.ShowHeaders := AcShowHeaders.Checked;
end;

procedure TMainFrm.AcSortColAscExecute(Sender: TObject);
var
  c, r: Cardinal;
  sortParams: TsSortParams;
begin
  r := WorksheetGrid.GetWorksheetRow(WorksheetGrid.Row);
  c := WorksheetGrid.GetWorksheetCol(WorksheetGrid.Col);
  sortParams := InitSortParams;
  WorksheetGrid.BeginUpdate;
  try
    with WorksheetGrid.Worksheet do
      Sort(sortParams, 0, c, GetLastOccupiedRowIndex, c);
  finally
    WorksheetGrid.EndUpdate;
  end;
end;

procedure TMainFrm.AcSortExecute(Sender: TObject);
var
  F: TSortParamsForm;
  r1,c1,r2,c2: Cardinal;
begin
  F := TSortParamsForm.Create(nil);
  try
    F.WorksheetGrid := WorksheetGrid;
    if F.ShowModal = mrOK then
    begin
      // Limits of the range to be sorted
      with WorksheetGrid do begin
        r1 := GetWorksheetRow(Selection.Top);
        c1 := GetWorksheetCol(Selection.Left);
        r2 := GetWorksheetRow(Selection.Bottom);
        c2 := GetWorksheetCol(Selection.Right);
      end;
      // Execute sorting. Use Begin/EndUpdate to avoid unnecessary redraws.
      WorksheetGrid.BeginUpdate;
      try
        WorksheetGrid.Worksheet.Sort(F.SortParams, r1, c1, r2, c2)
      finally
        WorksheetGrid.EndUpdate;
      end;
    end;
  finally
    F.Free;
  end;
end;

procedure TMainFrm.AcTextRotationExecute(Sender: TObject);
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

procedure TMainFrm.AcVertAlignmentExecute(Sender: TObject);
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

procedure TMainFrm.AcViewInspectorExecute(Sender: TObject);
begin
  InspectorPageControl.Visible := AcViewInspector.Checked;
  InspectorSplitter.Visible := AcViewInspector.Checked;
  InspectorSplitter.Left := 0;
end;

procedure TMainFrm.AcWordwrapExecute(Sender: TObject);
begin
  with WorksheetGrid do Wordwraps[Selection] := TAction(Sender).Checked;
end;

procedure TMainFrm.BeforeRun;
begin
  if ParamCount > 0 then
    LoadFile(ParamStr(1));
end;

procedure TMainFrm.CbAutoCalcFormulasChange(Sender: TObject);
begin
  WorksheetGrid.AutoCalc := CbAutoCalcFormulas.Checked;;
end;

procedure TMainFrm.CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
var
  clr: TColor;
  clrName: String;
  i: Integer;
begin
  if WorksheetGrid.Workbook <> nil then begin
    Items.Clear;
    Items.AddObject('no fill', TObject(PtrInt(clNone)));
    for i:=0 to WorksheetGrid.Workbook.GetPaletteSize-1 do begin
      clr := WorksheetGrid.Workbook.GetPaletteColor(i);
      clrName := WorksheetGrid.Workbook.GetColorName(i);
      Items.AddObject(Format('%d: %s', [i, clrName]), TObject(PtrInt(clr)));
    end;
  end;
end;

procedure TMainFrm.CbTextOverflowChange(Sender: TObject);
begin
  WorksheetGrid.TextOverflow := CbTextOverflow.Checked;
  WorksheetGrid.Invalidate;
end;

procedure TMainFrm.CbBackgroundColorSelect(Sender: TObject);
begin
  if CbBackgroundColor.ItemIndex <= 0 then
    with WorksheetGrid do BackgroundColors[Selection] := scNotDefined
  else
    with WorksheetGrid do BackgroundColors[Selection] := CbBackgroundColor.ItemIndex - 1;
end;

procedure TMainFrm.CbHeaderStyleChange(Sender: TObject);
begin
  WorksheetGrid.TitleStyle := TTitleStyle(CbHeaderStyle.ItemIndex);
end;

procedure TMainFrm.CbReadFormulasChange(Sender: TObject);
begin
  WorksheetGrid.ReadFormulas := CbReadFormulas.Checked;
end;

procedure TMainFrm.EdCellAddressEditingDone(Sender: TObject);
var
  c, r: cardinal;
begin
  if ParseCellString(EdCellAddress.Text, r, c) then begin
    WorksheetGrid.Row := WorksheetGrid.GetGridRow(r);
    WorksheetGrid.Col := WorksheetGrid.GetGridCol(c);
  end;
end;

procedure TMainFrm.EdFrozenColsChange(Sender: TObject);
begin
  WorksheetGrid.FrozenCols := EdFrozenCols.Value;
end;

procedure TMainFrm.EdFrozenRowsChange(Sender: TObject);
begin
  WorksheetGrid.FrozenRows := EdFrozenRows.Value;
end;

procedure TMainFrm.FontComboBoxSelect(Sender: TObject);
var
  fname: String;
begin
  fname := FontCombobox.Items[FontCombobox.ItemIndex];
  if fname <> '' then
    with WorksheetGrid do CellFontNames[Selection] := fName;
end;

procedure TMainFrm.FontSizeComboBoxSelect(Sender: TObject);
var
  sz: Integer;
begin
  sz := StrToInt(FontSizeCombobox.Items[FontSizeCombobox.ItemIndex]);
  if sz > 0 then
    with WorksheetGrid do CellFontSizes[Selection] := sz;
end;

procedure TMainFrm.FormActivate(Sender: TObject);
begin
  WorksheetGridSelection(nil, WorksheetGrid.Col, WorksheetGrid.Row);
end;

procedure TMainFrm.FormCreate(Sender: TObject);
begin
  // Adjust format toolbar height, looks strange at 120 dpi
//  FormatToolbar.Height := FontCombobox.Height + 2*FontCombobox.Top;
//  FormatToolbar.ButtonHeight := FormatToolbar.Height - 4;

  CbBackgroundColor.ItemHeight := FontCombobox.ItemHeight;
 {$IF LCL_FullVersion >= 1020000}
  CbBackgroundColor.ColorRectWidth := CbBackgroundColor.ItemHeight - 6; // to get a square box...
 {$ENDIF}

  InspectorPageControl.ActivePageIndex := 0;

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

  // Initialize a new empty workbook
  AcNewExecute(nil);

  ActiveControl := WorksheetGrid;
end;

procedure TMainFrm.InspectorPageControlChange(Sender: TObject);
begin
  CellInspector.Parent := InspectorPageControl.ActivePage;
  UpdateInspector;
end;

procedure TMainFrm.LoadFile(const AFileName: String);
// Loads first worksheet from file into grid
var
  err: String;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    try
      WorksheetGrid.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));
    except
      on E: Exception do begin
        // In an error occurs show at least an empty valid worksheet
        AcNewExecute(nil);
        MessageDlg(E.Message, mtError, [mbOk], 0);
        exit;
      end;
    end;

    // Update user interface
    Caption := Format('spready - %s (%s)', [
      AFilename,
      GetFileFormatName(WorksheetGrid.Workbook.FileFormat)
    ]);
    AcShowGridLines.Checked := WorksheetGrid.ShowGridLines;
    AcShowHeaders.Checked := WorksheetGrid.ShowHeaders;
    EdFrozenCols.Value := WorksheetGrid.FrozenCols;
    EdFrozenRows.Value := WorksheetGrid.FrozenRows;
    WorksheetGrid.TextOverflow := CbTextOverflow.Checked;
    SetupBackgroundColorBox;

    // Load names of worksheets into tabcontrol and show first sheet
    WorksheetGrid.GetSheets(TabControl.Tabs);
    TabControl.TabIndex := 0;
    // Update display
    WorksheetGridSelection(nil, WorksheetGrid.Col, WorksheetGrid.Row);

  finally
    Screen.Cursor := crDefault;

    err := WorksheetGrid.Workbook.ErrorMsg;
    if err <> '' then
      MessageDlg(err, mtError, [mbOK], 0);
  end;
end;

procedure TMainFrm.MemoFormulaEditingDone(Sender: TObject);
var
  r, c: Cardinal;
  s: String;
begin
  r := WorksheetGrid.GetWorksheetRow(WorksheetGrid.Row);
  c := WorksheetGrid.GetWorksheetCol(WorksheetGrid.Col);
  s := FormulaMemo.Lines.Text;
  if (s <> '') and (s[1] = '=') then
    WorksheetGrid.Worksheet.WriteFormula(r, c, Copy(s, 2, Length(s)), true)
  else
    WorksheetGrid.Worksheet.WriteCellValueAsString(r, c, s);
end;

procedure TMainFrm.SetupBackgroundColorBox;
begin
  // This change triggers re-reading of the workbooks palette by the OnGetColors
  // event of the ColorBox.
  CbBackgroundColor.Style := CbBackgroundColor.Style - [cbCustomColors];
  CbBackgroundColor.Style := CbBackgroundColor.Style + [cbCustomColors];
  Application.ProcessMessages;
end;

procedure TMainFrm.TabControlChange(Sender: TObject);
begin
  WorksheetGrid.SelectSheetByIndex(TabControl.TabIndex);
  WorksheetGridSelection(self, WorksheetGrid.Col, WorksheetGrid.Row);
end;

procedure TMainFrm.WorksheetGridHeaderClick(Sender: TObject; IsColumn: Boolean;
  Index: Integer);
begin
  Unused(Sender);
  Unused(IsColumn, Index);
  //ShowMessage('Header click');
end;

procedure TMainFrm.UpdateBackgroundColorIndex;
var
  sClr: TsColor;
begin
  with WorksheetGrid do sClr := BackgroundColors[Selection];
  if sClr = scNotDefined then
    CbBackgroundColor.ItemIndex := 0 // no fill
  else
    CbBackgroundColor.ItemIndex := sClr + 1;
end;

procedure TMainFrm.UpdateHorAlignmentActions;
var
  i: Integer;
  ac: TAction;
  hor_align: TsHorAlignment;
begin
  with WorksheetGrid do hor_align := HorAlignments[Selection];
  for i:=0 to ActionList.ActionCount-1 do
  begin
    ac := TAction(ActionList.Actions[i]);
    if (ac.Tag >= HORALIGN_TAG) and (ac.Tag < HORALIGN_TAG+10) then
      ac.Checked := ((ac.Tag - HORALIGN_TAG) = ord(hor_align));
  end;
end;

procedure TMainFrm.UpdateCellInfo(ACell: PCell);
var
  s: String;
  cb: TsCellBorder;
  r1,r2,c1,c2: Cardinal;
begin
  with CellInspector do
  begin
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
    if InspectorPageControl.ActivePage = PgSheet then
    begin
      if WorksheetGrid.Worksheet = nil then
      begin
        Strings.Add('First row=');
        Strings.Add('Last row=');
        Strings.Add('First column=');
        Strings.Add('Last column=');
      end else
      begin
        Strings.Add(Format('First row=%d', [WorksheetGrid.Worksheet.GetFirstRowIndex]));
        Strings.Add(Format('Last row=%d', [WorksheetGrid.Worksheet.GetLastRowIndex]));
        Strings.Add(Format('First column=%d', [WorksheetGrid.Worksheet.GetFirstColIndex]));
        Strings.Add(Format('Last column=%d', [WorksheetGrid.Worksheet.GetLastColIndex]));
      end;
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
      if not WorksheetGrid.Worksheet.IsMerged(ACell) then
        Strings.Add('Merged range=')
      else
      begin
        WorksheetGrid.Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
        Strings.Add('Merged range=' + GetCellRangeString(r1, c1, r2, c2));
      end;

    end;
  end;
end;

procedure TMainFrm.UpdateFontNameIndex;
var
  fname: String;
begin
  with WorksheetGrid do fname := CellFontNames[Selection];
  if fname = '' then
    FontCombobox.ItemIndex := -1
  else
    FontCombobox.ItemIndex := FontCombobox.Items.IndexOf(fname);
end;

procedure TMainFrm.UpdateFontSizeIndex;
var
  sz: Single;
begin
  with WorksheetGrid do sz := CellFontSizes[Selection];
  if sz < 0 then
    FontSizeCombobox.ItemIndex := -1
  else
    FontSizeCombobox.ItemIndex := FontSizeCombobox.Items.IndexOf(IntToStr(Round(sz)));
end;

procedure TMainFrm.UpdateFontStyleActions;
var
  style: TsFontStyles;
begin
  with WorksheetGrid do style := CellFontStyles[Selection];
  AcFontBold.Checked := fssBold in style;
  AcFontItalic.Checked := fssItalic in style;
  AcFontUnderline.Checked := fssUnderline in style;
  AcFontStrikeout.Checked := fssStrikeOut in style;
end;

procedure TMainFrm.UpdateInspector;
var
  r, c: Cardinal;
  cell: PCell;
begin
  r := WorksheetGrid.GetWorksheetRow(WorksheetGrid.Row);
  c := WorksheetGrid.GetWorksheetCol(WorksheetGrid.Col);
  cell := WorksheetGrid.Worksheet.FindCell(r, c);
  UpdateCellInfo(cell);
end;

procedure TMainFrm.UpdateNumFormatActions;
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

procedure TMainFrm.UpdateTextRotationActions;
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

procedure TMainFrm.UpdateVertAlignmentActions;
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

procedure TMainFrm.UpdateWordwraps;
var
  wrapped: Boolean;
begin
  with WorksheetGrid do wrapped := Wordwraps[Selection];
  AcWordwrap.Checked := wrapped;
end;

procedure TMainFrm.WorksheetGridSelection(Sender: TObject; ACol, ARow: Integer);
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
    s := WorksheetGrid.Worksheet.ReadFormulaAsString(cell, true);
    if s <> '' then begin
      if s[1] <> '=' then s := '=' + s;
      FormulaMemo.Lines.Text := s;
    end else
    begin
      case cell^.ContentType of
        cctNumber:
          s := FloatToStr(cell^.NumberValue);
        cctDateTime:
          if cell^.DateTimeValue < 1.0 then
            s := FormatDateTime('tt', cell^.DateTimeValue)
          else
            s := FormatDateTime('c', cell^.DateTimeValue);
        cctUTF8String:
          s := cell^.UTF8StringValue;
        else
          s := WorksheetGrid.Worksheet.ReadAsUTF8Text(cell);
      end;
      FormulaMemo.Lines.Text := s;
    end;
  end else
    FormulaMemo.Text := '';

  EdCellAddress.Text := GetCellString(r, c, [rfRelRow, rfRelCol]);
  AcMergeCells.Checked := WorksheetGrid.Worksheet.IsMerged(cell);

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


initialization
  {$I mainform.lrs}

end.

