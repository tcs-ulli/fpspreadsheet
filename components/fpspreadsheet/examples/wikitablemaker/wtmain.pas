unit wtMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ComCtrls, ActnList, Grids, ColorBox,
  SynEdit, SynEditHighlighter,
  SynHighlighterHTML, SynHighlighterMulti, SynHighlighterCss,
  fpspreadsheetgrid, fpspreadsheet, fpsallformats;

type

  { TMainFrm }

  TMainFrm = class(TForm)
    AcOpen: TAction;
    AcSaveAs: TAction;
    AcQuit: TAction;
    AcLeftAlign: TAction;
    AcHorCenterAlign: TAction;
    AcRightAlign: TAction;
    AcHorDefaultAlign: TAction;
    AcFontBold: TAction;
    AcFontItalic: TAction;
    AcFontStrikeout: TAction;
    AcFontUnderline: TAction;
    AcDefaultFont: TAction;
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
    AcCopyFormat: TAction;
    AcNew: TAction;
    AcAddColumn: TAction;
    AcAddRow: TAction;
    AcMergeCells: TAction;
    AcShowHeaders: TAction;
    AcShowGridlines: TAction;
    AcDeleteColumn: TAction;
    AcDeleteRow: TAction;
    AcCopyToClipboard: TAction;
    AcColumnTitles: TAction;
    AcRowTitles: TAction;
    AcVAlignDefault: TAction;
    AcVAlignTop: TAction;
    AcVAlignCenter: TAction;
    AcVAlignBottom: TAction;
    ActionList: TActionList;
    MnuBorderBottom: TMenuItem;
    MnuBorderBottomDbl: TMenuItem;
    MnuBorderBottomThick: TMenuItem;
    MnuBorderInner: TMenuItem;
    MnuBorderLeft: TMenuItem;
    MnuBorderRight: TMenuItem;
    MnuBordersAll: TMenuItem;
    MnuBordersInner: TMenuItem;
    MnuBordersOuter: TMenuItem;
    MnuBordersOuterThick: TMenuItem;
    MnuBordersSeparator1: TMenuItem;
    MnuBordersSeparator2: TMenuItem;
    MnuBordersSeparator3: TMenuItem;
    MnuBordersSeparator4: TMenuItem;
    MnuBordersSeparator5: TMenuItem;
    MnuBorderTop: TMenuItem;
    MnuBorderTopBottom: TMenuItem;
    MnuBorderTopBottomThick: TMenuItem;
    MnuBorderVCenter: TMenuItem;
    MnuFileSeparator1: TMenuItem;
    MnuNew: TMenuItem;
    MnuNoBorders: TMenuItem;
    MnuTableSeparator1: TMenuItem;
    ToolbarBevel: TBevel;
    CbBackgroundColor: TColorBox;
    FontComboBox: TComboBox;
    FontDialog: TFontDialog;
    FontSizeComboBox: TComboBox;
    ImageList: TImageList;
    MainMenu: TMainMenu;
    MnuRowHeaders: TMenuItem;
    MnuColHeaders: TMenuItem;
    MnuDeleteCol: TMenuItem;
    MnuTableSeparator2: TMenuItem;
    MnuAddRow: TMenuItem;
    MnuTableSeparator3: TMenuItem;
    MnuGridlines: TMenuItem;
    MnuAddCol: TMenuItem;
    MnuFormatSeparator: TMenuItem;
    MnuMergeCells: TMenuItem;
    MnuDeleteRow: TMenuItem;
    MnuLeftAlignment: TMenuItem;
    MnuCenterAlignment: TMenuItem;
    MnuRightAligment: TMenuItem;
    MnuHorAlignmentSeparator: TMenuItem;
    MnuVertAlignmentSeparator: TMenuItem;
    MnuVertBottom: TMenuItem;
    MnuVertCentered: TMenuItem;
    MnuVertTop: TMenuItem;
    MnuVertDefault: TMenuItem;
    MnuVertAlignment: TMenuItem;
    MnuFont: TMenuItem;
    MnuHorDefault: TMenuItem;
    MnuHorAlignment: TMenuItem;
    MnuFormat: TMenuItem;
    MnuTable: TMenuItem;
    MnuFile: TMenuItem;
    MnuOpen: TMenuItem;
    MnuQuit: TMenuItem;
    MnuSaveAs: TMenuItem;
    OpenDialog: TOpenDialog;
    BordersPopupMenu: TPopupMenu;
    PageControl: TPageControl;
    SaveDialog: TSaveDialog;
    SynCssSyn1: TSynCssSyn;
    SynEdit: TSynEdit;
    SynHTMLSyn1: TSynHTMLSyn;
    SynMultiSyn1: TSynMultiSyn;
    TabControl: TTabControl;
    PgTable: TTabSheet;
    PgCode: TTabSheet;
    CodeToolBar: TToolBar;
    TbDeleteColumn: TToolButton;
    TbAddRow: TToolButton;
    TbMergeCells: TToolButton;
    FormatToolBar: TToolBar;
    TbLeftAlign: TToolButton;
    TbFontStrikeout: TToolButton;
    TbHorCenterAlign: TToolButton;
    TbRightAlign: TToolButton;
    TbVAlignTop: TToolButton;
    TbVAlignCenter: TToolButton;
    TbVAlignBottom: TToolButton;
    TbBorders: TToolButton;
    TbCopyFormat: TToolButton;
    TbDefaultFont: TToolButton;
    TbDeleteRow: TToolButton;
    TbAddColumn: TToolButton;
    TbFontBold: TToolButton;
    TbFontItalic: TToolButton;
    TbFontUnderline: TToolButton;
    procedure AcAddColumnExecute(Sender: TObject);
    procedure AcAddRowExecute(Sender: TObject);
    procedure AcBorderExecute(Sender: TObject);
    procedure AcColumnTitlesExecute(Sender: TObject);
    procedure AcCopyFormatExecute(Sender: TObject);
    procedure AcCopyToClipboardExecute(Sender: TObject);
    procedure AcDeleteColumnExecute(Sender: TObject);
    procedure AcDeleteRowExecute(Sender: TObject);
    procedure AcDefaultFontExecute(Sender: TObject);
    procedure AcFontStyleExecute(Sender: TObject);
    procedure AcHorAlignmentExecute(Sender: TObject);
    procedure AcMergeCellsExecute(Sender: TObject);
    procedure AcNewExecute(Sender: TObject);
    procedure AcOpenExecute(Sender: TObject);
    procedure AcQuitExecute(Sender: TObject);
    procedure AcRowTitlesExecute(Sender: TObject);
    procedure AcSaveAsExecute(Sender: TObject);
    procedure AcShowGridlinesExecute(Sender: TObject);
    procedure AcShowHeadersExecute(Sender: TObject);
    procedure AcVertAlignmentExecute(Sender: TObject);
    procedure AcWordwrapExecute(Sender: TObject);
    procedure CbBackgroundColorSelect(Sender: TObject);
    procedure CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
    procedure FontComboBoxSelect(Sender: TObject);
    procedure FontSizeComboBoxSelect(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
    procedure WorksheetGridSelection(Sender: TObject; aCol, aRow: Integer);
  private
    WorksheetGrid: TsWorksheetGrid;
    FCopiedFormat: TCell;
    FHighlighter: TSynCustomHighlighter;
    procedure LoadFile(const AFileName: String);
    procedure SetupBackgroundColorBox;
    procedure UpdateBackgroundColorIndex;
    procedure UpdateFontNameIndex;
    procedure UpdateFontSizeIndex;
    procedure UpdateFontStyleActions;
    procedure UpdateHorAlignmentActions;
    procedure UpdateVertAlignmentActions;

  public
    procedure BeforeRun;

  end;

var
  MainFrm: TMainFrm;

implementation

uses
  TypInfo, LCLIntf, LCLType, clipbrd, fpcanvas,
  SynHighlighterWikiTable,
  fpsutils;

const
  DROPDOWN_COUNT = 24;

  HORALIGN_TAG   = 100;
  VERTALIGN_TAG  = 110;

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

procedure TMainFrm.AcColumnTitlesExecute(Sender: TObject);
begin
  if AcColumnTitles.Checked then
    WorksheetGrid.FrozenRows := 1
  else
    WorksheetGrid.FrozenRows := 0;
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

procedure TMainFrm.AcCopyToClipboardExecute(Sender: TObject);
begin
  if SynEdit.Lines.Count > 0 then
    Clipboard.AsText := SynEdit.Lines.Text;
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

{ Changes the default font of the workbook by calling a standard font dialog. }
procedure TMainFrm.AcDefaultFontExecute(Sender: TObject);
begin
  with WorksheetGrid do begin
    if Workbook = nil then
      exit;
    Convert_sFont_to_Font(Workbook.GetDefaultFont, FontDialog.Font);
    if FontDialog.Execute then begin
      Workbook.SetDefaultFont(FontDialog.Font.Name, FontDialog.Font.Size);
      Invalidate;
    end;
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

procedure TMainFrm.AcOpenExecute(Sender: TObject);
begin
  if OpenDialog.Execute then
    LoadFile(OpenDialog.FileName);
end;

procedure TMainFrm.AcQuitExecute(Sender: TObject);
begin
  Close;
end;

procedure TMainFrm.AcRowTitlesExecute(Sender: TObject);
begin
  if AcRowTitles.Checked then
    WorksheetGrid.FrozenCols := 1
  else
    WorksheetGrid.FrozenCols := 0;
end;

procedure TMainFrm.AcSaveAsExecute(Sender: TObject);
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

procedure TMainFrm.AcShowGridlinesExecute(Sender: TObject);
begin
  WorksheetGrid.ShowGridLines := AcShowGridLines.Checked;
end;

procedure TMainFrm.AcShowHeadersExecute(Sender: TObject);
begin
  WorksheetGrid.ShowHeaders := AcShowHeaders.Checked;
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

procedure TMainFrm.AcWordwrapExecute(Sender: TObject);
begin
  with WorksheetGrid do Wordwraps[Selection] := TAction(Sender).Checked;
end;

procedure TMainFrm.BeforeRun;
begin
  if ParamCount > 0 then
    LoadFile(ParamStr(1));
end;

procedure TMainFrm.CbBackgroundColorGetColors(Sender: TCustomColorBox; Items: TStrings);
var
  clr: TColor;
  clrName: String;
  i: Integer;
begin
  if (WorksheetGrid <> nil) and (WorksheetGrid.Workbook <> nil) then begin
    Items.Clear;
    Items.AddObject('no fill', TObject(PtrInt(clNone)));
    for i:=0 to WorksheetGrid.Workbook.GetPaletteSize-1 do begin
      clr := WorksheetGrid.Workbook.GetPaletteColor(i);
      clrName := WorksheetGrid.Workbook.GetColorName(i);
      Items.AddObject(Format('%d: %s', [i, clrName]), TObject(PtrInt(clr)));
    end;
  end;
end;

procedure TMainFrm.CbBackgroundColorSelect(Sender: TObject);
begin
  if CbBackgroundColor.ItemIndex <= 0 then
    with WorksheetGrid do BackgroundColors[Selection] := scNotDefined
  else
    with WorksheetGrid do BackgroundColors[Selection] := CbBackgroundColor.ItemIndex - 1;
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
  // Create the worksheet grid
  WorksheetGrid := TsWorksheetGrid.Create(self);
  with WorksheetGrid do begin
    Parent := TabControl;
    Align := alClient;
    AutoAdvance := aaDown;
    BorderStyle := bsNone;
    MouseWheelOption := mwGrid;
    Options := [goEditing, goFixedVertLine, goFixedHorzLine, goVertLine,
      goHorzLine, goRangeSelect, goRowSizing, goColSizing, goThumbTracking,
      goSmoothScroll, goFixedColSizing];
    TitleStyle := tsNative;
    OnSelection := @WorksheetGridSelection;
  end;

  // Create the syntax highlighter
  FHighlighter := TSynWikitableSyn.Create(self);
  SynEdit.Highlighter := FHighlighter;
//  SynEdit.Highlighter := SynCSSSyn1;

  // Adjust format toolbar height, looks strange at 120 dpi
  //FormatToolbar.Height := FontCombobox.Height + 2*FontCombobox.Top;
  //FormatToolbar.ButtonHeight := FormatToolbar.Height - 4;


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

  // Some initialization
  FontCombobox.Items.Assign(Screen.Fonts);          // Populate font combobox
  FontCombobox.DropDownCount := DROPDOWN_COUNT;
  FontSizeCombobox.DropDownCount := DROPDOWN_COUNT;
  CbBackgroundColor.DropDownCount := DROPDOWN_COUNT;
//  CbBackgroundColor.ItemHeight := FontCombobox.ItemHeight;
  CbBackgroundColor.ColorRectWidth := CbBackgroundColor.ItemHeight - 6; // to get a square box...

  // Initialize a new empty workbook
  AcNewExecute(nil);

  // Acitve control etc.
  PageControl.ActivePage := PgTable;
  ActiveControl := WorksheetGrid;
end;

procedure TMainFrm.PageControlChange(Sender: TObject);
var
  stream: TMemoryStream;
begin
  // Switch toolbars according to the selection of the pagecontrol
  CodeToolbar.Visible := PageControl.ActivePage = PgCode;
  FormatToolbar.Visible := PageControl.ActivePage = PgTable;
  ToolbarBevel.Top := Height;

  if (WorksheetGrid = nil) or (WorksheetGrid.Workbook = nil) then
    exit;

  if PageControl.ActivePage = PgCode then begin
    stream := TMemoryStream.Create;
    try
      WorksheetGrid.Workbook.WriteToStream(stream, sfWikitable_wikimedia);
      stream.Position := 0;
      SynEdit.Lines.LoadFromStream(stream);
    finally
      stream.Free;
    end;
  end;
end;

procedure TMainFrm.LoadFile(const AFileName: String);
// Loads first worksheet from file into grid
var
  pages: TStrings;
  i: Integer;
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
    AcRowTitles.Checked := WorksheetGrid.FrozenCols <> 0;
    AcColumnTitles.Checked := WorksheetGrid.FrozenRows <> 0;
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
  for i:=0 to ActionList.ActionCount-1 do begin
    ac := TAction(ActionList.Actions[i]);
    if (ac.Tag >= HORALIGN_TAG) and (ac.Tag < HORALIGN_TAG+10) then
      ac.Checked := ((ac.Tag - HORALIGN_TAG) = ord(hor_align));
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

procedure TMainFrm.WorksheetGridSelection(Sender: TObject; ACol, ARow: Integer);
var
  r, c: Cardinal;
  cell: PCell;
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
  AcMergeCells.Checked := WorksheetGrid.Worksheet.IsMerged(cell);

  UpdateHorAlignmentActions;
  UpdateVertAlignmentActions;
  UpdateBackgroundColorIndex;
  UpdateFontNameIndex;
  UpdateFontSizeIndex;
  UpdateFontStyleActions;
end;


initialization
  {$I wtmain.lrs}

end.

