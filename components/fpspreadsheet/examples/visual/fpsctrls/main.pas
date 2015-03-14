unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, ActnList, Menus, StdActns, Buttons,
  fpstypes, fpspreadsheet, fpspreadsheetctrls, fpspreadsheetgrid, fpsActions;

type

  { TMainForm }

  TMainForm = class(TForm)
    AcRowDelete: TAction;
    AcColDelete: TAction;
    AcRowAdd: TAction;
    AcColAdd: TAction;
    AcSettingsCSVParams: TAction;
    AcSettingsCurrency: TAction;
    AcSettingsFormatSettings: TAction;
    AcViewInspector: TAction;
    ActionList: TActionList;
    AcFileExit: TFileExit;
    AcFileOpen: TFileOpen;
    AcFileSaveAs: TFileSaveAs;
    ImageList: TImageList;
    MainMenu: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem100: TMenuItem;
    MenuItem101: TMenuItem;
    MenuItem102: TMenuItem;
    MenuItem103: TMenuItem;
    MenuItem104: TMenuItem;
    MenuItem105: TMenuItem;
    MenuItem106: TMenuItem;
    MenuItem107: TMenuItem;
    MenuItem108: TMenuItem;
    MenuItem109: TMenuItem;
    MnuSettings: TMenuItem;
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
    MenuItem3: TMenuItem;
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
    MenuItem4: TMenuItem;
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
    MenuItem5: TMenuItem;
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
    MenuItem75: TMenuItem;
    MenuItem76: TMenuItem;
    MenuItem77: TMenuItem;
    MenuItem78: TMenuItem;
    MenuItem79: TMenuItem;
    MenuItem80: TMenuItem;
    MenuItem81: TMenuItem;
    MenuItem82: TMenuItem;
    MenuItem83: TMenuItem;
    MenuItem84: TMenuItem;
    MenuItem85: TMenuItem;
    MenuItem86: TMenuItem;
    MenuItem87: TMenuItem;
    MenuItem88: TMenuItem;
    MenuItem89: TMenuItem;
    MenuItem9: TMenuItem;
    MenuItem90: TMenuItem;
    MenuItem91: TMenuItem;
    MenuItem95: TMenuItem;
    MenuItem96: TMenuItem;
    MenuItem97: TMenuItem;
    MenuItem98: TMenuItem;
    MenuItem99: TMenuItem;
    MnuColumn: TMenuItem;
    MenuItem93: TMenuItem;
    MenuItem94: TMenuItem;
    MnuAddWorksheet: TMenuItem;
    MnuRow: TMenuItem;
    MenuItem92: TMenuItem;
    MnuView: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MnuFormat: TMenuItem;
    MnuFile: TMenuItem;
    MnuWorksheet: TMenuItem;
    MnuAddSheet: TMenuItem;
    MnuEdit: TMenuItem;
    OpenDialog: TOpenDialog;
    OpenDialog1: TOpenDialog;
    CellEdit: TsCellEdit;
    CellIndicator: TsCellIndicator;
    AcFontBold: TsFontStyleAction;
    AcFontItalic: TsFontStyleAction;
    AcVertAlignTop: TsVertAlignmentAction;
    AcVertAlignCenter: TsVertAlignmentAction;
    AcVertAlignBottom: TsVertAlignmentAction;
    AcHorAlignLeft: TsHorAlignmentAction;
    AcHorAlignCenter: TsHorAlignmentAction;
    AcHorAlignRight: TsHorAlignmentAction;
    AcTextRotHor: TsTextRotationAction;
    AcTextRot90CW: TsTextRotationAction;
    AcTextRot90CCW: TsTextRotationAction;
    AcTextRotStacked: TsTextRotationAction;
    AcWordWrap: TsWordwrapAction;
    AcNumFormatFixed: TsNumberFormatAction;
    AcNumFormatFixedTh: TsNumberFormatAction;
    AcNumFormatPercentage: TsNumberFormatAction;
    AcNumFormatCurrency: TsNumberFormatAction;
    AcNumFormatCurrencyRed: TsNumberFormatAction;
    Panel2: TPanel;
    PuPaste: TPopupMenu;
    PuBorders: TPopupMenu;
    PuTimeFormat: TPopupMenu;
    PuDateFormat: TPopupMenu;
    PuCurrencyFormat: TPopupMenu;
    PuNumFormat: TPopupMenu;
    AcNumFormatGeneral: TsNumberFormatAction;
    AcNumFormatExp: TsNumberFormatAction;
    AcNumFormatDateTime: TsNumberFormatAction;
    AcNumFormatLongDate: TsNumberFormatAction;
    AcNumFormatShortDate: TsNumberFormatAction;
    AcNumFormatLongTime: TsNumberFormatAction;
    AcNumFormatShortTime: TsNumberFormatAction;
    AcNumFormatLongTimeAM: TsNumberFormatAction;
    AcNumFormatShortTimeAM: TsNumberFormatAction;
    AcNumFormatTimeInterval: TsNumberFormatAction;
    AcIncDecimals: TsDecimalsAction;
    AcDecDecimals: TsDecimalsAction;
    AcCellFontDialog: TsFontDialogAction;
    AcBackgroundColorDialog: TsBackgroundColorDialogAction;
    AcCellBorderTop: TsCellBorderAction;
    AcCellBorderBottom: TsCellBorderAction;
    AcCellBorderLeft: TsCellBorderAction;
    AcCellBorderRight: TsCellBorderAction;
    AcCellBorderInnerHor: TsCellBorderAction;
    AcCellBorderInnerVert: TsCellBorderAction;
    AcCellBorderAllHor: TsCellBorderAction;
    AcCellBorderBottomThick: TsCellBorderAction;
    AcCellBorderBottomDbl: TsCellBorderAction;
    AcCellBorderAllOuter: TsCellBorderAction;
    AcCellBorderNone: TsNoCellBordersAction;
    AcCellBorderAllOuterThick: TsCellBorderAction;
    AcCellBorderTopBottomThick: TsCellBorderAction;
    AcCellBorderTopBottomDbl: TsCellBorderAction;
    AcCellBorderAll: TsCellBorderAction;
    AcCellBorderAllVert: TsCellBorderAction;
    AcCopyFormat: TsCopyAction;
    FontColorCombobox: TsCellCombobox;
    BackgroundColorCombobox: TsCellCombobox;
    FontnameCombo: TsCellCombobox;
    FontsizeCombo: TsCellCombobox;
    AcMergeCells: TsMergeAction;
    AcCopyToClipboard: TsCopyAction;
    AcCutToClipboard: TsCopyAction;
    AcPasteAllFromClipboard: TsCopyAction;
    AcPasteValueFromClipboard: TsCopyAction;
    AcPasteFormatFromClipboard: TsCopyAction;
    AcPasteFormulaFromClipboard: TsCopyAction;
    AcCommentNew: TsCellCommentAction;
    AcCommentEdit: TsCellCommentAction;
    AcCommentDelete: TsCellCommentAction;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    ToolBar2: TToolBar;
    ToolBar3: TToolBar;
    ToolButton1: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    AcFontUnderline: TsFontStyleAction;
    AcFontStrikeout: TsFontStyleAction;
    InspectorSplitter: TSplitter;
    Inspector: TsSpreadsheetInspector;
    InspectorTabControl: TTabControl;
    AcAddWorksheet: TsWorksheetAddAction;
    AcDeleteWorksheet: TsWorksheetDeleteAction;
    acRenameWorksheet: TsWorksheetRenameAction;
    ToolBar1: TToolBar;
    ToolButton10: TToolButton;
    ToolButton2: TToolButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    ToolButton26: TToolButton;
    ToolButton27: TToolButton;
    ToolButton28: TToolButton;
    ToolButton29: TToolButton;
    ToolButton3: TToolButton;
    ToolButton30: TToolButton;
    ToolButton31: TToolButton;
    TbBorders: TToolButton;
    ToolButton32: TToolButton;
    ToolButton33: TToolButton;
    ToolButton34: TToolButton;
    ToolButton35: TToolButton;
    ToolButton36: TToolButton;
    ToolButton37: TToolButton;
    ToolButton38: TToolButton;
    ToolButton39: TToolButton;
    TbCommentAdd: TToolButton;
    ToolButton4: TToolButton;
    ToolButton40: TToolButton;
    ToolButton41: TToolButton;
    ToolButton42: TToolButton;
    ToolButton43: TToolButton;
    ToolButton44: TToolButton;
    ToolButton45: TToolButton;
    ToolButton46: TToolButton;
    ToolButton47: TToolButton;
    ToolButton48: TToolButton;
    ToolButton49: TToolButton;
    ToolButton5: TToolButton;
    TbCommentDelete: TToolButton;
    TbCommentEdit: TToolButton;
    ToolButton52: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    WorkbookSource: TsWorkbookSource;
    WorkbookTabControl: TsWorkbookTabControl;
    WorksheetGrid: TsWorksheetGrid;
    procedure AcColAddExecute(Sender: TObject);
    procedure AcColDeleteExecute(Sender: TObject);
    procedure AcFileOpenAccept(Sender: TObject);
    procedure AcFileSaveAsAccept(Sender: TObject);
    procedure AcRowAddExecute(Sender: TObject);
    procedure AcRowDeleteExecute(Sender: TObject);
    procedure AcSettingsCSVParamsExecute(Sender: TObject);
    procedure AcSettingsCurrencyExecute(Sender: TObject);
    procedure AcSettingsFormatSettingsExecute(Sender: TObject);
    procedure AcViewInspectorExecute(Sender: TObject);
    procedure InspectorTabControlChange(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure WorksheetGridClickHyperlink(Sender: TObject;
      const AHyperlink: TsHyperlink);
  private
    { private declarations }
    procedure UpdateCaption;
  public
    { public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

uses
  LCLProc, // debugln
  fpsUtils, fpsCSV,
  sCSVParamsForm, sCurrencyForm, sFormatSettingsForm, sSortParamsForm;


{ TMainForm }

{ Loads the spreadsheet file selected by the AcFileOpen action }
procedure TMainForm.AcFileOpenAccept(Sender: TObject);
var
  t: TTime;
begin
  WorkbookSource.AutodetectFormat := false;
  case AcFileOpen.Dialog.FilterIndex of
    1: WorkbookSource.AutoDetectFormat := true;      // All spreadsheet files
    2: WorkbookSource.AutoDetectFormat := true;      // All Excel files
    3: WorkbookSource.FileFormat := sfOOXML;         // Excel 2007+
    4: WorkbookSource.FileFormat := sfExcel8;        // Excel 97-2003
    5: WorkbookSource.FileFormat := sfExcel5;        // Excel 5.0
    6: WorkbookSource.FileFormat := sfExcel2;        // Excel 2.1
    7: WorkbookSource.FileFormat := sfOpenDocument;  // Open/LibreOffice
    8: WorkbookSource.FileFormat := sfCSV;           // Text files
  end;
  t := now;
  WorkbookSource.FileName := UTF8ToAnsi(AcFileOpen.Dialog.FileName);  // this loads the file
  t := (now - t)*24*3600;
  DebugLn(Format('Loading time for %s: %.3f sec', [AcFileOpen.Dialog.FileName, t]));
  UpdateCaption;
end;

{ Saves the spreadsheet to the file selected by the AcFileSaveAs action }
procedure TMainForm.AcFileSaveAsAccept(Sender: TObject);
var
  fmt: TsSpreadsheetFormat;
begin
  Screen.Cursor := crHourglass;
  try
    case AcFileSaveAs.Dialog.FilterIndex of
      1: fmt := sfOOXML;
      2: fmt := sfExcel8;
      3: fmt := sfExcel5;
      4: fmt := sfExcel2;
      5: fmt := sfOpenDocument;
      6: fmt := sfCSV;
      7: fmt := sfWikiTable_WikiMedia;
    end;
    WorkbookSource.SaveToSpreadsheetFile(UTF8ToAnsi(AcFileSaveAs.Dialog.FileName), fmt);
    UpdateCaption;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ Adds a column before the active cell }
procedure TMainForm.AcColAddExecute(Sender: TObject);
begin
  WorksheetGrid.InsertCol(WorksheetGrid.Col);
  WorksheetGrid.Col := WorksheetGrid.Col + 1;
end;

{ Deletes the column with the active cell }
procedure TMainForm.AcColDeleteExecute(Sender: TObject);
var
  c: Integer;
begin
  c := WorksheetGrid.Col;
  WorksheetGrid.DeleteCol(c);
  WorksheetGrid.Col := c;
end;

{ Adds a row before the active cell }
procedure TMainForm.AcRowAddExecute(Sender: TObject);
begin
  WorksheetGrid.InsertRow(WorksheetGrid.Row);
  WorksheetGrid.Row := WorksheetGrid.Row + 1;
end;

{ Deletes the row with the active cell }
procedure TMainForm.AcRowDeleteExecute(Sender: TObject);
var
  r: Integer;
begin
  r := WorksheetGrid.Row;
  WorksheetGrid.DeleteRow(r);
  WorksheetGrid.Row := r;
end;

procedure TMainForm.AcSettingsCSVParamsExecute(Sender: TObject);
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

procedure TMainForm.AcSettingsCurrencyExecute(Sender: TObject);
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

procedure TMainForm.AcSettingsFormatSettingsExecute(Sender: TObject);
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


{ Toggles the spreadsheet inspector on and off }
procedure TMainForm.AcViewInspectorExecute(Sender: TObject);
begin
  InspectorTabControl.Visible := AcViewInspector.Checked;
  InspectorSplitter.Visible := AcViewInspector.Checked;
  InspectorSplitter.Left := 0;
end;

{ Event handler to synchronize the mode of the spreadsheet inspector with the
  selected tab of the TabControl }
procedure TMainForm.InspectorTabControlChange(Sender: TObject);
begin
  Inspector.Mode := TsInspectorMode(InspectorTabControl.TabIndex);
end;

procedure TMainForm.ToolButton4Click(Sender: TObject);
begin
  WorkbookSource.Worksheet.WriteHyperlink(0, 0, '#Sheet2!B5', 'Go to B5');
end;

procedure TMainForm.WorksheetGridClickHyperlink(Sender: TObject;
  const AHyperlink: TsHyperlink);
begin
  ShowMessage('Hyperlink ' + AHyperlink.Target + ' clicked');
end;

procedure TMainForm.UpdateCaption;
begin
  if WorkbookSource = nil then
    Caption := 'demo_ctrls'
  else
    Caption := Format('demo_ctrls - "%s" [%s]', [
      AnsiToUTF8(WorkbookSource.Filename),
      GetFileFormatName(WorkbookSource.Workbook.FileFormat)
    ]);
end;

end.

