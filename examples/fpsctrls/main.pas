unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, ComCtrls, ActnList, Menus, StdActns, 
  fpspreadsheet, fpspreadsheetctrls, fpspreadsheetgrid, fpsActions;

type

  { TForm1 }

  TForm1 = class(TForm)
    ActionList: TActionList;
    Button1: TButton;
    AcFileExit: TFileExit;
    ImageList: TImageList;
    MainMenu: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MnuFile: TMenuItem;
    MnuWorksheet: TMenuItem;
    MnuAddSheet: TMenuItem;
    MnuEdit: TMenuItem;
    OpenDialog: TOpenDialog;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
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
    Splitter1: TSplitter;
    Inspector: TsSpreadsheetInspector;
    InspectorTabControl: TTabControl;
    AcAddWorksheet: TsWorksheetAddAction;
    AcDeleteWorksheet: TsWorksheetDeleteAction;
    acRenameWorksheet: TsWorksheetRenameAction;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    WorkbookSource: TsWorkbookSource;
    WorkbookTabControl: TsWorkbookTabControl;
    WorksheetGrid: TsWorksheetGrid;
    procedure Button1Click(Sender: TObject);
    procedure InspectorTabControlChange(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
begin
  if OpenDialog.Execute then begin
    WorkbookSource.AutodetectFormat := false;
    case OpenDialog.FilterIndex of
      1: WorkbookSource.AutoDetectFormat := true;      // All spreadsheet files
      2: WorkbookSource.AutoDetectFormat := true;      // All Excel files
      3: WorkbookSource.FileFormat := sfOOXML;         // Excel 2007+
      4: WorkbookSource.FileFormat := sfExcel8;        // Excel 97-2003
      5: WorkbookSource.FileFormat := sfExcel5;        // Excel 5.0
      6: WorkbookSource.FileFormat := sfExcel2;        // Excel 2.1
      7: WorkbookSource.FileFormat := sfOpenDocument;  // Open/LibreOffice
      8: WorkbookSource.FileFormat := sfCSV;           // Text files
    end;
    WorkbookSource.FileName := OpenDialog.FileName;    // this loads the file
  end;
end;

procedure TForm1.InspectorTabControlChange(Sender: TObject);
begin
  Inspector.Mode := TsInspectorMode(InspectorTabControl.TabIndex);
end;

end.

