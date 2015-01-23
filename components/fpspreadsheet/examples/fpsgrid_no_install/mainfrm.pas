unit mainfrm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, StdCtrls, Grids,
  fpspreadsheet, fpspreadsheetgrid, {%H-}fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    BtnNew: TButton;
    BtnLoad: TButton;
    BtnSave: TButton;
    ButtonPanel: TPanel;
    OpenDialog: TOpenDialog;
    SaveDialog: TSaveDialog;
    TabControl: TTabControl;
    procedure BtnLoadClick(Sender: TObject);
    procedure BtnNewClick(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
  private
    { private declarations }
    Grid: TsWorksheetGrid;
    procedure LoadFile(const AFileName: String);
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
begin
  Grid := TsWorksheetGrid.Create(self);

  // Put the grid into the TabControl and align it to fill the tabcontrol.
  Grid.Parent := TabControl;
  Grid.Align := alClient;

  // Useful options
  Grid.Options := Grid.Options + [goColSizing, goRowSizing,
    goFixedColSizing,    // useful if the spreadsheet contains frozen columns
    goEditing,           // needed for modifying cell content
    goThumbTracking,     // see the grid scroll while you drag the scrollbar
    goHeaderHotTracking, // hot-tracking of header cells
    goHeaderPushedLook,  // click at header cells --> pushed look
    goDblClickAutoSize   // optimum col width/row height after dbl click at header border
  ];
  Grid.AutoAdvance := aaDown;       // move active cell down on ENTER
  Grid.MouseWheelOption := mwGrid;  // mouse wheel scrolls the grid, not the active cell
  Grid.TextOverflow := true;        // too long text extends into neighbor cells

  // Create an empty worksheet
  Grid.NewWorkbook(26, 100);
end;

procedure TForm1.BtnLoadClick(Sender: TObject);
begin
  if OpenDialog.FileName <> '' then
  begin
    OpenDialog.InitialDir := ExtractFileDir(OpenDialog.FileName);
    OpenDialog.FileName := ChangeFileExt(ExtractFileName(OpenDialog.FileName), '');
  end;
  if OpenDialog.Execute then
  begin
    LoadFile(OpenDialog.FileName);
  end;
end;

procedure TForm1.BtnNewClick(Sender: TObject);
begin
  TabControl.Tabs.Clear;
  TabControl.Tabs.Add('Sheet1');
  Grid.NewWorkbook(26, 100);
end;

// Saves sheet in grid to file, overwriting existing file
procedure TForm1.BtnSaveClick(Sender: TObject);
var
  err: String;
begin
  if Grid.Workbook = nil then
    exit;

  if Grid.Workbook.Filename <>'' then
  begin
    SaveDialog.InitialDir := ExtractFileDir(Grid.Workbook.FileName);
    SaveDialog.FileName := ChangeFileExt(ExtractFileName(Grid.Workbook.FileName), '');
  end;

  if SaveDialog.Execute then
  begin
    Screen.Cursor := crHourglass;
    try
      Grid.SaveToSpreadsheetFile(SaveDialog.FileName);
    finally
      Screen.Cursor := crDefault;
      // Show a message in case of error(s)
      err := Grid.Workbook.ErrorMsg;
      if err <> '' then
        MessageDlg(err, mtError, [mbOK], 0);
    end;
  end;
end;

// Loads first worksheet from file into grid
procedure TForm1.LoadFile(const AFileName: String);
var
  err: String;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    try
      // Load file into workbook and grid
      Grid.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));

      // Update user interface
      Caption := Format('fpsGrid - %s (%s)', [
        AFilename,
        GetFileFormatName(Grid.Workbook.FileFormat)
      ]);

      // Collect the sheet names in the Tabs of the TabControl for switching sheets.
      Grid.GetSheets(TabControl.Tabs);
      TabControl.TabIndex := 0;
    except
      on E:Exception do begin
        // Empty worksheet instead of the loaded one
        Grid.NewWorkbook(26, 100);
        Caption := 'fpsGrid - no name';
        TabControl.Tabs.Clear;
        // Grab the error message
        Grid.Workbook.AddErrorMsg(E.Message);
      end;
    end;

  finally
    Screen.Cursor := crDefault;

    // Show a message in case of error(s)
    err := Grid.Workbook.ErrorMsg;
    if err <> '' then
      MessageDlg(err, mtError, [mbOK], 0);
  end;
end;

procedure TForm1.TabControlChange(Sender: TObject);
begin
  Grid.SelectSheetByIndex(TabControl.TabIndex);
end;


end.

