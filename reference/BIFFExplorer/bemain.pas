unit beMain;

{$mode objfpc}{$H+}

interface

uses
  ActnList, Classes, ComCtrls, ExtCtrls, Grids, Menus, StdCtrls, SysUtils,
  FileUtil, Forms, Controls, Graphics, Dialogs, Buttons, VirtualTrees,
  {$ifdef USE_NEW_OLE}
  fpolebasic,
  {$else}
  fpolestorage,
  {$endif}
  fpSpreadsheet,
  mrumanager, beBIFFGrid;

type
  { Virtual tree node data }
  TBiffNodeData = class
    Offset: Integer;
    RecordID: Integer;
    RecordName: String;
    RecordDescription: String;
    destructor Destroy; override;
  end;


  { TMainForm }
  TMainForm = class(TForm)
    AcFileOpen: TAction;
    AcFileQuit: TAction;
    AcFind: TAction;
    AcFindNext: TAction;
    AcFindPrev: TAction;
    AcAbout: TAction;
    AcFindClose: TAction;
    AcNodeExpand: TAction;
    AcNodeCollapse: TAction;
    ActionList: TActionList;
    BIFFTree: TVirtualStringTree;
    CbFind: TComboBox;
    HexGrid: TStringGrid;
    ImageList: TImageList;
    MainMenu: TMainMenu;
    AnalysisDetails: TMemo;
    MenuItem1: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MnuFind: TMenuItem;
    MnuRecord: TMenuItem;
    MnuFileReopen: TMenuItem;
    MenuItem4: TMenuItem;
    MnuHelp: TMenuItem;
    MenuItem2: TMenuItem;
    MnuFileQuit: TMenuItem;
    MnuFileOpen: TMenuItem;
    MnuFile: TMenuItem;
    OpenDialog: TOpenDialog;
    PageControl: TPageControl;
    DetailPanel: TPanel;
    HexPanel: TPanel;
    FindPanel: TPanel;
    TreePopupMenu: TPopupMenu;
    TreePanel: TPanel;
    BtnFindNext: TSpeedButton;
    BtnFindPrev: TSpeedButton;
    RecentFilesPopupMenu: TPopupMenu;
    BtnCloseFind: TSpeedButton;
    Splitter1: TSplitter;
    HexSplitter: TSplitter;
    AlphaGrid: TStringGrid;
    HexDumpSplitter: TSplitter;
    PgAnalysis: TTabSheet;
    PgValues: TTabSheet;
    DetailsSplitter: TSplitter;
    StatusBar: TStatusBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ValueGrid: TStringGrid;
    ToolBar: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    procedure AcAboutExecute(Sender: TObject);
    procedure AcFileOpenExecute(Sender: TObject);
    procedure AcFileQuitExecute(Sender: TObject);
    procedure AcFindCloseExecute(Sender: TObject);
    procedure AcFindExecute(Sender: TObject);
    procedure AcFindNextExecute(Sender: TObject);
    procedure AcFindPrevExecute(Sender: TObject);
    procedure AcNodeCollapseExecute(Sender: TObject);
    procedure AcNodeCollapseUpdate(Sender: TObject);
    procedure AcNodeExpandExecute(Sender: TObject);
    procedure AcNodeExpandUpdate(Sender: TObject);
    procedure AlphaGridSelection(Sender: TObject; aCol, aRow: Integer);
    procedure BIFFTreeBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
    procedure BIFFTreeFocusChanged(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Column: TColumnIndex);
    procedure BIFFTreeFreeNode(Sender: TBaseVirtualTree; Node: PVirtualNode);
    procedure BIFFTreeGetNodeDataSize(Sender: TBaseVirtualTree;
      var NodeDataSize: Integer);
    procedure BIFFTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
    procedure BIFFTreePaintText(Sender: TBaseVirtualTree;
      const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      TextType: TVSTTextType);
    procedure CbFindChange(Sender: TObject);
    procedure CbFindKeyPress(Sender: TObject; var Key: char);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GridClick(Sender: TObject);
    procedure HexGridPrepareCanvas(sender: TObject; aCol, aRow: Integer;
      aState: TGridDrawState);
    procedure HexGridSelection(Sender: TObject; aCol, aRow: Integer);
    procedure ListViewSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure PageControlChange(Sender: TObject);
    procedure ValueGridPrepareCanvas(sender: TObject; aCol, aRow: Integer;
      aState: TGridDrawState);

  private
    MemStream: TMemoryStream;
    OLEStorage: TOLEStorage;
    FFileName: String;
    FFormat: TsSpreadsheetFormat;
    FBuffer: TBIFFBuffer;
    FCurrOffset: Integer;
    FLockHexDumpGrids: Integer;
    FAnalysisGrid: TBIFFGrid;
    FMRUMenuManager : TMRUMenuManager;
    procedure AddToHistory(const AText: String);
    procedure AnalysisGridDetails(Sender: TObject; ADetails: TStrings);
    procedure AnalysisGridPrepareCanvas(sender: TObject; aCol, aRow: Integer;
      aState: TGridDrawState);
    procedure ExecFind(ANext, AKeep: Boolean);
    function  GetNodeData(ANode: PVirtualNode): TBiffNodeData;
    function  GetRecType: Word;
    procedure LoadFile(const AFileName: String); overload;
    procedure LoadFile(const AFileName: String; AFormat: TsSpreadsheetFormat); overload;
    procedure MRUMenuManagerRecentFile(Sender:TObject; const AFileName:string);
    procedure PopulateAnalysisGrid;
    procedure PopulateHexDump;
    procedure PopulateValueGrid;
    procedure ReadCmdLine;
    procedure ReadFromIni;
    procedure ReadFromStream(AStream: TStream);
    procedure UpdateCaption;
    procedure WriteToIni;

  public
    procedure BeforeRun;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

uses
  IniFiles, StrUtils, Math, lazutf8,
  fpsUtils,
  beUtils, beBIFFUtils, beAbout;

const
  VALUE_ROW_INDEX = 1;
  VALUE_ROW_BYTE = 2;
  VALUE_ROW_WORD = 3;
  VALUE_ROW_DWORD = 4;
  VALUE_ROW_QWORD = 5;
  VALUE_ROW_DOUBLE = 6;
  VALUE_ROW_ANSISTRING = 7;
  VALUE_ROW_WIDESTRING = 8;

  MAX_HISTORY = 16;


{ Virtual tree nodes }
type
  TObjectNodeData = record
    Data: TObject;
  end;
  PObjectNodeData = ^TObjectNodeData;


{ TBiffNodeData }

destructor TBiffNodeData.Destroy;
begin
  Finalize(RecordName);
  Finalize(RecordDescription);
  inherited;
end;


{ TMainForm }

procedure TMainForm.AcAboutExecute(Sender: TObject);
var
  F: TAboutForm;
begin
  F := TAboutForm.Create(nil);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;


procedure TMainForm.AcFileOpenExecute(Sender: TObject);
begin
  with OpenDialog do begin
    if Execute then LoadFile(FileName);
  end;
end;


procedure TMainForm.AcFileQuitExecute(Sender: TObject);
begin
  Close;
end;


procedure TMainForm.AcFindCloseExecute(Sender: TObject);
begin
  AcFind.Checked := false;
  FindPanel.Hide;
end;


procedure TMainForm.AcFindExecute(Sender: TObject);
begin
  if AcFind.Checked then begin
    FindPanel.Show;
    CbFind.SetFocus;
  end else begin
    FindPanel.Hide;
  end;
end;


procedure TMainForm.AcFindNextExecute(Sender: TObject);
begin
  ExecFind(true, false);
end;


procedure TMainForm.AcFindPrevExecute(Sender: TObject);
begin
  ExecFind(false, false);
end;


procedure TMainForm.AcNodeCollapseExecute(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
    BiffTree.Expanded[node] := false;
  end;
end;

procedure TMainForm.AcNodeCollapseUpdate(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
   end;
  AcNodeCollapse.Enabled := (node <> nil) and BiffTree.Expanded[node];
end;


procedure TMainForm.AcNodeExpandExecute(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
    BiffTree.Expanded[node] := true;
  end;
end;


procedure TMainForm.AcNodeExpandUpdate(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
  end;
  AcNodeExpand.Enabled := (node <> nil) and not BiffTree.Expanded[node];
end;

procedure TMainForm.AddToHistory(const AText: String);
begin
  if (AText <> '') and (CbFind.Items.IndexOf(AText) = -1) then begin
    CbFind.Items.Insert(0, AText);
    while CbFind.Items.Count > MAX_HISTORY do
      CbFind.Items.Delete(CbFind.Items.Count-1);
  end;
end;


procedure TMainForm.AlphaGridSelection(Sender: TObject; aCol, aRow: Integer);
begin
  if FLockHexDumpGrids > 0 then
    exit;
  FCurrOffset := (ARow - AlphaGrid.FixedRows)*16 + (ACol - AlphaGrid.FixedCols);
  if FCurrOffset < 0 then FCurrOffset := 0;
  inc(FLockHexDumpGrids);
  HexGrid.Col := aCol - AlphaGrid.FixedCols + HexGrid.FixedCols;
  HexGrid.Row := aRow - AlphaGrid.FixedRows + HexGrid.FixedRows;
  dec(FLockHexDumpGrids);
  if FCurrOffset > -1 then
    Statusbar.Panels[3].Text := Format('HexViewer offset: %d', [FCurrOffset])
  else
    Statusbar.Panels[3].Text := '';
end;


procedure TMainForm.AnalysisGridDetails(Sender: TObject; ADetails: TStrings);
begin
  AnalysisDetails.Lines.Assign(ADetails);
end;


procedure TMainForm.AnalysisGridPrepareCanvas(sender: TObject; aCol,
  aRow: Integer; aState: TGridDrawState);
begin
  if ARow = 0 then FAnalysisGrid.Canvas.Font.Style := [fsBold];
end;


procedure TMainForm.BeforeRun;
begin
  ReadFromIni;
  ReadCmdLine;
end;


procedure TMainForm.BIFFTreeBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
var
  s: String;
begin
  if (Sender.GetNodeLevel(Node) = 0) and (Column = 0) then begin
    // Left-align parent nodes (column 0 is right-aligned)
    BiffTreeGetText(Sender, Node, 0, ttNormal, s);
    TargetCanvas.Font.Style := [fsBold];
    ContentRect.Right := CellRect.Left + TargetCanvas.TextWidth(s) + 30;
  end;
end;


procedure TMainForm.BIFFTreeFocusChanged(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex);
var
  ptr: PObjectNodeData;
  data: TBiffNodeData;
  n: Word;
begin
  ptr := Sender.GetNodeData(Node);
  data := TBiffNodeData(ptr^.Data);

  // Move to start of record + 2 bytes to skip record type ID.
  MemStream.Position := PtrInt(data.Offset) + 2;

  // Read size of record
  n := WordLEToN(MemStream.ReadWord);

  // Read record data
  SetLength(FBuffer, n);
  if n > 0 then
    MemStream.ReadBuffer(FBuffer[0], n);

  // Update user interface
  if (BiffTree.FocusedNode <> nil) and (BiffTree.GetNodeLevel(BiffTree.FocusedNode) > 0)
  then begin
    Statusbar.Panels[0].Text := Format('Record ID: $%.4x', [data.RecordID]);
    Statusbar.Panels[1].Text := data.RecordName;
    Statusbar.Panels[2].Text := Format('Record size: %d bytes', [n]);
    Statusbar.Panels[3].Text := '';
  end else begin
    Statusbar.Panels[0].Text := '';
    Statusbar.Panels[1].Text := data.RecordName;
    Statusbar.Panels[2].Text := '';
    Statusbar.Panels[3].Text := '';
  end;
  PopulateHexDump;
  PageControlChange(nil);
end;


procedure TMainForm.BIFFTreeFreeNode(Sender: TBaseVirtualTree;
  Node: PVirtualNode);
var
  ptr: PObjectNodeData;
begin
  ptr := BiffTree.GetNodeData(Node);
  if ptr <> nil then
    FreeAndNil(ptr^.Data);
end;


procedure TMainForm.BIFFTreeGetNodeDataSize(Sender: TBaseVirtualTree;
  var NodeDataSize: Integer);
begin
  NodeDataSize := SizeOf(TObjectNodeData);
end;


procedure TMainForm.BIFFTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: String);
var
  ptr: PObjectNodeData;
  data: TBiffNodeData;
begin
  CellText := '';
  ptr := Sender.GetNodeData(Node);
  if ptr <> nil then begin
    data := TBiffNodeData(ptr^.Data);
    case Sender.GetNodeLevel(Node) of
      0: if Column = 0 then
           CellText := data.RecordName;
      1: case Column of
           0: CellText := IntToStr(data.Offset);
           1: CellText := Format('$%.4x', [data.RecordID]);
           2: CellText := data.RecordName;
           3: CellText := data.RecordDescription;
         end;
    end;
  end;
end;


procedure TMainForm.BIFFTreePaintText(Sender: TBaseVirtualTree;
  const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  TextType: TVSTTextType);
begin
  // Paint parent node in bold font.
  if (Sender.GetNodeLevel(Node) = 0) and (Column = 0) then
    TargetCanvas.Font.Style := [fsBold];
end;


procedure TMainForm.CbFindChange(Sender: TObject);
begin
  ExecFind(true, true);
end;


procedure TMainForm.CbFindKeyPress(Sender: TObject; var Key: char);
begin
  if Key = #13 then
    ExecFind(true, false);
end;


procedure TMainForm.ExecFind(ANext, AKeep: Boolean);
var
  s: String;
  node, node0: PVirtualNode;

  function GetRecordname(ANode: PVirtualNode; UseLowercase: Boolean = true): String;
  var
    data: TBIffNodeData;
  begin
    data := GetNodeData(ANode);
    if Assigned(data) then begin
      if UseLowercase then
        Result := lowercase(data.RecordName)
      else
        Result := data.RecordName;
    end else
      Result := '';
  end;

  function GetNextNode(ANode: PVirtualNode): PVirtualNode;
  var
    nextparent: PVirtualNode;
  begin
    Result := BiffTree.GetNextSibling(ANode);
    if (Result = nil) and (ANode <> nil) then begin
      nextparent := BiffTree.GetNextSibling(ANode^.Parent);
      if nextparent = nil then
        nextparent := BiffTree.GetFirst;
      Result := BiffTree.GetFirstChild(nextparent);
    end;
  end;

  function GetPrevNode(ANode: PVirtualNode): PVirtualNode;
  var
    prevparent: PVirtualNode;
  begin
    Result := BiffTree.GetPreviousSibling(ANode);
    if (Result = nil) and (ANode <> nil) then begin
      prevparent := BiffTree.GetPreviousSibling(ANode^.Parent);
      if prevparent = nil then
        prevparent := BiffTree.GetLast;
      Result := BiffTree.GetLastChild(prevparent);
    end;
  end;

begin
  if CbFind.Text = '' then
    exit;

  s := Lowercase(CbFind.Text);
  node0 := BiffTree.FocusedNode;
  if node0 = nil then
    node0 := BiffTree.GetFirst;
  if BiffTree.GetNodeLevel(node0) = 0 then
    node0 := BiffTree.GetFirstChild(node0);

  if ANext then begin
    if AKeep
      then node := node0
      else node := GetNextNode(node0);
    repeat
      if pos(s, GetRecordname(node)) > 0 then begin
        BiffTree.FocusedNode := node;
        BiffTree.Selected[node] := true;
        BiffTree.ScrollIntoView(node, true);
        AddToHistory(GetRecordname(node, false));
        exit;
      end;
      node := GetNextNode(node);
    until (node = node0) or (node = nil);
  end else begin
    if AKeep
      then node := node0
      else node := GetPrevNode(node0);
    repeat
      if pos(s, GetRecordName(node)) > 0 then begin
        BiffTree.FocusedNode := node;
        BiffTree.Selected[node] := true;
        BiffTree.ScrollIntoView(node, true);
        AddToHistory(GetRecordName(node, false));
        exit;
      end;
      node := GetPrevNode(node);
    until (node = node0) or (node = nil);
  end;
end;


procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if CanClose then
    try
      WriteToIni;
    except
      MessageDlg('Could not write setting to configuration file.', mtError, [mbOK], 0);
    end;
end;


procedure TMainForm.FormCreate(Sender: TObject);
begin
  FMRUMenuManager := TMRUMenuManager.Create(self);
  with FMRUMenuManager do begin
    Name := 'MRUMenuManager';
    IniFileName := GetAppConfigFile(false);
    IniSection := 'RecentFiles';
    MaxRecent := 16;
    MenuCaptionMask := '&%x - %s';    // & --> create hotkey
    MenuItem := MnuFileReopen;
    PopupMenu := RecentFilesPopupMenu;
    OnRecentFile := @MRUMenuManagerRecentFile;
  end;

  HexGrid.ColWidths[HexGrid.ColCount-1] := 5;
  HexGrid.DefaultRowHeight := HexGrid.Canvas.TextHeight('Tg') + 4;
  AlphaGrid.DefaultRowHeight := HexGrid.DefaultRowHeight;
  ValueGrid.DefaultRowHeight := HexGrid.DefaultRowHeight;
  BiffTree.DefaultNodeHeight := HexGrid.DefaultRowHeight;
  BiffTree.Header.DefaultHeight := HexGrid.DefaultRowHeight + 4;

  FAnalysisGrid := TBIFFGrid.Create(self);
  with FAnalysisGrid do begin
    Parent := PgAnalysis;
    Align := alClient;
    DefaultRowHeight := HexGrid.DefaultRowHeight;
    Options := Options + [goDrawFocusSelected];
    TitleStyle := tsNative;
    OnDetails := @AnalysisGridDetails;
    OnPrepareCanvas := @AnalysisGridPrepareCanvas;
  end;

  with ValueGrid do begin
    ColCount := 3;
    RowCount := VALUE_ROW_WIDESTRING+1;
    Cells[0, 0] := 'Data type';
    Cells[1, 0] := 'Value';
    Cells[2, 0] := 'Offset range';
    Cells[0, VALUE_ROW_INDEX] := 'Offset';
    Cells[0, VALUE_ROW_BYTE] := 'Byte';
    Cells[0, VALUE_ROW_WORD] := 'Word';
    Cells[0, VALUE_ROW_DWORD] := 'DWord';
    Cells[0, VALUE_ROW_QWORD] := 'QWord';
    Cells[0, VALUE_ROW_DOUBLE] := 'Double';
    Cells[0, VALUE_ROW_ANSISTRING] := 'AnsiString';
    Cells[0, VALUE_ROW_WIDESTRING] := 'WideString';
  end;
end;


procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if MemStream <> nil then
    FreeAndNil(MemStream);
  if OLEStorage <> nil then
    FreeAndNil(OLEStorage);
end;


procedure TMainForm.FormShow(Sender: TObject);
begin
  Width := Width + 1;     // remove black rectangle next to ValueGrid
  Width := Width - 1;
end;


function TMainForm.GetNodeData(ANode: PVirtualNode): TBiffNodeData;
var
  ptr: PObjectNodeData;
begin
  result := nil;
  if ANode <> nil then begin
    ptr := BiffTree.GetNodeData(ANode);
    if ptr <> nil then Result := TBiffNodeData(ptr^.Data);
  end;
end;


function TMainForm.GetRecType: Word;
var
  data: TBiffNodeData;
begin
  Result := Word(-1);
  if BiffTree.FocusedNode <> nil then begin
    data := GetNodeData(BiffTree.FocusedNode);
    if data <> nil then begin
      MemStream.Position := data.Offset;
      Result := WordLEToN(MemStream.ReadWord);
    end;
  end;
end;


procedure TMainForm.GridClick(Sender: TObject);
begin
  if PageControl.ActivePage = PgValues then
    PopulateValueGrid;
end;


procedure TMainForm.HexGridPrepareCanvas(sender: TObject; aCol, aRow: Integer;
  aState: TGridDrawState);
var
  ts: TTextStyle;
begin
  ts := HexGrid.Canvas.TextStyle;
  if ACol = 0 then
    ts.Alignment := taRightJustify
  else
    ts.Alignment := taCenter;
  ts.Opaque := false;
  ts.Layout := tlCenter;
  HexGrid.Canvas.TextStyle := ts;

  ts.Alignment := taCenter;
  AlphaGrid.Canvas.TextStyle := ts;
end;


procedure TMainForm.HexGridSelection(Sender: TObject; aCol, aRow: Integer);
begin
  if (FLockHexDumpGrids > 0) then
    exit;
  FCurrOffset := (ARow - HexGrid.FixedRows)*16 + (ACol - HexGrid.FixedCols);
  if FCurrOffset < 0 then
    FCurrOffset := 0;
  inc(FLockHexDumpGrids);
  AlphaGrid.Row := aRow - HexGrid.FixedRows + AlphaGrid.FixedRows;
  AlphaGrid.Col := aCol - HexGrid.FixedCols + AlphaGrid.FixedCols;
  dec(FLockHexDumpGrids);
  if FCurrOffset > -1 then
    Statusbar.Panels[3].Text := Format('HexViewer offset: %d', [FCurrOffset])
  else
    Statusbar.Panels[3].Text := '';
end;


procedure TMainForm.LoadFile(const AFileName: String);
var
  fmt: TsSpreadsheetFormat;
  valid: Boolean;
  excptn: Exception = nil;
begin
  if not FileExists(AFileName) then begin
    MessageDlg(Format('File "%s" not found.', [AFileName]), mtError, [mbOK], 0);
    exit;
  end;

  if not SameText(ExtractFileExt(AFileName), '.xls') then begin
    MessageDlg('BIFFExplorer can only process binary Excel files (extension ".xls")',
      mtError, [mbOK], 0);
    exit;
  end;

  fmt := sfExcel8;
  while True do begin
    try
      LoadFile(AFileName, fmt);
      valid := True;
    except
      on E: Exception do begin
        if fmt = sfExcel8 then excptn := E;
        valid := False
      end;
    end;
    if valid or (fmt = sfExcel2) then Break;
    fmt := Pred(fmt);
  end;

  // A failed attempt to read a file should bring an exception, so re-raise
  // the exception if necessary. We re-raise the exception brought by Excel 8,
  // since this is the most common format
  if (not valid) and (excptn <> nil) then
    raise excptn;

  FFormat := fmt;
end;


procedure TMainForm.LoadFile(const AFileName: String; AFormat: TsSpreadsheetFormat);
var
  OLEDocument: TOLEDocument;
begin
  if MemStream <> nil then
    FreeAndNil(MemStream);

  if OLEStorage <> nil then
    FreeAndNil(OLEStorage);

  MemStream := TMemoryStream.Create;

  if AFormat = sfExcel2 then begin
    MemStream.LoadFromFile(AFileName);
  end else begin
    OLEStorage := TOLEStorage.Create;

    // Only one stream is necessary for any number of worksheets
    OLEDocument.Stream := MemStream;
    OLEStorage.ReadOLEFile(AFileName, OLEDocument, 'Workbook');

    // Check if the operation succeded
    if MemStream.Size = 0 then
      raise Exception.Create('BIFF Explorer: Reading the OLE document failed');
  end;

  // Rewind the stream and read from it
  MemStream.Position := 0;
  FFileName := ExpandFileName(AFileName);
  ReadFromStream(MemStream);

  UpdateCaption;
  FMRUMenuManager.AddToRecent(AFileName);
end;


procedure TMainForm.ListViewSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  n: Word;
begin
  if Selected then begin
    // Move to start of record + 2 bytes to skip record type ID.
    MemStream.Position := PtrInt(Item.Data) + 2;

    // Read size of record
    n := WordLEToN(MemStream.ReadWord);

    // Read record data
    SetLength(FBuffer, n);
    MemStream.ReadBuffer(FBuffer[0], n);

    // Update user interface
    Statusbar.Panels[0].Text := Format('Record ID: %s', [Item.SubItems[0]]);
    Statusbar.Panels[1].Text := Item.SubItems[1];
    Statusbar.Panels[2].Text := Format('Record size: %s bytes', [Item.SubItems[3]]);
    PopulateHexDump;
    PageControlChange(nil);
  end;
end;


procedure TMainForm.MRUMenuManagerRecentFile(Sender: TObject;
  const AFileName: string);
begin
  LoadFile(AFileName);
end;


procedure TMainForm.PopulateAnalysisGrid;
begin
  FAnalysisGrid.SetRecordType(GetRecType, FBuffer, FFormat);
end;


procedure TMainForm.PopulateHexDump;
var
  n: Word;
  i,r,c, w: Integer;
begin
  n := Length(FBuffer);

  // Prepare hex viewer rows...
  r := HexGrid.FixedRows + n div 16;
  if n mod 16 <> 0 then inc(r);
  HexGrid.RowCount := r;
  AlphaGrid.RowCount := r;
  for i:=HexGrid.FixedRows to r-1 do begin
    HexGrid.Rows[i].Clear;
    HexGrid.Cells[0, i] := IntToStr((i - HexGrid.FixedRows)*16);
    AlphaGrid.Rows[i].Clear;
  end;

  // ... width of first column, ...
  w := HexGrid.Canvas.TextWidth(IntToStr(n)) + 4;
  if w > HexGrid.DefaultColWidth then
    HexGrid.ColWidths[0] := w;

  // ... and cells
  for i:=0 to Length(FBuffer)-1 do begin
    c := i mod 16;
    r := i div 16;
    with HexGrid do
      Cells[c+FixedCols, r+FixedRows] := Format('%.2x', [FBuffer[i]]);
    with AlphaGrid do
      if FBuffer[i] < 32 then
        Cells[c+FixedCols, r+FixedRows] := '.'
      else
        Cells[c+FixedCols, r+FixedRows] := AnsiToUTF8(AnsiChar(FBuffer[i]));
  end;

  // Clear value grid
  if PageControl.ActivePage = PgValues then
    with ValueGrid do begin
      Clean(1, 1, RowCount-1, ColCount-1, [gzNormal]);
    end;

  // Update status bar
  HexGridSelection(nil, HexGrid.Col, HexGrid.Row);
end;


procedure TMainForm.PopulateValueGrid;
var
  buf: array[0..1023] of Byte;
  w: word absolute buf;
  dw: DWord absolute buf;
  qw: QWord absolute buf;
  dbl: double absolute buf;
  idx: Integer;
  i, j: Integer;
  s: String;
  sw: WideString;
  ls: Integer;
begin
  idx := FCurrOffset;

  i := ValueGrid.RowCount;
  j := ValueGrid.ColCount;

  ValueGrid.Cells[1, VALUE_ROW_INDEX] := IntToStr(idx);

  if idx <= Length(FBuffer)-SizeOf(byte) then begin
    ValueGrid.Cells[1, VALUE_ROW_BYTE] := IntToStr(FBuffer[idx]);
    ValueGrid.Cells[2, VALUE_ROW_BYTE] := Format('%d ... %d', [idx, idx]);
  end
  else begin
    ValueGrid.Cells[1, VALUE_ROW_BYTE] := '';
    ValueGrid.Cells[2, VALUE_ROW_BYTE] := '';
  end;

  if idx <= Length(FBuffer)-SizeOf(word) then begin
    buf[0] := FBuffer[idx];
    buf[1] := FBuffer[idx+1];
    ValueGrid.Cells[1, VALUE_ROW_WORD] := IntToStr(WordLEToN(w));
    ValueGrid.Cells[2, VALUE_ROW_WORD] := Format('%d ... %d', [idx, idx+SizeOf(Word)-1]);
  end else begin
    ValueGrid.Cells[1, VALUE_Row_WORD] := '';
    ValueGrid.Cells[2, VALUE_ROW_WORD] := '';
  end;

  if idx <= Length(FBuffer) - SizeOf(DWord) then begin
    for i:=0 to SizeOf(DWord)-1 do buf[i] := FBuffer[idx+i];
    ValueGrid.Cells[1, VALUE_ROW_DWORD] := IntToStr(DWordLEToN(dw));
    ValueGrid.Cells[2, VALUE_ROW_DWORD] := Format('%d ... %d', [idx, idx+SizeOf(DWord)-1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_DWORD] := '';
    ValueGrid.Cells[2, VALUE_ROW_DWORD] := '';
  end;

  if idx <= Length(FBuffer) - SizeOf(QWord) then begin
    for i:=0 to SizeOf(QWord)-1 do buf[i] := FBuffer[idx+i];
    ValueGrid.Cells[1, VALUE_ROW_QWORD] := Format('%d', [qw]);
    ValueGrid.Cells[2, VALUE_ROW_QWORD] := Format('%d ... %d', [idx, idx+SizeOf(QWord)-1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_QWORD] := '';
    ValueGrid.Cells[2, VALUE_ROW_QWORD] := '';
  end;

  if idx <= Length(FBuffer) - SizeOf(double) then begin
    for i:=0 to SizeOf(double)-1 do buf[i] := FBuffer[idx+i];
    ValueGrid.Cells[1, VALUE_ROW_DOUBLE] := Format('%f', [dbl]);
    ValueGrid.Cells[2, VALUE_ROW_DOUBLE] := Format('%d ... %d', [idx, idx+SizeOf(Double)-1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_DOUBLE] := '';
    ValueGrid.Cells[2, VALUE_ROW_DOUBLE] := '';
  end;

  if idx < Length(FBuffer) then begin
    ls := FBuffer[idx];
    SetLength(s, ls);
    i := idx + 1;
    j := 0;
    while (i < Length(FBuffer)) and (j < Length(s)) do begin
      inc(j);
      s[j] := char(FBuffer[i]);
      inc(i);
    end;
    SetLength(s, j);
    ValueGrid.Cells[1, VALUE_ROW_ANSISTRING] := s;
    ValueGrid.Cells[2, VALUE_ROW_ANSISTRING] := Format('%d ... %d', [idx, ls * SizeOf(char) + 1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_ANSISTRING] := '';
    ValueGrid.Cells[2, VALUE_ROW_ANSISTRING] := '';
  end;

  if idx < Length(FBuffer) then begin
    ls := FBuffer[idx];
    SetLength(sw, ls);
    j := 0;
    i := idx + 2;
    while (i < Length(FBuffer)-1) and (j < Length(sw)) do begin
      buf[0] := FBuffer[i];
      buf[1] := FBuffer[i+1];
      inc(i, SizeOf(WideChar));
      inc(j);
      sw[j] := WideChar(w);
    end;
    SetLength(sw, j);
    ValueGrid.Cells[1, VALUE_ROW_WIDESTRING] := UTF8Decode(sw);
    ValueGrid.Cells[2, VALUE_ROW_WIDESTRING] := Format('%d ... %d', [idx, idx + ls*SizeOf(wideChar)+1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_WIDESTRING] := '';
    ValueGrid.Cells[2, VALUE_ROW_WIDESTRING] := '';
  end;
end;


procedure TMainForm.ReadCmdLine;
begin
  if ParamCount > 0 then
    LoadFile(ParamStr(1));
end;


procedure TMainForm.ReadFromIni;
var
  ini: TCustomIniFile;
  i: Integer;
begin
  ini := CreateIni;
  try
    ReadFormFromIni(ini, 'MainForm', self);

    BiffTree.Width := ini.ReadInteger('MainForm', 'RecordList_Width', BiffTree.Width);
    for i:=0 to BiffTree.Header.Columns.Count-1 do
      BiffTree.Header.Columns[i].Width := ini.ReadInteger('MainForm',
        Format('RecordList_ColWidth_%d', [i+1]), BiffTree.Header.Columns[i].Width);

    ValueGrid.Height := ini.ReadInteger('MainForm', 'ValueGrid_Height', ValueGrid.Height);
    for i:=0 to ValueGrid.ColCount-1 do
      ValueGrid.ColWidths[i] := ini.ReadInteger('MainForm',
        Format('ValueGrid_ColWidth_%d', [i+1]), ValueGrid.ColWidths[i]);

    AlphaGrid.Width := ini.ReadInteger('MainForm', 'AlphaGrid_Width', AlphaGrid.Width);
    for i:=0 to AlphaGrid.ColCount-1 do
      AlphaGrid.ColWidths[i] := ini.ReadInteger('MainForm',
        Format('AlphaGrid_ColWidth_%d', [i+1]), AlphaGrid.ColWidths[i]);

    for i:=0 to FAnalysisGrid.ColCount-1 do
      FAnalysisGrid.ColWidths[i] := ini.ReadInteger('MainForm',
        Format('AnalysisGrid_ColWidth_%d', [i+1]), FAnalysisGrid.ColWidths[i]);

    AnalysisDetails.Width := ini.ReadInteger('MainForm', 'AnalysisDetails_Width', AnalysisDetails.Width);

    PageControl.Height := ini.ReadInteger('MainForm', 'PageControl_Height', PageControl.Height);
    PageControl.ActivePageIndex := ini.ReadInteger('MainForm', 'PageIndex', PageControl.ActivePageIndex);
  finally
    ini.Free;
  end;
end;


procedure TMainForm.ReadFromStream(AStream: TStream);
var
  recType: Word;
  recSize: Word;
  p: Cardinal;
  p0: Cardinal;
  s: String;
  i: Integer;
  node: PVirtualNode;
  parentnode: PVirtualNode;
  ptr: PObjectNodeData;
  parentdata, data: TBiffNodeData;
  w: word;
  crs: TCursor;
begin
  crs := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;
    BiffTree.Clear;
    parentnode := nil;
    AStream.Position := 0;
    while AStream.Position < AStream.Size do begin
      p := AStream.Position;
      recType := WordLEToN(AStream.ReadWord);
      recSize := WordLEToN(AStream.ReadWord);
      s := RecTypeName(recType);
      i := pos(':', s);
      // in case of BOF record: create new parent node for this substream
      if (recType = $0009) or (recType = $0209) or (recType = $0409) or (recType = $0809)
      then begin
        // Read info on substream beginning here
        p0 := AStream.Position;
        AStream.Position := AStream.Position + 2;
        w := WordLEToN(AStream.ReadWord);
        AStream.Position := p0;
        parentdata := TBiffNodeData.Create;
        parentdata.Offset := p;
        parentdata.Recordname := BOFName(w);
        // add parent node for this substream
        parentnode := BIFFTree.AddChild(nil);
        ptr := BIFFTree.GetNodeData(parentnode);
        ptr^.Data := parentdata;
      end;
      // add node to parent node
      data := TBiffNodeData.Create;
      data.Offset := p;
      data.RecordID := recType;
      if i > 0 then begin
        data.RecordName := copy(s, 1, i-1);
        data.RecordDescription := copy(s, i+2, Length(s));
      end else begin
        data.RecordName := s;
        data.RecordDescription := '';
      end;
      node := BIFFTree.AddChild(parentnode);
      ptr := BIFFTree.GetNodeData(node);
      ptr^.Data := data;
      // advance stream pointer
      AStream.Position := AStream.Position + recSize;
    end;
    // expand all parent nodes
    node := BiffTree.GetFirst;
    while node <> nil do begin
      BiffTree.Expanded[node] := true;
      node := BiffTree.GetNextSibling(node);
    end;
    // Select first node
    BiffTree.FocusedNode := BiffTree.GetFirst;
    BiffTree.Selected[BiffTree.FocusedNode] := true;

  finally
    Screen.Cursor := crs;
  end;
end;


procedure TMainForm.PageControlChange(Sender: TObject);
begin
  if PageControl.ActivePage = PgAnalysis then
    PopulateAnalysisGrid
  else
  if PageControl.ActivePage = PgValues then
    PopulateValueGrid;
end;


procedure TMainForm.UpdateCaption;
begin
  Caption := Format('BIFF Explorer - "%s', [IfThen(FFileName <> '', FFileName, 'no file loaded')]);
end;


procedure TMainForm.ValueGridPrepareCanvas(sender: TObject; aCol,
  aRow: Integer; aState: TGridDrawState);
begin
  if ARow = 0 then ValueGrid.Canvas.Font.Style := [fsBold];
end;


procedure TMainForm.WriteToIni;
var
  ini: TCustomIniFile;
  i: Integer;
begin
  ini := CreateIni;
  try
    WriteFormToIni(ini, 'MainForm', self);

    ini.WriteInteger('MainForm', 'RecordList_Width', BiffTree.Width);
    for i:=0 to BiffTree.Header.Columns.Count-1 do
      ini.WriteInteger('MainForm', Format('RecordList_ColWidth_%d', [i+1]), BiffTree.Header.Columns[i].Width);

    ini.WriteInteger('MainForm', 'ValueGrid_Height', ValueGrid.Height);
    for i:=0 to ValueGrid.ColCount-1 do
      ini.WriteInteger('MainForm', Format('ValueGrid_ColWidth_%d', [i+1]), ValueGrid.ColWidths[i]);

    ini.WriteInteger('MainForm', 'AlphaGrid_Width', AlphaGrid.Width);
    for i:=0 to AlphaGrid.ColCount-1 do
      ini.WriteInteger('MainForm', Format('AlphaGrid_ColWidth_%d', [i+1]), AlphaGrid.ColWidths[i]);

    for i:=0 to FAnalysisGrid.ColCount-1 do
      ini.WriteInteger('MainForm', Format('AnalysisGrid_ColWidth_%d', [i+1]), FAnalysisGrid.ColWidths[i]);

    ini.WriteInteger('MainForm', 'AnalysisDetails_Width', AnalysisDetails.Width);

    ini.WriteInteger('MainForm', 'PageIndex', PageControl.ActivePageIndex);
    ini.WriteInteger('MainForm', 'PageControl_Height', PageControl.Height);
  finally
    ini.Free;
  end;
end;

end.

