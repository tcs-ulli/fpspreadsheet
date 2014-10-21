unit sSortParamsForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  ButtonPanel, Grids, ExtCtrls, Buttons, StdCtrls,
  fpspreadsheet, fpspreadsheetgrid;

type

  { TSortParamsForm }

  TSortParamsForm = class(TForm)
    BtnAdd: TBitBtn;
    BtnDelete: TBitBtn;
    ButtonPanel: TButtonPanel;
    CbSortColsRows: TComboBox;
    Panel1: TPanel;
    Grid: TStringGrid;
    procedure BtnAddClick(Sender: TObject);
    procedure BtnDeleteClick(Sender: TObject);
    procedure CbSortColsRowsChange(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
  private
    { private declarations }
    FWorksheetGrid: TsWorksheetGrid;
    function GetSortByCols: Boolean;
    function GetSortIndex: TsIndexArray;
    procedure SetWorksheetGrid(AValue: TsWorksheetGrid);
    procedure UpdateColRowList;
    procedure UpdateCmds;
    function ValidParams(out AMsg: String): Boolean;
  public
    { public declarations }
    property SortByCols: Boolean read GetSortByCols;
    property SortIndex: TsIndexArray read GetSortIndex;
    property WorksheetGrid: TsWorksheetGrid read FWorksheetGrid write SetWorksheetGrid;
  end;

var
  SortParamsForm: TSortParamsForm;

implementation

uses
  fpsutils;

procedure TSortParamsForm.CbSortColsRowsChange(Sender: TObject);
begin
  UpdateColRowList;
  UpdateCmds;
end;

procedure TSortParamsForm.OKButtonClick(Sender: TObject);
var
  msg: String;
begin
  if not ValidParams(msg) then begin
    MessageDlg(msg, mtError, [mbOK], 0);
    ModalResult := mrNone;
  end;
end;

procedure TSortParamsForm.BtnAddClick(Sender: TObject);
var
  numConditions: Integer;
begin
  case CbSortColsRows.ItemIndex of
    0: numConditions := FWorksheetGrid.Selection.Right - FWorksheetGrid.Selection.Left + 1;
    1: numConditions := FWorksheetGrid.Selection.Bottom - FWorksheetGrid.Selection.Top + 1;
  end;
  if Grid.RowCount - Grid.FixedRows >= numConditions then
    exit;  // there can't be more conditions than defined by the worksheetgrid selection
  Grid.RowCount := Grid.RowCount + 1;
  Grid.Cells[0, Grid.RowCount-1] := 'Then by';
  UpdateCmds;
end;

procedure TSortParamsForm.BtnDeleteClick(Sender: TObject);
begin
  if Grid.RowCount = Grid.FixedRows + 1 then
    exit;  // 1 condition must remain
  Grid.DeleteRow(Grid.Row);
  Grid.Cells[0, 1] := 'Sort by';
  UpdateCmds;
end;

function TSortParamsForm.GetSortByCols: Boolean;
begin
  Result := CbSortColsRows.ItemIndex = 0;
end;

function TSortParamsForm.GetSortIndex: TsIndexArray;
var
  i, p: Integer;
  s: String;
  n: Cardinal;
begin
  SetLength(Result, 0);
  s:= Grid.Cells[0, 0];
  s := Grid.Cells[0, 1];
  for i:= Grid.FixedRows to Grid.RowCount-1 do
  begin
    s := Grid.Cells[1, i];
    if s <> '' then
    begin
      p := pos(' ', s);
      s := Copy(s, p+1, Length(s));
      case CbSortColsRows.ItemIndex of
        0: if not ParseCellColString(s, n) then continue;     // row index
        1: if not TryStrToInt(s, LongInt(n)) then continue else dec(n);   // column index
      end;
      SetLength(Result, Length(Result)+1);
      Result[Length(Result)-1] := n;
    end;
  end;
end;

procedure TSortParamsForm.SetWorksheetGrid(AValue: TsWorksheetGrid);
begin
  FWorksheetGrid := AValue;
  UpdateColRowList;
  UpdateCmds;
end;

procedure TSortParamsForm.UpdateColRowList;
var
  r,c, r1,c1, r2,c2: Cardinal;
  L: TStrings;
begin
  with FWorksheetGrid do begin
    r1 := GetWorksheetRow(Selection.Top);
    c1 := GetWorksheetCol(Selection.Left);
    r2 := GetWorksheetRow(Selection.Bottom);
    c2 := GetWorksheetCol(Selection.Right);
  end;
  L := TStringList.Create;
  try
    case CbSortColsRows.ItemIndex of
      0: begin
           Grid.RowCount := Grid.FixedRows + 1;
           Grid.Columns[0].Title.Caption := 'Columns';
           for c := c1 to c2 do
             L.Add('Column ' + GetColString(c));
         end;
      1: begin
           Grid.RowCount := Grid.FixedRows + 1;
           Grid.Columns[0].Title.Caption := 'Rows';
           for r := r1 to r2 do
             L.Add('Row ' + IntToStr(r+1));
         end;
    end;
    Grid.Columns[0].PickList.Assign(L);
    for r := Grid.FixedRows to Grid.RowCount-1 do
    begin
      Grid.Cells[1, r] := '';
      Grid.Cells[2, r] := ''
    end;
  finally
    L.Free;
  end;
end;

procedure TSortParamsForm.UpdateCmds;
var
  r1,c1,r2,c2: Cardinal;
  numConditions: Integer;
begin
  with FWorksheetGrid do begin
    r1 := GetWorksheetRow(Selection.Top);
    c1 := GetWorksheetCol(Selection.Left);
    r2 := GetWorksheetRow(Selection.Bottom);
    c2 := GetWorksheetCol(Selection.Right);
  end;
  numConditions := Grid.RowCount - Grid.FixedRows;
  case CbSortColsRows.ItemIndex of
    0: BtnAdd.Enabled := numConditions < c2-c1+1;
    1: BtnAdd.Enabled := numConditions < r2-r1+1;
  end;
  BtnDelete.Enabled := numConditions > 1;
end;

function TSortParamsForm.ValidParams(out AMsg: String): Boolean;
begin
  Result := false;
  if Length(SortIndex) = 0 then
  begin
    AMsg := 'No sorting criteria selected.';
    Grid.SetFocus;
    exit;
  end;
  Result := true;
end;

initialization
  {$I ssortparamsform.lrs}

end.

