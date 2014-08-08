unit errortests;

{$mode objfpc}{$H+}

{ Tests for error logging by readers / writers }

interface

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet {and a project requirement for lclbase for utf8 handling},
  fpsutils, testsutility;

type
  { TSpreadErrorTests }

  TSpreadErrorTests= class(TTestCase)
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteErrorMessages(AFormat: TsSpreadsheetFormat);

  published
    // Tests collection of error messages during writing
    procedure TestWriteErrorMessages_BIFF2;
    procedure TestWriteErrorMessages_BIFF5;
    procedure TestWriteErrorMessages_BIFF8;
    procedure TestWriteErrorMessages_ODS;
    procedure TestWriteErrorMessages_OOXML;
  end;

implementation

uses
  StrUtils;

const
  ERROR_SHEET = 'ErrorTest'; //worksheet name

procedure TSpreadErrorTests.SetUp;
begin
end;

procedure TSpreadErrorTests.TearDown;
begin
end;

procedure TSpreadErrorTests.TestWriteErrorMessages(AFormat: TsSpreadsheetFormat);
type
  TTestFormat = (sfExcel2, sfExcel5, sfExcel8, sfOOXML, sfOpenDocument);
const
  MAX_ROW_COUNT: array[TTestFormat] of Cardinal = (65536, 65536, 65536, 1048576, 1048576);
  MAX_COL_COUNT: array[TTestFormat] of Cardinal = (256, 256, 256, 16384, 1024);
  MAX_CELL_LEN: array[TTestFormat] of Cardinal = (255, 255, 32767, cardinal(-1), Cardinal(-1));
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  row, col: Cardinal;
  row1, row2: Cardinal;
  col1, col2: Cardinal;
  formula: TsFormula;
  s: String;
  TempFile: String;
  ErrList: TStringList;
begin
  formula.FormulaStr := '=A1';
  formula.DoubleValue := 0.0;

  ErrList := TStringList.Create;
  try
    // Test 1: Too many rows
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkSheet:= MyWorkBook.AddWorksheet(ERROR_SHEET);
      row1 := MAX_ROW_COUNT[TTestFormat(AFormat)] - 5;
      row2 := MAX_ROW_COUNT[TTestFormat(AFormat)] + 5;
      for row :=row1 to row2 do begin
        MyWorksheet.WriteBlank(row, 0);
        MyWorksheet.WriteNumber(row, 1, 1.0);
        MyWorksheet.WriteUTF8Text(row, 2, 'A');
        MyWorksheet.WriteFormula(Row, 3, formula);
        MyWorksheet.WriteRPNFormula(row, 4, CreateRPNFormula(
          RPNCellValue('A1', nil)));
      end;
      TempFile:=NewTempFile;
      MyWorkBook.WriteToFile(TempFile, AFormat, true);
      ErrList.Text := MyWorkbook.ErrorMsg;
      CheckEquals(1, ErrList.Count, 'Error count mismatch in test 1');
    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

    // Test 2: Too many columns
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkSheet:= MyWorkBook.AddWorksheet(ERROR_SHEET);
      col1 := MAX_COL_COUNT[TTestFormat(AFormat)] - 5;
      col2 := MAX_COL_COUNT[TTestFormat(AFormat)] + 5;
      for col := col1 to col2 do begin
        MyWorksheet.WriteBlank(0, col);
        MyWorksheet.WriteNumber(1, col, 1.0);
        MyWorksheet.WriteUTF8Text(2, col, 'A');
        MyWorksheet.WriteFormula(3, col, formula);
        MyWorksheet.WriteRPNFormula(4, col, CreateRPNFormula(
          RPNCellValue('A1', nil)));
      end;
      TempFile:=NewTempFile;
      MyWorkBook.WriteToFile(TempFile, AFormat, true);
      ErrList.Text := MyWorkbook.ErrorMsg;
      CheckEquals(1, ErrList.Count, 'Error count mismatch in test 2');
    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

    // Test 3: Too long cell label
    if MAX_CELL_LEN[TTestFormat(AFormat)] <> Cardinal(-1) then begin
      s := DupeString('A', MAX_CELL_LEN[TTestFormat(AFormat)] + 10);
      MyWorkbook := TsWorkbook.Create;
      try
        MyWorkSheet:= MyWorkBook.AddWorksheet(ERROR_SHEET);
        MyWorksheet.WriteUTF8Text(0, 0, s);
        TempFile:=NewTempFile;
        MyWorkBook.WriteToFile(TempFile, AFormat, true);
        ErrList.Text := MyWorkbook.ErrorMsg;
        CheckEquals(1, ErrList.Count, 'Error count mismatch in test 3');
      finally
        MyWorkbook.Free;
        DeleteFile(TempFile);
      end;
    end;

  finally
    ErrList.Free;
  end;
end;

procedure TSpreadErrorTests.TestWriteErrorMessages_BIFF2;
begin
  TestWriteErrorMessages(sfExcel2);
end;

procedure TSpreadErrorTests.TestWriteErrorMessages_BIFF5;
begin
  TestWriteErrorMessages(sfExcel5);
end;

procedure TSpreadErrorTests.TestWriteErrorMessages_BIFF8;
begin
  TestWriteErrorMessages(sfExcel8);
end;

procedure TSpreadErrorTests.TestWriteErrorMessages_ODS;
begin
  TestWriteErrorMessages(sfOpenDocument);
end;

procedure TSpreadErrorTests.TestWriteErrorMessages_OOXML;
begin
  TestWriteErrorMessages(sfOOXML);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadErrorTests);

end.

