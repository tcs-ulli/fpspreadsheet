{ Tests for insertion and deletion of columns and rows
  This unit test is writing out to and reading back from files.
}

unit insertdeletetests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpspreadsheet, xlsbiff8, {and a project requirement for lclbase for utf8 handling}
  testsutility;

type
  TInsDelTestDataItem = record
    Layout: string;
    InsertCol: Integer;
    InsertRow: Integer;
    DeleteCol: Integer;
    DeleteRow: Integer;
    Formula: String;
    SharedFormulaRowCount: Integer;
    SharedFormulaColCount: Integer;
    MergedColCount: Integer;
    MergedRowCount: Integer;
    SollLayout: String;
  end;

var
  InsDelTestData: array[0..5] of TInsDelTestDataItem;

  procedure InitTestData;

type
  { TSpreadWriteReadInsertColRowTests }
  TSpreadWriteRead_InsDelColRow_Tests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_InsDelColRow(ATestIndex: Integer);

  published
    // Writes out simple cell layout and inserts columns
    procedure TestWriteRead_InsDelColRow_0;     // before first
    procedure TestWriteRead_InsDelColRow_1;     // middle
    procedure TestWriteRead_InsDelColRow_2;     // before last
    // Writes out simple cell layout and deletes columns
    procedure TestWriteRead_InsDelColRow_3;     // first
    procedure TestWriteRead_InsDelColRow_4;     // middle
    procedure TestWriteRead_InsDelColRow_5;     // last
  end;

implementation

uses
  StrUtils;

const
  InsertColRowSheet = 'Insert_Columns_Rows';

procedure InitTestData;
var
  i: Integer;
begin
  for i := 0 to High(InsDelTestData) do
    with InsDelTestData[i] do
    begin
      Layout := '';
      InsertCol := -1;
      InsertRow := -1;
      DeleteCol := -1;
      DeleteRow := -1;
      Formula := '';
      SharedFormulaColCount := 0;
      SharedFormulaRowCount := 0;
      MergedColCount := 0;
      MergedRowCount := 0;
    end;

  // Insert a column before col 0
  with InsDelTestData[0] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    InsertCol := 0;
    SollLayout := ' 12345678|'+
                  ' 23456789|'+
                  ' 34567890|'+
                  ' 45678901';
  end;

  // Insert a column before col 2
  with InsDelTestData[1] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    InsertCol := 2;
    SollLayout := '12 345678|'+
                  '23 456789|'+
                  '34 567890|'+
                  '45 678901';
  end;

  // Insert a column before last col
  with InsDelTestData[2] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    InsertCol := 7;
    SollLayout := '1234567 8|'+
                  '2345678 9|'+
                  '3456789 0|'+
                  '4567890 1';
  end;

  // Delete column 0
  with InsDelTestData[3] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    DeleteCol := 0;
    SollLayout := '2345678|'+
                  '3456789|'+
                  '4567890|'+
                  '5678901';
  end;

  // Delete column 2
  with InsDelTestData[4] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    DeleteCol := 2;
    SollLayout := '1245678|'+
                  '2356789|'+
                  '3467890|'+
                  '4578901';
  end;

  // Delete last column
  with InsDelTestData[5] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    DeleteCol := 7;
    SollLayout := '1234567|'+
                  '2345678|'+
                  '3456789|'+
                  '4567890';
  end;
end;


{ TSpreadWriteRead_InsDelColRowTests }

procedure TSpreadWriteRead_InsDelColRow_Tests.SetUp;
begin
  inherited SetUp;
  InitTestData;
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow(
  ATestIndex: Integer);
const
  AFormat = sfExcel8;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  L: TStringList;
  s: String;
  expected: String;
  actual: String;

begin
  TempFile := GetTempFileName;

  L := TStringList.Create;
  try
    L.Delimiter := '|';
    L.StrictDelimiter := true;
    L.DelimitedText := InsDelTestData[ATestIndex].Layout;

    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkSheet:= MyWorkBook.AddWorksheet(InsertColRowSheet);

      // Write out cells
      for row := 0 to L.Count-1 do
      begin
        s := L[row];
        for col := 0 to Length(s)-1 do
          case s[col+1] of
            '0'..'9': MyWorksheet.WriteNumber(row, col, StrToInt(s[col+1]));
            ' '     : ;
          end;
      end;

      if InsDelTestData[ATestIndex].InsertCol >= 0 then
        MyWorksheet.InsertCol(InsDelTestData[ATestIndex].InsertCol);

      if InsDelTestData[ATestIndex].InsertRow >= 0 then
        MyWorksheet.InsertRow(InsDelTestData[ATestIndex].InsertRow);

      if InsDelTestData[ATestIndex].DeleteCol >= 0 then
        MyWorksheet.DeleteCol(InsDelTestData[ATestIndex].DeleteCol);

      if InsDelTestData[ATestIndex].DeleteRow >= 0 then
        MyWorksheet.DeleteRow(InsDelTestData[ATestIndex].DeleteRow);

      MyWorkBook.WriteToFile(TempFile, AFormat, true);
    finally
      MyWorkbook.Free;
    end;

    L.DelimitedText := InsDelTestData[ATestIndex].SollLayout;

    // Open the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, InsertColRowSheet);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet');

      for row := 0 to MyWorksheet.GetLastRowIndex do begin
        expected := L[row];
        actual := '';
        for col := 0 to MyWorksheet.GetLastColIndex do
        begin
          MyCell := MyWorksheet.FindCell(row, col);
          if MyCell = nil then
            actual := actual + ' '
          else
            case MyCell^.ContentType of
              cctEmpty : actual := actual + ' ';
              cctNumber: actual := actual + IntToStr(Round(Mycell^.NumberValue));
            end;
        end;
        CheckEquals(actual, expected,
          'Test empty cell layout mismatch, cell '+CellNotation(MyWorksheet, Row, Col));
      end;
    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

  finally
    L.Free;
  end;
end;


procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_0;
// insert a column before the first one
begin
  TestWriteRead_InsDelColRow(0);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_1;
// insert a column before column 2
begin
  TestWriteRead_InsDelColRow(1);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_2;
// insert a column before the last one
begin
  TestWriteRead_InsDelColRow(2);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_3;
// delete column 0
begin
  TestWriteRead_InsDelColRow(3);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_4;
// delete column 2
begin
  TestWriteRead_InsDelColRow(4);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_5;
// delete last column
begin
  TestWriteRead_InsDelColRow(5);
end;


initialization
  RegisterTest(TSpreadWriteRead_InsDelColRow_Tests);
  InitTestData;

end.

