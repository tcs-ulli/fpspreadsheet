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
    SollFormula: String;
    SharedFormulaRowCount: Integer;
    SharedFormulaColCount: Integer;
    MergedColCount: Integer;
    MergedRowCount: Integer;
    SollLayout: String;
  end;

var
  InsDelTestData: array[0..25] of TInsDelTestDataItem;

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
    // Writes out simple cell layout and inserts rows
    procedure TestWriteRead_InsDelColRow_6;     // before first
    procedure TestWriteRead_InsDelColRow_7;     // middle
    procedure TestWriteRead_InsDelColRow_8;     // before last
    // Writes out simple cell layout and deletes rows
    procedure TestWriteRead_InsDelColRow_9;     // first
    procedure TestWriteRead_InsDelColRow_10;    // middle
    procedure TestWriteRead_InsDelColRow_11;    // last

    // Writes out cell layout with formula and inserts columns
    procedure TestWriteRead_InsDelColRow_12;    // before formula cell
    procedure TestWriteRead_InsDelColRow_13;    // after formula cell
    // Writes out cell layout with formula and inserts rows
    procedure TestWriteRead_InsDelColRow_14;    // before formula cell
    procedure TestWriteRead_InsDelColRow_15;    // after formula cell
    // Writes out cell layout with formula and deletes columns
    procedure TestWriteRead_InsDelColRow_16;    // before formula cell
    procedure TestWriteRead_InsDelColRow_17;    // after formula cell
    procedure TestWriteRead_InsDelColRow_18;    // cell in formula
    // Writes out cell layout with formula and deletes rows
    procedure TestWriteRead_InsDelColRow_19;    // before formula cell
    procedure TestWriteRead_InsDelColRow_20;    // after formula cell
    procedure TestWriteRead_InsDelColRow_21;    // cell in formula

    // Writes out cell layout with shared formula
    procedure TestWriteRead_InsDelColRow_22;    // no insert/delete; just test shared formula
    // ... and inserts columns
    procedure TestWriteRead_InsDelColRow_23;    // column before shared formula cells
    procedure TestWriteRead_InsDelColRow_24;    // column after shared formula cells
    procedure TestWriteRead_InsDelColRow_25;    // column through cells addressed by shared formula
  end;

implementation

uses
  StrUtils;

const
  InsertColRowSheet = 'InsertDelete_ColumnsRows';

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
      SollFormula := '';
      SharedFormulaColCount := 0;
      SharedFormulaRowCount := 0;
      MergedColCount := 0;
      MergedRowCount := 0;
    end;

  { ---------------------------------------------------------------------------}
  {  Simple layouts                                                            }
  { ---------------------------------------------------------------------------}

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

  // Insert a ROW before row 0
  with InsDelTestData[6] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    InsertRow := 0;
    SollLayout := '     |'+
                  '12345|'+
                  '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Insert a ROW before row 2
  with InsDelTestData[7] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    InsertRow := 2;
    SollLayout := '12345|'+
                  '23456|'+
                  '     |'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Insert a ROW before last row
  with InsDelTestData[8] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    InsertRow := 5;
    SollLayout := '12345|'+
                  '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '     |'+
                  '67890|';
  end;

  // Delete the first row
  with InsDelTestData[9] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    DeleteRow := 0;
    SollLayout := '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Delete row #2
  with InsDelTestData[10] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    DeleteRow := 2;
    SollLayout := '12345|'+
                  '23456|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Delete last row
  with InsDelTestData[11] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    DeleteRow := 5;
    SollLayout := '12345|'+
                  '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789';
  end;

  { ---------------------------------------------------------------------------}
  {  Layouts with formula                                                      }
  { ---------------------------------------------------------------------------}

  // Insert a column before #1, i.e. before formula cell
  with InsDelTestData[12] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertCol := 1;
    Formula := 'C3';
    SollFormula := 'D3';           // col index increases due to inserted col
    SollLayout := '1 2345678|'+
                  '2 3456789|'+
                  '3 4565890|'+
                  '4 5678901|'+
                  '5 6789012|'+
                  '6 7890123';
  end;

  // Insert a column before #3, i.e. after formula cell
  with InsDelTestData[13] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertCol := 3;
    Formula := 'C3';
    SollFormula := 'C3';           // no change of cell because insertion is behind
    SollLayout := '123 45678|'+
                  '234 56789|'+
                  '345 65890|'+
                  '456 78901|'+
                  '567 89012|'+
                  '678 90123';
  end;

  // Insert a row before #1, i.e. before formula cell
  with InsDelTestData[14] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertRow := 1;
    Formula := 'E4';
    SollFormula := 'E5';         // row index increaes due to inserted row
    SollLayout := '12345678|'+
                  '        |'+
                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
                  '56789012|'+
                  '67890123';
  end;

  // Insert a row before #4, i.e. after formula cell
  with InsDelTestData[15] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertRow := 5;
    Formula := 'E4';
    SollFormula := 'E4';         // row index not changed dur to insert after cell
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
                  '56789012|'+
                  '        |'+
                  '67890123';
  end;

  // Deletes column #1, i.e. before formula cell
  with InsDelTestData[16] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteCol := 1;
    Formula := 'C3';
    SollFormula := 'B3';           // col index decreases due to delete before cell
    SollLayout := '1345678|'+
                  '2456789|'+
                  '3565890|'+
                  '4678901|'+
                  '5789012|'+
                  '6890123';
  end;

  // Deletes column #5, i.e. after formula cell
  with InsDelTestData[17] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteCol := 5;
    Formula := 'C3';
    SollFormula := 'C3';         // col index unchanged due to deleted after cell
    SollLayout := '1234578|'+
                  '2345689|'+
                  '3456590|'+
                  '4567801|'+
                  '5678912|'+
                  '6789023';
  end;

  // Deletes column #2, i.e. cell appearing in formula is gone --> #REF! error
  with InsDelTestData[18] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteCol := 2;
    Formula := 'C3';
    SollFormula := '#REF!';         // col index unchanged due to deletion after cell
    SollLayout := '1245678|'+
                  '2356789|'+
                  '346E890|'+    // "E" = error
                  '4578901|'+
                  '5689012|'+
                  '6790123';
  end;

  // Deletes row #1, i.e. before formula cell
  with InsDelTestData[19] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteRow := 1;
    Formula := 'E4';
    SollFormula := 'E3';           // row index decreases due to delete before cell
    SollLayout := '12345678|'+
//                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
                  '56789012|'+
                  '67890123';
  end;

  // Deletes row #4, i.e. after formula cell
  with InsDelTestData[20] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteRow := 4;
    Formula := 'E4';
    SollFormula := 'E4';           // row index unchanged (delete is after cell)
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
//                  '56789012|'+
                  '67890123';
  end;

  // Deletes row #2, i.e. row containing cell used in formula --> #REF! error!
  with InsDelTestData[21] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteRow := 3;
    Formula := 'E4';
    SollFormula := '#REF!';
    SollLayout := '12345678|'+
                  '23456789|'+
                  '3456E890|'+    // "E" = error
//                  '45678901|'+
                  '56789012|'+
                  '67890123';
  end;

  { ---------------------------------------------------------------------------}
  {  Layouts with shared formula                                                      }
  { ---------------------------------------------------------------------------}

  // No insert/delete, just to test the shared formula
  with InsDelTestData[22] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SollFormula := 'A3-$B$2,B3-$B$2;'+
                   'A4-$B$2,B4-$B$2;'+
                   'A5-$B$2,B5-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34501890|'+
                  '45612901|'+
                  '56723012|'+
                  '67890123';
  end;

  // Insert column before any cell referred to by the shared formula (col = 0)
  with InsDelTestData[23] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 0;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SollFormula := 'B3-$C$2,C3-$C$2;'+   // all column indexes increase by 1 due to added col in front
                   'B4-$C$2,C4-$C$2;'+
                   'B5-$C$2,C5-$C$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := ' 12345678|'+
                  ' 23456789|'+
                  ' 34501890|'+
                  ' 45612901|'+
                  ' 56723012|'+
                  ' 67890123';
  end;

  // Insert column after last cell addressed by the shared formula
  with InsDelTestData[24] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 7;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SollFormula := 'A3-$B$2,B3-$B$2;'+    // formulas unchanged by insert
                   'A4-$B$2,B4-$B$2;'+
                   'A5-$B$2,B5-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '1234567 8|'+
                  '2345678 9|'+
                  '3450189 0|'+
                  '4561290 1|'+
                  '5672301 2|'+
                  '6789012 3';
  end;

  // Insert column between cells referred to by the shared formula (col = 1)
  with InsDelTestData[25] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 1;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SollFormula := 'A3-$C$2,C3-$C$2;'+   // some column indexes increase by 1, some unchanged
                   'A4-$C$2,C4-$C$2;'+
                   'A5-$C$2,C5-$C$2';
    SollLayout := '1 2345678|'+
                  '2 3456789|'+
                  '3 4501890|'+
                  '4 5612901|'+
                  '5 6723012|'+
                  '6 7890123';
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
  L, LL: TStringList;
  s: String;
  expected: String;
  actual: String;
  expectedFormulas: array of array of String;

begin
  TempFile := GetTempFileName;

  L := TStringList.Create;
  try
    // Extract soll formulas into a 2D array in case of shared formulas
    if (InsDelTestData[ATestIndex].SharedFormulaRowCount > 0) or
       (InsDelTestData[ATestIndex].SharedFormulaColCount > 0) then
    begin
      with InsDelTestData[ATestIndex] do
        SetLength(expectedFormulas, SharedFormulaRowCount, SharedFormulaColCount);
      L.Delimiter := ';';
      L.DelimitedText := InsDelTestData[ATestIndex].SollFormula;
      LL := TStringList.Create;
      try
        LL.Delimiter := ',';
        for row := 0 to InsDelTestData[ATestIndex].SharedFormulaRowCount-1 do
        begin
          s := L[row];
          LL.DelimitedText := L[row];
          for col := 0 to InsDelTestData[ATestIndex].SharedFormulaColCount-1 do
            expectedFormulas[row, col] := LL[col];
        end;
      finally
        LL.Free;
      end;
    end;

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
            ' '     : ; // Leave cell empty
            '0'..'9': MyWorksheet.WriteNumber(row, col, StrToInt(s[col+1]));
            'F'     : MyWorksheet.WriteFormula(row, col, InsDelTestData[ATestIndex].Formula);
            'S'     : MyWorksheet.WriteSharedFormula(
                        row,
                        col,
                        row + InsDelTestData[ATestIndex].SharedFormulaRowCount-1,
                        col + InsDelTestData[ATestIndex].SharedFormulaColCount-1,
                        InsDelTestData[ATestIndex].Formula
                      );
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
      MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas, boAutoCalc];
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
              cctError : actual := actual + 'E';
            end;
          if HasFormula(MyCell) then
          begin
            if MyCell^.SharedFormulaBase <> nil then
              CheckEquals(
                expectedFormulas[row-MyCell^.SharedFormulaBase^.Row, col-MyCell^.SharedFormulaBase^.Col],
                MyWorksheet.ReadFormulaAsString(MyCell),
                'Shared formula mismatch, cell ' + CellNotation(MyWorksheet, Row, Col)
              )
            else
              CheckEquals(
                InsDelTestData[ATestIndex].SollFormula,
                MyWorksheet.ReadFormulaAsString(MyCell),
                'Formula mismatch, cell '+CellNotation(MyWorksheet, Row, Col)
              );
          end;
        end;
        CheckEquals(expected, actual,
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

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_6;
// insert row before first one
begin
  TestWriteRead_InsDelColRow(6);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_7;
// insert row before #2
begin
  TestWriteRead_InsDelColRow(7);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_8;
// insert row before last one
begin
  TestWriteRead_InsDelColRow(8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_9;
// delete first row
begin
  TestWriteRead_InsDelColRow(9);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_10;
// delete row #2
begin
  TestWriteRead_InsDelColRow(10);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_11;
// delete last row
begin
  TestWriteRead_InsDelColRow(11);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_12;
// insert column before formula cell
begin
  TestWriteRead_InsDelColRow(12);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_13;
// insert column after formula cell
begin
  TestWriteRead_InsDelColRow(13);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_14;
// insert row before formula cell
begin
  TestWriteRead_InsDelColRow(14);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_15;
// insert row after formula cell
begin
  TestWriteRead_InsDelColRow(15);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_16;
// delete column before formula cell
begin
  TestWriteRead_InsDelColRow(16);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_17;
// delete column after formula cell
begin
  TestWriteRead_InsDelColRow(17);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_18;
// delete column containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(18);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_19;
// delete row before formula cell
begin
  TestWriteRead_InsDelColRow(19);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_20;
// delete row after formula cell
begin
  TestWriteRead_InsDelColRow(20);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_21;
// delete row containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(21);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_22;
// no insert/delete; just test shared formula
begin
  TestWriteRead_InsDelColRow(22);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_23;
// insert column before any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(23);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_24;
// insert column after any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(24);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_25;
// column through cells addressed by shared formula
begin
  TestWriteRead_InsDelColRow(25);
end;


initialization
  RegisterTest(TSpreadWriteRead_InsDelColRow_Tests);
  InitTestData;

end.

