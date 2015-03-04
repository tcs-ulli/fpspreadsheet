{ Tests for iteration through cells by means of the enumerator of the cells tree.
  This unit test is not writing anything to file.
}

unit enumeratortests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, fpsclasses, {and a project requirement for lclbase for utf8 handling}
  testsutility;

type
  { TSpreadEnumeratorTests }
    TSpreadEnumeratorTests = class(TTestCase)
    private

    protected
      procedure SetUp; override;
      procedure TearDown; override;
      procedure Test_EnumCells(what: Integer; reverse, withGaps: Boolean);
      procedure Test_EnumComments(what: Integer; withGaps: Boolean);

    published
      procedure Test_Enum_Cells_All;
      procedure Test_Enum_Cells_All_Reverse;
      procedure Test_Enum_Cells_FullRow;
      procedure Test_Enum_Cells_FullRow_Reverse;
      procedure Test_Enum_Cells_FullCol;
      procedure Test_Enum_Cells_FullCol_Reverse;
      procedure Test_Enum_Cells_PartialRow;
      procedure Test_Enum_Cells_PartialRow_Reverse;
      procedure Test_Enum_Cells_PartialCol;
      procedure Test_Enum_Cells_PartialCol_Reverse;
      procedure Test_Enum_Cells_Range;
      procedure Test_Enum_Cells_Range_Reverse;

      procedure Test_Enum_Cells_WithGaps_All;
      procedure Test_Enum_Cells_WithGaps_All_Reverse;
      procedure Test_Enum_Cells_WithGaps_FullRow;
      procedure Test_Enum_Cells_WithGaps_FullRow_Reverse;
      procedure Test_Enum_Cells_WithGaps_FullCol;
      procedure Test_Enum_Cells_WithGaps_FullCol_Reverse;
      procedure Test_Enum_Cells_WithGaps_PartialRow;
      procedure Test_Enum_Cells_WithGaps_PartialRow_Reverse;
      procedure Test_Enum_Cells_WithGaps_PartialCol;
      procedure Test_Enum_Cells_WithGaps_PartialCol_Reverse;
      procedure Test_Enum_Cells_WithGaps_Range;
      procedure Test_Enum_Cells_WithGaps_Range_Reverse;

      procedure Test_Enum_Comments_All;
      procedure Test_Enum_Comments_Range;

      procedure Test_Enum_Comments_WithGaps_All;
      procedure Test_Enum_Comments_WithGaps_Range;

    end;

implementation

const
  NUM_ROWS = 100;
  NUM_COLS = 100;
  TEST_ROW = 10;
  TEST_COL = 20;
  TEST_ROW1 = 20;
  TEST_ROW2 = 50;
  TEST_COL1 = 30;
  TEST_COL2 = 60;

procedure TSpreadEnumeratorTests.Setup;
begin
end;

procedure TSpreadEnumeratorTests.TearDown;
begin
end;

procedure TSpreadEnumeratorTests.Test_EnumCells(what: Integer; reverse: Boolean;
  withGaps: Boolean);
{ what = 1 ---> iterate through entire worksheet
  what = 2 ---> iterate along full row
  what = 3 ---> iterate along full column
  what = 4 ---> iterate along partial row
  what = 5 ---> iterate along partial column
  what = 6 ---> iterate through rectangular cell range

  The test writes numbers into the worksheet calculated by <row>*10000 + <col>.
  Then the test iterates through the designed range (according to "what") and
  compares the read number with the soll values.

  If "withGaps" is true then numbers are only written at cells where
  <col>+<row> is odd. }
var
  row, col: Cardinal;
  cell: PCell;
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  expected, actual: Double;
  enumerator: TsCellEnumerator;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkbook.AddWorksheet('Sheet1');
    for row := 0 to NUM_ROWS-1 do
      for col := 0 to NUM_COLS-1 do
        if (withGaps and odd(row + col)) or (not withGaps) then
          MyWorksheet.WriteNumber(row, col, row*10000.0 + col);

    if reverse then
      case what of
        1: enumerator := MyWorksheet.Cells.GetReverseRangeEnumerator(0, 0, $7FFFFFFF, $7FFFFFFF);
        2: enumerator := Myworksheet.Cells.GetReverseRowEnumerator(TEST_ROW);
        3: enumerator := MyWorksheet.Cells.GetReverseColEnumerator(TEST_COL);
        4: enumerator := MyWorksheet.Cells.GetReverseRowEnumerator(TEST_ROW, TEST_COL1, TEST_COL2);
        5: enumerator := Myworksheet.Cells.GetReverseColEnumerator(TEST_COL, TEST_ROW1, TEST_ROW2);
        6: enumerator := MyWorksheet.Cells.GetReverseRangeEnumerator(TEST_ROW1, TEST_COL1, TEST_ROW2, TEST_COL2);
      end
    else
      case what of
        1: enumerator := MyWorksheet.Cells.GetEnumerator;
        2: enumerator := Myworksheet.Cells.GetRowEnumerator(TEST_ROW);
        3: enumerator := MyWorksheet.Cells.GetColEnumerator(TEST_COL);
        4: enumerator := MyWorksheet.Cells.GetRowEnumerator(TEST_ROW, TEST_COL1, TEST_COL2);
        5: enumerator := Myworksheet.Cells.GetColEnumerator(TEST_COL, TEST_ROW1, TEST_ROW2);
        6: enumerator := MyWorksheet.Cells.GetRangeEnumerator(TEST_ROW1, TEST_COL1, TEST_ROW2, TEST_COL2);
      end;

    for cell in enumerator do
    begin
      row := cell^.Row;
      col := cell^.Col;
      if (withgaps and odd(row + col)) or (not withgaps) then
        expected := row * 10000.0 + col
      else
        expected := 0.0;
      actual := MyWorksheet.ReadAsNumber(cell);
      CheckEquals(expected, actual,
        'Enumerated cell value mismatch, cell '+CellNotation(MyWorksheet, row, col));
    end;

    // for debugging, to see the data file
    // MyWorkbook.WriteToFile('enumerator-test.xlsx', sfOOXML, true);

  finally
    MyWorkbook.Free;
  end;
end;

procedure TSpreadEnumeratorTests.Test_EnumComments(what: Integer;
  withGaps: Boolean);
{ what = 1 ---> iterate through entire worksheet
  what = 2 ---> iterate through rectangular cell range

  The test writes comments into the worksheet calculated by <row>*10000 + <col>.
  Then the test iterates through the designed range (according to "what") and
  compares the read comments with the soll values.

  if "withGaps" is true then comments are only written at cells where
  <col>+<row> is odd. }
var
  row, col: Cardinal;
  comment: PsComment;
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  expected, actual: string;
  enumerator: TsCommentEnumerator;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkbook.AddWorksheet('Sheet1');
    for row := 0 to NUM_ROWS-1 do
      for col := 0 to NUM_COLS-1 do
        if (withGaps and odd(row + col)) or (not withGaps) then
          MyWorksheet.WriteComment(row, col, IntToStr(row*10000 + col));

    case what of
      1: enumerator := MyWorksheet.Comments.GetEnumerator;
      2: enumerator := MyWorksheet.Comments.GetRangeEnumerator(TEST_ROW1, TEST_COL1, TEST_ROW2, TEST_COL2);
    end;

    for comment in enumerator do
    begin
      row := comment^.Row;
      col := comment^.Col;
      if (withgaps and odd(row + col)) or (not withgaps) then
        expected := IntToStr(row * 10000 + col)
      else
        expected := '';
      actual := MyWorksheet.ReadComment(row, col);
      CheckEquals(expected, actual,
        'Enumerated comment mismatch, cell '+CellNotation(MyWorksheet, row, col));
    end;

    // for debugging, to see the data file
    // MyWorkbook.WriteToFile('enumerator-test.xlsx', sfOOXML, true);

  finally
    MyWorkbook.Free;
  end;
end;


{ Fully filled worksheet }
procedure TSpreadEnumeratorTests.Test_Enum_Cells_All;
begin
  Test_Enumcells(1, false, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_All_Reverse;
begin
  Test_EnumCells(1, true, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_FullRow;
begin
  Test_EnumCells(2, false, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_FullRow_Reverse;
begin
  Test_EnumCells(2, true, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_FullCol;
begin
  Test_EnumCells(3, false, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_FullCol_Reverse;
begin
  Test_EnumCells(3, true, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_PartialRow;
begin
  Test_EnumCells(4, false, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_PartialRow_Reverse;
begin
  Test_EnumCells(4, true, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_PartialCol;
begin
  Test_EnumCells(5, false, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_PartialCol_Reverse;
begin
  Test_EnumCells(5, true, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_Range;
begin
  Test_EnumCells(6, false, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_Range_Reverse;
begin
  Test_EnumCells(6, true, false);
end;


{ Worksheet with gaps}

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_All;
begin
  Test_Enumcells(1, false, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_All_Reverse;
begin
  Test_EnumCells(1, true, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_FullRow;
begin
  Test_EnumCells(2, false, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_FullRow_Reverse;
begin
  Test_EnumCells(2, true, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_FullCol;
begin
  Test_EnumCells(3, false, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_FullCol_Reverse;
begin
  Test_EnumCells(3, true, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_PartialRow;
begin
  Test_EnumCells(4, false, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_PartialRow_Reverse;
begin
  Test_EnumCells(4, true, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_PartialCol;
begin
  Test_EnumCells(5, false, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_PartialCol_Reverse;
begin
  Test_EnumCells(5, true, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_Range;
begin
  Test_EnumCells(6, false, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Cells_WithGaps_Range_Reverse;
begin
  Test_EnumCells(6, true, true);
end;


{ Fully filled worksheet }

procedure TSpreadEnumeratorTests.Test_Enum_Comments_All;
begin
  Test_EnumComments(1, false);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Comments_Range;
begin
  Test_EnumComments(2, false);
end;

{ Every other cell empty }

procedure TSpreadEnumeratorTests.Test_Enum_Comments_WithGaps_All;
begin
  Test_EnumComments(1, true);
end;

procedure TSpreadEnumeratorTests.Test_Enum_Comments_WithGaps_Range;
begin
  Test_EnumComments(2, true);
end;


initialization
  RegisterTest(TSpreadEnumeratorTests);

end.

