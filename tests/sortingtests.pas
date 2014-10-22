unit sortingtests;

{$mode objfpc}{$H+}

interface
{ Tests for sorting cells
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8, fpsopendocument, {and a project requirement for lclbase for utf8 handling}
  testsutility;

var
  // Norm to test against - list of numbers and strings that will be sorted
  SollSortNumbers: array[0..9] of Double;
  SollSortStrings: array[0..9] of String;

  procedure InitUnsortedData;

type
  { TSpreadSortingTests }
  TSpreadSortingTests = class(TTestCase)
  private

  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;

    procedure Test_Sorting_1(     // one column or row
      ASortByCols: Boolean;
      AMode: Integer  // AMode = 0: number, 1: strings, 2: mixed
    );
    procedure Test_Sorting_2(     // two columns/rows, primary keys equal
      ASortByCols: Boolean
    );

  published
    procedure Test_SortingByCols1_Numbers;
    procedure Test_SortingByCols1_Strings;
    procedure Test_SortingByCols1_NumbersStrings;

    procedure Test_SortingByRows1_Numbers;
    procedure Test_SortingByRows1_Strings;
    procedure Test_SortingByRows1_NumbersStrings;

    procedure Test_SortingByCols2;
    procedure Test_SortingByRows2;

  end;

implementation

uses
  fpsutils;

const
  SortingTestSheet = 'Sorting';

procedure InitUnsortedData;
// The logics of the detection requires equal count of numbers and strings.
begin
  // When sorted the value is equal to the index
  SollSortNumbers[0] := 9;
  SollSortNumbers[1] := 8;
  SollSortNumbers[2] := 5;
  SollSortNumbers[3] := 2;
  SollSortNumbers[4] := 6;
  SollSortNumbers[5] := 7;
  SollSortNumbers[6] := 1;
  SollSortNumbers[7] := 3;
  SollSortNumbers[8] := 4;
  SollSortNumbers[9] := 0;

  // When sorted the value is equal to 'A' + index
  SollSortStrings[0] := 'C';
  SollSortStrings[1] := 'G';
  SollSortStrings[2] := 'F';
  SollSortStrings[3] := 'I';
  SollSortStrings[4] := 'B';
  SollSortStrings[5] := 'D';
  SollSortStrings[6] := 'J';
  SollSortStrings[7] := 'H';
  SollSortStrings[8] := 'E';
  SollSortStrings[9] := 'A';
end;


{ TSpreadSortingTests }

procedure TSpreadSortingTests.SetUp;
begin
  inherited SetUp;
  InitUnsortedData;
end;

procedure TSpreadSortingTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadSortingTests.Test_Sorting_1(ASortByCols: Boolean;
  AMode: Integer);
const
  AFormat = sfExcel8;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  i, ilast, n, row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  L: TStringList;
  s: String;
  sortParams: TsSortParams;
  sortDir: TsSortOrder;
  r1,r2,c1,c2: Cardinal;
  actualNumber: Double;
  actualString: String;
  expectedNumber: Double;
  expectedString: String;

begin
  sortParams := InitSortParams(ASortByCols, 1);

  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(SortingTestSheet);

    col := 0;
    row := 0;
    if ASortByCols then begin
      case AMode of
        0: for i :=0 to High(SollSortNumbers) do
             MyWorksheet.WriteNumber(i, col, SollSortNumbers[i]);
        1: for i := 0 to High(SollSortStrings) do
             Myworksheet.WriteUTF8Text(i, col, SollSortStrings[i]);
        2: begin
             for i := 0 to High(SollSortNumbers) do
               MyWorkSheet.WriteNumber(i*2, col, SollSortNumbers[i]);
             for i := 0 to High(SollSortStrings) do
               MyWorksheet.WriteUTF8Text(i*2+1, col, SollSortStrings[i]);
           end;
      end
    end
    else begin
      case AMode of
        0: for i := 0 to High(SollSortNumbers) do
             MyWorksheet.WriteNumber(row, i, SollSortNumbers[i]);
        1: for i := 0 to High(SollSortStrings) do
             MyWorksheet.WriteUTF8Text(row, i, SollSortStrings[i]);
        2: begin
             for i := 0 to High(SollSortNumbers) do
               myWorkSheet.WriteNumber(row, i*2, SollSortNumbers[i]);
             for i:=0 to High(SollSortStrings) do
               MyWorksheet.WriteUTF8Text(row, i*2+1, SollSortStrings[i]);
           end;
      end;
    end;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Test ascending and descending sort orders
  for sortDir in TsSortOrder do
  begin
    MyWorkbook := TsWorkbook.Create;
    try
      // Read spreadsheet file...
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, SortingTestSheet);
      if MyWorksheet = nil then
        fail('Error in test code. Failed to get named worksheet');

      // ... and sort it.
      case AMode of
        0: iLast:= High(SollSortNumbers);
        1: iLast := High(SollSortStrings);
        2: iLast := Length(SollSortNumbers) + Length(SollSortStrings) - 1;
      end;
      r1 := 0;
      r2 := 0;
      c1 := 0;
      c2 := 0;
      if ASortByCols then
        r2 := iLast
      else
        c2 := iLast;
      sortParams.Keys[0].Order := sortDir;
      MyWorksheet.Sort(sortParams, r1,c1, r2, c2);

      // for debugging, to see the sorted data
      // MyWorkbook.WriteToFile('sorted.xls', AFormat, true);

      row := 0;
      col := 0;
      for i:=0 to iLast do
      begin
        if ASortByCols then
          case sortDir of
            ssoAscending : row := i;
            ssoDescending: row := iLast - i;
          end
        else
          case sortDir of
            ssoAscending : col := i;
            ssoDescending: col := iLast - i;
          end;
        case AMode of
          0: begin
               actualNumber := MyWorksheet.ReadAsNumber(row, col);
               expectedNumber := i;
               CheckEquals(expectednumber, actualnumber,
                 'Sorted cell number mismatch, cell '+CellNotation(MyWorksheet, row, col));
             end;
          1: begin
               actualString := MyWorksheet.ReadAsUTF8Text(row, col);
               expectedString := char(ord('A') + i);
               CheckEquals(expectedstring, actualstring,
                 'Sorted cell string mismatch, cell '+CellNotation(MyWorksheet, row, col));
             end;
          2: begin  // with increasing i, we see first the numbers, then the strings
               if i <= High(SollSortNumbers) then begin
                 actualnumber := MyWorksheet.ReadAsNumber(row, col);
                 expectedNumber := i;
                 CheckEquals(expectednumber, actualnumber,
                   'Sorted cell number mismatch, cell '+CellNotation(MyWorksheet, row, col));
               end else begin
                 actualstring := MyWorksheet.ReadAsUTF8Text(row, col);
                 expectedstring := char(ord('A') + i - Length(SollSortNumbers));
                 CheckEquals(expectedstring, actualstring,
                   'Sorted cell string mismatch, cell '+CellNotation(MyWorksheet, row, col));
               end;
             end;
        end;
      end;

    finally
      MyWorkbook.Free;
    end;
  end;  // for sortDir

  DeleteFile(TempFile);
end;

procedure TSpreadSortingTests.Test_Sorting_2(ASortByCols: Boolean);
const
  AFormat = sfExcel8;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  i, ilast, n, row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  L: TStringList;
  s: String;
  sortParams: TsSortParams;
  sortDir: TsSortOrder;
  r1,r2,c1,c2: Cardinal;
  actualNumber: Double;
  actualString: String;
  expectedNumber: Double;
  expectedString: String;

begin
  sortParams := InitSortParams(ASortByCols, 2);
  sortParams.Keys[0].ColRowIndex := 0;
  sortParams.Keys[1].ColRowIndex := 1;

  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(SortingTestSheet);

    col := 0;
    row := 0;
    if ASortByCols then
    begin
      // Always 2 numbers in the first column are equal
      for i:=0 to High(SollSortNumbers) do
        MyWorksheet.WriteNumber(i, col, SollSortNumbers[(i mod 2)*2]);
      // All strings in the second column are distinct
      for i:=0 to High(SollSortStrings) do
        MyWorksheet.WriteUTF8Text(i, col+1, SollSortStrings[i]);
    end else
    begin
      for i:=0 to High(SollSortNumbers) do
        MyWorksheet.WriteNumber(row, i, SollSortNumbers[(i mod 2)*2]);
      for i:=0 to High(SollSortStrings) do
        MyWorksheet.WriteUTF8Text(row+1, i, SollSortStrings[i]);
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Test ascending and descending sort orders
  for sortDir in TsSortOrder do
  begin
    MyWorkbook := TsWorkbook.Create;
    try
      // Read spreadsheet file...
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, SortingTestSheet);
      if MyWorksheet = nil then
        fail('Error in test code. Failed to get named worksheet');

      // ... and sort it.
      iLast := High(SollSortNumbers);  //must be the same as for SollSortStrings
      r1 := 0;  c1 := 0;
      if ASortByCols then begin
        c2 := 1;
        r2 := iLast;
      end else
      begin
        c2 := iLast;
        r2 := 1;
      end;
      sortParams.Keys[0].Order := sortDir;
      sortParams.Keys[1].Order := sortDir;
      MyWorksheet.Sort(sortParams, r1,c1, r2, c2);

      // for debugging, to see the sorted data
      MyWorkbook.WriteToFile('sorted.xls', AFormat, true);

      for i:=0 to iLast do
      begin
        row := 0;
        col := 0;
        if ASortByCols then
          case sortDir of
            ssoAscending : row := i;
            ssoDescending: row := iLast - i;
          end
        else
          case sortDir of
            ssoAscending : col := i;
            ssoDescending: col := iLast - i;
          end;
        actualNumber := MyWorksheet.ReadAsNumber(row, col);
        expectedNumber := (i mod 2) * 2;
        CheckEquals(expectednumber, actualnumber,
          'Sorted cell number mismatch, cell '+CellNotation(MyWorksheet, row, col));

        if ASortByCols then
          inc(col)
        else
          inc(row);
        actualString := MyWorksheet.ReadAsUTF8Text(row, col);
        expectedString := char(ord('A') + i);
        CheckEquals(expectedstring, actualstring,
          'Sorted cell string mismatch, cell '+CellNotation(MyWorksheet, row, col));
      end;
    finally
      MyWorkbook.Free;
    end;
  end;    // for sortDir

  DeleteFile(TempFile);
end;


procedure TSpreadSortingTests.Test_SortingByCols1_Numbers;
begin
  Test_Sorting_1(true, 0);
end;

procedure TSpreadSortingTests.Test_SortingByCols1_Strings;
begin
  Test_Sorting_1(true, 1);
end;

procedure TSpreadSortingTests.Test_SortingByCols1_NumbersStrings;
begin
  Test_Sorting_1(true, 2);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_Numbers;
begin
  Test_Sorting_1(false, 0);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_Strings;
begin
  Test_Sorting_1(false, 1);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_NumbersStrings;
begin
  Test_Sorting_1(false, 2);
end;

procedure TSpreadSortingTests.Test_SortingByCols2;
begin
  Test_Sorting_2(true);
end;

procedure TSpreadSortingTests.Test_SortingByRows2;
begin
  Test_Sorting_2(false);
end;

initialization
  RegisterTest(TSpreadSortingTests);
  InitUnsortedData;

end.

