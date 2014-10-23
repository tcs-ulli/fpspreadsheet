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
      ADescending: Boolean;       // true: desending order
      AWhat: Integer  // What = 0: number, 1: strings, 2: mixed
    );
    procedure Test_Sorting_2(     // two columns/rows, primary keys equal
      ASortByCols: Boolean;
      ADescending: Boolean
    );

  published
    procedure Test_SortingByCols1_Numbers_Asc;
    procedure Test_SortingByCols1_Numbers_Desc;
    procedure Test_SortingByCols1_Strings_Asc;
    procedure Test_SortingByCols1_Strings_Desc;
    procedure Test_SortingByCols1_NumbersStrings_Asc;
    procedure Test_SortingByCols1_NumbersStrings_Desc;

    procedure Test_SortingByRows1_Numbers_Asc;
    procedure Test_SortingByRows1_Numbers_Desc;
    procedure Test_SortingByRows1_Strings_Asc;
    procedure Test_SortingByRows1_Strings_Desc;
    procedure Test_SortingByRows1_NumbersStrings_Asc;
    procedure Test_SortingByRows1_NumbersStrings_Desc;

    procedure Test_SortingByCols2_Asc;
    procedure Test_SortingByCols2_Desc;
    procedure Test_SortingByRows2_Asc;
    procedure Test_SortingByRows2_Desc;

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
  ADescending: Boolean; AWhat: Integer);
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
  sortOptions: TsSortOptions;
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
      case AWhat of
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
      case AWhat of
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

    // ... set up sorting direction
    case ADescending of
      false: sortParams.Keys[0].Options := [];               // Ascending sort
      true : sortParams.Keys[0].Options := [ssoDescending];  // Descending sort
    end;

    // ... and sort it.
    case AWhat of
      0: iLast:= High(SollSortNumbers);
      1: iLast := High(SollSortStrings);
      2: iLast := Length(SollSortNumbers) + Length(SollSortStrings) - 1;
    end;
    if ASortByCols then
      MyWorksheet.Sort(sortParams, 0, 0, iLast, 0)
    else
      MyWorksheet.Sort(sortParams, 0, 0, 0, iLast);

    // for debugging, to see the sorted data
    // MyWorkbook.WriteToFile('sorted.xls', AFormat, true);

    row := 0;
    col := 0;
    for i:=0 to iLast do
    begin
      if ASortByCols then
        case ADescending of
          false: row := i;            // ascending
          true : row := iLast - i;    // descending
        end
      else
        case ADescending of
          false: col := i;            // ascending
          true : col := iLast - i;    // descending
        end;
      case AWhat of
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

  DeleteFile(TempFile);
end;

procedure TSpreadSortingTests.Test_Sorting_2(ASortByCols: Boolean;
  ADescending: Boolean);
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
  sortOptions: TsSortOptions;
  r1,r2,c1,c2: Cardinal;
  actualNumber: Double;
  actualString: String;
  expectedNumber: Double;
  expectedString: String;

begin
  sortParams := InitSortParams(ASortByCols, 2);
  sortParams.Keys[0].ColRowIndex := 0;    // col/row 0 is primary key
  sortParams.Keys[1].ColRowIndex := 1;    // col/row 1 is second key

  iLast := High(SollSortNumbers);

  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(SortingTestSheet);

    col := 0;
    row := 0;
    if ASortByCols then
    begin
      // Write all randomized numbers to column B
      for i:=0 to iLast do
        MyWorksheet.WriteNumber(i, col+1, SollSortNumbers[i]);
      // divide each number by 2 and calculate the character assigned to it
      // and write it to column A
      // We will sort primarily according to column A, and seconarily according
      // to B. The construction allows us to determine if the sorting is correct.
      for i:=0 to iLast do
        MyWorksheet.WriteUTF8Text(i, col, char(ord('A')+round(SollSortNumbers[i]) div 2));
    end else
    begin
      // The same with the rows...
      for i:=0 to iLast do
        MyWorksheet.WriteNumber(row+1, i, SollSortNumbers[i]);
      for i:=0 to iLast do
        MyWorksheet.WriteUTF8Text(row, i, char(ord('A')+round(SollSortNumbers[i]) div 2));
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

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

    // ... set up sort direction
    if ADescending then
    begin
      sortParams.Keys[0].Options := [ssoDescending];
      sortParams.Keys[1].Options := [ssoDescending];
    end else
    begin
      sortParams.Keys[0].Options := [];
      sortParams.Keys[1].Options := [];
    end;

    // ... and sort it.
    if ASortByCols then
      MyWorksheet.Sort(sortParams, 0, 0, iLast, 1)
    else
      MyWorksheet.Sort(sortParams, 0, 0, 1, iLast);

    // for debugging, to see the sorted data
    MyWorkbook.WriteToFile('sorted.xls', AFormat, true);

    for i:=0 to iLast do
    begin
      if ASortByCols then
      begin
        // Read the number first, they must be in order 0...9 (if ascending).
        col := 1;
        case ADescending of
          false : row := i;
          true  : row := iLast - i;
        end;
        actualNumber := MyWorksheet.ReadAsNumber(row, col);  // col B is the number, must be 0...9 here
        expectedNumber := i;
        CheckEquals(expectednumber, actualnumber,
          'Sorted cell number mismatch, cell '+CellNotation(MyWorksheet, row, col));

        // Now read the string. It must be the character corresponding to the
        // half of the number
        col := 0;
        actualString := MyWorksheet.ReadAsUTF8Text(row, col);
        expectedString := char(ord('A') + round(expectedNumber) div 2);
        CheckEquals(expectedstring, actualstring,
          'Sorted cell string mismatch, cell '+CellNotation(MyWorksheet, row, col));
      end else
      begin
        row := 1;
        case ADescending of
          false : col := i;
          true  : col := iLast - i;
        end;
        actualNumber := MyWorksheet.ReadAsNumber(row, col);
        expectedNumber := i;
        CheckEquals(expectednumber, actualnumber,
          'Sorted cell number mismatch, cell '+CellNotation(MyWorksheet, row, col));

        row := 0;
        actualstring := MyWorksheet.ReadAsUTF8Text(row, col);
        expectedString := char(ord('A') + round(expectedNumber) div 2);
        CheckEquals(expectedstring, actualstring,
          'Sorted cell string mismatch, cell '+CellNotation(MyWorksheet, row, col));
      end;
    end;
  finally
    MyWorkbook.Free;
  end;

  DeleteFile(TempFile);
end;

{ Sort 1 column }
procedure TSpreadSortingTests.Test_SortingByCols1_Numbers_ASC;
begin
  Test_Sorting_1(true, false, 0);
end;

procedure TSpreadSortingTests.Test_SortingByCols1_Numbers_DESC;
begin
  Test_Sorting_1(true, true, 0);
end;

procedure TSpreadSortingTests.Test_SortingByCols1_Strings_ASC;
begin
  Test_Sorting_1(true, false, 1);
end;

procedure TSpreadSortingTests.Test_SortingByCols1_Strings_DESC;
begin
  Test_Sorting_1(true, true, 1);
end;

procedure TSpreadSortingTests.Test_SortingByCols1_NumbersStrings_ASC;
begin
  Test_Sorting_1(true, false, 2);
end;

procedure TSpreadSortingTests.Test_SortingByCols1_NumbersStrings_DESC;
begin
  Test_Sorting_1(true, true, 2);
end;

{ Sort 1 row }
procedure TSpreadSortingTests.Test_SortingByRows1_Numbers_asc;
begin
  Test_Sorting_1(false, false, 0);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_Numbers_Desc;
begin
  Test_Sorting_1(false, true, 0);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_Strings_Asc;
begin
  Test_Sorting_1(false, false, 1);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_Strings_Desc;
begin
  Test_Sorting_1(false, true, 1);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_NumbersStrings_Asc;
begin
  Test_Sorting_1(false, false, 2);
end;

procedure TSpreadSortingTests.Test_SortingByRows1_NumbersStrings_Desc;
begin
  Test_Sorting_1(false, true, 2);
end;

{ two columns }
procedure TSpreadSortingTests.Test_SortingByCols2_Asc;
begin
  Test_Sorting_2(true, false);
end;

procedure TSpreadSortingTests.Test_SortingByCols2_Desc;
begin
  Test_Sorting_2(true, true);
end;

procedure TSpreadSortingTests.Test_SortingByRows2_Asc;
begin
  Test_Sorting_2(false, false);
end;

procedure TSpreadSortingTests.Test_SortingByRows2_Desc;
begin
  Test_Sorting_2(false, true);
end;


initialization
  RegisterTest(TSpreadSortingTests);
  InitUnsortedData;

end.

