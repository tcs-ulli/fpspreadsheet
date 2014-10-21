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
    procedure Test_Sorting(
      ASortByCols: Boolean;
      AMode: Integer  // AMode = 0: number, 1: strings, 2: mixed
    );

  published
    procedure Test_SortingByCols_Numbers;
    procedure Test_SortingByCols_Strings;
    procedure Test_SortingByCols_Mixed;
              {
    procedure Test_SortingByRows_Numbers;
    procedure Test_SortingByRows_Strings;
    procedure Test_SortingByRows_Mixed;
    }
  end;

implementation

const
  SortingTestSheet = 'Sorting';

procedure InitUnsortedData;
// When sorted the value is equal to the index
begin
  SollSortNumbers[0] := 9;       // Equal count of numbers and strings needed
  SollSortNumbers[1] := 8;
  SollSortNumbers[2] := 5;
  SollSortNumbers[3] := 2;
  SollSortNumbers[4] := 6;
  SollSortNumbers[5] := 7;
  SollSortNumbers[6] := 1;
  SollSortNumbers[7] := 3;
  SollSortNumbers[8] := 4;
  SollSortNumbers[9] := 0;

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

procedure TSpreadSortingTests.Test_Sorting(ASortByCols: Boolean;
  AMode: Integer);
const
  AFormat = sfExcel8;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  i, row, col: Integer;
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
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(SortingTestSheet);

    col := 0;
    row := 0;
    SetLength(sortParams.Keys, 1);
    sortparams.Keys[0].ColRowIndex := 0;
    if ASortByCols then begin
      sortParams.SortByCols := true;
      r1 := 0;
      r2 := High(SollSortNumbers);
      c1 := 0;
      c2 := 0;
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
      sortParams.SortByCols := false;
      r1 := 0;
      r2 := 0;
      c1 := 0;
      c2 := High(SollSortNumbers);
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
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, SortingTestSheet);
      if MyWorksheet = nil then
        fail('Error in test code. Failed to get named worksheet');

      sortParams.Keys[0].Order := sortDir;
      MyWorksheet.Sort(sortParams, r1,c1, r2, c2);

      if ASortByCols then
        case AMode of
          0: for i:=0 to MyWorksheet.GetLastColIndex do
             begin
               actualNumber := MyWorksheet.ReadAsNumber(i, col);
               if sortDir = ssoAscending then expectedNumber := i
                 else expectedNumber := High(SollSortNumbers)-i;
               CheckEquals(actualnumber, expectedNumber,
                 'Sorted number cells mismatch, cell '+CellNotation(MyWorksheet, i, col));
             end;
          1: for i:=0 to Myworksheet.GetLastColIndex do
             begin
               actualString := MyWorksheet.ReadAsUTF8Text(i, col);
               if sortDir = ssoAscending then expectedString := char(ord('A') + i)
                 else expectedString := char(ord('A') + High(SollSortStrings)-i);
               CheckEquals(actualString, expectedString,
                 'Sorted string cells mismatch, cell '+CellNotation(MyWorksheet, i, col));
             end;
          2: begin       (*  to be done...
               for i:=0 to High(SollNumbers) do
               begin
                 actualNumber := MyWorkbook.ReadAsNumber(i*2, col);
                 if sortdir =ssoAscending then
                   expectedNumber := i
                 CheckEquals(actualnumber, expectedNumber,
                   'Sorted number cells mismatch, cell '+CellNotation(MyWorksheet, i*2, col));
               end;
               for i:=0 to High(SollStrings) do
               begin
                 actualString := MyWorkbook.ReadAsUTF8String(i*2+1, col);
                 expectedString := SollStrings[i];
                 CheckEquals(actualString, expectedString,
                   'Sorted string cells mismatch, cell '+CellNotation(MyWorksheet, i*2+1, col));
               end;
               *)
             end;
        end  // case
      else
        case AMode of
          0: for i:=0 to MyWorksheet.GetLastColIndex do
             begin
               actualNumber := MyWorksheet.ReadAsNumber(row, i);
               if sortDir = ssoAscending then expectedNumber := i
                 else expectedNumber := High(SollSortNumbers)-i;
               CheckEquals(actualnumber, expectedNumber,
                 'Sorted number cells mismatch, cell '+CellNotation(MyWorksheet, row, i));
             end;
        1: for i:=0 to MyWorksheet.GetLastColIndex do
           begin
             actualString := MyWorksheet.ReadAsUTF8Text(row, i);
             if sortDir = ssoAscending then expectedString := char(ord('A')+i)
               else expectedString := char(ord('A') + High(SollSortStrings)-i);
             CheckEquals(actualString, expectedString,
               'Sorted string cells mismatch, cell '+CellNotation(MyWorksheet, row, i));
           end;
        2: begin{
             for i:=0 to High(SollNumbers) do
             begin
               actualNumber := MyWorkbook.ReadAsNumber(row, i*2);
               expectedNumber := SollNumbers[i];
               CheckEquals(actualnumber, expectedNumber,
                 'Sorted number cells mismatch, cell '+CellNotation(MyWorksheet, row, i*2));
             end;
             for i:=0 to High(SollStrings) do
             begin
               actualString := MyWorkbook.ReadAsUTF8String(row, i*2+1);
               expectedString := SollStrings[i];
               CheckEquals(actualString, expectedString,
                 'Sorted string cells mismatch, cell '+CellNotation(MyWorksheet, row, i*2+1));
             end;
             }
           end;
        end;  // case

    finally
      MyWorkbook.Free;
    end;
  end;  // for sortDir

  DeleteFile(TempFile);
end;


procedure TSpreadSortingTests.Test_SortingByCols_Numbers;
begin
  Test_Sorting(true, 0);
end;

procedure TSpreadSortingTests.Test_SortingByCols_Strings;
begin
  Test_Sorting(true, 1);
end;

procedure TSpreadSortingTests.Test_SortingByCols_Mixed;
begin
  //Test_Sorting(true, 2);
end;

initialization
  RegisterTest(TSpreadSortingTests);
  InitUnsortedData;

end.

