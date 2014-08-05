unit emptycelltests;

{$mode objfpc}{$H+}

interface
{ Tests for correct location of empty cells
This unit test is writing out to and reading back from files.
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8, fpsopendocument, {and a project requirement for lclbase for utf8 handling}
  testsutility;

var
  // Norm to test against - list of strings that show the layout of empty and occupied cells
  SollLayoutStrings: array[0..5] of string;

  procedure InitSollLayouts;

type
  { TSpreadWriteReadEmptyCellTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadEmptyCellTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteReadEmptyCells(AFormat: TsSpreadsheetFormat;
      ALayout: Integer; AInverted: Boolean);

  published
    // Writes out cell layouts

    { BIFF2 file format tests }
    procedure TestWriteReadEmptyCells_BIFF2_0;
    procedure TestWriteReadEmptyCells_BIFF2_0_inv;
    procedure TestWriteReadEmptyCells_BIFF2_1;
    procedure TestWriteReadEmptyCells_BIFF2_1_inv;
    procedure TestWriteReadEmptyCells_BIFF2_2;
    procedure TestWriteReadEmptyCells_BIFF2_2_inv;
    procedure TestWriteReadEmptyCells_BIFF2_3;
    procedure TestWriteReadEmptyCells_BIFF2_3_inv;
    procedure TestWriteReadEmptyCells_BIFF2_4;
    procedure TestWriteReadEmptyCells_BIFF2_4_inv;
    procedure TestWriteReadEmptyCells_BIFF2_5;
    procedure TestWriteReadEmptyCells_BIFF2_5_inv;

    { BIFF5 file format tests }
    procedure TestWriteReadEmptyCells_BIFF5_0;
    procedure TestWriteReadEmptyCells_BIFF5_0_inv;
    procedure TestWriteReadEmptyCells_BIFF5_1;
    procedure TestWriteReadEmptyCells_BIFF5_1_inv;
    procedure TestWriteReadEmptyCells_BIFF5_2;
    procedure TestWriteReadEmptyCells_BIFF5_2_inv;
    procedure TestWriteReadEmptyCells_BIFF5_3;
    procedure TestWriteReadEmptyCells_BIFF5_3_inv;
    procedure TestWriteReadEmptyCells_BIFF5_4;
    procedure TestWriteReadEmptyCells_BIFF5_4_inv;
    procedure TestWriteReadEmptyCells_BIFF5_5;
    procedure TestWriteReadEmptyCells_BIFF5_5_inv;

    { BIFF8 file format tests }
    procedure TestWriteReadEmptyCells_BIFF8_0;
    procedure TestWriteReadEmptyCells_BIFF8_0_inv;
    procedure TestWriteReadEmptyCells_BIFF8_1;
    procedure TestWriteReadEmptyCells_BIFF8_1_inv;
    procedure TestWriteReadEmptyCells_BIFF8_2;
    procedure TestWriteReadEmptyCells_BIFF8_2_inv;
    procedure TestWriteReadEmptyCells_BIFF8_3;
    procedure TestWriteReadEmptyCells_BIFF8_3_inv;
    procedure TestWriteReadEmptyCells_BIFF8_4;
    procedure TestWriteReadEmptyCells_BIFF8_4_inv;
    procedure TestWriteReadEmptyCells_BIFF8_5;
    procedure TestWriteReadEmptyCells_BIFF8_5_inv;

    { OpenDocument file format tests }
    procedure TestWriteReadEmptyCells_ODS_0;
    procedure TestWriteReadEmptyCells_ODS_0_inv;
    procedure TestWriteReadEmptyCells_ODS_1;
    procedure TestWriteReadEmptyCells_ODS_1_inv;
    procedure TestWriteReadEmptyCells_ODS_2;
    procedure TestWriteReadEmptyCells_ODS_2_inv;
    procedure TestWriteReadEmptyCells_ODS_3;
    procedure TestWriteReadEmptyCells_ODS_3_inv;
    procedure TestWriteReadEmptyCells_ODS_4;
    procedure TestWriteReadEmptyCells_ODS_4_inv;
    procedure TestWriteReadEmptyCells_ODS_5;
    procedure TestWriteReadEmptyCells_ODS_5_inv;

  end;

implementation

const
  EmptyCellsSheet = 'EmptyCells';

procedure InitSollLayouts;
begin
  SollLayoutStrings[0] := 'x      x|'+
                          '        |'+
                          '   x    |'+
                          '        |'+
                          'x      x|';

  SollLayoutStrings[1] := 'xx  xx  |'+
                          '  xx  xx|'+
                          'xx  xx  |';

  SollLayoutStrings[2] := '        |'+
                          'xxxxxxxx|'+
                          '        |';

  SollLayoutStrings[3] := '        |'+
                          'xxxxxxxx';

  SollLayoutStrings[4] := 'xxxxxxxx|'+
                          '        |'+
                          '        ';

  SollLayoutStrings[5] := 'xxxxxxxx|'+
                          '  x  x  |'+
                          '        |';
end;


{ TSpreadWriteReadEmptyCellTests }

procedure TSpreadWriteReadEmptyCellTests.SetUp;
begin
  inherited SetUp;
  InitSollLayouts;
end;

procedure TSpreadWriteReadEmptyCellTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells(
  AFormat: TsSpreadsheetFormat; ALayout: Integer; AInverted: Boolean);
const
  CELLTEXT = 'x';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  L: TStringList;
  s: String;
begin
  TempFile := GetTempFileName;

  L := TStringList.Create;
  try
    L.Delimiter := '|';
    L.StrictDelimiter := true;
    L.DelimitedText := SollLayoutStrings[ALayout];

    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkSheet:= MyWorkBook.AddWorksheet(EmptyCellsSheet);

        // Write out cells
        for row := 0 to L.Count-1 do begin
          s := L[row];
          for col := 0 to Length(s)-1 do begin
            if AInverted then begin
              if s[col+1] = ' ' then s[col+1] := 'x' else s[col+1] := ' ';
            end;
            if s[col+1] = 'x' then
              MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
          end;
        end;
        MyWorkBook.WriteToFile(TempFile, AFormat, true);
      finally
        MyWorkbook.Free;
      end;

    // Open the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, EmptyCellsSheet);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet');

      for row := 0 to MyWorksheet.GetLastRowIndex do begin
        SetLength(s, MyWorksheet.GetLastColIndex + 1);
        for col := 0 to MyWorksheet.GetLastColIndex do begin
          MyCell := MyWorksheet.FindCell(row, col);
          if MyCell = nil then s[col+1] := ' ' else s[col+1] := 'x';
          if AInverted then begin
            if s[col+1] = ' ' then s[col+1] := 'x' else s[col+1] := ' ';
          end;
        end;
        if AInverted then
          while Length(s) < Length(L[row]) do s := s + 'x'
        else
          while Length(s) < Length(L[row]) do s := s + ' ';
        CheckEquals(L[row], s,
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

{ BIFF2 tests }

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_0;
begin
  TestWriteReadEmptyCells(sfExcel2, 0, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_0_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 0, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_1;
begin
  TestWriteReadEmptyCells(sfExcel2, 1, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_1_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 1, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_2;
begin
  TestWriteReadEmptyCells(sfExcel2, 2, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_2_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 2, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_3;
begin
  TestWriteReadEmptyCells(sfExcel2, 3, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_3_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 3, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_4;
begin
  TestWriteReadEmptyCells(sfExcel2, 4, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_4_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 4, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_5;
begin
  TestWriteReadEmptyCells(sfExcel2, 5, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_5_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 5, true);
end;


{ BIFF5 tests }

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_0;
begin
  TestWriteReadEmptyCells(sfExcel5, 0, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_0_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 0, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_1;
begin
  TestWriteReadEmptyCells(sfExcel5, 1, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_1_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 1, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_2;
begin
  TestWriteReadEmptyCells(sfExcel5, 2, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_2_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 2, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_3;
begin
  TestWriteReadEmptyCells(sfExcel5, 3, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_3_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 3, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_4;
begin
  TestWriteReadEmptyCells(sfExcel5, 4, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_4_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 4, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_5;
begin
  TestWriteReadEmptyCells(sfExcel5, 5, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_5_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 5, true);
end;


{ BIFF8 tests }

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_0;
begin
  TestWriteReadEmptyCells(sfExcel8, 0, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_0_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 0, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_1;
begin
  TestWriteReadEmptyCells(sfExcel8, 1, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_1_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 1, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_2;
begin
  TestWriteReadEmptyCells(sfExcel8, 2, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_2_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 2, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_3;
begin
  TestWriteReadEmptyCells(sfExcel8, 3, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_3_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 3, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_4;
begin
  TestWriteReadEmptyCells(sfExcel8, 4, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_4_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 4, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_5;
begin
  TestWriteReadEmptyCells(sfExcel8, 5, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_5_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 5, true);
end;

{ OpenDocument tests }

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_0;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 0, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_0_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 0, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_1;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 1, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_1_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 1, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_2;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 2, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_2_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 2, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_3;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 3, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_3_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 3, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_4;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 4, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_4_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 4, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_5;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 5, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_5_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 5, true);
end;


initialization
  RegisterTest(TSpreadWriteReadEmptyCellTests);
  InitSollLayouts;

end.

