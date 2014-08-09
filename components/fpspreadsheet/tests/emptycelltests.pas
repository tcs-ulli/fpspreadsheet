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
  SollLayoutStrings: array[0..9] of string;

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
    procedure TestWriteReadEmptyCells_BIFF2_6;
    procedure TestWriteReadEmptyCells_BIFF2_6_inv;
    procedure TestWriteReadEmptyCells_BIFF2_7;
    procedure TestWriteReadEmptyCells_BIFF2_7_inv;
    procedure TestWriteReadEmptyCells_BIFF2_8;
    procedure TestWriteReadEmptyCells_BIFF2_8_inv;
    procedure TestWriteReadEmptyCells_BIFF2_9;
    procedure TestWriteReadEmptyCells_BIFF2_9_inv;

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
    procedure TestWriteReadEmptyCells_BIFF5_6;
    procedure TestWriteReadEmptyCells_BIFF5_6_inv;
    procedure TestWriteReadEmptyCells_BIFF5_7;
    procedure TestWriteReadEmptyCells_BIFF5_7_inv;
    procedure TestWriteReadEmptyCells_BIFF5_8;
    procedure TestWriteReadEmptyCells_BIFF5_8_inv;
    procedure TestWriteReadEmptyCells_BIFF5_9;
    procedure TestWriteReadEmptyCells_BIFF5_9_inv;

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
    procedure TestWriteReadEmptyCells_BIFF8_6;
    procedure TestWriteReadEmptyCells_BIFF8_6_inv;
    procedure TestWriteReadEmptyCells_BIFF8_7;
    procedure TestWriteReadEmptyCells_BIFF8_7_inv;
    procedure TestWriteReadEmptyCells_BIFF8_8;
    procedure TestWriteReadEmptyCells_BIFF8_8_inv;
    procedure TestWriteReadEmptyCells_BIFF8_9;
    procedure TestWriteReadEmptyCells_BIFF8_9_inv;

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
    procedure TestWriteReadEmptyCells_ODS_6;
    procedure TestWriteReadEmptyCells_ODS_6_inv;
    procedure TestWriteReadEmptyCells_ODS_7;
    procedure TestWriteReadEmptyCells_ODS_7_inv;
    procedure TestWriteReadEmptyCells_ODS_8;
    procedure TestWriteReadEmptyCells_ODS_8_inv;
    procedure TestWriteReadEmptyCells_ODS_9;
    procedure TestWriteReadEmptyCells_ODS_9_inv;

    { OOXML file format tests }
    procedure TestWriteReadEmptyCells_OOXML_0;
    procedure TestWriteReadEmptyCells_OOXML_0_inv;
    procedure TestWriteReadEmptyCells_OOXML_1;
    procedure TestWriteReadEmptyCells_OOXML_1_inv;
    procedure TestWriteReadEmptyCells_OOXML_2;
    procedure TestWriteReadEmptyCells_OOXML_2_inv;
    procedure TestWriteReadEmptyCells_OOXML_3;
    procedure TestWriteReadEmptyCells_OOXML_3_inv;
    procedure TestWriteReadEmptyCells_OOXML_4;
    procedure TestWriteReadEmptyCells_OOXML_4_inv;
    procedure TestWriteReadEmptyCells_OOXML_5;
    procedure TestWriteReadEmptyCells_OOXML_5_inv;
    procedure TestWriteReadEmptyCells_OOXML_6;
    procedure TestWriteReadEmptyCells_OOXML_6_inv;
    procedure TestWriteReadEmptyCells_OOXML_7;
    procedure TestWriteReadEmptyCells_OOXML_7_inv;
    procedure TestWriteReadEmptyCells_OOXML_8;
    procedure TestWriteReadEmptyCells_OOXML_8_inv;
    procedure TestWriteReadEmptyCells_OOXML_9;
    procedure TestWriteReadEmptyCells_OOXML_9_inv;

  end;

implementation

const
  EmptyCellsSheet = 'EmptyCells';

procedure InitSollLayouts;
begin
  SollLayoutStrings[0] := 'x      x|'+
                          '        |'+
                          '  ox    |'+
                          '        |'+
                          'x      x|';

  SollLayoutStrings[1] := 'xx  xx  |'+
                          '  xx  xx|'+
                          'xx  xx  |';

  SollLayoutStrings[2] := 'xxooxxoo|'+
                          'ooxxooxx|'+
                          'xxooxxoo|';

  SollLayoutStrings[3] := '        |'+
                          'xxxxxxxx|'+
                          '        |';

  SollLayoutStrings[4] := '        |'+
                          'xxxxxxxx';

  SollLayoutStrings[5] := 'xxxxxxxx|'+
                          '        |'+
                          '        ';

  SollLayoutStrings[6] := 'xxxxxxxx|'+
                          '  x  x  |'+
                          '        |';

  SollLayoutStrings[7] := '        |'+
                          '        |'+
                          '   xx   |'+
                          '   xx   |';

  SollLayoutStrings[8] := '        |'+
                          '        |'+
                          '   x x  |'+
                          '   x x  |';

  SollLayoutStrings[9] := 'oooooooo|'+
                          'oooooooo|'+
                          'oooxoxoo|'+
                          '   x x  |';

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

  function FixODS(s: String): String;
  // In this test, ODS cannot distinguish between a blank and a nonexisting cell
  var
    i: Integer;
  begin
    Result := s;
    for i := 1 to Length(Result) do
      if Result[i] = 'o' then begin
        if AInverted then Result[i] := 'x' else Result[i] := ' ';
      end;
  end;

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
        // 'x' --> write label cell
        // 'o' --> write blank cell
        // ' ' --> do not write a cell
        for row := 0 to L.Count-1 do begin
          s := L[row];
          for col := 0 to Length(s)-1 do begin
            if AInverted then begin
              if s[col+1] = ' ' then s[col+1] := 'x'
              else
              if s[col+1] = 'x' then s[col+1] := ' ';
            end;
            if s[col+1] = 'x' then
              MyWorksheet.WriteUTF8Text(row, col, CELLTEXT)
            else
            if s[col+1] = 'o' then
              MyWorksheet.WriteBlank(row, col);
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
          if MyCell = nil then
            s[col+1] := ' '
          else
          if MyCell^.ContentType = cctEmpty then
            s[col+1] := 'o'
          else
            s[col+1] := 'x';
          if AInverted then begin
            if s[col+1] = ' ' then s[col+1] := 'x'
            else
            if s[col+1] = 'x' then s[col+1] := ' ';
          end;
        end;
        if AInverted then
          while Length(s) < Length(L[row]) do s := s + 'x'
        else
          while Length(s) < Length(L[row]) do s := s + ' ';
        if AFormat = sfOpenDocument then
          CheckEquals(FixODS(L[row]), s,
            'Test empty cell layout mismatch, cell '+CellNotation(MyWorksheet, Row, Col))
        else
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

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_6;
begin
  TestWriteReadEmptyCells(sfExcel2, 6, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_6_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 6, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_7;
begin
  TestWriteReadEmptyCells(sfExcel2, 7, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_7_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 7, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_8;
begin
  TestWriteReadEmptyCells(sfExcel2, 8, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_8_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 8, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_9;
begin
  TestWriteReadEmptyCells(sfExcel2, 9, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF2_9_inv;
begin
  TestWriteReadEmptyCells(sfExcel2, 9, true);
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

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_6;
begin
  TestWriteReadEmptyCells(sfExcel5, 6, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_6_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 6, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_7;
begin
  TestWriteReadEmptyCells(sfExcel5, 7, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_7_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 7, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_8;
begin
  TestWriteReadEmptyCells(sfExcel5, 8, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_8_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 8, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_9;
begin
  TestWriteReadEmptyCells(sfExcel5, 9, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF5_9_inv;
begin
  TestWriteReadEmptyCells(sfExcel5, 9, true);
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

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_6;
begin
  TestWriteReadEmptyCells(sfExcel8, 6, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_6_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 6, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_7;
begin
  TestWriteReadEmptyCells(sfExcel8, 7, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_7_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 7, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_8;
begin
  TestWriteReadEmptyCells(sfExcel8, 8, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_8_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 8, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_9;
begin
  TestWriteReadEmptyCells(sfExcel8, 9, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_BIFF8_9_inv;
begin
  TestWriteReadEmptyCells(sfExcel8, 9, true);
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

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_6;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 6, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_6_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 6, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_7;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 7, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_7_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 7, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_8;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 8, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_8_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 8, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_9;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 9, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_ODS_9_inv;
begin
  TestWriteReadEmptyCells(sfOpenDocument, 9, true);
end;


{ OOXML tests }

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_0;
begin
  TestWriteReadEmptyCells(sfOOXML, 0, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_0_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 0, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_1;
begin
  TestWriteReadEmptyCells(sfOOXML, 1, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_1_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 1, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_2;
begin
  TestWriteReadEmptyCells(sfOOXML, 2, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_2_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 2, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_3;
begin
  TestWriteReadEmptyCells(sfOOXML, 3, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_3_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 3, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_4;
begin
  TestWriteReadEmptyCells(sfOOXML, 4, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_4_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 4, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_5;
begin
  TestWriteReadEmptyCells(sfOOXML, 5, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_5_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 5, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_6;
begin
  TestWriteReadEmptyCells(sfOOXML, 6, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_6_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 6, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_7;
begin
  TestWriteReadEmptyCells(sfOOXML, 7, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_7_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 7, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_8;
begin
  TestWriteReadEmptyCells(sfOOXML, 8, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_8_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 8, true);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_9;
begin
  TestWriteReadEmptyCells(sfOOXML, 9, false);
end;

procedure TSpreadWriteReadEmptyCellTests.TestWriteReadEmptyCells_OOXML_9_inv;
begin
  TestWriteReadEmptyCells(sfOOXML, 9, true);
end;


initialization
  RegisterTest(TSpreadWriteReadEmptyCellTests);
  InitSollLayouts;

end.

