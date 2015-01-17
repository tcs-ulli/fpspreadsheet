unit celltypetests;

{$mode objfpc}{$H+}

interface
{ Cell type tests
This unit tests writing the various cell data types out to and reading them 
back from files.
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadCellTypeTests }
  // Write cell types to xls/xml file and read back
  TSpreadWriteReadCellTypeTests = class(TTestCase)
  private

  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_Bool(AFormat: TsSpreadsheetFormat);

  published
    // BIFF2 test cases
    procedure TestWriteRead_Bool_BIFF2;

    // BIFF5 test cases
    procedure TestWriteRead_Bool_BIFF5;

    // BIFF8 test cases
    procedure TestWriteRead_Bool_BIFF8;

    // ODS test cases
    procedure TestWriteRead_Bool_ODS;

    // OOXML test cases
    procedure TestWriteRead_Bool_OOXML;

    // CSV test cases
    procedure TestWriteRead_Bool_CSV;
  end;

implementation

uses
  TypInfo;

const
  SheetName = 'CellTypes';


{ TSpreadWriteReadCellTypeTests }

procedure TSpreadWriteReadCellTypeTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadCellTypeTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadCellTypeTests.TestWriteRead_Bool(AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  value: Boolean;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];

    MyWorkSheet:= MyWorkBook.AddWorksheet(SheetName);
    row := 0;

    // direct cells with TRUE and FALSE
    MyWorksheet.WriteBoolValue(row, 0, false);       // A1
    Myworksheet.WriteBoolValue(row, 1, true);        // B1
    inc(row);

    // cells with TRUE and FALSE as formula results
    MyWorksheet.WriteFormula(row, 0, '=FALSE()');    // A2
    MyWorksheet.WriteFormula(row, 1, '=TRUE()');     // B2
    inc(row);

    // Merged cells with TRUE and FALSE
    MyWorksheet.MergeCells(row, 0, row+1, 0);        // A3
    Myworksheet.WriteBoolValue(row, 0, false);
    MyWorksheet.MergeCells(row, 1, row+1, 1);        // B3
    MyWorksheet.WriteBoolValue(row, 1, true);
    inc(row, 2);

    // Merged cells with TRUE and FALSE function results
    MyWorksheet.MergeCells(row, 0, row+1, 0);        // A5
    MyWorksheet.WriteFormula(row, 0, '=FALSE()');
    MyWorksheet.MergeCells(row, 1, row+1, 1);        // B5
    MyWorksheet.WriteFormula(row, 1, '=TRUE()');

    TempFile := NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat in [sfExcel2, sfCSV] then
      MyWorksheet := MyWorkbook.GetFirstWorksheet  // only 1 sheet for BIFF2
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SheetName);
    if MyWorksheet = nil then
      fail('Error in test code. Failed to get named worksheet');

    // Try to read cell
    row := 0;
    repeat
      for col:=0 to 1 do
      begin
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell '+CellNotation(MyWorksheet, row, col));
        CheckEquals(
          GetEnumName(TypeInfo(TCellContentType), ord(cctBool)),
          GetEnumName(TypeInfo(TCellContentType), ord(MyCell^.ContentType)),
          'Test saved content type mismatch, cell '+CellNotation(MyWorksheet, row, col)
        );
        value := MyCell^.BoolValue;
        CheckEquals(
          Boolean(col),
          MyCell^.BoolValue,
          'Test saved boolean value mismatch, cell '+CellNotation(MyWorksheet, row, col));
      end;

      case row of
        0, 1: inc(row);
        2   : inc(row, 2);
        else  break;
      end;

    until false;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


{ BIFF2 }
procedure TSpreadWriteReadCellTypeTests.TestWriteRead_Bool_BIFF2;
begin
  TestWriteRead_Bool(sfExcel2);
end;

{ BIFF5 }
procedure TSpreadWriteReadCellTypeTests.TestWriteRead_Bool_BIFF5;
begin
  TestWriteRead_Bool(sfExcel5);
end;

{ BIFF8 }
procedure TSpreadWriteReadCellTypeTests.TestWriteRead_Bool_BIFF8;
begin
  TestWriteRead_Bool(sfExcel8);
end;

{ ODS }
procedure TSpreadWriteReadCellTypeTests.TestWriteRead_Bool_ODS;
begin
  TestWriteRead_Bool(sfOpenDocument);
end;

{ OOXML }
procedure TSpreadWriteReadCellTypeTests.TestWriteRead_Bool_OOXML;
begin
  TestWriteRead_Bool(sfOOXML);
end;

{ CSV }
procedure TSpreadWriteReadCellTypeTests.TestWriteRead_Bool_CSV;
begin
  TestWriteRead_Bool(sfCSV);
end;


initialization
  RegisterTest(TSpreadWriteReadCellTypeTests);

end.

