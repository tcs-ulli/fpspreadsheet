unit formulatests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadFormula }
  //Write to xls/xml file and read back
  TSpreadWriteReadFormulaTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    // Test formula strings
    procedure TestWriteReadFormulaStrings(AFormat: TsSpreadsheetFormat);

  published
    // Writes out numbers & reads back.
    // If previous read tests are ok, this effectively tests writing.
    { BIFF8 Tests }
    procedure TestWriteReadBIFF8_FormulaStrings;
  end;

implementation

uses
  rpnFormulaUnit;

{ TSpreadWriteReadFormatTests }

procedure TSpreadWriteReadFormulaTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadFormulaTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadFormulaTests.TestWriteReadFormulaStrings(AFormat: TsSpreadsheetFormat);
const
  SHEET = 'Sheet1';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: String;
  Row, Col: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
  expected: String;
  actual: String;
  cell: PCell;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);

  // Write out all test formulas
  // All formulas are in column B
  WriteRPNFormulaSamples(MyWorksheet, AFormat);
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for Row := 0 to MyWorksheet.GetLastRowNumber do begin
    cell := MyWorksheet.FindCell(Row, 1);
    if (cell <> nil) and (Length(cell^.RPNFormulaValue) > 0) then begin
      actual := MyWorksheet.ReadRPNFormulaAsString(cell);
      expected := MyWorksheet.ReadAsUTF8Text(Row, 0);
      CheckEquals(actual, expected, 'Test read formula mismatch, cell '+CellNotation(MyWorkSheet,Row,Col));
    end;
  end;

  // Finalization
  MyWorkbook.Free;
  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteReadBIFF8_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel8);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadWriteReadFormulaTests);


end.

