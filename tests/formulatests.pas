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
    { BIFF2 Tests }
    procedure TestWriteReadBIFF2_FormulaStrings;
    { BIFF5 Tests }
    procedure TestWriteReadBIFF5_FormulaStrings;
    { BIFF8 Tests }
    procedure TestWriteReadBIFF8_FormulaStrings;
  end;

implementation

uses
  fpsUtils, rpnFormulaUnit;

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
  fs: TFormatSettings;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
//  MyWorkbook.FormatSettings.DecimalSeparator := '.';
//  MyWorkbook.FormatSettings.ShortDateFormat := 'yyyy-mm-dd';
  MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);

  // Write out all test formulas
  // All formulas are in column B
  WriteRPNFormulaSamples(MyWorksheet, AFormat, true);
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFormulas := true;

  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for Row := 0 to MyWorksheet.GetLastRowIndex do begin
    cell := MyWorksheet.FindCell(Row, 1);
    if (cell <> nil) and (Length(cell^.RPNFormulaValue) > 0) then begin
      actual := MyWorksheet.ReadRPNFormulaAsString(cell);
      expected := MyWorksheet.ReadAsUTF8Text(Row, 0);
      CheckEquals(expected, actual, 'Test read formula mismatch, cell '+CellNotation(MyWorkSheet,Row,1));
    end;
  end;

  // Finalization
  MyWorkbook.Free;
  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteReadBIFF2_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteReadBIFF5_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteReadBIFF8_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel8);
end;

initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadWriteReadFormulaTests);


end.

