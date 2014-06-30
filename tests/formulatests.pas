unit formulatests;

{$mode objfpc}{$H+}

{ Deactivate this define in order to bypass tests which will raise an exception
  when the corresponding rpn formula is calculated. }
{.$DEFINE ENABLE_CALC_RPN_EXCEPTIONS}


interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, fpsmath,
  xlsbiff8 {and a project requirement for lclbase for utf8 handling},
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
    // Test calculation of rpn formulas
    procedure TestCalcRPNFormulas(AFormat: TsSpreadsheetformat);

  published
    // Writes out numbers & reads back.
    // If previous read tests are ok, this effectively tests writing.
    { BIFF2 Tests }
    procedure TestWriteRead_BIFF2_FormulaStrings;
    { BIFF5 Tests }
    procedure TestWriteRead_BIFF5_FormulaStrings;
    { BIFF8 Tests }
    procedure TestWriteRead_BIFF8_FormulaStrings;

    // Writes out and calculates formulas, read back
    { BIFF8 Tests }
    procedure TestWriteRead_BIFF8_CalcRPNFormula;
  end;

implementation

uses
  math, typinfo, lazUTF8, fpsUtils, rpnFormulaUnit;

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
  Row: Integer;
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

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF2_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF5_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF8_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel8);
end;


{ Test calculation of rpn formulas }

procedure TSpreadWriteReadFormulaTests.TestCalcRPNFormulas(AFormat: TsSpreadsheetFormat);
const
  SHEET = 'Sheet1';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Integer;
  TempFile: string;    //write xls/xml to this file and read back from it
  actual: TsArgument;
  expected: TsArgument;
  cell: PCell;
  sollValues: array of TsArgument;
  formula: String;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);
  MyWorkSheet.Options := MyWorkSheet.Options + [soCalcBeforeSaving];
  // Calculation of rpn formulas must be activated expicitely!

  { Write out test formulas.
    This include file creates various rpn formulas and stores the expected
    results in array "sollValues".
    The test file contains the text representation in column A, and the
    formula in column B. }
  Row := 0;
  {$I testcases_calcrpnformula.inc}

  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the workbook
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  for Row := 0 to MyWorksheet.GetLastRowIndex do begin
    formula := MyWorksheet.ReadAsUTF8Text(Row, 0);
    cell := MyWorksheet.FindCell(Row, 1);
    if (cell = nil) then
      fail('Error in test code: Failed to get cell ' + CellNotation(MyWorksheet, Row, 1));
    case cell^.ContentType of
      cctBool       : actual := CreateBool(cell^.BoolValue);
      cctNumber     : actual := CreateNumber(cell^.NumberValue);
      cctError      : actual := CreateError(cell^.ErrorValue);
      cctUTF8String : actual := CreateString(cell^.UTF8StringValue);
      else            fail('ContentType not supported');
    end;
    expected := SollValues[row];
    CheckEquals(ord(expected.ArgumentType), ord(actual.ArgumentType),
      'Test read calculated formula data type mismatch, formula "' + formula +
      '", cell '+CellNotation(MyWorkSheet,Row,1));
    case actual.ArgumentType of
      atBool:
        CheckEquals(BoolToStr(expected.BoolValue), BoolToStr(actual.BoolValue),
          'Test read calculated formula result mismatch, formula "' + formula +
          '", cell '+CellNotation(MyWorkSheet,Row,1));
      atNumber:
        CheckEquals(expected.NumberValue, actual.NumberValue,
          'Test read calculated formula result mismatch, formula "' + formula +
          '", cell '+CellNotation(MyWorkSheet,Row,1));
      atString:
        CheckEquals(expected.StringValue, actual.StringValue,
          'Test read calculated formula result mismatch, formula "' + formula +
          '", cell '+CellNotation(MyWorkSheet,Row,1));
      atError:
        CheckEquals(
          GetEnumName(TypeInfo(TsErrorValue), ord(expected.ErrorValue)),
          GetEnumname(TypeInfo(TsErrorValue), ord(actual.ErrorValue)),
          'Test read calculated formula error value mismatch, formula ' + formula +
          ', cell '+CellNotation(MyWorkSheet,Row,1));
    end;
  end;

  // Finalization
  MyWorkbook.Free;
  DeleteFile(TempFile);
end;


procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF8_CalcRPNFormula;
begin
  TestCalcRPNFormulas(sfExcel8);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadWriteReadFormulaTests);


end.

