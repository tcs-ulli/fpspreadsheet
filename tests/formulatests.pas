unit formulatests;

{$mode objfpc}{$H+}

{ Deactivate this define in order to bypass tests which will raise an exception
  when the corresponding rpn formula is calculated. }
{.$DEFINE ENABLE_CALC_RPN_EXCEPTIONS}

{ Deactivate this define to include errors in the structure of the rpn formulas.
  Note that Excel report a corrupted file when trying to read this file }
{.DEFINE ENABLE_DEFECTIVE_FORMULAS }


interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, fpsexprparser,
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
    procedure TestWriteReadFormulaStrings(AFormat: TsSpreadsheetFormat;
      UseRPNFormula: Boolean);
    // Test calculation of rpn formulas
    procedure TestCalcFormulas(AFormat: TsSpreadsheetformat; UseRPNFormula: Boolean);

  published
    // Writes out numbers & reads back.
    // If previous read tests are ok, this effectively tests writing.
    { BIFF2 Tests }
    procedure Test_Write_Read_FormulaStrings_BIFF2;
    { BIFF5 Tests }
    procedure Test_Write_Read_FormulaStrings_BIFF5;
    { BIFF8 Tests }
    procedure Test_Write_Read_FormulaStrings_BIFF8;
    { OOXML Tests }
    procedure Test_Write_Read_FormulaStrings_OOXML;
    { ODS Tests }
    procedure Test_Write_Read_FormulaStrings_ODS;

    // Writes out and calculates rpn formulas, read back
    { BIFF2 Tests }
    procedure Test_Write_Read_CalcRPNFormula_BIFF2;
    { BIFF5 Tests }
    procedure Test_Write_Read_CalcRPNFormula_BIFF5;
    { BIFF8 Tests }
    procedure Test_Write_Read_CalcRPNFormula_BIFF8;
    { OOXML Tests }
    procedure Test_Write_Read_CalcRPNFormula_OOXML;
    { ODSL Tests }
    procedure Test_Write_Read_CalcRPNFormula_ODS;

    // Writes out and calculates string formulas, read back
    { BIFF2 Tests }
    procedure Test_Write_Read_CalcStringFormula_BIFF2;
    { BIFF5 Tests }
    procedure Test_Write_Read_CalcStringFormula_BIFF5;
    { BIFF8 Tests }
    procedure Test_Write_Read_CalcStringFormula_BIFF8;
    { OOXML Tests }
    procedure Test_Write_Read_CalcStringFormula_OOXML;
    { ODS Tests }
    procedure Test_Write_Read_CalcStringFormula_ODS;
  end;

implementation

uses
  math, typinfo, lazUTF8, fpsUtils, rpnFormulaUnit;

var
  // Array containing the "true" results of the formulas, for comparison
  SollValues: array of TsExpressionResult;

// Helper for statistics tests
const
  STATS_NUMBERS: Array[0..4] of Double = (1.0, 1.1, 1.2, 0.9, 0.8);
var
  numberArray: array[0..4] of Double;



{ TSpreadWriteReadFormatTests }

procedure TSpreadWriteReadFormulaTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadFormulaTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadFormulaTests.TestWriteReadFormulaStrings(
  AFormat: TsSpreadsheetFormat; UseRPNFormula: Boolean);
{ If UseRPNFormula is true the test formulas are generated from RPN formulas.
  Otherwise they are generated from string formulas. }
const
  SHEET = 'Sheet1';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
  formula: String;
  expected: String;
  actual: String;
  cell: PCell;
  cellB1: Double;
  cellB2: Double;
  number: Double;
  s: String;
  hr, min, sec, msec: Word;
  k: Integer;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];
    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);

    // Write out all test formulas
    // All formulas are in column B
    {$I testcases_calcrpnformula.inc}
//    WriteRPNFormulaSamples(MyWorksheet, AFormat, true, UseRPNFormula);
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];

    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    for Row := 0 to MyWorksheet.GetLastRowIndex do
    begin
      cell := MyWorksheet.FindCell(Row, 1);
      if HasFormula(cell) then begin
        actual := MyWorksheet.ReadFormulaAsString(cell);
        expected := MyWorksheet.ReadAsUTF8Text(Row, 0);
        CheckEquals(expected, actual, 'Test read formula mismatch, cell '+CellNotation(MyWorkSheet,Row,1));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_BIFF2;
begin
  TestWriteReadFormulaStrings(sfExcel2, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_BIFF5;
begin
  TestWriteReadFormulaStrings(sfExcel5, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_BIFF8;
begin
  TestWriteReadFormulaStrings(sfExcel8, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_OOXML;
begin
  TestWriteReadFormulaStrings(sfOOXML, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_ODS;
begin
  TestWriteReadFormulaStrings(sfOpenDocument, true);
end;


{ Test calculation of formulas }

procedure TSpreadWriteReadFormulaTests.TestCalcFormulas(AFormat: TsSpreadsheetFormat;
  UseRPNFormula: Boolean);
{ If UseRPNFormula is TRUE, the test formulas are generated from RPN syntax,
  otherwise string formulas are used. }
const
  SHEET = 'Sheet1';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Integer;
  TempFile: string;    //write xls/xml to this file and read back from it
  actual: TsExpressionResult;
  expected: TsExpressionResult;
  cell: PCell;
  sollValues: array of TsExpressionResult;
  formula: String;
  s: String;
  hr,min,sec,msec: Word;
  ErrorMargin: double;
  k: Integer;
  { When comparing soll and formula values we must make sure that the soll
    values are calculated from double precision numbers, they are used in
    the formula calculation as well. The next variables, along with STATS_NUMBERS
    above, hold the arguments for the direction function calls. }
  number: Double;
  cellB1: Double;
  cellB2: Double;
begin
  ErrorMargin:=0; //1.44E-7;
  //1.44E-7 for SUMSQ formula
  //6.0E-8 for SUM formula
  //4.8E-8 for MAX formula
  //2.4E-8 for now formula
  //about 1E-15 is needed for some trig functions

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];
    // Calculation of rpn formulas must be activated explicitly!

    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);
    { Write out test formulas.
      This include file creates various rpn formulas and stores the expected
      results in array "sollValues".
      The test file contains the text representation in column A, and the
      formula in column B. }
    Row := 0;
    TempFile:=GetTempFileName;
    {$I testcases_calcrpnformula.inc}
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := Myworkbook.Options + [boReadFormulas];
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    for Row := 0 to MyWorksheet.GetLastRowIndex do
    begin
      formula := MyWorksheet.ReadAsUTF8Text(Row, 0);
      cell := MyWorksheet.FindCell(Row, 1);
      if (cell = nil) then
        fail('Error in test code: Failed to get cell ' + CellNotation(MyWorksheet, Row, 1));

      case cell^.ContentType of
        cctBool       : actual := BooleanResult(cell^.BoolValue);
        cctNumber     : actual := FloatResult(cell^.NumberValue);
        cctDateTime   : actual := DateTimeResult(cell^.DateTimeValue);
        cctUTF8String : actual := StringResult(cell^.UTF8StringValue);
        cctError      : actual := ErrorResult(cell^.ErrorValue);
        cctEmpty      : actual := EmptyResult;
        else            fail('ContentType not supported');
      end;

      expected := SollValues[row];
      // Cell does not store integers!
      if expected.ResultType = rtInteger then expected := FloatResult(expected.ResInteger);

      CheckEquals(
        GetEnumName(TypeInfo(TsExpressionResult), ord(expected.ResultType)),
        GetEnumName(TypeInfo(TsExpressionResult), ord(actual.ResultType)),
        'Test read calculated formula data type mismatch, formula "' + formula +
        '", cell '+CellNotation(MyWorkSheet,Row,1));

      // The now function result is volatile, i.e. changes continuously. The
      // time for the soll value was created such that we can expect to have
      // the file value in the same second. Therefore we neglect the milliseconds.
      if formula = '=NOW()' then begin
        // Round soll value to seconds
        DecodeTime(expected.ResDateTime, hr,min,sec,msec);
        expected.ResDateTime := EncodeTime(hr, min, sec, 0);
        // Round formula value to seconds
        DecodeTime(actual.ResDateTime, hr,min,sec,msec);
        actual.ResDateTime := EncodeTime(hr,min,sec,0);
      end;

      case actual.ResultType of
        rtBoolean:
          CheckEquals(BoolToStr(expected.ResBoolean), BoolToStr(actual.ResBoolean),
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
        rtFloat:
          {$if (defined(mswindows)) or (FPC_FULLVERSION>=20701)}
          // FPC 2.6.x and trunk on Windows need this, also FPC trunk on Linux x64
          CheckEquals(expected.ResFloat, actual.ResFloat, ErrorMargin,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
          {$else}
          // Non-Windows: test without error margin
          CheckEquals(expected.NumberValue, actual.NumberValue,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
          {$endif}
        rtString:
          CheckEquals(expected.ResString, actual.ResString,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
        rtError:
          CheckEquals(
            GetEnumName(TypeInfo(TsErrorValue), ord(expected.ResError)),
            GetEnumname(TypeInfo(TsErrorValue), ord(actual.ResError)),
            'Test read calculated formula error value mismatch, formula ' + formula +
            ', cell '+CellNotation(MyWorkSheet,Row,1));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_BIFF2;
begin
  TestCalcFormulas(sfExcel2, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_BIFF5;
begin
  TestCalcFormulas(sfExcel5, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_BIFF8;
begin
  TestCalcFormulas(sfExcel8, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_OOXML;
begin
  TestCalcFormulas(sfOOXML, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_ODS;
begin
  TestCalcFormulas(sfOpenDocument, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_BIFF2;
begin
  TestCalcFormulas(sfExcel2, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_BIFF5;
begin
  TestCalcFormulas(sfExcel5, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_BIFF8;
begin
  TestCalcFormulas(sfExcel8, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_OOXML;
begin
  TestCalcFormulas(sfOOXML, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_ODS;
begin
  TestCalcFormulas(sfOpenDocument, false);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadWriteReadFormulaTests);


end.

