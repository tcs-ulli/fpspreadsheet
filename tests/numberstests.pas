unit numberstests;

{$mode objfpc}{$H+}

interface

{
Adding tests/test data:
1. Add a new value to column A in the relevant worksheet, and save the spreadsheet read-only
   (for dates, there are 2 files, with different datemodes. Use them both...)
2. Increase SollNumbers array size
3. Add value from 1) to InitNormVariables so you can test against it
4. Add your read test(s), read and check read value against SollDates[<added number>]
}

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of numbers/times that should occur in spreadsheet
  SollNumbers: array[0..12] of double; //"Soll" is a German word in Dutch accountancy circles meaning "normative value to check against". There ;)
  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollNumbers;

type
  { TSpreadReadNumberTests }
  // Read from xls/xml file with known values
  TSpreadReadNumberTests= class(TTestCase)
  private
    // Tries to read number in column A, specified (0-based) row
    procedure TestReadNumber(FileName: string; Row: integer);
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Reads numbers values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestReadNumber0; //number tests
    procedure TestReadNumber1; //number and time
    procedure TestReadNumber2;
    procedure TestReadNumber3;
    procedure TestReadNumber4; //time only tests start here
    procedure TestReadNumber5;
    procedure TestReadNumber6;
    procedure TestReadNumber7;
    procedure TestReadNumber8;
    procedure TestReadNumber9;
    procedure TestReadNumber10;
    procedure TestReadNumber11;
  end;

  { TSpreadWriteReadNumberTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadNumberTests= class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Reads numbers values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestWriteReadNumbers;
  end;


implementation

const
  TestFileBIFF8='testbiff8.xls'; //with 1904 numbermode number system
  TestFileBIFF8_1899='testbiff8_1899.xls'; //with 1899/1900 numbermode number system
  NumbersSheet = 'Numbers'; //worksheet name

// Initialize array with variables that represent the values
// we expect to be in the test spreadsheet files.
//
// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollNumbers;
begin
  // Set up norm - MUST match spreadsheet cells exactly
  SollNumbers[0]:=-59000000; //minus 59 million
  SollNumbers[1]:=-988;
  SollNumbers[2]:=-124.23432;
  SollNumbers[3]:=-81.9028508730274;
  SollNumbers[4]:=-15;
  SollNumbers[5]:=-0.002934; //minus small fraction
  SollNumbers[6]:=-0; //minus zero
  SollNumbers[7]:=0; //zero
  SollNumbers[8]:=0.000000005; //small fraction
  SollNumbers[9]:=0.982394; //almost 1
  SollNumbers[10]:=3.14159265358979; //some parts of pi
  SollNumbers[11]:=59000000; //59 million
  SollNumbers[12]:=59000000.1; //same + a tenth
end;

{ TSpreadWriteReadNumberTests }

procedure TSpreadWriteReadNumberTests.SetUp;
begin
  inherited SetUp;
  InitSollNumbers;
end;

procedure TSpreadWriteReadNumberTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadNumberTests.TestWriteReadNumbers;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualNumber: double;
  Row: Cardinal;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:=MyWorkBook.AddWorksheet(NumbersSheet);
  for Row := Low(SollNumbers) to High(SollNumbers) do
  begin
    MyWorkSheet.WriteNumber(Row,0,SollNumbers[Row]);
    // Some checks inside worksheet itself
    ActualNumber:=MyWorkSheet.ReadAsNumber(Row,0);
    CheckEquals(SollNumbers[Row],ActualNumber,'Test value mismatch cell '+CellNotation(MyWorkSheet,Row));
  end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook,NumbersSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  // Read test data from A column & compare if written=original
  for Row := Low(SollNumbers) to High(SollNumbers) do
  begin
    ActualNumber:=MyWorkSheet.ReadAsNumber(Row,0);
    CheckEquals(SollNumbers[Row],ActualNumber,'Test value mismatch cell '+CellNotation(MyWorkSheet,Row));
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadReadNumberTests.TestReadNumber(FileName: string; Row: integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualNumber: double;
begin
  if Row>High(SollNumbers) then
    fail('Error in test code: array bounds overflow. Check array size is correct.');

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(FileName, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook,NumbersSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  ActualNumber:=MyWorkSheet.ReadAsNumber(Row, 0);
  CheckEquals(SollNumbers[Row],ActualNumber,'Test value mismatch '
    +'cell '+CellNotation(MyWorkSheet,Row));

  // Finalization
  MyWorkbook.Free;
end;

procedure TSpreadReadNumberTests.SetUp;
begin
  InitSollNumbers;
end;

procedure TSpreadReadNumberTests.TearDown;
begin

end;

procedure TSpreadReadNumberTests.TestReadNumber0;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,0);
end;

procedure TSpreadReadNumberTests.TestReadNumber1;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,1);
end;

procedure TSpreadReadNumberTests.TestReadNumber2;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,2);
end;

procedure TSpreadReadNumberTests.TestReadNumber3;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,3);
end;

procedure TSpreadReadNumberTests.TestReadNumber4;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,4);
end;

procedure TSpreadReadNumberTests.TestReadNumber5;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,5);
end;

procedure TSpreadReadNumberTests.TestReadNumber6;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,6);
end;

procedure TSpreadReadNumberTests.TestReadNumber7;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,7);
end;

procedure TSpreadReadNumberTests.TestReadNumber8;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,8);
end;

procedure TSpreadReadNumberTests.TestReadNumber9;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,9);
end;

procedure TSpreadReadNumberTests.TestReadNumber10;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,10);
end;

procedure TSpreadReadNumberTests.TestReadNumber11;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,11);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadReadNumberTests);
  RegisterTest(TSpreadWriteReadNumberTests);
  InitSollNumbers; //useful to have norm data if other code want to use this unit
end.


