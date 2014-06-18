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
  SollNumbers: array[0..23] of double; //"Soll" is a German word in Dutch accountancy jargon meaning "normative value to check against". There ;)
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
    procedure TestReadNumber12;
    procedure TestReadNumber13;
    procedure TestReadNumber14;
    procedure TestReadNumber15;
    procedure TestReadNumber16;
    procedure TestReadNumber17;
    procedure TestReadNumber18;
    procedure TestReadNumber19;
    procedure TestReadNumber20;
    procedure TestReadNumber21;
    procedure TestReadNumber22;
    procedure TestReadNumber23;
    procedure TestReadODFNumber0; //number tests using ODF/LibreOffice file format
    procedure TestReadODFNumber1; //number and time
    procedure TestReadODFNumber2;
    procedure TestReadODFNumber3;
    procedure TestReadODFNumber4; //time only tests start here
    procedure TestReadODFNumber5;
    procedure TestReadODFNumber6;
    procedure TestReadODFNumber7;
    procedure TestReadODFNumber8;
    procedure TestReadODFNumber9;
    procedure TestReadODFNumber10;
    procedure TestReadODFNumber11;
    procedure TestReadODFNumber12;
    procedure TestReadODFNumber13;
    procedure TestReadODFNumber14;
    procedure TestReadODFNumber15;
    procedure TestReadODFNumber16;
    procedure TestReadODFNumber17;
    procedure TestReadODFNumber18;
    procedure TestReadODFNumber19;
    procedure TestReadODFNumber20;
    procedure TestReadODFNumber21;
    procedure TestReadODFNumber22;
    procedure TestReadODFNumber23;
  end;

  { TSpreadWriteReadNumberTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadNumberTests= class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    // Reads numbers values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestWriteReadNumbers(AFormat: TsSpreadsheetFormat);
  published
    procedure TestWriteReadNumbers_BIFF2;
    procedure TestWriteReadNumbers_BIFF5;
    procedure TestWriteReadNumbers_BIFF8;
    procedure TestWriteReadNumbers_ODS;
  end;


implementation

var
  TestWorksheet: TsWorksheet = nil;
  TestWorkbook: TsWorkbook = nil;
  TestFileName: String = '';

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
  SollNumbers[13]:=0.3536;  // 0.3536 formatted as percentage, no decimals
  SollNumbers[14]:=0.3536;  // 0.3536 formatted as percentage, 2 decimals
  SollNumbers[15]:=59000000.1234;  // 59 million + 0.1234 formatted with thousand separator, no decimals
  SollNumbers[16]:=59000000.1234;  // 59 million + 0.1234 formatted with thousand separator, 2 decimals
  SollNumbers[17]:=-59000000.1234; // minus 59 million + 0.1234, formatted as "scientific" with 1 decimal
  SollNumbers[18]:=-59000000.1234; // minus 59 million + 0.1234, formatted as "exp" with 2 decimals
  SollNumbers[19]:=59000000.1234;  // 59 million + 0.1234 formatted as currrency (EUROs, at end), 2 decimals
  SollNumbers[20]:=59000000.1234;  // 59 million + 0.1234 formatted as currrency (Dollars, at end), 2 decimals
  SollNumbers[21]:=-59000000.1234; // minus 59 million + 0.1234 formatted as currrency (EUROs, at end), 2 decimals
  SollNumbers[22]:=-59000000.1234; // minus 59 million + 0.1234 formatted as currrency (Dollars, at end), 2 decimals
  SollNumbers[23]:=-59000000.1234; // minus 59 million + 0.1234 formatted as currrency (Dollars, at end, neg red), 2 decimals
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

procedure TSpreadWriteReadNumberTests.TestWriteReadNumbers(AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualNumber: double;
  Row: Cardinal;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  //todo: add support for ODF format
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet := MyWorkBook.AddWorksheet(NumbersSheet);
  for Row := Low(SollNumbers) to High(SollNumbers) do
  begin
    MyWorkSheet.WriteNumber(Row,0,SollNumbers[Row]);
    // Some checks inside worksheet itself
    ActualNumber:=MyWorkSheet.ReadAsNumber(Row,0);
    CheckEquals(SollNumbers[Row],ActualNumber,'Test value mismatch cell '+CellNotation(MyWorkSheet,Row));
  end;
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook,NumbersSheet);
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

procedure TSpreadWriteReadNumberTests.TestWriteReadNumbers_BIFF2;
begin
  TestWriteReadNumbers(sfExcel2);
end;

procedure TSpreadWriteReadNumberTests.TestWriteReadNumbers_BIFF5;
begin
  TestWriteReadNumbers(sfExcel5);
end;

procedure TSpreadWriteReadNumberTests.TestWriteReadNumbers_BIFF8;
begin
  TestWriteReadNumbers(sfExcel8);
end;

procedure TSpreadWriteReadNumberTests.TestWriteReadNumbers_ODS;
begin
  TestWriteReadNumbers(sfOpenDocument);
end;


{ TSpreadReadNumberTests }

procedure TSpreadReadNumberTests.TestReadNumber(FileName: string; Row: integer);
var
  ActualNumber: double;
begin
  if Row>High(SollNumbers) then
    fail('Error in test code: array bounds overflow. Check array size is correct.');

  // Load the file only if is the file name changes.
  if (FileName <> TestFileName) then begin
    if TestWorkbook <> nil then
      TestWorkbook.Free;

    // Open the spreadsheet
    TestWorkbook := TsWorkbook.Create;
    case UpperCase(ExtractFileExt(FileName)) of
      '.XLSX': TestWorkbook.ReadFromFile(FileName, sfOOXML);
      '.ODS': TestWorkbook.ReadFromFile(FileName, sfOpenDocument);
      // Excel XLS/BIFF
      else TestWorkbook.ReadFromFile(FileName, sfExcel8);
    end;
    TestWorksheet:=GetWorksheetByName(TestWorkBook,NumbersSheet);
    if TestWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    TestFileName := FileName;
  end;

  ActualNumber := TestWorkSheet.ReadAsNumber(Row, 0);
  CheckEquals(SollNumbers[Row], ActualNumber,'Test value mismatch, '
    +'cell '+CellNotation(TestWorkSheet,Row));

  // Don't free the workbook here - it will be reused. It is destroyed at finalization.
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

procedure TSpreadReadNumberTests.TestReadNumber12;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,12);
end;

procedure TSpreadReadNumberTests.TestReadNumber13;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,13);
end;

procedure TSpreadReadNumberTests.TestReadNumber14;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,14);
end;

procedure TSpreadReadNumberTests.TestReadNumber15;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,15);
end;

procedure TSpreadReadNumberTests.TestReadNumber16;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,16);
end;

procedure TSpreadReadNumberTests.TestReadNumber17;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,17);
end;

procedure TSpreadReadNumberTests.TestReadNumber18;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,18);
end;

procedure TSpreadReadNumberTests.TestReadNumber19;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,19);
end;

procedure TSpreadReadNumberTests.TestReadNumber20;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,20);
end;

procedure TSpreadReadNumberTests.TestReadNumber21;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,21);
end;

procedure TSpreadReadNumberTests.TestReadNumber22;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,22);
end;

procedure TSpreadReadNumberTests.TestReadNumber23;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,23);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber0;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,0);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber1;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,1);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber2;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,2);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber3;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,3);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber4;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,4);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber5;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,5);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber6;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,6);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber7;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,7);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber8;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,8);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber9;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,9);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber10;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,10);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber11;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,11);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber12;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,12);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber13;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,13);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber14;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,14);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber15;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,15);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber16;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,16);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber17;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,17);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber18;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,18);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber19;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,19);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber20;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,20);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber21;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,21);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber22;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,22);
end;

procedure TSpreadReadNumberTests.TestReadODFNumber23;
begin
  TestReadNumber(ExtractFilePath(ParamStr(0)) + TestFileODF,23);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadReadNumberTests);
  RegisterTest(TSpreadWriteReadNumberTests);
  InitSollNumbers; //useful to have norm data if other code wants to use this unit

finalization
  FreeAndNil(TestWorkbook);

end.


