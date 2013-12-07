unit datetests;

{$mode objfpc}{$H+}

{
Adding tests/test data:
1. Add a new value to column A in the relevant worksheet, and save the spreadsheet read-only
   (for dates, there are 2 files, with different datemodes. Use them both...)
2. Increase SollDates array size
3. Add value from 1) to InitNormVariables so you can test against it
4. Add your read test(s), read and check read value against SollDates[<added number>]
}

interface

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of dates/times that should occur in spreadsheet
  SollDates: array[0..11] of TDateTime; //"Soll" is a German word in Dutch accountancy circles meaning "normative value to check against". There ;)
  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollDates;

type
  { TSpreadReadDateTests }
  // Read from xls/xml file with known values
  TSpreadReadDateTests= class(TTestCase)
  private
    // Tries to read date in column A, specified (0-based) row
    procedure TestReadDate(FileName: string; Row: integer);
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Reads dates, date/time and time values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestReadDate0; //date tests
    procedure TestReadDate1; //date and time
    procedure TestReadDate2;
    procedure TestReadDate3;
    procedure TestReadDate4; //time only tests start here
    procedure TestReadDate5;
    procedure TestReadDate6;
    procedure TestReadDate7;
    procedure TestReadDate8;
    procedure TestReadDate9;
    procedure TestReadDate10;
    procedure TestReadDate11;
    procedure TestReadDate1899_0; //same as above except with the 1899/1900 date system set
    procedure TestReadDate1899_1;
    procedure TestReadDate1899_2;
    procedure TestReadDate1899_3;
    procedure TestReadDate1899_4;
    procedure TestReadDate1899_5;
    procedure TestReadDate1899_6;
    procedure TestReadDate1899_7;
    procedure TestReadDate1899_8;
    procedure TestReadDate1899_9;
    procedure TestReadDate1899_10;
    procedure TestReadDate1899_11;
  end;

  { TSpreadWriteReadDateTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadDateTests= class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Reads dates, date/time and time values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestWriteReadDates;
  end;


implementation

const
  TestFileBIFF8='testbiff8.xls'; //with 1904 datemode date system
  TestFileBIFF8_1899='testbiff8_1899.xls'; //with 1899/1900 datemode date system
  DatesSheet = 'Dates'; //worksheet name

// Initialize array with variables that represent the values
// we expect to be in the test spreadsheet files.
//
// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollDates;
begin
  // Set up norm - MUST match spreadsheet cells exactly
  SollDates[0]:=EncodeDate(1905,09,12); //FPC number 2082
  SollDates[1]:=EncodeDate(1908,09,12)+EncodeTime(12,0,0,0); //noon
  SollDates[2]:=EncodeDate(2013,11,24);
  SollDates[3]:=EncodeDate(2030,12,31);
  SollDates[4]:=EncodeTime(0,0,0,0);
  SollDates[5]:=EncodeTime(0,0,1,0);
  SollDates[6]:=EncodeTime(1,0,0,0);
  SollDates[7]:=EncodeTime(3,0,0,0);
  SollDates[8]:=EncodeTime(12,0,0,0);
  SollDates[9]:=EncodeTime(18,0,0,0);
  SollDates[10]:=EncodeTime(23,59,0,0);
  SollDates[11]:=EncodeTime(23,59,59,0);
end;

{ TSpreadWriteReadDateTests }

procedure TSpreadWriteReadDateTests.SetUp;
begin
  inherited SetUp;
  InitSollDates;
end;

procedure TSpreadWriteReadDateTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualDateTime: TDateTime;
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
  MyWorkSheet:=MyWorkBook.AddWorksheet(DatesSheet);
  for Row := Low(SollDates) to High(SollDates) do
  begin
    MyWorkSheet.WriteDateTime(Row,0,SollDates[Row]);
    // Some checks inside worksheet itself
    if not(MyWorkSheet.ReadAsDateTime(Row,0,ActualDateTime)) then
      Fail('Failed writing date time for cell '+CellNotation(Row));
    CheckEquals(SollDates[Row],ActualDateTime,'Test date/time value mismatch cell '+CellNotation(Row));
  end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook,DatesSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  // Read test data from A column & compare if written=original
  for Row := Low(SollDates) to High(SollDates) do
  begin
    if not(MyWorkSheet.ReadAsDateTime(Row,0,ActualDateTime)) then
      Fail('Could not read date time for cell '+CellNotation(Row));
    CheckEquals(SollDates[Row],ActualDateTime,'Test date/time value mismatch cell '+CellNotation(Row));
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadReadDateTests.TestReadDate(FileName: string; Row: integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualDateTime: TDateTime;
begin
  if Row>High(SollDates) then
    fail('Error in test code: array bounds overflow. Check array size is correct.');

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(FileName, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook,DatesSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  if not(MyWorkSheet.ReadAsDateTime(Row, 0, ActualDateTime)) then
    Fail('Could not read date time for cell '+CellNotation(Row));
  CheckEquals(SollDates[Row],ActualDateTime,'Test date/time value mismatch '
    +'cell '+CellNotation(Row));

  // Finalization
  MyWorkbook.Free;
end;

procedure TSpreadReadDateTests.SetUp;
begin
  InitSollDates;
end;

procedure TSpreadReadDateTests.TearDown;
begin

end;

procedure TSpreadReadDateTests.TestReadDate0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,0);
end;

procedure TSpreadReadDateTests.TestReadDate1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,1);
end;

procedure TSpreadReadDateTests.TestReadDate2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,2);
end;

procedure TSpreadReadDateTests.TestReadDate3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,3);
end;

procedure TSpreadReadDateTests.TestReadDate4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,4);
end;

procedure TSpreadReadDateTests.TestReadDate5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,5);
end;

procedure TSpreadReadDateTests.TestReadDate6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,6);
end;

procedure TSpreadReadDateTests.TestReadDate7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,7);
end;

procedure TSpreadReadDateTests.TestReadDate8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,8);
end;

procedure TSpreadReadDateTests.TestReadDate9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,9);
end;

procedure TSpreadReadDateTests.TestReadDate10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,10);
end;

procedure TSpreadReadDateTests.TestReadDate11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,11);
end;

procedure TSpreadReadDateTests.TestReadDate1899_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,0);
end;

procedure TSpreadReadDateTests.TestReadDate1899_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,1);
end;

procedure TSpreadReadDateTests.TestReadDate1899_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,2);
end;

procedure TSpreadReadDateTests.TestReadDate1899_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,3);
end;

procedure TSpreadReadDateTests.TestReadDate1899_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,4);
end;

procedure TSpreadReadDateTests.TestReadDate1899_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,5);
end;

procedure TSpreadReadDateTests.TestReadDate1899_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,6);
end;

procedure TSpreadReadDateTests.TestReadDate1899_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,7);
end;

procedure TSpreadReadDateTests.TestReadDate1899_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,8);
end;

procedure TSpreadReadDateTests.TestReadDate1899_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,9);
end;

procedure TSpreadReadDateTests.TestReadDate1899_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,10);
end;

procedure TSpreadReadDateTests.TestReadDate1899_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,11);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadReadDateTests);
  RegisterTest(TSpreadWriteReadDateTests);
  InitSollDates; //useful to have norm data if other code want to use this unit
end.


