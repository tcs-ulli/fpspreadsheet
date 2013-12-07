unit stringtests;

{$mode objfpc}{$H+}

{
Adding tests/test data:
1. Add a new value to column A in the relevant worksheet, and save the spreadsheet read-only
   (for dates, there are 2 files, with different datemodes. Use them both...)
2. Increase SollStrings array size
3. Add value from 1) to InitNormVariables so you can test against it
4. Add your read test(s), read and check read value against SollStrings[<added number>]
}

interface

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of strings that should occur in spreadsheet
  SollStrings: array[0..6] of string; //"Soll" is a German word in Dutch accountancy circles meaning "normative value to check against". There ;)
  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollStrings;

type
  { TSpreadReadStringTests }
  // Read from xls/xml file with known values
  TSpreadReadStringTests= class(TTestCase)
  private
    // Tries to read string in column A, specified (0-based) row
    procedure TestReadString(FileName: string; Row: integer);
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Reads string values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestReadString0; //empty string
    procedure TestReadString1;
    procedure TestReadString2;
    procedure TestReadString3;
    procedure TestReadString4;
    procedure TestReadString5;
    procedure TestReadString6;
  end;

  { TSpreadWriteReadStringTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadStringTests= class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Writes out norm strings & reads back.
    // If previous read tests are ok, this effectively tests writing.
    procedure TestWriteReadStrings;
    // Testing some limits & exception handling
    procedure TestWriteReadStringsLimits;
  end;


implementation

// Initialize array with variables that represent the values
// we expect to be in the test spreadsheet files.
//
// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollStrings;
begin
  // Set up norm - MUST match spreadsheet cells exactly
  SollStrings[0]:='';
  SollStrings[1]:='a';
  SollStrings[2]:='1';
  SollStrings[3]:='The quick brown fox jumps over the lazy dog';
  SollStrings[4]:='café au lait'; //accent aigue on the e
  SollStrings[5]:='водка'; //cyrillic
  SollStrings[6]:='wódka'; //Polish o accent aigue
end;

{ TSpreadWriteReadStringTests }

procedure TSpreadWriteReadStringTests.SetUp;
begin
  inherited SetUp;
  InitSollStrings; //just for security: make sure the variables are reset to default
end;

procedure TSpreadWriteReadStringTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadStringTests.TestWriteReadStrings;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: String;
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
  MyWorkSheet:=MyWorkBook.AddWorksheet(StringsSheet);
  for Row := Low(SollStrings) to High(SollStrings) do
  begin
    MyWorkSheet.WriteUTF8Text(Row,0,SollStrings[Row]);
    // Some checks inside worksheet itself
    ActualString:=MyWorkSheet.ReadAsUTF8Text(Row,0);
    CheckEquals(SollStrings[Row],ActualString,'Test value mismatch cell '+CellNotation(Row));
  end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook,StringsSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  // Read test data from A column & compare if written=original
  for Row := Low(SollStrings) to High(SollStrings) do
  begin
    ActualString:=MyWorkSheet.ReadAsUTF8Text(Row,0);
    CheckEquals(SollStrings[Row],ActualString,'Test value mismatch cell '+CellNotation(Row));
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadStringTests.TestWriteReadStringsLimits;
const
  MaxBytesBIFF8=32758; //limit for strings in this file format
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: String;
  ExceptionMessage: string;
  LocalNormStrings: array[0..3] of string;
  Row: Cardinal;
  TempFile: string; //write xls/xml to this file and read back from it
  TestResult: boolean;
begin
  LocalNormStrings[0]:=StringOfChar('a',MaxBytesBIFF8-1);
  LocalNormStrings[1]:=StringOfChar('b',MaxBytesBIFF8);
  LocalNormStrings[2]:=StringOfChar('z',MaxBytesBiff8+1); //problems should occur here
  LocalNormStrings[3]:='this text should be readable'; //whatever happens, this text should be ok

  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:=MyWorkBook.AddWorksheet(StringsSheet);

  for Row := Low(LocalNormStrings) to High(LocalNormStrings) do
  begin
    // We could use CheckException but then you can't pass parameters
    TestResult:=true;
    try
      MyWorkSheet.WriteUTF8Text(Row,0,LocalNormStrings[Row]);
      // Some checks inside worksheet itself
      ActualString:=MyWorkSheet.ReadAsUTF8Text(Row,0);
      CheckEquals(length(LocalNormStrings[Row]),length(ActualString),
        'Test value mismatch cell '+CellNotation(Row)+
        ' for string length.');
    except
      { When over size limit we expect to hit this:
        if TextTooLong then
          Raise Exception.CreateFmt('Text value exceeds %d character limit in cell [%d,%d]. Text has been truncated.',[MaxBytes,ARow,ACol]);
      }
      //todo: rewrite when/if the fpspreadsheet exception class changes
      on E: Exception do
      begin
        if Row=2 then
          TestResult:=true
        else
        begin
          TestResult:=false;
          ExceptionMessage:=E.Message;
        end;
      end;
    end;
    // Notify user of exception if it happened where we didn't expect it:
    CheckTrue(TestResult,'Exception: '+ExceptionMessage);
  end;
  TestResult:=true;
  try
    MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  except
    //todo: rewrite when/if the fpspreadsheet exception class changes
    on E: Exception do
    begin
      if Row=2 then
        TestResult:=true
      else
      begin
        TestResult:=false;
        ExceptionMessage:=E.Message;
      end;
    end;
  end;
  // Notify user of exception if it happened where we didn't expect it:
  CheckTrue(TestResult,'Exception: '+ExceptionMessage);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook,StringsSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  // Read test data from A column & compare if written=original
  for Row := Low(LocalNormStrings) to High(LocalNormStrings) do
  begin
    ActualString:=MyWorkSheet.ReadAsUTF8Text(Row,0);
    // Allow for truncation of excessive strings by fpspreadsheet
    if length(LocalNormStrings[Row])>MaxBytesBIFF8 then
      CheckEquals(MaxBytesBIFF8,length(ActualString),
        'Test value mismatch cell '+CellNotation(Row)+
        ' for string length.')
    else
    CheckEquals(length(LocalNormStrings[Row]),length(ActualString),
      'Test value mismatch cell '+CellNotation(Row)+
      ' for string length.');
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadReadStringTests.TestReadString(FileName: string; Row: integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: string;
begin
  if Row>High(SollStrings) then
    fail('Error in test code: array bounds overflow. Check array size is correct.');

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(FileName, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook,StringsSheet);
  if MyWorksheet=nil then
    fail('Error in test code: could not retrieve worksheet.');

  ActualString:=MyWorkSheet.ReadAsUTF8Text(Row,0);
  CheckEquals(SollStrings[Row],ActualString,'Test value mismatch '
    +'cell '+CellNotation(Row));

  // Finalization
  MyWorkbook.Free;
end;

procedure TSpreadReadStringTests.SetUp;
begin
  InitSollStrings;
end;

procedure TSpreadReadStringTests.TearDown;
begin

end;

procedure TSpreadReadStringTests.TestReadString0;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,0);
end;

procedure TSpreadReadStringTests.TestReadString1;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,1);
end;

procedure TSpreadReadStringTests.TestReadString2;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,2);
end;

procedure TSpreadReadStringTests.TestReadString3;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,3);
end;

procedure TSpreadReadStringTests.TestReadString4;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,4);
end;

procedure TSpreadReadStringTests.TestReadString5;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,5);
end;

procedure TSpreadReadStringTests.TestReadString6;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,6);
end;

initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadReadStringTests);
  RegisterTest(TSpreadWriteReadStringTests);
  // Initialize the norm variables in case other units want to use it:
  InitSollStrings;
end.
