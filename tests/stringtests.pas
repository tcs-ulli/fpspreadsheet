{$ifdef fpc}
  {$if FPC_FULLVERSION>20701}
    //Explicitly specify this is an UTF8 encoded file.
    //Alternative would be UTF8 with BOM but writing UTF8 BOM is bad practice.
    //See http://wiki.lazarus.freepascal.org/FPC_Unicode_support#String_constants 
    {$codepage UTF8} //Win 65001
   {$endif} //fpc_fullversion
{$endif fpc}
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
  fpstypes, fpsallformats, fpsutils, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of strings that should occur in spreadsheet
  SollStrings: array[0..13] of string; //"Soll" is a German word in Dutch accountancy jargon meaning "normative value to check against". There ;)
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
    procedure TestReadString0; //biff8 empty string
    procedure TestReadString1;
    procedure TestReadString2;
    procedure TestReadString3;
    procedure TestReadString4;
    procedure TestReadString5;
    procedure TestReadString6;
    procedure TestReadString7;
    procedure TestReadString8;
    procedure TestReadString9;
    procedure TestReadString10;
    procedure TestReadString11;
    procedure TestReadString12;
    procedure TestReadString13;

    procedure TestReadODFString0; //OpenDocument/LibreOffice format empty string
    procedure TestReadODFString1;
    procedure TestReadODFString2;
    procedure TestReadODFString3;
    procedure TestReadODFString4;
    procedure TestReadODFString5;
    procedure TestReadODFString6;
    procedure TestReadODFString7;
    procedure TestReadODFString8;
    procedure TestReadODFString9;
    procedure TestReadODFString10;
    procedure TestReadODFString11;
    procedure TestReadODFString12;
    procedure TestReadODFString13;

    procedure TestReadOOXMLString0; //Excel xlsx format empty string
    procedure TestReadOOXMLString1;
    procedure TestReadOOXMLString2;
    procedure TestReadOOXMLString3;
    procedure TestReadOOXMLString4;
    procedure TestReadOOXMLString5;
    procedure TestReadOOXMLString6;
    procedure TestReadOOXMLString7;
    procedure TestReadOOXMLString8;
    procedure TestReadOOXMLString9;
    procedure TestReadOOXMLString10;
    procedure TestReadOOXMLString11;
    procedure TestReadOOXMLString12;
    procedure TestReadOOXMLString13;

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

var
  TestWorksheet: TsWorksheet = nil;
  TestWorkbook: TsWorkbook = nil;
  TestFileName: String = '';

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
  SollStrings[7]:='35%';  // 0.3536 formatted as percentage, no decimals
  SollStrings[8]:=FormatFloat('0.00', 35.36)+'%';  // 0.3536 formatted as percentage, 2 decimals
  SollStrings[9]:=FormatFloat('#,##0', 59000000.1234);  // 59 million + 0.1234 formatted with thousand separator, no decimals
  SollStrings[10]:=FormatFloat('#,##0.00', 59000000.1234);  // 59 million + 0.1234 formatted with thousand separator, 2 decimals
  SollStrings[11]:=FormatFloat('0.00E+00', -59000000.1234); // minus 59 million + 0.1234, formatted as "exp" with 2 decimals
  SollStrings[12]:=FormatFloat('#,##0.00 "EUR";(#,##0.00 "EUR")', 59000000.1234); // 59 million + 0.1234, formatted as "currencyRed" with 2 decimals, brackets and EUR
  SollStrings[13]:=FormatFloat('#,##0.00 "EUR";(#,##0.00 "EUR")', -59000000.1234); // minus 59 million + 0.1234, formatted as "currencyRed" with 2 decimals, brackets and EUR
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
  //todo: add support for ODF/LibreOffice format
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(StringsSheet);
    for Row := Low(SollStrings) to High(SollStrings) do
    begin
      MyWorkSheet.WriteUTF8Text(Row,0,SollStrings[Row]);
      // Some checks inside worksheet itself
      ActualString:=MyWorkSheet.ReadAsUTF8Text(Row,0);
      CheckEquals(SollStrings[Row],ActualString,'Test value mismatch cell '+CellNotation(MyWorkSheet,Row));
    end;
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, sfExcel8, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, sfExcel8);
    MyWorksheet:=GetWorksheetByName(MyWorkBook,StringsSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    // Read test data from A column & compare if written=original
    for Row := Low(SollStrings) to High(SollStrings) do
    begin
      ActualString:=MyWorkSheet.ReadAsUTF8Text(Row,0);
      CheckEquals(SollStrings[Row],ActualString,'Test value mismatch, cell '+CellNotation(MyWorkSheet,Row));
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
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

  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
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
          'Test value mismatch cell '+CellNotation(MyWorkSheet,Row)+
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
    TempFile:=NewTempFile;
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
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  try
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
          'Test value mismatch cell '+CellNotation(MyWorkSheet,Row)+
          ' for string length.')
      else
      CheckEquals(length(LocalNormStrings[Row]),length(ActualString),
        'Test value mismatch cell '+CellNotation(MyWorkSheet,Row)+
        ' for string length.');
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


{ TSpreadReadStringTests }

procedure TSpreadReadStringTests.TestReadString(FileName: string; Row: integer);
var
  ActualString: string;
  AFormat: TsSpreadsheetFormat;
begin
  if Row>High(SollStrings) then
    fail('Error in test code: array bounds overflow. Check array size is correct.');

  // Load the file only if is the file name changes.
  if (FileName <> TestFileName) then begin
    if TestWorkbook <> nil then
      TestWorkbook.Free;

    // Open the spreadsheet
    TestWorkbook := TsWorkbook.Create;
    case Uppercase(ExtractFileExt(FileName)) of
      '.XLSX': AFormat := sfOOXML;
      '.ODS' : AFormat := sfOpenDocument;
      else     AFormat := sfExcel8;
    end;
    TestWorkbook.ReadFromFile(FileName, AFormat);
    TestWorksheet := GetWorksheetByName(TestWorkBook, StringsSheet);
    if TestWorksheet=nil then
      fail('Error in test code: could not retrieve worksheet.');
  end;

  ActualString := TestWorkSheet.ReadAsUTF8Text(Row,0);
  if (Row = 11) and (AFormat = sfOpenDocument) then
    // SciFloat is not supported by Biff2 and ODS --> we just compare the value
    CheckEquals(StrToFloat(SollStrings[Row]), StrToFloat(ActualString),
      'Test value mismatch, cell ' + CellNotation(TestWorksheet, Row))
  else
    CheckEquals(SollStrings[Row], ActualString, 'Test value mismatch, '
      +'cell '+CellNotation(TestWorkSheet, Row));

  // Don't free the workbook here - it will be reused. It is destroyed at finalization.
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

procedure TSpreadReadStringTests.TestReadString7;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,7);
end;

procedure TSpreadReadStringTests.TestReadString8;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,8);
end;

procedure TSpreadReadStringTests.TestReadString9;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,9);
end;

procedure TSpreadReadStringTests.TestReadString10;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,10);
end;

procedure TSpreadReadStringTests.TestReadString11;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,11);
end;

procedure TSpreadReadStringTests.TestReadString12;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,12);
end;

procedure TSpreadReadStringTests.TestReadString13;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,13);
end;

{ ODF Tests }
procedure TSpreadReadStringTests.TestReadODFString0;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,0);
end;

procedure TSpreadReadStringTests.TestReadODFString1;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,1);
end;

procedure TSpreadReadStringTests.TestReadODFString2;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,2);
end;

procedure TSpreadReadStringTests.TestReadODFString3;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,3);
end;

procedure TSpreadReadStringTests.TestReadODFString4;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,4);
end;

procedure TSpreadReadStringTests.TestReadODFString5;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,5);
end;

procedure TSpreadReadStringTests.TestReadODFString6;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,6);
end;

procedure TSpreadReadStringTests.TestReadODFString7;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,7);
end;

procedure TSpreadReadStringTests.TestReadODFString8;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,8);
end;

procedure TSpreadReadStringTests.TestReadODFString9;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,9);
end;

procedure TSpreadReadStringTests.TestReadODFString10;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,10);
end;

procedure TSpreadReadStringTests.TestReadODFString11;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,11);
end;

procedure TSpreadReadStringTests.TestReadODFString12;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,12);
end;

procedure TSpreadReadStringTests.TestReadODFString13;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileODF,13);
end;

{ ODF Tests }
procedure TSpreadReadStringTests.TestReadOOXMLString0;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,0);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString1;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,1);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString2;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,2);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString3;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,3);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString4;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,4);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString5;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,5);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString6;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,6);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString7;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,7);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString8;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,8);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString9;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,9);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString10;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,10);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString11;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,11);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString12;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,12);
end;

procedure TSpreadReadStringTests.TestReadOOXMLString13;
begin
  TestReadString(ExtractFilePath(ParamStr(0)) + TestFileOOXML,13);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadReadStringTests);
  RegisterTest(TSpreadWriteReadStringTests);
  // Initialize the norm variables in case other units want to use it:
  InitSollStrings;

finalization
  FreeAndNil(TestWorkbook);

end.
