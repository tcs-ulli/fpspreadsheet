unit datetests;

{$mode objfpc}{$H+}

{
Adding tests/test data:
1. Add a new value to column A in the relevant worksheet, and save the spreadsheet read-only
   (for dates, there are 2 files, with different datemodes. Use them both...)
   Repeat this for all supported spreadsheet formats (Excel XLS, ODF, etc)
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
  SollDates: array[0..37] of TDateTime; //"Soll" is a German word in Dutch accountancy jargon meaning "normative value to check against". There ;)
  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollDates;

type
  { TSpreadReadDateTests }
  // Read from xls/xml file with known values to test interoperability with Excel/LibreOffice/OpenOffice
  TSpreadReadDateTests= class(TTestCase)
  private
    // Tries to read date from the external file in column A, specified (0-based) row
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
    procedure TestReadDate12;
    procedure TestReadDate13;
    procedure TestReadDate14;
    procedure TestReadDate15;
    procedure TestReadDate16;
    procedure TestReadDate17;
    procedure TestReadDate18;
    procedure TestReadDate19;
    procedure TestReadDate20;
    procedure TestReadDate21;
    procedure TestReadDate22;
    procedure TestReadDate23;
    procedure TestReadDate24;
    procedure TestReadDate25;
    procedure TestReadDate26;
    procedure TestReadDate27;
    procedure TestReadDate28;
    procedure TestReadDate29;
    procedure TestReadDate30;
    procedure TestReadDate31;
    procedure TestReadDate32;
    procedure TestReadDate33;
    procedure TestReadDate34;
    procedure TestReadDate35;
    procedure TestReadDate36;
    procedure TestReadDate37;
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
    procedure TestReadDate1899_12;
    procedure TestReadDate1899_13;
    procedure TestReadDate1899_14;
    procedure TestReadDate1899_15;
    procedure TestReadDate1899_16;
    procedure TestReadDate1899_17;
    procedure TestReadDate1899_18;
    procedure TestReadDate1899_19;
    procedure TestReadDate1899_20;
    procedure TestReadDate1899_21;
    procedure TestReadDate1899_22;
    procedure TestReadDate1899_23;
    procedure TestReadDate1899_24;
    procedure TestReadDate1899_25;
    procedure TestReadDate1899_26;
    procedure TestReadDate1899_27;
    procedure TestReadDate1899_28;
    procedure TestReadDate1899_29;
    procedure TestReadDate1899_30;
    procedure TestReadDate1899_31;
    procedure TestReadDate1899_32;
    procedure TestReadDate1899_33;
    procedure TestReadDate1899_34;
    procedure TestReadDate1899_35;
    procedure TestReadDate1899_36;
    procedure TestReadDate1899_37;
    procedure TestReadODFDate0; // same as above except OpenDocument/ODF format
    procedure TestReadODFDate1; //date and time
    procedure TestReadODFDate2;
    procedure TestReadODFDate3;
    procedure TestReadODFDate4; //time only tests start here
    procedure TestReadODFDate5;
    procedure TestReadODFDate6;
    procedure TestReadODFDate7;
    procedure TestReadODFDate8;
    procedure TestReadODFDate9;
    procedure TestReadODFDate10;
    procedure TestReadODFDate11;
    procedure TestReadODFDate12;
    procedure TestReadODFDate13;
    procedure TestReadODFDate14;
    procedure TestReadODFDate15;
    procedure TestReadODFDate16;
    procedure TestReadODFDate17;
    procedure TestReadODFDate18;
    procedure TestReadODFDate19;
    procedure TestReadODFDate20;
    procedure TestReadODFDate21;
    procedure TestReadODFDate22;
    procedure TestReadODFDate23;
    procedure TestReadODFDate24;
    procedure TestReadODFDate25;
    procedure TestReadODFDate26;
    procedure TestReadODFDate27;
    procedure TestReadODFDate28;
    procedure TestReadODFDate29;
    procedure TestReadODFDate30;
    procedure TestReadODFDate31;
    procedure TestReadODFDate32;
    procedure TestReadODFDate33;
    procedure TestReadODFDate34;
    procedure TestReadODFDate35;
    procedure TestReadODFDate36;
    procedure TestReadODFDate37;
    procedure TestReadODFDate1899_0; //same as above except with the 1899/1900 date system set
    procedure TestReadODFDate1899_1;
    procedure TestReadODFDate1899_2;
    procedure TestReadODFDate1899_3;
    procedure TestReadODFDate1899_4;
    procedure TestReadODFDate1899_5;
    procedure TestReadODFDate1899_6;
    procedure TestReadODFDate1899_7;
    procedure TestReadODFDate1899_8;
    procedure TestReadODFDate1899_9;
    procedure TestReadODFDate1899_10;
    procedure TestReadODFDate1899_11;
    procedure TestReadODFDate1899_12;
    procedure TestReadODFDate1899_13;
    procedure TestReadODFDate1899_14;
    procedure TestReadODFDate1899_15;
    procedure TestReadODFDate1899_16;
    procedure TestReadODFDate1899_17;
    procedure TestReadODFDate1899_18;
    procedure TestReadODFDate1899_19;
    procedure TestReadODFDate1899_20;
    procedure TestReadODFDate1899_21;
    procedure TestReadODFDate1899_22;
    procedure TestReadODFDate1899_23;
    procedure TestReadODFDate1899_24;
    procedure TestReadODFDate1899_25;
    procedure TestReadODFDate1899_26;
    procedure TestReadODFDate1899_27;
    procedure TestReadODFDate1899_28;
    procedure TestReadODFDate1899_29;
    procedure TestReadODFDate1899_30;
    procedure TestReadODFDate1899_31;
    procedure TestReadODFDate1899_32;
    procedure TestReadODFDate1899_33;
    procedure TestReadODFDate1899_34;
    procedure TestReadODFDate1899_35;
    procedure TestReadODFDate1899_36;
    procedure TestReadODFDate1899_37;
  end;

  { TSpreadWriteReadDateTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadDateTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    // Reads dates, date/time and time values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestWriteReadDates(AFormat: TsSpreadsheetFormat);
  published
    procedure TestWriteReadDates_BIFF5;
    procedure TestWriteReadDates_BIFF8;
  end;


implementation

const
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

  SollDates[12]:=SollDates[1];  // #1 formatted as nfShortDateTime
  SollDates[13]:=SollDates[1];  // #1 formatted as nfShortTime
  SollDates[14]:=SollDates[1];  // #1 formatted as nfLongTime
  SollDates[15]:=SollDates[1];  // #1 formatted as nfShortTimeAM
  SollDates[16]:=SollDates[1];  // #1 formatted as nfLongTimeAM
  SollDates[17]:=SollDates[1];  // #1 formatted as nfFmtDateTime dm
  SollDates[18]:=SollDates[1];  // #1 formatted as nfFmtDateTime my
  SollDates[19]:=SollDates[1];  // #1 formatted as nfFmtDateTime ms

  SollDates[20]:=SollDates[5];  // #5 formatted as nfShortDateTime
  SollDates[21]:=SollDates[5];  // #5 formatted as nfShortTime
  SollDates[22]:=SollDates[5];  // #5 formatted as nfLongTime
  SollDates[23]:=SollDates[5];  // #5 formatted as nfShortTimeAM
  SollDates[24]:=SollDates[5];  // #5 formatted as nfLongTimeAM
  SollDates[25]:=SollDates[5];  // #5 formatted as nfFmtDateTime dm
  SollDates[26]:=SollDates[5];  // #5 formatted as nfFmtDateTime my
  SollDates[27]:=SollDates[5];  // #5 formatted as nfFmtDateTime ms

  SollDates[28]:=SollDates[11];  // #11 formatted as nfShortDateTime
  SollDates[29]:=SollDates[11];  // #11 formatted as nfShortTime
  SollDates[30]:=SollDates[11];  // #11 formatted as nfLongTime
  SollDates[31]:=SollDates[11];  // #11 formatted as nfShortTimeAM
  SollDates[32]:=SollDates[11];  // #11 formatted as nfLongTimeAM
  SollDates[33]:=SollDates[11];  // #11 formatted as nfFmtDateTime dm
  SollDates[34]:=SollDates[11];  // #11 formatted as nfFmtDateTime my
  SollDates[35]:=SollDates[11];  // #11 formatted as nfFmtDateTime ms

  SollDates[36]:=EncodeTime(3,45,12,0);     // formatted as nfTimeDuration
  SollDates[37]:=EncodeTime(3,45,12,0) + 1  // formatted as nfTimeDuration
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

procedure TSpreadWriteReadDateTests.TestWriteReadDates(AFormat: TsSpreadsheetFormat);
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
      Fail('Failed writing date time for cell '+CellNotation(MyWorkSheet,Row));
    CheckEquals(SollDates[Row],ActualDateTime,'Test date/time value mismatch cell '+CellNotation(MyWorksheet,Row));
  end;
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook,DatesSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  // Read test data from A column & compare if written=original
  for Row := Low(SollDates) to High(SollDates) do
  begin
    if not(MyWorkSheet.ReadAsDateTime(Row,0,ActualDateTime)) then
      Fail('Could not read date time for cell '+CellNotation(MyWorkSheet,Row));
    CheckEquals(SollDates[Row],ActualDateTime,'Test date/time value mismatch cell '+CellNotation(MyWorkSheet,Row));
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_BIFF5;
begin
  TestWriteReadDates(sfExcel5);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_BIFF8;
begin
  TestWriteReadDates(sfExcel8);
end;

procedure TSpreadReadDateTests.TestReadDate(FileName: string; Row: integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualDateTime: TDateTime;
begin
  if Row>High(SollDates) then
    fail('Error in test code: array bounds overflow. Check array size is correct.');

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  case UpperCase(ExtractFileExt(FileName)) of
    '.XLSX': MyWorkbook.ReadFromFile(FileName, sfOOXML);
    '.ODS': MyWorkbook.ReadFromFile(FileName, sfOpenDocument);
    // Excel XLS/BIFF
    else MyWorkbook.ReadFromFile(FileName, sfExcel8);
  end;
  MyWorksheet:=GetWorksheetByName(MyWorkBook,DatesSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  // We know these are valid time/date/datetime values....
  // Just test for empty string; we'll probably end up in a maze of localized date/time stuff
  // if we don't.
  CheckNotEquals(MyWorkSheet.ReadAsUTF8Text(Row, 0), '','Could not read date time as string for cell '+CellNotation(MyWorkSheet,Row));

  if not(MyWorkSheet.ReadAsDateTime(Row, 0, ActualDateTime)) then
    Fail('Could not read date time for cell '+CellNotation(MyWorkSheet,Row));
  CheckEquals(SollDates[Row],ActualDateTime,'Test date/time value mismatch '
    +'cell '+CellNotation(MyWorksheet,Row));

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

procedure TSpreadReadDateTests.TestReadDate12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,12);
end;

procedure TSpreadReadDateTests.TestReadDate13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,13);
end;

procedure TSpreadReadDateTests.TestReadDate14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,14);
end;

procedure TSpreadReadDateTests.TestReadDate15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,15);
end;

procedure TSpreadReadDateTests.TestReadDate16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,16);
end;

procedure TSpreadReadDateTests.TestReadDate17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,17);
end;

procedure TSpreadReadDateTests.TestReadDate18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,18);
end;

procedure TSpreadReadDateTests.TestReadDate19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,19);
end;

procedure TSpreadReadDateTests.TestReadDate20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,20);
end;

procedure TSpreadReadDateTests.TestReadDate21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,21);
end;

procedure TSpreadReadDateTests.TestReadDate22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,22);
end;

procedure TSpreadReadDateTests.TestReadDate23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,23);
end;

procedure TSpreadReadDateTests.TestReadDate24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,24);
end;

procedure TSpreadReadDateTests.TestReadDate25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,25);
end;

procedure TSpreadReadDateTests.TestReadDate26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,26);
end;

procedure TSpreadReadDateTests.TestReadDate27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,27);
end;

procedure TSpreadReadDateTests.TestReadDate28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,28);
end;

procedure TSpreadReadDateTests.TestReadDate29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,29);
end;

procedure TSpreadReadDateTests.TestReadDate30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,30);
end;

procedure TSpreadReadDateTests.TestReadDate31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,31);
end;

procedure TSpreadReadDateTests.TestReadDate32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,32);
end;

procedure TSpreadReadDateTests.TestReadDate33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,33);
end;

procedure TSpreadReadDateTests.TestReadDate34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,34);
end;

procedure TSpreadReadDateTests.TestReadDate35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,35);
end;

procedure TSpreadReadDateTests.TestReadDate36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,36);
end;

procedure TSpreadReadDateTests.TestReadDate37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8,37);
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

procedure TSpreadReadDateTests.TestReadDate1899_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,12);
end;

procedure TSpreadReadDateTests.TestReadDate1899_13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,13);
end;

procedure TSpreadReadDateTests.TestReadDate1899_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,14);
end;

procedure TSpreadReadDateTests.TestReadDate1899_15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,15);
end;

procedure TSpreadReadDateTests.TestReadDate1899_16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,16);
end;

procedure TSpreadReadDateTests.TestReadDate1899_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,17);
end;

procedure TSpreadReadDateTests.TestReadDate1899_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,18);
end;

procedure TSpreadReadDateTests.TestReadDate1899_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,19);
end;

procedure TSpreadReadDateTests.TestReadDate1899_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,20);
end;

procedure TSpreadReadDateTests.TestReadDate1899_21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,21);
end;

procedure TSpreadReadDateTests.TestReadDate1899_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,22);
end;

procedure TSpreadReadDateTests.TestReadDate1899_23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,23);
end;

procedure TSpreadReadDateTests.TestReadDate1899_24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,24);
end;

procedure TSpreadReadDateTests.TestReadDate1899_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,25);
end;

procedure TSpreadReadDateTests.TestReadDate1899_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,26);
end;

procedure TSpreadReadDateTests.TestReadDate1899_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,27);
end;

procedure TSpreadReadDateTests.TestReadDate1899_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,28);
end;

procedure TSpreadReadDateTests.TestReadDate1899_29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,29);
end;

procedure TSpreadReadDateTests.TestReadDate1899_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,30);
end;

procedure TSpreadReadDateTests.TestReadDate1899_31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,31);
end;

procedure TSpreadReadDateTests.TestReadDate1899_32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,32);
end;

procedure TSpreadReadDateTests.TestReadDate1899_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,33);
end;

procedure TSpreadReadDateTests.TestReadDate1899_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,34);
end;

procedure TSpreadReadDateTests.TestReadDate1899_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,35);
end;

procedure TSpreadReadDateTests.TestReadDate1899_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,36);
end;

procedure TSpreadReadDateTests.TestReadDate1899_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,37);
end;

procedure TSpreadReadDateTests.TestReadODFDate0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,0);
end;

procedure TSpreadReadDateTests.TestReadODFDate1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,1);
end;

procedure TSpreadReadDateTests.TestReadODFDate2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,2);
end;

procedure TSpreadReadDateTests.TestReadODFDate3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,3);
end;

procedure TSpreadReadDateTests.TestReadODFDate4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,4);
end;

procedure TSpreadReadDateTests.TestReadODFDate5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,5);
end;

procedure TSpreadReadDateTests.TestReadODFDate6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,6);
end;

procedure TSpreadReadDateTests.TestReadODFDate7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,7);
end;

procedure TSpreadReadDateTests.TestReadODFDate8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,8);
end;

procedure TSpreadReadDateTests.TestReadODFDate9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,9);
end;

procedure TSpreadReadDateTests.TestReadODFDate10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,10);
end;

procedure TSpreadReadDateTests.TestReadODFDate11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,11);
end;

procedure TSpreadReadDateTests.TestReadODFDate12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,12);
end;

procedure TSpreadReadDateTests.TestReadODFDate13;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,13);
end;

procedure TSpreadReadDateTests.TestReadODFDate14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,14);
end;

procedure TSpreadReadDateTests.TestReadODFDate15;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,15);
end;

procedure TSpreadReadDateTests.TestReadODFDate16;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,16);
end;

procedure TSpreadReadDateTests.TestReadODFDate17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,17);
end;

procedure TSpreadReadDateTests.TestReadODFDate18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,18);
end;

procedure TSpreadReadDateTests.TestReadODFDate19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,19);
end;

procedure TSpreadReadDateTests.TestReadODFDate20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,20);
end;

procedure TSpreadReadDateTests.TestReadODFDate21;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,21);
end;

procedure TSpreadReadDateTests.TestReadODFDate22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,22);
end;

procedure TSpreadReadDateTests.TestReadODFDate23;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,23);
end;

procedure TSpreadReadDateTests.TestReadODFDate24;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,24);
end;

procedure TSpreadReadDateTests.TestReadODFDate25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,25);
end;

procedure TSpreadReadDateTests.TestReadODFDate26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,26);
end;

procedure TSpreadReadDateTests.TestReadODFDate27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,27);
end;

procedure TSpreadReadDateTests.TestReadODFDate28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,28);
end;

procedure TSpreadReadDateTests.TestReadODFDate29;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,29);
end;

procedure TSpreadReadDateTests.TestReadODFDate30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,30);
end;

procedure TSpreadReadDateTests.TestReadODFDate31;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,31);
end;

procedure TSpreadReadDateTests.TestReadODFDate32;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,32);
end;

procedure TSpreadReadDateTests.TestReadODFDate33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,33);
end;

procedure TSpreadReadDateTests.TestReadODFDate34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,34);
end;

procedure TSpreadReadDateTests.TestReadODFDate35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,35);
end;

procedure TSpreadReadDateTests.TestReadODFDate36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,36);
end;

procedure TSpreadReadDateTests.TestReadODFDate37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF,37);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,0);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,1);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,2);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,3);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,4);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,5);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,6);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,7);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,8);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,9);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,10);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,11);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,12);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_13;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,13);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,14);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_15;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,15);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_16;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,16);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,17);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,18);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,19);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,20);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_21;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,21);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,22);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_23;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,23);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_24;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,24);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,25);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,26);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,27);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,28);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_29;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,29);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,30);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_31;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,31);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_32;
begin
  Ignore('ODF code does not support custom date format');
  //TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,32);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,33);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,34);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,35);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,36);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,37);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadReadDateTests);
  RegisterTest(TSpreadWriteReadDateTests);
  InitSollDates; //useful to have norm data if other code want to use this unit
end.


