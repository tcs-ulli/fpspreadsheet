unit internaltests;

{ Other units test file read/write capability.
This unit tests functions, procedures and properties that fpspreadsheet provides.
}
{$mode objfpc}{$H+}

interface

{
Adding tests/test data:
- just add your new test procedure
}

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadReadInternalTests }
  // Read from xls/xml file with known values

  { TSpreadInternalTests }

  TSpreadInternalTests= class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    //todo: add more calls, rename sheets, try to get sheets with invalid indexes etc
    //(see strings tests for how to deal with expected exceptions)
    procedure GetSheetByIndex;
    // Verify GetSheetByName returns the correct sheet number
    // GetSheetByName was implemented in SVN revision 2857
    procedure GetSheetByName;
    // Write out date cell and try to read as UTF8; verify if contents the same
    procedure ReadDateAsUTF8;
  end;

implementation

const
  InternalSheet = 'Internal'; //worksheet name


procedure TSpreadInternalTests.GetSheetByIndex;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Cardinal;
begin
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
  MyWorkSheet:=nil;
  MyWorkSheet:=MyWorkBook.GetWorksheetByIndex(0);
  CheckFalse((MyWorksheet=nil),'GetWorksheetByIndex should return a valid index');
  MyWorkbook.Free;
end;

procedure TSpreadInternalTests.GetSheetByName;
const
  AnotherSheet='AnotherSheet';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Cardinal;
begin
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
  MyWorkSheet:=MyWorkBook.AddWorksheet(AnotherSheet);
  MyWorkSheet:=nil;
  MyWorkSheet:=MyWorkBook.GetWorksheetByName(InternalSheet);
  CheckFalse((MyWorksheet=nil),'GetWorksheetByName should return a valid index');
  CheckEquals(MyWorksheet.Name,InternalSheet,'GetWorksheetByName should return correct name.');
  MyWorkbook.Free;
end;

procedure TSpreadInternalTests.ReadDateAsUTF8;
var
  ActualDT: TDateTime;
  ActualDTString: string; //Result from ReadAsUTF8Text
  Cell: PCell;
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row,Column: Cardinal;
  TestDT: TDateTime;
begin
  Row:=0;
  Column:=0;
  TestDT:=EncodeDate(1969,7,21)+EncodeTime(2,56,0,0);
  MyWorkbook:=TsWorkbook.Create;
  MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
  MyWorkSheet.WriteDateTime(Row,Column,TestDT); //write date

  // Reading as date/time should just work
  if not(MyWorksheet.ReadAsDateTime(Row,Column,ActualDT)) then
    Fail('Could not read date time for cell '+CellNotation(MyWorkSheet,Row,Column));
  CheckEquals(TestDT,ActualDT,'Test date/time value mismatch '
    +'cell '+CellNotation(MyWorkSheet,Row,Column));

  //Check reading as string, convert to date & compare
  ActualDTString:=MyWorkSheet.ReadAsUTF8Text(Row,Column);
  ActualDT:=StrToDateTimeDef(ActualDTString,EncodeDate(1906,1,1));
  CheckEquals(TestDT,ActualDT,'Date/time mismatch using ReadAsUTF8Text');

  MyWorkbook.Free;
end;


procedure TSpreadInternalTests.SetUp;
begin

end;

procedure TSpreadInternalTests.TearDown;
begin

end;



initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadInternalTests);
end.


