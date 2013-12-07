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


