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
  fpsutils, testsutility, md5;

type
  { TSpreadReadInternalTests }
  // Tests fpspreadsheet functionality, especially internal functions
  // Excel/LibreOffice/OpenOffice import/export compatibility should *NOT* be tested here

  { TSpreadInternalTests }

  TSpreadInternalTests= class(TTestCase)
  private
    procedure NeedVirtualCellData(Sender: TObject; ARow, ACol: Cardinal;
      var AValue:Variant; var AStyleCell: PCell);
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestVirtualMode(AFormat: TsSpreadsheetFormat; SaveMemoryMode: Boolean);
  published
    // Tests getting Excel style A1 cell locations from row/column based locations.
    // Bug 26447
    procedure TestCellString;
    //todo: add more calls, rename sheets, try to get sheets with invalid indexes etc
    //(see strings tests for how to deal with expected exceptions)
    procedure GetSheetByIndex;
    // Verify GetSheetByName returns the correct sheet number
    // GetSheetByName was implemented in SVN revision 2857
    procedure GetSheetByName;
    // Tests whether overwriting existing file works
    procedure OverwriteExistingFile;
    // Write out date cell and try to read as UTF8; verify if contents the same
    procedure ReadDateAsUTF8;

    // Virtual mode tests for all file formats
    procedure TestVirtualMode_BIFF2;
    procedure TestVirtualMode_BIFF5;
    procedure TestVirtualMode_BIFF8;
    procedure TestVirtualMode_OOXML;

    procedure TestVirtualMode_BIFF2_SaveMemory;
    procedure TestVirtualMode_BIFF5_SaveMemory;
    procedure TestVirtualMode_BIFF8_SaveMemory;
    procedure TestVirtualMode_OOXML_SaveMemory;
  end;

implementation

uses
  numberstests, stringtests;

const
  InternalSheet = 'Internal'; //worksheet name


procedure TSpreadInternalTests.GetSheetByIndex;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
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

procedure TSpreadInternalTests.OverwriteExistingFile;
const
  FirstFileCellText='Old version';
  SecondFileCellText='New version';
var
  FirstFileHash: string;
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string;
begin
  TempFile:=GetTempFileName;
  if fileexists(TempFile) then
    DeleteFile(TempFile);

  // Write out first file
  MyWorkbook:=TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
    MyWorkSheet.WriteUTF8Text(0,0,FirstFileCellText);
    MyWorkBook.WriteToFile(TempFile,sfExcel8,false);
  finally
    MyWorkbook.Free;
  end;

  if not(FileExists(TempFile)) then
    fail('Trying to write first file did not work.');
  FirstFileHash:=MD5Print(MD5File(TempFile));

  // Now overwrite with second file
  MyWorkbook:=TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
    MyWorkSheet.WriteUTF8Text(0,0,SecondFileCellText);
    MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  finally
    MyWorkbook.Free;
  end;
  if FirstFileHash=MD5Print(MD5File(TempFile)) then
    fail('File contents are still those of the first file.');
end;

procedure TSpreadInternalTests.ReadDateAsUTF8;
var
  ActualDT: TDateTime;
  ActualDTString: string; //Result from ReadAsUTF8Text
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

procedure TSpreadInternalTests.TestCellString;
var
  r,c: Cardinal;
  s: String;
  flags: TsRelFlags;
begin
  CheckEquals('$A$1',GetCellString(0,0,[]));
  CheckEquals('$Z$1',GetCellString(0,25,[])); //bug 26447
  CheckEquals('$AA$2',GetCellString(1,26,[])); //just past the last letter
  CheckEquals('$GW$5',GetCellString(4,204,[])); //some big value
  CheckEquals('$IV$1',GetCellString(0,255,[])); //the last column of xls
  CheckEquals('$IW$1',GetCellString(0,256,[])); //the first column beyond xls
  CheckEquals('$XFD$1',GetCellString(0,16383,[])); // the last column of xlsx
  CheckEquals('$XFE$1',GetCellString(0,16384,[])); // the first column beyond xlsx

  // Something VERY big, beyond xlsx
  s := 'ZZZZ1';
  ParseCellString(s, r, c, flags);
  CheckEquals(s, GetCellString(r, c, flags));
end;


procedure TSpreadInternalTests.SetUp;
begin
end;

procedure TSpreadInternalTests.TearDown;
begin

end;

procedure TSpreadInternalTests.NeedVirtualCellData(Sender: TObject;
  ARow, ACol: Cardinal; var AValue:Variant; var AStyleCell: PCell);
begin
  // First read the SollNumbers, then the first 4 SollStrings
  // See comment in TestVirtualMode().
  if ARow < Length(SollNumbers) then
    AValue := SollNumbers[ARow]
  else
    AValue := SollStrings[ARow - Length(SollNumbers)];
end;

procedure TSpreadInternalTests.TestVirtualMode(AFormat: TsSpreadsheetFormat;
  SaveMemoryMode: Boolean);
var
  tempFile: String;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  row, col: Integer;
  value: Double;
  s: String;
begin
  TempFile := GetTempFileName;
  if FileExists(TempFile) then
    DeleteFile(TempFile);

  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('VirtualMode');
    workbook.WritingOptions := workbook.WritingOptions + [woVirtualMode];
    if SaveMemoryMode then
      workbook.WritingOptions := workbook.WritingOptions + [woSaveMemory];
    workbook.VirtualColCount := 1;
    workbook.VirtualRowCount := Length(SollNumbers) + 4;
    // We'll use only the first 4 SollStrings, the others cause trouble due to utf8 and formatting.
    workbook.OnNeedCellData := @NeedVirtualCellData;
    workbook.WriteToFile(tempfile, AFormat, true);
  finally
    workbook.Free;
  end;

  if AFormat <> sfOOXML then begin     // No reader support for OOXML
    workbook := TsWorkbook.Create;
    try
      workbook.ReadFromFile(tempFile, AFormat);
      worksheet := workbook.GetWorksheetByIndex(0);
      col := 0;
      CheckEquals(Length(SollNumbers) + 4, worksheet.GetLastRowIndex+1,
        'Row count mismatch');
      for row := 0 to Length(SollNumbers)-1 do begin
        value := worksheet.ReadAsNumber(row, col);
        CheckEquals(SollNumbers[row], value,
          'Test number value mismatch, cell '+CellNotation(workSheet, row, col))
      end;
      for row := Length(SollNumbers) to worksheet.GetLastRowIndex do begin
        s := worksheet.ReadAsUTF8Text(row, col);
        CheckEquals(SollStrings[row - Length(SollNumbers)], s,
          'Test string value mismatch, cell '+CellNotation(workSheet, row, col));
      end;
    finally
      workbook.Free;
    end;
  end;

  DeleteFile(tempFile);
end;

procedure TSpreadInternalTests.TestVirtualMode_BIFF2;
begin
  TestVirtualMode(sfExcel2, false);
end;

procedure TSpreadInternalTests.TestVirtualMode_BIFF5;
begin
  TestVirtualMode(sfExcel5, false);
end;

procedure TSpreadInternalTests.TestVirtualMode_BIFF8;
begin
  TestVirtualMode(sfExcel8, false);
end;

procedure TSpreadInternalTests.TestVirtualMode_OOXML;
begin
  TestVirtualMode(sfOOXML, false);
end;

procedure TSpreadInternalTests.TestVirtualMode_BIFF2_SaveMemory;
begin
  TestVirtualMode(sfExcel2, True);
end;

procedure TSpreadInternalTests.TestVirtualMode_BIFF5_SaveMemory;
begin
  TestVirtualMode(sfExcel5, true);
end;

procedure TSpreadInternalTests.TestVirtualMode_BIFF8_SaveMemory;
begin
  TestVirtualMode(sfExcel8, true);
end;

procedure TSpreadInternalTests.TestVirtualMode_OOXML_SaveMemory;
begin
  TestVirtualMode(sfOOXML, true);
end;

initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadInternalTests);
end.


