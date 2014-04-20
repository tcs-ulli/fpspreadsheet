unit formattests;

{$mode objfpc}{$H+}

interface
{ Formatted date/time/number tests
This unit tests writing out to and reading back from files.
Tests that verify reading from an Excel/LibreOffice/OpenOffice file are located in other
units (e.g. datetests).
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of strings that should occur in spreadsheet
  SollNumberStrings: array[0..6, 0..8] of string;
  SollNumbers: array[0..6] of Double;
  SollNumberFormats: array[0..8] of TsNumberFormat;
  SollNumberDecimals: array[0..8] of word;

  SollDateTimeStrings: array[0..4, 0..9] of string;
  SollDateTimes: array[0..4] of TDateTime;
  SollDateTimeFormats: array[0..9] of TsNumberFormat;
  SollDateTimeFormatStrings: array[0..9] of String;

  SollColWidths: array[0..1] of Single;

  procedure InitSollFmtData;

type
  { TSpreadWriteReadFormatTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadFormatTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Writes out numbers & reads back.
    // If previous read tests are ok, this effectively tests writing.
    procedure TestWriteReadNumberFormats;
    // Repeat with date/times
    procedure TestWriteReadDateTimeFormats;
    // Test column width
    procedure TestWriteReadColWidths;
    // Test word wrapping
    procedure TestWriteReadWordWrap;
    // Test alignments
    procedure TestWriteReadAlignments;
  end;

implementation

const
  FmtNumbersSheet = 'NumbersFormat'; //let's distinguish it from the regular numbers sheet
  FmtDateTimesSheet = 'DateTimesFormat';

// Initialize array with variables that represent the values
// we expect to be in the test spreadsheet files.
//
// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollFmtData;
var
  i: Integer;
begin
  // Set up norm - MUST match spreadsheet cells exactly

  // Numbers
  SollNumbers[0] := 0.0;
  SollNumbers[1] := 1.0;
  SollNumbers[2] := -1.0;
  SollNumbers[3] := 1.2345E6;
  SollNumbers[4] := -1.23456E6;
  SollNumbers[5] := 1.23456E-6;
  SollNumbers[6] := -1.23456E-6;

  SollNumberFormats[0] := nfGeneral;      SollNumberDecimals[0] := 0;
  SollNumberFormats[1] := nfFixed;        SollNumberDecimals[1] := 0;
  SollNumberFormats[2] := nfFixed;        SollNumberDecimals[2] := 2;
  SollNumberFormats[3] := nfFixedTh;      SollNumberDecimals[3] := 0;
  SollNumberFormats[4] := nfFixedTh;      SollNumberDecimals[4] := 2;
  SollNumberFormats[5] := nfExp;          SollNumberDecimals[5] := 2;
  SollNumberFormats[6] := nfSci;          SollNumberDecimals[6] := 1;
  SollNumberFormats[7] := nfPercentage;   SollNumberDecimals[7] := 0;
  SollNumberFormats[8] := nfPercentage;   SollNumberDecimals[8] := 2;

  for i:=Low(SollNumbers) to High(SollNumbers) do begin
    SollNumberStrings[i, 0] := FloatToStr(SollNumbers[i]);
    SollNumberStrings[i, 1] := FormatFloat('0', SollNumbers[i]);
    SollNumberStrings[i, 2] := FormatFloat('0.00', SollNumbers[i]);
    SollNumberStrings[i, 3] := FormatFloat('#,##0', SollNumbers[i]);
    SollNumberStrings[i, 4] := FormatFloat('#,##0.00', SollNumbers[i]);
    SollNumberStrings[i, 5] := FormatFloat('0.00E+00', SollNumbers[i]);
    SollNumberStrings[i, 6] := SciFloat(SollNumbers[i], 1);
    SollNumberStrings[i, 7] := FormatFloat('0', SollNumbers[i]*100) + '%';
    SollNumberStrings[i, 8] := FormatFloat('0.00', SollNumbers[i]*100) + '%';
  end;

  // Date/time values
  SollDateTimes[0] := EncodeDate(2012, 1, 12) + EncodeTime(13, 14, 15, 567);
  SolLDateTimes[1] := EncodeDate(2012, 2, 29) + EncodeTime(0, 0, 0, 1);
  SollDateTimes[2] := EncodeDate(2040, 12, 31) + EncodeTime(12, 0, 0, 0);
  SollDateTimes[3] := 1 + EncodeTime(3,45, 0, 0);
  SollDateTimes[4] := EncodeTime(12, 0, 0, 0);

  SollDateTimeFormats[0] := nfShortDateTime;   SollDateTimeFormatStrings[0] := '';
  SollDateTimeFormats[1] := nfShortDate;       SollDateTimeFormatStrings[1] := '';
  SollDateTimeFormats[2] := nfShortTime;       SollDateTimeFormatStrings[2] := '';
  SollDateTimeFormats[3] := nfLongTime;        SollDateTimeFormatStrings[3] := '';
  SollDateTimeFormats[4] := nfShortTimeAM;     SollDateTimeFormatStrings[4] := '';
  SollDateTimeFormats[5] := nfLongTimeAM;      SollDateTimeFormatStrings[5] := '';
  SollDateTimeFormats[6] := nfFmtDateTime;     SollDateTimeFormatStrings[6] := 'dm';
  SolLDateTimeFormats[7] := nfFmtDateTime;     SollDateTimeFormatStrings[7] := 'my';
  SollDateTimeFormats[8] := nfFmtDateTime;     SollDateTimeFormatStrings[8] := 'ms';
  SollDateTimeFormats[9] := nfTimeInterval;    SollDateTimeFormatStrings[9] := '';

  for i:=Low(SollDateTimes) to High(SollDateTimes) do begin
    SollDateTimeStrings[i, 0] := DateToStr(SollDateTimes[i]) + ' ' + FormatDateTime('t', SollDateTimes[i]);
    SollDateTimeStrings[i, 1] := DateToStr(SollDateTimes[i]);
    SollDateTimeStrings[i, 2] := FormatDateTime('t', SollDateTimes[i]);
    SolLDateTimeStrings[i, 3] := FormatDateTime('tt', SollDateTimes[i]);
    SollDateTimeStrings[i, 4] := FormatDateTime('t am/pm', SollDateTimes[i]);
    SollDateTimeStrings[i, 5] := FormatDateTime('tt am/pm', SollDateTimes[i]);
    SollDateTimeStrings[i, 6] := FormatDateTime('dd/mmm', SollDateTimes[i]);
    SollDateTimeStrings[i, 7] := FormatDateTime('mmm/yy', SollDateTimes[i]);
    SollDateTimeStrings[i, 8] := FormatDateTime('nn:ss', SollDateTimes[i]);
    SollDateTimeStrings[i, 9] := TimeIntervalToString(SollDateTimes[i]);
  end;

  // Column width
  SollColWidths[0] := 20;  // characters based on width of "0"
  SollColWidths[1] := 40;
end;

{ TSpreadWriteReadFormatTests }

procedure TSpreadWriteReadFormatTests.SetUp;
begin
  inherited SetUp;
  InitSollFmtData; //just for security: make sure the variables are reset to default
end;

procedure TSpreadWriteReadFormatTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadNumberFormats;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: String;
  Row, Col: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(FmtNumbersSheet);
  for Row := Low(SollNumbers) to High(SollNumbers) do
    for Col := ord(Low(SollNumberFormats)) to ord(High(SollNumberFormats)) do begin
      MyWorksheet.WriteNumber(Row, Col, SollNumbers[Row], SollNumberFormats[Col], SollNumberDecimals[Col]);
      ActualString := MyWorksheet.ReadAsUTF8Text(Row, Col);
      CheckEquals(SollNumberStrings[Row, Col], ActualString, 'Test unsaved string mismatch cell ' + CellNotation(MyWorksheet,Row,Col));
    end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook, FmtNumbersSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for Row := Low(SollNumbers) to High(SollNumbers) do
    for Col := Low(SollNumberFormats) to High(SollNumberFormats) do begin
      ActualString := MyWorkSheet.ReadAsUTF8Text(Row,Col);
      CheckEquals(SollNumberStrings[Row,Col],ActualString,'Test saved string mismatch cell '+CellNotation(MyWorkSheet,Row,Col));
    end;

  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadDateTimeFormats;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: String;
  Row,Col: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet(FmtDateTimesSheet);
  for Row := Low(SollDateTimes) to High(SollDateTimes) do
    for Col := Low(SollDateTimeFormats) to High(SollDateTimeFormats) do begin
      MyWorksheet.WriteDateTime(Row, Col, SollDateTimes[Row], SollDateTimeFormats[Col], SollDateTimeFormatStrings[Col]);
      ActualString := MyWorksheet.ReadAsUTF8Text(Row, Col);
      CheckEquals(SollDateTimeStrings[Row, Col], ActualString, 'Test unsaved string mismatch cell ' + CellNotation(MyWorksheet,Row,Col));
    end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet := GetWorksheetByName(MyWorkbook, FmtDateTimesSheet);
  if MyWorksheet = nil then
    fail('Error in test code. Failed to get named worksheet');
  for Row := Low(SollDateTimes) to High(SollDateTimes) do
    for Col := Low(SollDateTimeFormats) to High(SollDateTimeFormats) do begin
      ActualString := myWorksheet.ReadAsUTF8Text(Row,Col);
      CheckEquals(SollDateTimeStrings[Row, Col], ActualString, 'Test saved string mismatch cell '+CellNotation(MyWorksheet,Row,Col));
    end;

  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadColWidths;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualColWidth: Single;
  Col: Integer;
  lpCol: PCol;
  lCol: TCol;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(FmtNumbersSheet);
  for Col := Low(SollColWidths) to High(SollColWidths) do begin
    lCol.Width := SollColWidths[Col];
    MyWorksheet.WriteColInfo(Col, lCol);
  end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook, FmtNumbersSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for Col := Low(SollColWidths) to High(SollColWidths) do begin
    lpCol := MyWorksheet.GetCol(Col);
    if lpCol = nil then
      fail('Error in test code. Failed to return saved column width');
    ActualColWidth := lpCol^.Width;
    CheckEquals(SollColWidths[Col], ActualColWidth, 'Test saved colwidth mismatch column '+ColNotation(MyWorkSheet,Col));
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadWordWrap;
const
  LONGTEXT = 'This is a very, very, very, very long text.';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values:
  // Cell A1 is word-wrapped, Cell B1 is NOT word-wrapped
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(FmtNumbersSheet);
  MyWorksheet.WriteUTF8Text(0, 0, LONGTEXT);
  MyWorksheet.WriteUsedFormatting(0, 0, [uffWordwrap]);
  MyCell := MyWorksheet.FindCell(0, 0);
  if MyCell = nil then
    fail('Error in test code. Failed to get word-wrapped cell.');
  CheckEquals((uffWordWrap in MyCell^.UsedFormattingFields), true, 'Test unsaved word wrap mismatch cell ' + CellNotation(MyWorksheet,0,0));
  MyWorksheet.WriteUTF8Text(1, 0, LONGTEXT);
  MyWorksheet.WriteUsedFormatting(1, 0, []);
  MyCell := MyWorksheet.FindCell(1, 0);
  if MyCell = nil then
    fail('Error in test code. Failed to get word-wrapped cell.');
  CheckEquals((uffWordWrap in MyCell^.UsedFormattingFields), false, 'Test unsaved non-wrapped cell mismatch, cell ' + CellNotation(MyWorksheet,0,0));
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook, FmtNumbersSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  MyCell := MyWorksheet.FindCell(0, 0);
  if MyCell = nil then
    fail('Error in test code. Failed to get word-wrapped cell.');
  CheckEquals((uffWordWrap in MyCell^.UsedFormattingFields), true, 'failed to return correct word-wrap flag, cell ' + CellNotation(MyWorksheet,0,0));
  MyCell := MyWorksheet.FindCell(1, 0);
  if MyCell = nil then
    fail('Error in test code. Failed to get non-wrapped cell.');
  CheckEquals((uffWordWrap in MyCell^.UsedFormattingFields), false, 'failed to return correct word-wrap flag, cell ' + CellNotation(MyWorksheet,0,0));
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadAlignments;
const
  CELLTEXT = 'This is a text.';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  horAlign: TsHorAlignment;
  vertAlign: TsVertAlignment;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values: HorAlignments along columns, VertAlignments along rows
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(FmtNumbersSheet);

  row := 0;
  for horAlign in TsHorAlignment do begin
    col := 0;
    for vertAlign in TsVertAlignment do begin
      MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
      MyWorksheet.WriteHorAlignment(row, col, horAlign);
      MyWorksheet.WriteVertAlignment(row, col, vertAlign);
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell.');
      CheckEquals(vertAlign = MyCell^.VertAlignment, true,
        'Test unsaved word vertical alignment, cell ' + CellNotation(MyWorksheet,0,0));
      CheckEquals(horAlign = MyCell^.HorAlignment, true,
        'Test unsaved word horizontal alignment, cell ' + CellNotation(MyWorksheet,0,0));
      inc(col);
    end;
    inc(row);
  end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet:=GetWorksheetByName(MyWorkBook, FmtNumbersSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for row := 0 to MyWorksheet.GetLastRowNumber do
    for col := 0 to MyWorksheet.GetLastColNumber do begin
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell.');
      vertAlign := TsVertAlignment(col);
      if vertAlign = vaDefault then vertAlign := vaBottom;
      CheckEquals(vertAlign = MyCell^.VertAlignment, true,
        'Test saved vertical alignment mismatch, cell '+CellNotation(MyWorksheet,Row,Col));
      horAlign := TsHorAlignment(row);
      CheckEquals(horAlign = MyCell^.HorAlignment, true,
        'Test saved horizontal mismatch, cell '+CellNotation(MyWorksheet,Row,Col));
    end;
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;


initialization
  RegisterTest(TSpreadWriteReadFormatTests);
  InitSollFmtData;

end.
