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
  SollBorders: array[0..15] of TsCellBorders;

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

    // Test alignments
    procedure TestWriteReadAlignment(AFormat: TsSpreadsheetFormat);
    // Test border
    procedure TestWriteReadBorder(AFormat: TsSpreadsheetFormat);
    // Test column widths
    procedure TestWriteReadColWidths(AFormat: TsSpreadsheetFormat);
    // Test text rotation
    procedure TestWriteReadTextRotation(AFormat:TsSpreadsheetFormat);
    // Test word wrapping
    procedure TestWriteReadWordWrap(AFormat: TsSpreadsheetFormat);

  published
    // Writes out numbers & reads back.
    // If previous read tests are ok, this effectively tests writing.

    { BIFF2 Tests }
    procedure TestWriteReadBIFF2_Alignment;
    procedure TestWriteReadBIFF2_Border;
    procedure TestWriteReadBIFF2_ColWidths;
    // These features are not supported by Excel2 --> no test cases required!
    // - TextRotation
    // - Wordwrap

    { BIFF5 Tests }
    procedure TestWriteReadBIFF5_Alignment;
    procedure TestWriteReadBIFF5_Border;
    procedure TestWriteReadBIFF5_ColWidths;
    procedure TestWriteReadBIFF5_TextRotation;
    procedure TestWriteReadBIFF5_WordWrap;

    { BIFF8 Tests }
    procedure TestWriteReadBIFF8_Alignment;
    procedure TestWriteReadBIFF8_Border;
    procedure TestWriteReadBIFF8_ColWidths;
    procedure TestWriteReadBIFF8_TextRotation;
    procedure TestWriteReadBIFF8_WordWrap;
    procedure TestWriteReadNumberFormats;
    // Repeat with date/times
    procedure TestWriteReadDateTimeFormats;
  end;

implementation

uses
  TypInfo;

const
  FmtNumbersSheet = 'NumbersFormat'; //let's distinguish it from the regular numbers sheet
  FmtDateTimesSheet = 'DateTimesFormat';
  ColWidthSheet = 'ColWidths';
  BordersSheet = 'CellBorders';
  AlignmentSheet = 'TextAlignments';
  TextRotationSheet = 'TextRotation';
  WordwrapSheet = 'Wordwrap';

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

  // Cell borders
  SollBorders[0] := [];
  SollBorders[1] := [cbEast];
  SollBorders[2] := [cbSouth];
  SollBorders[3] := [cbWest];
  SollBorders[4] := [cbNorth];
  SollBorders[5] := [cbEast, cbSouth];
  SollBorders[6] := [cbEast, cbWest];
  SollBorders[7] := [cbEast, cbNorth];
  SollBorders[8] := [cbSouth, cbWest];
  SollBorders[9] := [cbSouth, cbNorth];
  SollBorders[10] := [cbWest, cbNorth];
  SollBorders[11] := [cbEast, cbSouth, cbWest];
  SollBorders[12] := [cbEast, cbSouth, cbNorth];
  SollBorders[13] := [cbSouth, cbWest, cbNorth];
  SollBorders[14] := [cbWest, cbNorth, cbEast];
  SollBorders[15] := [cbEast, cbSouth, cbWest, cbNorth];
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

procedure TSpreadWriteReadFormatTests.TestWriteReadAlignment(AFormat: TsSpreadsheetFormat);
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
  MyWorkSheet:= MyWorkBook.AddWorksheet(AlignmentSheet);

  row := 0;
  for horAlign in TsHorAlignment do begin
    col := 0;
    if AFormat = sfExcel2 then begin
      MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
      MyWorksheet.WriteHorAlignment(row, col, horAlign);
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell.');
      CheckEquals(horAlign = MyCell^.HorAlignment, true,
        'Test unsaved horizontal alignment, cell ' + CellNotation(MyWorksheet,0,0));
    end else
      for vertAlign in TsVertAlignment do begin
        MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
        MyWorksheet.WriteHorAlignment(row, col, horAlign);
        MyWorksheet.WriteVertAlignment(row, col, vertAlign);
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        CheckEquals(vertAlign = MyCell^.VertAlignment, true,
          'Test unsaved vertical alignment, cell ' + CellNotation(MyWorksheet,0,0));
        CheckEquals(horAlign = MyCell^.HorAlignment, true,
          'Test unsaved horizontal alignment, cell ' + CellNotation(MyWorksheet,0,0));
        inc(col);
      end;
    inc(row);
  end;
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, AlignmentSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for row := 0 to MyWorksheet.GetLastRowNumber do
    if AFormat = sfExcel2 then begin
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failded to get cell.');
      horAlign := TsHorAlignment(row);
      CheckEquals(horAlign = MyCell^.HorAlignment, true,
        'Test saved horizontal alignment mismatch, cell '+CellNotation(MyWorksheet,row,col));
    end else
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
          'Test saved horizontal alignment mismatch, cell '+CellNotation(MyWorksheet,Row,Col));
      end;
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF2_Alignment;
begin
  TestWriteReadAlignment(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF5_Alignment;
begin
  TestWriteReadAlignment(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF8_Alignment;
begin
  TestWriteReadAlignment(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBorder(AFormat: TsSpreadsheetFormat);
const
  row = 0;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  ActualColWidth: Single;
  col: Integer;
  expected: String;
  current: String;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(BordersSheet);
  for col := Low(SollBorders) to High(SollBorders) do begin
    MyWorksheet.WriteUsedFormatting(row, col, [uffBorder]);
    MyCell := MyWorksheet.GetCell(row, col);
    Include(MyCell^.UsedFormattingFields, uffBorder);
    MyCell^.Border := SollBorders[col];
  end;
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, BordersSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for col := 0 to MyWorksheet.GetLastColNumber do begin
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failed to get cell');
    current := GetEnumName(TypeInfo(TsCellBorders), byte(MyCell^.Border));
    expected := GetEnumName(TypeInfo(TsCellBorders), byte(SollBorders[col]));
    CheckEquals(current, expected,
      'Test saved border mismatch, cell ' + CellNotation(MyWorksheet, row, col));
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF2_Border;
begin
  TestWriteReadBorder(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF5_Border;
begin
  TestWriteReadBorder(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF8_Border;
begin
  TestWriteReadBorder(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadColWidths(AFormat: TsSpreadsheetFormat);
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
  MyWorkSheet:= MyWorkBook.AddWorksheet(ColWidthSheet);
  for Col := Low(SollColWidths) to High(SollColWidths) do begin
    lCol.Width := SollColWidths[Col];
    MyWorksheet.WriteColInfo(Col, lCol);
  end;
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, ColWidthSheet);
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

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF2_ColWidths;
begin
  TestWriteReadColWidths(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF5_ColWidths;
begin
  TestWriteReadColWidths(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF8_ColWidths;
begin
  TestWriteReadColWidths(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadTextRotation(AFormat: TsSpreadsheetFormat);
const
  col = 0;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  ActualColWidth: Single;
  tr: TsTextRotation;
  row: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(TextRotationSheet);
  for tr := Low(TsTextRotation) to High(TsTextRotation) do begin
    row := ord(tr);
    MyWorksheet.WriteTextRotation(row, col, tr);
    MyCell := MyWorksheet.GetCell(row, col);
    CheckEquals(ord(tr), ord(MyCell^.TextRotation),
      'Test unsaved textrotation mismatch, cell ' + CellNotation(MyWorksheet, row, col));
  end;
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, TextRotationSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for row := 0 to MyWorksheet.GetLastRowNumber do begin
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failed to get cell');
    tr := MyCell^.TextRotation;
    CheckEquals(ord(TsTextRotation(row)), ord(MyCell^.TextRotation),
      'Test saved textrotation mismatch, cell ' + CellNotation(MyWorksheet, row, col));
  end;
  // Finalization
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF5_TextRotation;
begin
  TestWriteReadTextRotation(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF8_TextRotation;
begin
  TestWriteReadTextRotation(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadWordWrap(AFormat: TsSpreadsheetFormat);
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
  MyWorkSheet:= MyWorkBook.AddWorksheet(WordwrapSheet);
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
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, WordwrapSheet);
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

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF5_Wordwrap;
begin
  TestWriteReadWordwrap(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteReadBIFF8_Wordwrap;
begin
  TestWriteReadWordwrap(sfExcel8);
end;

initialization
  RegisterTest(TSpreadWriteReadFormatTests);
  InitSollFmtData;

end.

