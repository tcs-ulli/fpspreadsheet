unit fonttests;

{$mode objfpc}{$H+}

interface
{ Font tests
This unit tests writing out to and reading back from files.
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of font sizes that should occur in spreadsheet
  SollSizes: array[0..12] of single; //"Soll" is a German word in Dutch accountancy jargon meaning "normative value to check against". There ;)
  SollStyles: array[0..15] of TsFontStyles;

  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollSizes;
  procedure InitSollStyles;

type
  { TSpreadWriteReadFontTests }
  // Write to xls/xml file and read back
  TSpreadWriteReadFontTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteReadBold(AFormat: TsSpreadsheetFormat);
    procedure TestWriteReadFont(AFormat: TsSpreadsheetFormat; AFontName: String);

  published
    // BIFF2 test cases
    procedure TestWriteReadBoldBIFF2;
    procedure TestWriteReadFontBIFF2_Arial;
    procedure TestWriteReadFontBIFF2_TimesNewRoman;
    procedure TestWriteReadFontBIFF2_CourierNew;

    // BIFF5 test cases
    procedure TestWriteReadBoldBIFF5;
    procedure TestWriteReadFontBIFF5_Arial;
    procedure TestWriteReadFontBIFF5_TimesNewRoman;
    procedure TestWriteReadFontBIFF5_CourierNew;

    // BIFF8 test cases
    procedure TestWriteReadBoldBIFF8;
    procedure TestWriteReadFontBIFF8_Arial;
    procedure TestWriteReadFontBIFF8_TimesNewRoman;
    procedure TestWriteReadFontBIFF8_CourierNew;
  end;

implementation

uses
  TypInfo;

const
  FontSheet = 'Font';

// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollSizes;
begin
  // Set up norm - MUST match spreadsheet cells exactly
  SollSizes[0]:=8.0;
  SollSizes[1]:=9.0;
  SollSizes[2]:=10.0;
  SollSizes[3]:=11.0;
  SollSizes[4]:=12.0;
  SollSizes[5]:=13.0;
  SollSizes[6]:=14.0;
  SollSizes[7]:=16.0;
  SollSizes[8]:=18.0;
  SollSizes[9]:=20.0;
  SollSizes[10]:=24.0;
  SollSizes[11]:=32.0;
  SollSizes[12]:=48.0;
end;

procedure InitSollStyles;
begin
  SollStyles[0] := [];
  SollStyles[1] := [fssBold];
  SolLStyles[2] := [fssItalic];
  SollStyles[3] := [fssBold, fssItalic];
  SollStyles[4] := [fssUnderline];
  SollStyles[5] := [fssUnderline, fssBold];
  SollStyles[6] := [fssUnderline, fssItalic];
  SollStyles[7] := [fssUnderline, fssBold, fssItalic];
  SollStyles[8] := [fssStrikeout];
  SollStyles[9] := [fssStrikeout, fssBold];
  SolLStyles[10] := [fssStrikeout, fssItalic];
  SollStyles[11] := [fssStrikeout, fssBold, fssItalic];
  SollStyles[12] := [fssStrikeout, fssUnderline];
  SollStyles[13] := [fssStrikeout, fssUnderline, fssBold];
  SollStyles[14] := [fssStrikeout, fssUnderline, fssItalic];
  SollStyles[15] := [fssStrikeout, fssUnderline, fssBold, fssItalic];
end;

{ TSpreadWriteReadFontTests }

procedure TSpreadWriteReadFontTests.SetUp;
begin
  inherited SetUp;
  InitSollSizes;
  InitSollStyles;
end;

procedure TSpreadWriteReadFontTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadFontTests.TestWriteReadBold(AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  currValue: String;
  expectedValue: String;
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(FontSheet);

  // Write out a cell without "bold"formatting style
  row := 0;
  col := 0;
  MyWorksheet.WriteUTF8Text(row, col, 'not bold');
  MyCell := MyWorksheet.FindCell(row, col);
  if MyCell = nil then
    fail('Error in test code. Failed to get cell.');
  CheckEquals(uffBold in MyCell^.UsedFormattingFields, false,
    'Test unsaved bold attribute, cell '+CellNotation(MyWorksheet,Row,Col));

  // Write out a cell with "bold"formatting style
  inc(row);
  MyWorksheet.WriteUTF8Text(row, col, 'bold');
  MyWorksheet.WriteUsedFormatting(row, col, [uffBold]);
  MyCell := MyWorksheet.FindCell(row, col);
  if MyCell = nil then
    fail('Error in test code. Failded to get cell.');
  CheckEquals(uffBold in MyCell^.UsedFormattingFields, true,
    'Test unsaved bold attribute, cell '+CellNotation(MyWorksheet,Row, Col));

  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet  // only 1 sheet for BIFF2
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, FontSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');

  // Try to read cell without "bold"
  row := 0;
  col := 0;
  MyCell := MyWorksheet.FindCell(row, col);
  if MyCell = nil then
    fail('Error in test code. Failed to get cell.');
  CheckEquals(uffBold in MyCell^.UsedFormattingFields, false,
    'Test saved bold attribute, cell '+CellNotation(MyWorksheet,row,col));

  // Try to read cell with "bold"
  inc(row);
  MyCell := MyWorksheet.FindCell(row, col);
  if MyCell = nil then
    fail('Error in test code. Failed to get cell.');
  CheckEquals(uffBold in MyCell^.UsedFormattingFields, true,
    'Test saved bold attribute, cell '+CellNotation(MyWorksheet,row,col));

  MyWorkbook.Free;
  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadFontTests.TestWriteReadBoldBIFF2;
begin
  TestWriteReadBold(sfExcel2);
end;

procedure TSpreadWriteReadFontTests.TestWriteReadBoldBIFF5;
begin
  TestWriteReadBold(sfExcel5);
end;

procedure TSpreadWriteReadFontTests.TestWriteReadBoldBIFF8;
begin
  TestWriteReadBold(sfExcel8);
end;

procedure TSpreadWriteReadFontTests.TestWriteReadFont(AFormat: TsSpreadsheetFormat;
  AFontName: String);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  cellText: String;
  font: TsFont;
  currValue: String;
  expectedValue: String;
  counter: Integer;
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(FontSheet);

  // Write out all font styles at various sizes
  for row := 0 to High(SollSizes) do begin
    for col := 0 to High(SollStyles) do begin
      cellText := Format('%s, %.1f-pt', [AFontName, SollSizes[row]]);
      MyWorksheet.WriteUTF8Text(row, col, celltext);
      MyWorksheet.WriteFont(row, col, AFontName, SollSizes[row], SollStyles[col], scBlack);

      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell.');
      font := MyWorkbook.GetFont(MyCell^.FontIndex);
      CheckEquals(SollSizes[row], font.Size,
        'Test unsaved font size, cell ' + CellNotation(MyWorksheet,0,0));
      currValue := GetEnumName(TypeInfo(TsFontStyles), byte(font.Style));
      expectedValue := GetEnumName(TypeInfo(TsFontStyles), byte(SollStyles[col]));
      CheckEquals(currValue, expectedValue,
        'Test unsaved font style, cell ' + CellNotation(MyWorksheet,0,0));
    end;
  end;
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet  // only 1 sheet for BIFF2
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, FontSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  counter := 0;
  for row := 0 to MyWorksheet.GetLastRowNumber do
    for col := 0 to MyWorksheet.GetLastColNumber do begin
      if (AFormat = sfExcel2) and (counter = 4) then
        break;  // Excel 2 allows only 4 fonts
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell.');
      font := MyWorkbook.GetFont(MyCell^.FontIndex);
      if abs(SollSizes[row] - font.Size) > 1e-6 then  // safe-guard against rounding errors
        CheckEquals(SollSizes[row], font.Size,
          'Test saved font size, cell '+CellNotation(MyWorksheet,Row,Col));
      currValue := GetEnumName(TypeInfo(TsFontStyles), byte(font.Style));
      expectedValue := GetEnumName(TypeInfo(TsFontStyles), byte(SollStyles[col]));
      CheckEquals(currValue, expectedValue,
        'Test unsaved font style, cell ' + CellNotation(MyWorksheet,0,0));
      inc(counter);
    end;
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

{ BIFF2 }
procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF2_Arial;
begin
  TestWriteReadFont(sfExcel2, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF2_TimesNewRoman;
begin
  TestWriteReadFont(sfExcel2, 'TimesNewRoman');
end;

procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF2_CourierNew;
begin
  TestWriteReadFont(sfExcel2, 'CourierNew');
end;

{ BIFF5 }
procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF5_Arial;
begin
  TestWriteReadFont(sfExcel5, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF5_TimesNewRoman;
begin
  TestWriteReadFont(sfExcel5, 'TimesNewRoman');
end;

procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF5_CourierNew;
begin
  TestWriteReadFont(sfExcel5, 'CourierNew');
end;

{ BIFF8 }
procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF8_Arial;
begin
  TestWriteReadFont(sfExcel8, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF8_TimesNewRoman;
begin
  TestWriteReadFont(sfExcel8, 'TimesNewRoman');
end;

procedure TSpreadWriteReadFontTests.TestWriteReadFontBIFF8_CourierNew;
begin
  TestWriteReadFont(sfExcel8, 'CourierNew');
end;


initialization
  RegisterTest(TSpreadWriteReadFontTests);

end.

