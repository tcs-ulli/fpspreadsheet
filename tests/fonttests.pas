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
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
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
    procedure TestWriteRead_BIFF2_Bold;
    procedure TestWriteRead_BIFF2_Font_Arial;
    procedure TestWriteRead_BIFF2_Font_TimesNewRoman;
    procedure TestWriteRead_BIFF2_Font_CourierNew;

    // BIFF5 test cases
    procedure TestWriteRead_BIFF5_Bold;
    procedure TestWriteRead_BIFF5_Font_Arial;
    procedure TestWriteRead_BIFF5_Font_TimesNewRoman;
    procedure TestWriteRead_BIFF5_Font_CourierNew;

    // BIFF8 test cases
    procedure TestWriteRead_BIFF8_Bold;
    procedure TestWriteRead_BIFF8_Font_Arial;
    procedure TestWriteRead_BIFF8_Font_TimesNewRoman;
    procedure TestWriteRead_BIFF8_Font_CourierNew;

    // ODS test cases
    procedure TestWriteRead_ODS_Bold;
    procedure TestWriteRead_ODS_Font_Arial;
    procedure TestWriteRead_ODS_Font_TimesNewRoman;
    procedure TestWriteRead_ODS_Font_CourierNew;

    // OOXML test cases
    procedure TestWriteRead_OOXML_Bold;
    procedure TestWriteRead_OOXML_Font_Arial;
    procedure TestWriteRead_OOXML_Font_TimesNewRoman;
    procedure TestWriteRead_OOXML_Font_CourierNew;
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
begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(FontSheet);

    // Write out a cell without "bold" formatting style
    row := 0;
    col := 0;
    MyWorksheet.WriteUTF8Text(row, col, 'not bold');
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failed to get cell.');
    CheckEquals(uffBold in MyCell^.UsedFormattingFields, false,
      'Test unsaved bold attribute, cell '+CellNotation(MyWorksheet,Row,Col));

    // Write out a cell with "bold" formatting style
    inc(row);
    MyWorksheet.WriteUTF8Text(row, col, 'bold');
    MyWorksheet.WriteUsedFormatting(row, col, [uffBold]);
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failded to get cell.');
    CheckEquals(uffBold in MyCell^.UsedFormattingFields, true,
      'Test unsaved bold attribute, cell '+CellNotation(MyWorksheet,Row, Col));

    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
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
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
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

  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(FontSheet);

    // Write out all font styles at various sizes
    for row := 0 to High(SollSizes) do
      begin
      for col := 0 to High(SollStyles) do
      begin
        cellText := Format('%s, %.1f-pt', [AFontName, SollSizes[row]]);
        MyWorksheet.WriteUTF8Text(row, col, celltext);
        MyWorksheet.WriteFont(row, col, AFontName, SollSizes[row], SollStyles[col], scBlack);

        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        font := MyWorkbook.GetFont(MyCell^.FontIndex);
        CheckEquals(SollSizes[row], font.Size,
          'Test unsaved font size, cell ' + CellNotation(MyWorksheet,0,0));
        currValue := GetEnumName(TypeInfo(TsFontStyles), integer(font.Style));
        expectedValue := GetEnumName(TypeInfo(TsFontStyles), integer(SollStyles[col]));
        CheckEquals(currValue, expectedValue,
          'Test unsaved font style, cell ' + CellNotation(MyWorksheet,0,0));
      end;
    end;
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet  // only 1 sheet for BIFF2
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, FontSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    counter := 0;
    for row := 0 to MyWorksheet.GetLastRowIndex do
      for col := 0 to MyWorksheet.GetLastColIndex do
      begin
        if (AFormat = sfExcel2) and (counter = 4) then
          break;  // Excel 2 allows only 4 fonts
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        font := MyWorkbook.GetFont(MyCell^.FontIndex);
        if abs(SollSizes[row] - font.Size) > 1e-6 then  // safe-guard against rounding errors
          CheckEquals(SollSizes[row], font.Size,
            'Test saved font size, cell '+CellNotation(MyWorksheet,Row,Col));
        currValue := GetEnumName(TypeInfo(TsFontStyles), integer(font.Style));
        expectedValue := GetEnumName(TypeInfo(TsFontStyles), integer(SollStyles[col]));
        CheckEquals(currValue, expectedValue,
          'Test unsaved font style, cell ' + CellNotation(MyWorksheet,0,0));
        inc(counter);
      end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

{ BIFF2 }

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF2_Bold;
begin
  TestWriteReadBold(sfExcel2);
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF2_Font_Arial;
begin
  TestWriteReadFont(sfExcel2, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF2_Font_TimesNewRoman;
begin
  TestWriteReadFont(sfExcel2, 'Times New Roman');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF2_Font_CourierNew;
begin
  TestWriteReadFont(sfExcel2, 'Courier New');
end;

{ BIFF5 }
procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF5_Bold;
begin
  TestWriteReadBold(sfExcel5);
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF5_Font_Arial;
begin
  TestWriteReadFont(sfExcel5, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF5_Font_TimesNewRoman;
begin
  TestWriteReadFont(sfExcel5, 'Times New Roman');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF5_Font_CourierNew;
begin
  TestWriteReadFont(sfExcel5, 'Courier New');
end;

{ BIFF8 }
procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF8_Bold;
begin
  TestWriteReadBold(sfExcel8);
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF8_Font_Arial;
begin
  TestWriteReadFont(sfExcel8, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF8_Font_TimesNewRoman;
begin
  TestWriteReadFont(sfExcel8, 'Times New Roman');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_BIFF8_Font_CourierNew;
begin
  TestWriteReadFont(sfExcel8, 'Courier New');
end;

{ ODS }
procedure TSpreadWriteReadFontTests.TestWriteRead_ODS_Bold;
begin
  TestWriteReadBold(sfOpenDocument);
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_ODS_Font_Arial;
begin
  TestWriteReadFont(sfOpenDocument, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_ODS_Font_TimesNewRoman;
begin
  TestWriteReadFont(sfOpenDocument, 'Times New Roman');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_ODS_Font_CourierNew;
begin
  TestWriteReadFont(sfOpenDocument, 'Courier New');
end;

{ OOXML }
procedure TSpreadWriteReadFontTests.TestWriteRead_OOXML_Bold;
begin
  TestWriteReadBold(sfOOXML);
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_OOXML_Font_Arial;
begin
  TestWriteReadFont(sfOOXML, 'Arial');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_OOXML_Font_TimesNewRoman;
begin
  TestWriteReadFont(sfOOXML, 'Times New Roman');
end;

procedure TSpreadWriteReadFontTests.TestWriteRead_OOXML_Font_CourierNew;
begin
  TestWriteReadFont(sfOOXML, 'Courier New');
end;

initialization
  RegisterTest(TSpreadWriteReadFontTests);

end.

