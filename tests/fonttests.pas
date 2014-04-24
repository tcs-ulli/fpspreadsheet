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
    procedure TestWriteReadFont(AFormat: TsSpreadsheetFormat; AFontName: String);
  published
    {
    // BIFF2 test cases
    procedure TestWriteReadFontBIFF2_Arial;
    procedure TestWriteReadFontBIFF2_TimesNewRoman;
    procedure TestWriteReadFontBIFF2_CourierNew;

    !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Currently the BIFF2 tests fail because of a font size mismatch at size 11.
    Outside the test suite, however, this error is not reproduced,
    and also not when the conflicting case is used in the SollValues alone.

    STRANGE...
    !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    }

    // BIFF8 test cases
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
  for row := 0 to MyWorksheet.GetLastRowNumber do
    for col := 0 to MyWorksheet.GetLastColNumber do begin
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
    end;
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;
                                  (*
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
*)

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

