unit colortests;

{$mode objfpc}{$H+}

interface
{ Color tests
This unit tests writing out to and reading back from files.
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpspreadsheet, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadColorTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadColorTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteReadBackgroundColors(WhichPalette: Integer);
    procedure TestWriteReadFontColors(WhichPalette: Integer);
  published
    // Writes out colors & reads back.
    // Background colors...
    procedure TestWriteRead_Background_Internal;  // internal palette
    procedure TestWriteRead_Background_Biff5;     // official biff5 palette
    procedure TestWriteRead_Background_Biff8;     // official biff8 palette
    // Font colors...
    procedure TestWriteRead_Font_Internal;        // internal palette
    procedure TestWriteRead_Font_Biff5;           // official biff5 palette
    procedure TestWriteRead_Font_Biff8;           // official biff8 palette
  end;

implementation

const
  ColorsSheet = 'Colors';

{ TSpreadWriteReadColorTests }

procedure TSpreadWriteReadColorTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadColorTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBackgroundColors(WhichPalette: Integer);
// WhichPalette = 5: BIFF5 palette
//                8: BIFF8 palette
//              else internal palette
// see also "manualtests".
const
  CELLTEXT = 'Color test';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  color: TsColor;
  expectedRGB: DWord;
  currentRGB: DWord;
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(ColorsSheet);

  // Define palette
  case whichPalette of
    5: MyWorkbook.UsePalette(@PALETTE_BIFF5, High(PALETTE_BIFF5)+1, true);
    8: MyWorkbook.UsePalette(@PALETTE_BIFF8, High(PALETTE_BIFF8)+1, true);
    // else use default palette
  end;

  // Write out all colors
  row := 0;
  col := 0;
  for color := 0 to MyWorkbook.GetPaletteSize-1 do begin
    MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
    MyWorksheet.WriteBackgroundColor(row, col, color);
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failed to get cell.');
    currentRGB := MyWorkbook.GetPaletteColor(MyCell^.BackgroundColor);
    expectedRGB := MyWorkbook.GetPaletteColor(color);
    CheckEquals(currentRGB, expectedRGB,
      'Test unsaved background color, cell ' + CellNotation(MyWorksheet,0,0));
    inc(row);
  end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet := GetWorksheetByName(MyWorkBook, ColorsSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for row := 0 to MyWorksheet.GetLastRowNumber do begin
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failed to get cell.');
    color := TsColor(row);
    currentRGB := MyWorkbook.GetPaletteColor(MyCell^.BackgroundColor);
    expectedRGB := MyWorkbook.GetPaletteColor(color);
    CheckEquals(currentRGB, expectedRGB,
      'Test saved background color, cell '+CellNotation(MyWorksheet,Row,Col));
  end;
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadFontColors(WhichPalette: Integer);
// WhichPalette = 5: BIFF5 palette
//                8: BIFF8 palette
//              else internal palette
// see also "manualtests".
const
  CELLTEXT = 'Color test';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  color, colorInFile: TsColor;
  expectedRGB, currentRGB: DWord;
begin
  TempFile:=GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  MyWorkbook := TsWorkbook.Create;
  MyWorkSheet:= MyWorkBook.AddWorksheet(ColorsSheet);

  // Define palette
  case whichPalette of
    5: MyWorkbook.UsePalette(@PALETTE_BIFF5, High(PALETTE_BIFF5)+1, true);
    8: MyWorkbook.UsePalette(@PALETTE_BIFF8, High(PALETTE_BIFF8)+1, true);
    // else use default palette
  end;

  // Write out all colors
  row := 0;
  col := 0;
  for color := 0 to MyWorkbook.GetPaletteSize-1 do begin
    MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
    MyWorksheet.WriteFontColor(row, col, color);
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failed to get cell.');
    colorInFile := MyWorkbook.GetFont(MyCell^.FontIndex).Color;
    currentRGB := MyWorkbook.GetPaletteColor(colorInFile);
    expectedRGB := MyWorkbook.GetPaletteColor(color);
    CheckEquals(currentRGB, expectedRGB,
      'Test unsaved font color, cell ' + CellNotation(MyWorksheet,0,0));
    inc(row);
  end;
  MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, sfExcel8);
  MyWorksheet := GetWorksheetByName(MyWorkBook, ColorsSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  for row := 0 to MyWorksheet.GetLastRowNumber do begin
    MyCell := MyWorksheet.FindCell(row, col);
    if MyCell = nil then
      fail('Error in test code. Failed to get cell.');
    color := TsColor(row);
    colorInFile := MyWorkbook.GetFont(MyCell^.FontIndex).Color;
    currentRGB := MyWorkbook.GetPaletteColor(colorInFile);
    expectedRGB := MyWorkbook.GetPaletteColor(color);
    CheckEquals(currentRGB, expectedRGB,
      'Test saved font color, cell '+CellNotation(MyWorksheet,Row,Col));
  end;
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_Background_Internal;
begin
  TestWriteReadBackgroundColors(0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_Background_Biff5;
begin
  TestWriteReadBackgroundColors(5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_Background_Biff8;
begin
  TestWriteReadBackgroundColors(8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_Font_Internal;
begin
  TestWriteReadFontColors(0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_Font_Biff5;
begin
  TestWriteReadFontColors(5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_Font_Biff8;
begin
  TestWriteReadFontColors(8);
end;

initialization
  RegisterTest(TSpreadWriteReadColorTests);

end.

