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
  fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
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
    procedure TestWriteReadBackgroundColors(AFormat: TsSpreadsheetFormat; WhichPalette: Integer);
    procedure TestWriteReadFontColors(AFormat: TsSpreadsheetFormat; WhichPalette: Integer);
  published
    // Writes out colors & reads back.

    { BIFF2 file format tests }
    procedure TestWriteReadBIFF2_Font_InternalPal;        // internal palette for BIFF2 file format

    { BIFF8 file format tests }
    // Background colors...
    procedure TestWriteReadBIFF8_Background_InternalPal;  // internal palette
    procedure TestWriteReadBIFF8_Background_Biff5Pal;     // official biff5 palette
    procedure TestWriteReadBIFF8_Background_Biff8Pal;     // official biff8 palette
    procedure TestWriteReadBIFF8_Background_RandomPal;    // palette 64, top 56 entries random
    // Font colors...
    procedure TestWriteReadBIFF8_Font_InternalPal;        // internal palette for BIFF8 file format
    procedure TestWriteReadBIFF8_Font_Biff5Pal;           // official biff5 palette in BIFF8 file format
    procedure TestWriteReadBIFF8_Font_Biff8Pal;           // official biff8 palette in BIFF8 file format
    procedure TestWriteReadBIFF8_Font_RandomPal;          // palette 64, top 56 entries random
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

procedure TSpreadWriteReadColorTests.TestWriteReadBackgroundColors(AFormat: TsSpreadsheetFormat;
  WhichPalette: Integer);
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
  pal: Array of TsColorValue;
  i: Integer;
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
    5: MyWorkbook.UsePalette(@PALETTE_BIFF5, Length(PALETTE_BIFF5));
    8: MyWorkbook.UsePalette(@PALETTE_BIFF8, Length(PALETTE_BIFF8));
  999: begin  // Random palette: testing of color replacement
         MyWorkbook.UsePalette(@PALETTE_BIFF8, Length(PALETTE_BIFF8));
         for i:=8 to 63 do  // first 8 colors cannot be changed
           MyWorkbook.SetPaletteColor(i, random(256) + random(256) shr 8 + random(256) shr 16);
       end;
{
  999: begin  // Random palette
         SetLength(pal, 64);
         for i:=0 to 67 do pal[i] := PALETTE_BIFF8[i];
         for i:=8 to 63 do pal[i] := Random(256) + Random(256) shr 8 + random(256) shr 16;
         MyWorkbook.UsePalette(@pal[0], 64);
       end;        }

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
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
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

procedure TSpreadWriteReadColorTests.TestWriteReadFontColors(AFormat: TsSpreadsheetFormat;
  WhichPalette: Integer);
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
  pal: Array of TsColorValue;
  i: Integer;
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
   999: begin  // Random palette: testing of color replacement
          MyWorkbook.UsePalette(@PALETTE_BIFF8, Length(PALETTE_BIFF8));
          for i:=8 to 63 do  // first 8 colors cannot be changed
            MyWorkbook.SetPaletteColor(i, random(256) + random(256) shr 8 + random(256) shr 16);
        end;
{
  999: begin
         SetLength(pal, 64);
         for i:=0 to 7 do pal[i] := PALETTE_BIFF8[i];
         for i:=8 to 63 do pal[i] := Random(256) + Random(256) shr 8 + random(256) shr 16;
         MyWorkbook.UsePalette(@pal[0], 64);
       end;
        }
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
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
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

{ Tests for BIFF2 file format }
{ BIFF2 supports only a fixed palette, and no background color --> test only
  internal palette for font color }
procedure TSpreadWriteReadColorTests.TestWriteReadBIFF2_Font_InternalPal;
begin
  TestWriteReadFontColors(sfExcel2, 0);
end;

{ Tests for BIFF8 file format }
procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Background_InternalPal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Background_Biff5Pal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Background_Biff8Pal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Background_RandomPal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 999);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Font_InternalPal;
begin
  TestWriteReadFontColors(sfExcel8, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Font_Biff5Pal;
begin
  TestWriteReadFontColors(sfExcel8, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Font_Biff8Pal;
begin
  TestWriteReadFontColors(sfExcel8, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteReadBIFF8_Font_RandomPal;
begin
  TestWriteReadFontColors(sfExcel8, 999);
end;

initialization
  RegisterTest(TSpreadWriteReadColorTests);

end.

