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
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
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
    procedure TestWriteRead_BIFF2_Font_InternalPal;        // internal palette for BIFF2 file format

    { BIFF5 file format tests }
    // Background colors...
    procedure TestWriteRead_BIFF5_Background_InternalPal;  // internal palette
    procedure TestWriteRead_BIFF5_Background_Biff5Pal;     // official biff5 palette
    procedure TestWriteRead_BIFF5_Background_Biff8Pal;     // official biff8 palette
    procedure TestWriteRead_BIFF5_Background_RandomPal;    // palette 64, top 56 entries random
    // Font colors...
    procedure TestWriteRead_BIFF5_Font_InternalPal;        // internal palette for BIFF8 file format
    procedure TestWriteRead_BIFF5_Font_Biff5Pal;           // official biff5 palette in BIFF8 file format
    procedure TestWriteRead_BIFF5_Font_Biff8Pal;           // official biff8 palette in BIFF8 file format
    procedure TestWriteRead_BIFF5_Font_RandomPal;          // palette 64, top 56 entries random

    { BIFF8 file format tests }
    // Background colors...
    procedure TestWriteRead_BIFF8_Background_InternalPal;  // internal palette
    procedure TestWriteRead_BIFF8_Background_Biff5Pal;     // official biff5 palette
    procedure TestWriteRead_BIFF8_Background_Biff8Pal;     // official biff8 palette
    procedure TestWriteRead_BIFF8_Background_RandomPal;    // palette 64, top 56 entries random
    // Font colors...
    procedure TestWriteRead_BIFF8_Font_InternalPal;        // internal palette for BIFF8 file format
    procedure TestWriteRead_BIFF8_Font_Biff5Pal;           // official biff5 palette in BIFF8 file format
    procedure TestWriteRead_BIFF8_Font_Biff8Pal;           // official biff8 palette in BIFF8 file format
    procedure TestWriteRead_BIFF8_Font_RandomPal;          // palette 64, top 56 entries random

    { OpenDocument file format tests }
    // Background colors...
    procedure TestWriteRead_ODS_Background_InternalPal;    // internal palette
    procedure TestWriteRead_ODS_Background_Biff5Pal;       // official biff5 palette
    procedure TestWriteRead_ODS_Background_Biff8Pal;       // official biff8 palette
    procedure TestWriteRead_ODS_Background_RandomPal;      // palette 64, top 56 entries random
    // Font colors...
    procedure TestWriteRead_ODS_Font_InternalPal;          // internal palette for BIFF8 file format
    procedure TestWriteRead_ODS_Font_Biff5Pal;             // official biff5 palette in BIFF8 file format
    procedure TestWriteRead_ODS_Font_Biff8Pal;             // official biff8 palette in BIFF8 file format
    procedure TestWriteRead_ODS_Font_RandomPal;            // palette 64, top 56 entries random

    { OOXML file format tests }
    // Background colors...
    procedure TestWriteRead_OOXML_Background_InternalPal;  // internal palette
    procedure TestWriteRead_OOXML_Background_Biff5Pal;     // official biff5 palette
    procedure TestWriteRead_OOXML_Background_Biff8Pal;     // official biff8 palette
    procedure TestWriteRead_OOXML_Background_RandomPal;    // palette 64, top 56 entries random
    // Font colors...
    procedure TestWriteRead_OOXML_Font_InternalPal;        // internal palette for BIFF8 file format
    procedure TestWriteRead_OOXML_Font_Biff5Pal;           // official biff5 palette in BIFF8 file format
    procedure TestWriteRead_OOXML_Font_Biff8Pal;           // official biff8 palette in BIFF8 file format
    procedure TestWriteRead_OOXML_Font_RandomPal;          // palette 64, top 56 entries random
  end;


implementation

uses
  fpsPalette;

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
  expectedRGB: DWord;
  currentRGB: DWord;
  palette: TsPalette;
  i: Integer;
begin
  TempFile:=GetTempFileName;

  // Define palette
  palette := TsPalette.Create;
  try
    case whichPalette of
      5: palette.UseColors(PALETTE_BIFF5);
      8: palette.UseColors(PALETTE_BIFF8);
    999: begin  // random palette: testing of color replacement
           palette.UseColors(PALETTE_BIFF8);
           for i:=8 to 63 do   // first 8 colors must not be changed in Excel
             palette[i] := random(256) + random(256) shr 8 + random(256) shr 16;
         end;
    else palette.AddBuiltinColors;
    end;

    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkSheet:= MyWorkBook.AddWorksheet(ColorsSheet);

      // Write out all colors
      row := 0;
      col := 0;
      for i := 0 to palette.Count-1 do begin
        MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
        MyCell := MyWorksheet.WriteBackgroundColor(row, col, palette[i]);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        currentRGB := MyWorksheet.ReadBackgroundColor(MyCell);
        expectedRGB := palette[i];
        CheckEquals(expectedRGB, currentRGB,
          'Test unsaved background color, cell ' + CellNotation(MyWorksheet,0,0));
        inc(row);
      end;
      MyWorkBook.WriteToFile(TempFile, AFormat, true);
    finally
      MyWorkbook.Free;
    end;

    // Open the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, ColorsSheet);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet');
      for row := 0 to MyWorksheet.GetLastRowIndex do begin
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        currentRGB := MyWorksheet.ReadBackgroundColor(MyCell);
        expectedRGB := palette[row];
        CheckEquals(expectedRGB, currentRGB,
          'Test saved background color, cell '+CellNotation(MyWorksheet,Row,Col));
      end;
    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

  finally
    palette.Free
  end;
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
  expectedRGB, currentRGB: DWord;
  palette: TsPalette;
  i: Integer;
begin
  TempFile:=GetTempFileName;

  // Define palette
  palette := TsPalette.Create;
  try
    case whichPalette of
      5: palette.UseColors(PALETTE_BIFF5);
      8: palette.UseColors(PALETTE_BIFF8);
    999: begin  // random palette: testing of color replacement
           palette.UseColors(PALETTE_BIFF8);
           for i:=8 to 63 do   // first 8 colors must not be changed in Excel
             palette[i] := random(256) + random(256) shr 8 + random(256) shr 16;
         end;
    else palette.AddBuiltinColors;
    end;

    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkSheet:= MyWorkBook.AddWorksheet(ColorsSheet);

      // Write out all colors
      row := 0;
      col := 0;
      for i := 0 to palette.Count-1 do begin
        MyWorksheet.WriteUTF8Text(row, col, CELLTEXT);
        MyWorksheet.WriteFontColor(row, col, palette[i]);
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        currentRGB := MyWorksheet.ReadCellFont(MyCell).Color;
        expectedRGB := palette[i];
        CheckEquals(expectedRGB, currentRGB,
          'Test unsaved font color, cell ' + CellNotation(MyWorksheet,row, col));
        inc(row);
      end;
      MyWorkBook.WriteToFile(TempFile, AFormat, true);
    finally
      MyWorkbook.Free;
    end;

    // Open the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, ColorsSheet);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet');
      col := 0;
      for row := 0 to MyWorksheet.GetLastRowIndex do begin
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        expectedRGB := palette[row];
        currentRGB := MyWorksheet.ReadCellFont(MyCell).Color;

        // Excel2 cannot write the entire palette. We have to look for the
        // closest color.
        if (AFormat = sfExcel2) then
          expectedRGB := palette[palette.FindClosestColorIndex(expectedRGB, BIFF2_MAX_PALETTE_SIZE)];
        CheckEquals(expectedRGB, currentRGB,
          'Test saved font color, cell '+CellNotation(MyWorksheet,Row,Col));
      end;
    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

  finally
    palette.Free;
  end;
end;

{ Tests for BIFF2 file format }
{ BIFF2 supports only a fixed palette, and no background color --> test only
  internal palette for font color }
procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF2_Font_InternalPal;
begin
  TestWriteReadFontColors(sfExcel2, 0);
end;

{ Tests for BIFF5 file format }
procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Background_InternalPal;
begin
  TestWriteReadBackgroundColors(sfExcel5, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Background_Biff5Pal;
begin
  TestWriteReadBackgroundColors(sfExcel5, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Background_Biff8Pal;
begin
  TestWriteReadBackgroundColors(sfExcel5, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Background_RandomPal;
begin
  TestWriteReadBackgroundColors(sfExcel5, 999);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Font_InternalPal;
begin
  TestWriteReadFontColors(sfExcel5, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Font_Biff5Pal;
begin
  TestWriteReadFontColors(sfExcel5, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Font_Biff8Pal;
begin
  TestWriteReadFontColors(sfExcel5, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF5_Font_RandomPal;
begin
  TestWriteReadFontColors(sfExcel5, 999);
end;

{ Tests for BIFF8 file format }
procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Background_InternalPal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Background_Biff5Pal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Background_Biff8Pal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Background_RandomPal;
begin
  TestWriteReadBackgroundColors(sfExcel8, 999);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Font_InternalPal;
begin
  TestWriteReadFontColors(sfExcel8, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Font_Biff5Pal;
begin
  TestWriteReadFontColors(sfExcel8, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Font_Biff8Pal;
begin
  TestWriteReadFontColors(sfExcel8, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_BIFF8_Font_RandomPal;
begin
  TestWriteReadFontColors(sfExcel8, 999);
end;

{ Tests for Open Document file format }
procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Background_InternalPal;
begin
  TestWriteReadBackgroundColors(sfOpenDocument, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Background_Biff5Pal;
begin
  TestWriteReadBackgroundColors(sfOpenDocument, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Background_Biff8Pal;
begin
  TestWriteReadBackgroundColors(sfOpenDocument, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Background_RandomPal;
begin
  TestWriteReadBackgroundColors(sfOpenDocument, 999);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Font_InternalPal;
begin
  TestWriteReadFontColors(sfOpenDocument, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Font_Biff5Pal;
begin
  TestWriteReadFontColors(sfOpenDocument, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Font_Biff8Pal;
begin
  TestWriteReadFontColors(sfOpenDocument, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_ODS_Font_RandomPal;
begin
  TestWriteReadFontColors(sfOpenDocument, 999);
end;

{ Tests for OOXML file format }
procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Background_InternalPal;
begin
  TestWriteReadBackgroundColors(sfOOXML, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Background_Biff5Pal;
begin
  TestWriteReadBackgroundColors(sfOOXML, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Background_Biff8Pal;
begin
  TestWriteReadBackgroundColors(sfOOXML, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Background_RandomPal;
begin
  TestWriteReadBackgroundColors(sfOOXML, 999);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Font_InternalPal;
begin
  TestWriteReadFontColors(sfOOXML, 0);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Font_Biff5Pal;
begin
  TestWriteReadFontColors(sfOOXML, 5);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Font_Biff8Pal;
begin
  TestWriteReadFontColors(sfOOXML, 8);
end;

procedure TSpreadWriteReadColorTests.TestWriteRead_OOXML_Font_RandomPal;
begin
  TestWriteReadFontColors(sfOOXML, 999);
end;


initialization
  RegisterTest(TSpreadWriteReadColorTests);

end.

