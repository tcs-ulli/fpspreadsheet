{ PageLayout tests
  These unit tests are writing out to and reading back from file.
}

unit pagelayouttests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadHyperlinkTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadPageLayoutTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_PageLayout(AFormat: TsSpreadsheetFormat; ANumSheets, ATestMode: Integer);
    procedure TestWriteRead_PageMargins(AFormat: TsSpreadsheetFormat; ANumSheets, AHeaderFooterMode: Integer);
      
  published
    { BIFF2 page layout tests }
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_0;
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_1;
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_2;
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_3;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_0;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_1;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_2;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_3;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_0;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_1;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_2;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_3;

    // no BIFF2 page orientation tests because this info is not readily available in the file


    { BIFF5 page layout tests }
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_0;
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_1;
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_2;
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_3;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_0;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_1;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_2;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_3;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_0;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_1;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_2;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_3;

    procedure TestWriteRead_BIFF5_PageOrientation_1sheet;
    procedure TestWriteRead_BIFF5_PageOrientation_2sheets;
    procedure TestWriteRead_BIFF5_PageOrientation_3sheets;

    procedure TestWriteRead_BIFF5_PaperSize_1sheet;
    procedure TestWriteRead_BIFF5_PaperSize_2sheets;
    procedure TestWriteRead_BIFF5_PaperSize_3sheets;

    procedure TestWriteRead_BIFF5_ScalingFactor_1sheet;
    procedure TestWriteRead_BIFF5_ScalingFactor_2sheets;
    procedure TestWriteRead_BIFF5_ScalingFactor_3sheets;

    { BIFF8 page layout tests }
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_0;
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_1;
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_2;
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_3;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_0;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_1;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_2;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_3;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_0;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_1;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_2;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_3;

    procedure TestWriteRead_BIFF8_PageOrientation_1sheet;
    procedure TestWriteRead_BIFF8_PageOrientation_2sheets;
    procedure TestWriteRead_BIFF8_PageOrientation_3sheets;

    procedure TestWriteRead_BIFF8_PaperSize_1sheet;
    procedure TestWriteRead_BIFF8_PaperSize_2sheets;
    procedure TestWriteRead_BIFF8_PaperSize_3sheets;

    procedure TestWriteRead_BIFF8_ScalingFactor_1sheet;
    procedure TestWriteRead_BIFF8_ScalingFactor_2sheets;
    procedure TestWriteRead_BIFF8_ScalingFactor_3sheets;

    { OOXML page layout tests }
    procedure TestWriteRead_OOXML_PageMargins_1sheet_0;
    procedure TestWriteRead_OOXML_PageMargins_1sheet_1;
    procedure TestWriteRead_OOXML_PageMargins_1sheet_2;
    procedure TestWriteRead_OOXML_PageMargins_1sheet_3;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_0;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_1;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_2;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_3;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_0;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_1;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_2;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_3;

    procedure TestWriteRead_OOXML_PageOrientation_1sheet;
    procedure TestWriteRead_OOXML_PageOrientation_2sheets;
    procedure TestWriteRead_OOXML_PageOrientation_3sheets;

    procedure TestWriteRead_OOXML_PaperSize_1sheet;
    procedure TestWriteRead_OOXML_PaperSize_2sheets;
    procedure TestWriteRead_OOXML_PaperSize_3sheets;

    procedure TestWriteRead_OOXML_ScalingFactor_1sheet;
    procedure TestWriteRead_OOXML_ScalingFactor_2sheets;
    procedure TestWriteRead_OOXML_ScalingFactor_3sheets;

    { OpenDocument page layout tests }
    procedure TestWriteRead_ODS_PageMargins_1sheet_0;
    procedure TestWriteRead_ODS_PageMargins_1sheet_1;
    procedure TestWriteRead_ODS_PageMargins_1sheet_2;
    procedure TestWriteRead_ODS_PageMargins_1sheet_3;
    procedure TestWriteRead_ODS_PageMargins_2sheets_0;
    procedure TestWriteRead_ODS_PageMargins_2sheets_1;
    procedure TestWriteRead_ODS_PageMargins_2sheets_2;
    procedure TestWriteRead_ODS_PageMargins_2sheets_3;
    procedure TestWriteRead_ODS_PageMargins_3sheets_0;
    procedure TestWriteRead_ODS_PageMargins_3sheets_1;
    procedure TestWriteRead_ODS_PageMargins_3sheets_2;
    procedure TestWriteRead_ODS_PageMargins_3sheets_3;

    procedure TestWriteRead_ODS_PageOrientation_1sheet;
    procedure TestWriteRead_ODS_PageOrientation_2sheets;
    procedure TestWriteRead_ODS_PageOrientation_3sheets;

    procedure TestWriteRead_ODS_PaperSize_1sheet;
    procedure TestWriteRead_ODS_PaperSize_2sheets;
    procedure TestWriteRead_ODS_PaperSize_3sheets;

    procedure TestWriteRead_ODS_ScalingFactor_1sheet;
    procedure TestWriteRead_ODS_ScalingFactor_2sheets;
    procedure TestWriteRead_ODS_ScalingFactor_3sheets;

  end;

implementation

uses
  typinfo, fpsutils;
//  uriparser, lazfileutils, fpsutils;

const
  PageLayoutSheet = 'PageLayout';


{ TSpreadWriteReadPageLayoutTests }

procedure TSpreadWriteReadPageLayoutTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadPageLayoutTests.TearDown;
begin
  inherited TearDown;
end;

{ AHeaderFooterMode = 0 ... no header, no footer
                      1 ... header, no footer
                      2 ... no header, footer
	                  3 ... header, footer }
procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins(
  AFormat: TsSpreadsheetFormat; ANumSheets, AHeaderFooterMode: Integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col, p: Integer;
  sollPageLayout, actualPageLayout: TsPageLayout;
  expected, actual: String;
  cell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;

  InitPageLayout(sollPageLayout);
  with SollPageLayout do
  begin
    TopMargin := 20;
    BottomMargin := 30;
    LeftMargin := 21;
    RightMargin := 22;
    HeaderMargin := 10;
    FooterMargin := 11;
    case AHeaderFooterMode of
      0: ;  // header and footer already are empty strings
      1: Headers[HEADER_FOOTER_INDEX_ALL] := 'Test header';
      2: Footers[HEADER_FOOTER_INDEX_ALL] := 'Test footer';
      3: begin 
           Headers[HEADER_FOOTER_INDEX_ALL] := 'Test header';
	       Footers[HEADER_FOOTER_INDEX_ALL] := 'Test footer';
         end;
    end;
  end;
  
  MyWorkbook := TsWorkbook.Create;
  try
    col := 0;
    for p := 1 to ANumSheets do
    begin
      MyWorkSheet:= MyWorkBook.AddWorksheet(PageLayoutSheet+IntToStr(p));
      for row := 0 to 9 do
        Myworksheet.WriteNumber(row, 0, row+col*100+p*10000 );      
      MyWorksheet.PageLayout := SollPageLayout;
    end;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    for p := 0 to MyWorkbook.GetWorksheetCount-1 do
    begin
      MyWorksheet := MyWorkBook.GetWorksheetByIndex(p);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get worksheet by index');
	
      actualPageLayout := MyWorksheet.PageLayout;
      CheckEquals(sollPageLayout.TopMargin, actualPageLayout.TopMargin, 'Top margin mismatch, sheet "'+MyWorksheet.Name+'"');
      CheckEquals(sollPageLayout.BottomMargin, actualPageLayout.Bottommargin, 'Bottom margin mismatch, sheet "'+MyWorksheet.Name+'"');
      CheckEquals(sollPageLayout.LeftMargin, actualPageLayout.LeftMargin, 'Left margin mismatch, sheet "'+MyWorksheet.Name+'"');
      CheckEquals(sollPageLayout.RightMargin, actualPageLayout.RightMargin, 'Right margin mismatch, sheet "'+MyWorksheet.Name+'"');
      if (AFormat <> sfExcel2) then  // No header/footer margin in BIFF2
      begin
        if AHeaderFooterMode in [1, 3] then
          CheckEquals(sollPageLayout.HeaderMargin, actualPageLayout.HeaderMargin, 'Header margin mismatch, sheet "'+MyWorksheet.Name+'"');
        if AHeaderFooterMode in [2, 3] then
          CheckEquals(sollPageLayout.FooterMargin, actualPageLayout.FooterMargin, 'Footer margin mismatch, sheet "'+MyWorksheet.Name+'"');
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

{ ------------------------------------------------------------------------------
 Main page layout test: it writes a file with a specific page layout and reads it
 back. The written pagelayout ("Solllayout") must match the read pagelayout.

 ATestMode:
   0 - Landscape page orientation for sheets 0 und 2, sheet 1 is portrait
   1 - Paper size: sheet 1 "Letter" (8.5" x 11"), sheets 0 and 2 "A5" (148 mm x 210 mm)
   2 - Scaling factor: sheet 1 50%, sheet 2 200%, sheet 3 100%
-------------------------------------------------------------------------------}
procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageLayout(
  AFormat: TsSpreadsheetFormat; ANumSheets, ATestMode: Integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col, p: Integer;
  sollPageLayout: Array of TsPageLayout;
  actualPageLayout: TsPageLayout;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;

  SetLength(SollPageLayout, ANumSheets);
  for p:=0 to High(SollPageLayout) do
  begin
    InitPageLayout(sollPageLayout[p]);
    with SollPageLayout[p] do
    begin
      case ATestMode of
        0: // Page orientation test: sheets 0 and 2 are portrait, sheet 1 is landscape
           if p <> 1 then Orientation := spoLandscape;
        1: // Paper size test: sheets 0 and 2 are A5, sheet 1 is LETTER
           if odd(p) then
           begin
             PageWidth := 8.5*2.54; PageHeight := 11*2.54;
           end else
           begin
             PageWidth := 148; PageHeight := 210;
           end;
        2: // Scaling factor: sheet 1 50%, sheet 2 200%, sheet 3 100%
          begin
            if p = 0 then ScalingFactor := 50 else
            if p = 1 then ScalingFactor := 200;
            Exclude(Options, poFitPages);
          end;
      end;
    end;
  end;

  MyWorkbook := TsWorkbook.Create;
  try
    col := 0;
    for p := 0 to ANumSheets-1 do
    begin
      MyWorkSheet:= MyWorkBook.AddWorksheet(PageLayoutSheet+IntToStr(p+1));
      for row := 0 to 99 do
        Myworksheet.WriteNumber(row, 0, (row+1)+(col+1)*100+(p+1)*10000 );
      MyWorksheet.PageLayout := SollPageLayout[p];
    end;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    for p := 0 to MyWorkbook.GetWorksheetCount-1 do
    begin
      MyWorksheet := MyWorkBook.GetWorksheetByIndex(p);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get worksheet by index');

      actualPageLayout := MyWorksheet.PageLayout;
      case ATestMode of
        0: // Page orientation test
          CheckEquals(GetEnumName(TypeInfo(TsPageOrientation), ord(sollPageLayout[p].Orientation)),
            GetEnumName(TypeInfo(TsPageOrientation), ord(actualPageLayout.Orientation)),
           'Page orientation mismatch, sheet "'+MyWorksheet.Name+'"'
          );
        1: // Paper size test
          begin
            CheckEquals(sollPagelayout[p].PageHeight, actualPageLayout.PageHeight, 0.1,
              'Page height mismatch, sheet "' + MyWorksheet.Name + '"');
            CheckEquals(sollPageLayout[p].PageWidth, actualPageLayout.PageWidth, 0.1,
              'Page width mismatch, sheet "' + MyWorksheet.name + '"');
          end;
        2: // Scaling factor
          begin
            CheckEquals(poFitPages in sollPageLayout[p].Options, poFitPages in actualPageLayout.Options,
              '"poFitPages" option mismatch, sheet "' + MyWorksheet.name + '"');
            CheckEquals(sollPageLayout[p].ScalingFactor, actualPageLayout.ScalingFactor,
              'Scaling factor mismatch, sheet "' + MyWorksheet.Name + '"');
          end;
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


{ Tests for BIFF2 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 3);
end;


{ Tests for BIFF5 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 2);
end;

{ Tests for BIFF8 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 2);
end;

{ Tests for OOXML file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 2);
end;


{ Tests for Open Document file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_ODS_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 2);
end;


initialization
  RegisterTest(TSpreadWriteReadPageLayoutTests);

end.

