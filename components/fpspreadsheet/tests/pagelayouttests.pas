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
    procedure TestWriteRead_PageMargins(AFormat: TsSpreadsheetFormat; ANumSheets, AHeaderFooterMode: Integer);
      
  published
    { BIFF2 page layout tests }
    procedure TestWriteRead_PageMargins_BIFF2_1sheet_0;
    procedure TestWriteRead_PageMargins_BIFF2_1sheet_1;
    procedure TestWriteRead_PageMargins_BIFF2_1sheet_2;
    procedure TestWriteRead_PageMargins_BIFF2_1sheet_3;
    procedure TestWriteRead_PageMargins_BIFF2_2sheets_0;
    procedure TestWriteRead_PageMargins_BIFF2_2sheets_1;
    procedure TestWriteRead_PageMargins_BIFF2_2sheets_2;
    procedure TestWriteRead_PageMargins_BIFF2_2sheets_3;
    procedure TestWriteRead_PageMargins_BIFF2_3sheets_0;
    procedure TestWriteRead_PageMargins_BIFF2_3sheets_1;
    procedure TestWriteRead_PageMargins_BIFF2_3sheets_2;
    procedure TestWriteRead_PageMargins_BIFF2_3sheets_3;

    { BIFF5 page layout tests }
    procedure TestWriteRead_PageMargins_BIFF5_1sheet_0;
    procedure TestWriteRead_PageMargins_BIFF5_1sheet_1;
    procedure TestWriteRead_PageMargins_BIFF5_1sheet_2;
    procedure TestWriteRead_PageMargins_BIFF5_1sheet_3;
    procedure TestWriteRead_PageMargins_BIFF5_2sheets_0;
    procedure TestWriteRead_PageMargins_BIFF5_2sheets_1;
    procedure TestWriteRead_PageMargins_BIFF5_2sheets_2;
    procedure TestWriteRead_PageMargins_BIFF5_2sheets_3;
    procedure TestWriteRead_PageMargins_BIFF5_3sheets_0;
    procedure TestWriteRead_PageMargins_BIFF5_3sheets_1;
    procedure TestWriteRead_PageMargins_BIFF5_3sheets_2;
    procedure TestWriteRead_PageMargins_BIFF5_3sheets_3;

    { BIFF8 page layout tests }
    procedure TestWriteRead_PageMargins_BIFF8_1sheet_0;
    procedure TestWriteRead_PageMargins_BIFF8_1sheet_1;
    procedure TestWriteRead_PageMargins_BIFF8_1sheet_2;
    procedure TestWriteRead_PageMargins_BIFF8_1sheet_3;
    procedure TestWriteRead_PageMargins_BIFF8_2sheets_0;
    procedure TestWriteRead_PageMargins_BIFF8_2sheets_1;
    procedure TestWriteRead_PageMargins_BIFF8_2sheets_2;
    procedure TestWriteRead_PageMargins_BIFF8_2sheets_3;
    procedure TestWriteRead_PageMargins_BIFF8_3sheets_0;
    procedure TestWriteRead_PageMargins_BIFF8_3sheets_1;
    procedure TestWriteRead_PageMargins_BIFF8_3sheets_2;
    procedure TestWriteRead_PageMargins_BIFF8_3sheets_3;

    { OOXML page layout tests }
    procedure TestWriteRead_PageMargins_OOXML_1sheet_0;
    procedure TestWriteRead_PageMargins_OOXML_1sheet_1;
    procedure TestWriteRead_PageMargins_OOXML_1sheet_2;
    procedure TestWriteRead_PageMargins_OOXML_1sheet_3;
    procedure TestWriteRead_PageMargins_OOXML_2sheets_0;
    procedure TestWriteRead_PageMargins_OOXML_2sheets_1;
    procedure TestWriteRead_PageMargins_OOXML_2sheets_2;
    procedure TestWriteRead_PageMargins_OOXML_2sheets_3;
    procedure TestWriteRead_PageMargins_OOXML_3sheets_0;
    procedure TestWriteRead_PageMargins_OOXML_3sheets_1;
    procedure TestWriteRead_PageMargins_OOXML_3sheets_2;
    procedure TestWriteRead_PageMargins_OOXML_3sheets_3;

    { OpenDocument page layout tests }
    procedure TestWriteRead_PageMargins_ODS_1sheet_0;
    procedure TestWriteRead_PageMargins_ODS_1sheet_1;
    procedure TestWriteRead_PageMargins_ODS_1sheet_2;
    procedure TestWriteRead_PageMargins_ODS_1sheet_3;
    procedure TestWriteRead_PageMargins_ODS_2sheets_0;
    procedure TestWriteRead_PageMargins_ODS_2sheets_1;
    procedure TestWriteRead_PageMargins_ODS_2sheets_2;
    procedure TestWriteRead_PageMargins_ODS_2sheets_3;
    procedure TestWriteRead_PageMargins_ODS_3sheets_0;
    procedure TestWriteRead_PageMargins_ODS_3sheets_1;
    procedure TestWriteRead_PageMargins_ODS_3sheets_2;
    procedure TestWriteRead_PageMargins_ODS_3sheets_3;

  end;

implementation

uses
  uriparser, lazfileutils, fpsutils;

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


{ Tests for BIFF8 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_PageMargins_BIFF2_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF2_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 3);
end;

{ Tests for BIFF8 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_PageMargins_BIFF5_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF5_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 3);
end;


{ Tests for BIFF8 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_PageMargins_BIFF8_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_BIFF8_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 3);
end;


{ Tests for OOXML file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_1sheet_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_1sheet_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_1sheet_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_1sheet_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_PageMargins_OOXML_2sheets_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_2sheets_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_2sheets_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_2sheets_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_3sheets_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_3sheets_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_3sheets_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_OOXML_3sheets_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 3);
end;


{ Tests for Open Document file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_1sheet_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_1sheet_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_1sheet_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_1sheet_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_PageMargins_ODS_2sheets_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_2sheets_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_2sheets_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_2sheets_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_3sheets_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_3sheets_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_3sheets_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins_ODS_3sheets_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 3);
end;


initialization
  RegisterTest(TSpreadWriteReadPageLayoutTests);

end.

