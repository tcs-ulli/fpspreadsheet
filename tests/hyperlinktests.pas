{ Hyperlink tests
  These unit tests are writing out to and reading back from file.
}

unit hyperlinktests;

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
  TSpreadWriteReadHyperlinkTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_Hyperlink(AFormat: TsSpreadsheetFormat;
      ATestMode, ATooltipMode: Integer);

  published
    { BIFF2 comment tests - nothing to do: BIFF2 does not support hyperlinks }
    { BIFF5 comment tests - nothing to do: BIFF5 does not support hyperlinks }

    { BIFF8 comment tests }
    procedure TestWriteRead_Hyperlink_BIFF8_HTTPLink1;
    procedure TestWriteRead_Hyperlink_BIFF8_HTTPLink1_Tooltip1;
    procedure TestWriteRead_Hyperlink_BIFF8_HTTPLink1_Tooltip2;
    procedure TestWriteRead_Hyperlink_BIFF8_HTTPLink2;
    procedure TestWriteRead_Hyperlink_BIFF8_HTTPLink2_Tooltip1;
    procedure TestWriteRead_Hyperlink_BIFF8_HTTPLink2_Tooltip2;
    procedure TestWriteRead_Hyperlink_BIFF8_FileLink;
    procedure TestWriteRead_Hyperlink_BIFF8_FileLink_Tooltip1;
    procedure TestWriteRead_Hyperlink_BIFF8_FileLink_Tooltip2;
    procedure TestWriteRead_Hyperlink_BIFF8_RelFileLink1;
    procedure TestWriteRead_Hyperlink_BIFF8_RelFileLink1_Tooltip1;
    procedure TestWriteRead_Hyperlink_BIFF8_RelFileLink1_Tooltip2;
    procedure TestWriteRead_Hyperlink_BIFF8_RelFileLink2;
    procedure TestWriteRead_Hyperlink_BIFF8_RelFileLink2_Tooltip1;
    procedure TestWriteRead_Hyperlink_BIFF8_RelFileLink2_Tooltip2;
    procedure TestWriteRead_Hyperlink_BIFF8_InternalLink;
    procedure TestWriteRead_Hyperlink_BIFF8_InternalLink_Tooltip1;
    procedure TestWriteRead_Hyperlink_BIFF8_InternalLink_Tooltip2;

    { OpenDocument comment tests }
    procedure TestWriteRead_Hyperlink_ODS_HTTPLink1;
    procedure TestWriteRead_Hyperlink_ODS_HTTPLink1_Tooltip1;
    procedure TestWriteRead_Hyperlink_ODS_HTTPLink1_Tooltip2;
    procedure TestWriteRead_Hyperlink_ODS_HTTPLink2;
    procedure TestWriteRead_Hyperlink_ODS_HTTPLink2_Tooltip1;
    procedure TestWriteRead_Hyperlink_ODS_HTTPLink2_Tooltip2;
    procedure TestWriteRead_Hyperlink_ODS_FileLink;
    procedure TestWriteRead_Hyperlink_ODS_FileLink_Tooltip1;
    procedure TestWriteRead_Hyperlink_ODS_FileLink_Tooltip2;
    procedure TestWriteRead_Hyperlink_ODS_RelFileLink1;
    procedure TestWriteRead_Hyperlink_ODS_RelFileLink1_Tooltip1;
    procedure TestWriteRead_Hyperlink_ODS_RElFileLink1_Tooltip2;
    procedure TestWriteRead_Hyperlink_ODS_RelFileLink2;
    procedure TestWriteRead_Hyperlink_ODS_RelFileLink2_Tooltip1;
    procedure TestWriteRead_Hyperlink_ODS_RelFileLink2_Tooltip2;
    procedure TestWriteRead_Hyperlink_ODS_InternalLink;
    procedure TestWriteRead_Hyperlink_ODS_InternalLink_Tooltip1;
    procedure TestWriteRead_Hyperlink_ODS_InternalLink_Tooltip2;

    { OOXML comment tests }
    procedure TestWriteRead_Hyperlink_OOXML_HTTPLink1;
    procedure TestWriteRead_Hyperlink_OOXML_HTTPLink1_Tooltip1;
    procedure TestWriteRead_Hyperlink_OOXML_HTTPLink1_Tooltip2;
    procedure TestWriteRead_Hyperlink_OOXML_HTTPLink2;
    procedure TestWriteRead_Hyperlink_OOXML_HTTPLink2_Tooltip1;
    procedure TestWriteRead_Hyperlink_OOXML_HTTPLink2_Tooltip2;
    procedure TestWriteRead_Hyperlink_OOXML_FileLink;
    procedure TestWriteRead_Hyperlink_OOXML_FileLink_Tooltip1;
    procedure TestWriteRead_Hyperlink_OOXML_FileLink_Tooltip2;
    procedure TestWriteRead_Hyperlink_OOXML_RelFileLink1;
    procedure TestWriteRead_Hyperlink_OOXML_RelFileLink1_Tooltip1;
    procedure TestWriteRead_Hyperlink_OOXML_RelFileLink1_Tooltip2;
    procedure TestWriteRead_Hyperlink_OOXML_RelFileLink2;
    procedure TestWriteRead_Hyperlink_OOXML_RelFileLink2_Tooltip1;
    procedure TestWriteRead_Hyperlink_OOXML_RelFileLink2_Tooltip2;
    procedure TestWriteRead_Hyperlink_OOXML_InternalLink;
    procedure TestWriteRead_Hyperlink_OOXML_InternalLink_Tooltip1;
    procedure TestWriteRead_Hyperlink_OOXML_InternalLink_Tooltip2;
  end;

implementation

uses
  lazfileutils, fpsutils;

const
  HyperlinkSheet = 'Hyperlinks';

var
  SollLinks: array[0..5] of String = (
    'http://wiki.lazarus.freepascal.org/Lazarus_Documentation',
    'http://wiki.lazarus.freepascal.org/Lazarus_Documentation#The_Lazarus_User_Guides',
    'file:///',   // file link: path of test file will be added
    'testbiff8_1899.xls',
    'testbiff8_1899.xls#Texts!A2',
    '#A10'
  );
  SollCellContent: array[0..3] of string = (
    '',
    'Label',    // Label cell
    '1',        // Number cell
    '12:00:00'  // Date/time cell
  );
  SollTooltip: array[0..2] of String = (
    '',  // no tooltip
    'This is the tooltip for a hyperlink.',
    '<< Special characters äöüÄÖÜ >>'
  );


{ TSpreadWriteReadHyperlinkTests }

procedure TSpreadWriteReadHyperlinkTests.SetUp;
var
  i: Integer;
begin
  inherited SetUp;
  for i:=Low(SollLinks) to High(SollLinks) do
    if SollLinks[i] = 'file:///' then
    begin
      SollLinks[i] := 'file:///' + ExpandFileName('testbiff8_1899.xls');
      exit;
    end;
end;

procedure TSpreadWriteReadHyperlinkTests.TearDown;
begin
  inherited TearDown;
end;

{ Tests differ by "TestMode" (http link, file link, internal link) and usage of
  tooltip (no tooltip, tootip with "normal" characters, tooltip with special
  characters - "ToolTipMode"). All cells have hyperlinks based on the same
  combination of TestMode and ToolTipMode, but they differ in their content
  (SollCellContent): blank, string, number, date/time. }
procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink(
  AFormat: TsSpreadsheetFormat; ATestMode, ATooltipMode: Integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  hyperlink: TsHyperlink;
  expected, actual: String;
  cell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(HyperlinkSheet);

    col := 0;
    for row := 0 to High(SollCellContent) do
    begin
      Myworksheet.WriteHyperlink(row, col, SollLinks[ATestMode], SollTooltip[AToolTipMode]);
      if SollCellContent[row] <> '' then
        MyWorksheet.WriteCellValueAsString(row, col, SollCellContent[row]);
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
    // To see the file also in the test folder uncomment the next line
    // MyWorkBook.WriteToFile(Format('hyperlink_Test_%d_%d%s', [ATestMode, AToolTipMode, GetFileFormatExt(AFormat)]), AFormat, true);

  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    MyWorksheet := GetWorksheetByName(MyWorkBook, HyperlinkSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    col := 0;
    for row := 0 to High(SollCellContent) do
    begin
      cell := MyWorksheet.FindCell(row, col);

      // Open document can attach hyperlinks only to label cells --> skip this test
      if (AFormat = sfOpenDocument) and (cell^.ContentType <> cctUTF8String) then
        continue;

      hyperlink := MyWorksheet.ReadHyperlink(cell);

      actual := hyperlink.Target;
      expected := SollLinks[ATestMode];
      // Make sure that the same path delimiter is used in the comparison (fps accepts both)
      FixHyperlinkPathDelims(actual);
      FixHyperlinkPathDelims(expected);
      CheckEquals(expected, actual,
        'Test saved hyperlink target, cell '+CellNotation(MyWorksheet, row, col));

      actual := MyWorksheet.ReadAsUTF8Text(cell);
      if row = 0 then begin
        // an originally blank cell shows the hyperlink.Target. But Worksheet.WriteHyperlink
        // removes the "file:///" protocol
        expected := hyperlink.Target;
        if pos('file:', SollLinks[ATestMode])=1 then begin
          Delete(expected, 1, Length('file:///'));
          ForcePathDelims(expected);
        end;
      end else
        expected := SollCellContent[row];
      CheckEquals(expected, actual,
        'Test saved hyperlink cell text, cell '+ CellNotation(MyWorksheet, row, col));

      // Tooltips are not supported by ODS --> don't check
      if AFormat <> sfOpenDocument then
        CheckEquals(SollToolTip[AToolTipMode], hyperlink.Tooltip,
          'Test saved hyperlink tooltip, cell ' + CellNotation(MyWorksheet, row, col));
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


{ Tests for BIFF8 file format }
procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_HttpLink1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 0, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_HttpLink1_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 0, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_HttpLink1_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 0, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_HttpLink2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_HttpLink2_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 1, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_HttpLink2_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_FileLink;
begin
  TestWriteRead_Hyperlink(sfExcel8, 2, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_FileLink_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 2, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_FileLink_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 2, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_RelFileLink1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 3, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_RelFileLink1_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 3, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_RelFileLink1_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 3, 2);
end;
procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_RelFileLink2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 4, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_RelFileLink2_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 4, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_RelFileLink2_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 4, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_InternalLink;
begin
  TestWriteRead_Hyperlink(sfExcel8, 5, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_InternalLink_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfExcel8, 5, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_BIFF8_InternalLink_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfExcel8, 5, 2);
end;


{ Tests for Open Document file format }
procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_HttpLink1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 0, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_HttpLink1_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 0, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_HttpLink1_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 0, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_HttpLink2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 1, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_HttpLink2_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 1, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_HttpLink2_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 1, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_FileLink;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 2, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_FileLink_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 2, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_FileLink_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 2, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_RelFileLink1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 3, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_RelFileLink1_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 3, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_RelFileLink1_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 3, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_RelFileLink2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 4, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_RelFileLink2_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 4, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_RelFileLink2_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 4, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_InternalLink;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 5, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_InternalLink_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 5, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_ODS_InternalLink_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOpenDocument, 5, 2);
end;


{ Tests for OOXML file format }
procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_HttpLink1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 0, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_HttpLink1_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 0, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_HttpLink1_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 0, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_HttpLink2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 1, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_HttpLink2_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 1, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_HttpLink2_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 1, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_FileLink;
begin
  TestWriteRead_Hyperlink(sfOOXML, 2, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_FileLink_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 2, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_FileLink_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 2, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_RelFileLink1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 3, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_RelFileLink1_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 3, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_RelFileLink1_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 3, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_RelFileLink2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 4, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_RelFileLink2_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 4, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_RelFileLink2_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 4, 2);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_InternalLink;
begin
  TestWriteRead_Hyperlink(sfOOXML, 5, 0);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_InternalLink_ToolTip1;
begin
  TestWriteRead_Hyperlink(sfOOXML, 5, 1);
end;

procedure TSpreadWriteReadHyperlinkTests.TestWriteRead_Hyperlink_OOXML_InternalLink_ToolTip2;
begin
  TestWriteRead_Hyperlink(sfOOXML, 5, 2);
end;

initialization
  RegisterTest(TSpreadWriteReadHyperlinkTests);

end.

