unit optiontests;

{$mode objfpc}{$H+}

interface
{ Tests for spreadsheet options
  This unit tests writing out to and reading back from files.
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadOptionTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadOptionsTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteReadPanes(AFormat: TsSpreadsheetFormat;
      ALeftPaneWidth, ATopPaneHeight: Integer);
    procedure TestWriteReadGridHeaders(AFormat: TsSpreadsheetFormat;
      AShowGridLines, AShowHeaders: Boolean);

  published
    // Writes out sheet options & reads back.

    { BIFF2 tests }
    procedure TestWriteRead_BIFF2_ShowGridLines_ShowHeaders;
    procedure TestWriteRead_BIFF2_ShowGridLines_HideHeaders;
    procedure TestWriteRead_BIFF2_HideGridLines_ShowHeaders;
    procedure TestWriteRead_BIFF2_HideGridLines_HideHeaders;

    procedure TestWriteRead_BIFF2_Panes_HorVert;
    procedure TestWriteRead_BIFF2_Panes_Hor;
    procedure TestWriteRead_BIFF2_Panes_Vert;
    procedure TestWriteRead_BIFF2_Panes_None;

    { BIFF5 tests }
    procedure TestWriteRead_BIFF5_ShowGridLines_ShowHeaders;
    procedure TestWriteRead_BIFF5_ShowGridLines_HideHeaders;
    procedure TestWriteRead_BIFF5_HideGridLines_ShowHeaders;
    procedure TestWriteRead_BIFF5_HideGridLines_HideHeaders;

    procedure TestWriteRead_BIFF5_Panes_HorVert;
    procedure TestWriteRead_BIFF5_Panes_Hor;
    procedure TestWriteRead_BIFF5_Panes_Vert;
    procedure TestWriteRead_BIFF5_Panes_None;

    { BIFF8 tests }
    procedure TestWriteRead_BIFF8_ShowGridLines_ShowHeaders;
    procedure TestWriteRead_BIFF8_ShowGridLines_HideHeaders;
    procedure TestWriteRead_BIFF8_HideGridLines_ShowHeaders;
    procedure TestWriteRead_BIFF8_HideGridLines_HideHeaders;

    procedure TestWriteRead_BIFF8_Panes_HorVert;
    procedure TestWriteRead_BIFF8_Panes_Hor;
    procedure TestWriteRead_BIFF8_Panes_Vert;
    procedure TestWriteRead_BIFF8_Panes_None;

    { ODS tests }
    procedure TestWriteRead_ODS_ShowGridLines_ShowHeaders;
    procedure TestWriteRead_ODS_ShowGridLines_HideHeaders;
    procedure TestWriteRead_ODS_HideGridLines_ShowHeaders;
    procedure TestWriteRead_ODS_HideGridLines_HideHeaders;

    procedure TestWriteRead_ODS_Panes_HorVert;
    procedure TestWriteRead_ODS_Panes_Hor;
    procedure TestWriteRead_ODS_Panes_Vert;
    procedure TestWriteRead_ODS_Panes_None;

    { OOXML tests }
    procedure TestWriteRead_OOXML_ShowGridLines_ShowHeaders;
    procedure TestWriteRead_OOXML_ShowGridLines_HideHeaders;
    procedure TestWriteRead_OOXML_HideGridLines_ShowHeaders;
    procedure TestWriteRead_OOXML_HideGridLines_HideHeaders;

    procedure TestWriteRead_OOXML_Panes_HorVert;
    procedure TestWriteRead_OOXML_Panes_Hor;
    procedure TestWriteRead_OOXML_Panes_Vert;
    procedure TestWriteRead_OOXML_Panes_None;
  end;

implementation

const
  OptionsSheet = 'Options';

{ TSpreadWriteReadOptions }

procedure TSpreadWriteReadOptionsTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadOptionsTests.TearDown;
begin
  inherited TearDown;
end;

{ Test for grid lines and sheet headers }

procedure TSpreadWriteReadOptionsTests.TestWriteReadGridHeaders(AFormat: TsSpreadsheetFormat;
  AShowGridLines, AShowHeaders: Boolean);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }

  // Write out show/hide grid lines/sheet headers
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(OptionsSheet);
    if AShowGridLines then
      MyWorksheet.Options := MyWorksheet.Options + [soShowGridLines]
    else
      MyWorksheet.Options := MyWorksheet.Options - [soShowGridLines];
    if AShowHeaders then
      MyWorksheet.Options := MyWorksheet.Options + [soShowHeaders]
    else
      MyWorksheet.Options := MyWorksheet.Options - [soShowHeaders];

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Read back presence of grid lines/sheet headers
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, OptionsSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    CheckEquals(soShowGridLines in MyWorksheet.Options, AShowGridLines,
      'Test saved show grid lines mismatch');
    CheckEquals(soShowHeaders in MyWorksheet.Options, AShowHeaders,
      'Test saved show headers mismatch');
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

{ Tests for BIFF2 grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, false, false);
end;

{ Tests for BIFF5 grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, false, false);
end;

{ Tests for BIFF8 grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, false, false);
end;

{ Tests for ODS grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfOpenDocument, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfOpenDocument, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfOpenDocument, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfOpenDocument, false, false);
end;

{ Tests for OOXML grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfOOXML, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfOOXML, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfOOXML, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfOOXML, false, false);
end;


{ Test for frozen panes }

procedure TSpreadWriteReadOptionsTests.TestWriteReadPanes(AFormat: TsSpreadsheetFormat;
  ALeftPaneWidth, ATopPaneHeight: Integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }

  // Write out pane sizes
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(OptionsSheet);
    MyWorksheet.LeftPaneWidth := ALeftPaneWidth;
    MyWorksheet.TopPaneHeight := ATopPaneHeight;
    MyWorksheet.Options := MyWorksheet.Options + [soHasFrozenPanes];
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Read back pane sizes
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, OptionsSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    CheckEquals(
      (AleftPaneWidth > 0) or (ATopPaneHeight > 0),
      (soHasFrozenPanes in MyWorksheet.Options)
        and ((MyWorksheet.LeftPaneWidth > 0) or (MyWorksheet.TopPaneHeight > 0)),
      'Test saved frozen panes mismatch');
    CheckEquals(ALeftPaneWidth, MyWorksheet.LeftPaneWidth,
      'Test saved left pane width mismatch');
    CheckEquals(ATopPaneHeight, MyWorksheet.TopPaneHeight,
      'Test save top pane height mismatch');
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

{ Tests for BIFF2 frozen panes }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_Panes_HorVert;
begin
  TestWriteReadPanes(sfExcel2, 1, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_Panes_Hor;
begin
  TestWriteReadPanes(sfExcel2, 1, 0);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_Panes_Vert;
begin
  TestWriteReadPanes(sfExcel2, 0, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF2_Panes_None;
begin
  TestWriteReadPanes(sfExcel2, 0, 0);
end;


{ Tests for BIFF5 frozen panes }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_Panes_HorVert;
begin
  TestWriteReadPanes(sfExcel5, 1, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_Panes_Hor;
begin
  TestWriteReadPanes(sfExcel5, 1, 0);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_Panes_Vert;
begin
  TestWriteReadPanes(sfExcel5, 0, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF5_Panes_None;
begin
  TestWriteReadPanes(sfExcel5, 0, 0);
end;

{ Tests for BIFF8 frozen panes }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_Panes_HorVert;
begin
  TestWriteReadPanes(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_Panes_Hor;
begin
  TestWriteReadPanes(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_Panes_Vert;
begin
  TestWriteReadPanes(sfExcel8, 0, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_BIFF8_Panes_None;
begin
  TestWriteReadPanes(sfExcel8, 0, 0);
end;

{ Tests for ODS frozen panes }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_Panes_HorVert;
begin
  TestWriteReadPanes(sfOpenDocument, 1, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_Panes_Hor;
begin
  TestWriteReadPanes(sfOpenDocument, 1, 0);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_Panes_Vert;
begin
  TestWriteReadPanes(sfOpenDocument, 0, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_ODS_Panes_None;
begin
  TestWriteReadPanes(sfOpenDocument, 0, 0);
end;

{ Tests for OOXML frozen panes }
procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_Panes_HorVert;
begin
  TestWriteReadPanes(sfOOXML, 1, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_Panes_Hor;
begin
  TestWriteReadPanes(sfOOXML, 1, 0);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_Panes_Vert;
begin
  TestWriteReadPanes(sfOOXML, 0, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteRead_OOXML_Panes_None;
begin
  TestWriteReadPanes(sfOOXML, 0, 0);
end;


initialization
  RegisterTest(TSpreadWriteReadOptionsTests);

end.

