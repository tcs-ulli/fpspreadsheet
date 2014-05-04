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
  fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
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
    procedure TestWriteReadBIFF2_ShowGridLines_ShowHeaders;
    procedure TestWriteReadBIFF2_ShowGridLines_HideHeaders;
    procedure TestWriteReadBIFF2_HideGridLines_ShowHeaders;
    procedure TestWriteReadBIFF2_HideGridLines_HideHeaders;

    { BIFF5 tests }
    procedure TestWriteReadBIFF5_ShowGridLines_ShowHeaders;
    procedure TestWriteReadBIFF5_ShowGridLines_HideHeaders;
    procedure TestWriteReadBIFF5_HideGridLines_ShowHeaders;
    procedure TestWriteReadBIFF5_HideGridLines_HideHeaders;

    procedure TestWriteReadBIFF5_Panes_HorVert;
    procedure TestWriteReadBIFF5_Panes_Hor;
    procedure TestWriteReadBIFF5_Panes_Vert;
    procedure TestWriteReadBIFF5_Panes_None;

    { BIFF8 tests }
    procedure TestWriteReadBIFF8_ShowGridLines_ShowHeaders;
    procedure TestWriteReadBIFF8_ShowGridLines_HideHeaders;
    procedure TestWriteReadBIFF8_HideGridLines_ShowHeaders;
    procedure TestWriteReadBIFF8_HideGridLines_HideHeaders;

    procedure TestWriteReadBIFF8_Panes_HorVert;
    procedure TestWriteReadBIFF8_Panes_Hor;
    procedure TestWriteReadBIFF8_Panes_Vert;
    procedure TestWriteReadBIFF8_Panes_None;
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
  MyWorkbook.Free;

  // Read back presence of grid lines/sheet headers
  MyWorkbook := TsWorkbook.Create;
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
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

{ Tests for BIFF2 grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF2_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF2_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF2_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF2_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel2, false, false);
end;

{ Tests for BIFF5 grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel5, false, false);
end;

{ Tests for BIFF8 grid lines and/or headers }
procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_ShowGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, true, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_ShowGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, true, false);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_HideGridLines_ShowHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, false, true);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_HideGridLines_HideHeaders;
begin
  TestWriteReadGridHeaders(sfExcel8, false, false);
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
  MyWorkSheet:= MyWorkBook.AddWorksheet(OptionsSheet);
  MyWorksheet.LeftPaneWidth := ALeftPaneWidth;
  MyWorksheet.TopPaneHeight := ATopPaneHeight;
  MyWorksheet.Options := MyWorksheet.Options + [soHasFrozenPanes];
  MyWorkBook.WriteToFile(TempFile, AFormat, true);
  MyWorkbook.Free;

  // Read back pane sizes
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(TempFile, AFormat);
  if AFormat = sfExcel2 then
    MyWorksheet := MyWorkbook.GetFirstWorksheet
  else
    MyWorksheet := GetWorksheetByName(MyWorkBook, OptionsSheet);
  if MyWorksheet=nil then
    fail('Error in test code. Failed to get named worksheet');
  CheckEquals(soHasFrozenPanes in MyWorksheet.Options, true,
    'Test saved frozen panes mismatch');
  CheckEquals(MyWorksheet.LeftPaneWidth, ALeftPaneWidth,
    'Test saved left pane width mismatch');
  CheckEquals(MyWorksheet.TopPaneHeight, ATopPaneHeight,
    'Test save top pane height mismatch');
  MyWorkbook.Free;

  DeleteFile(TempFile);
end;

{ Tests for BIFF5 frozen panes }
procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_Panes_HorVert;
begin
  TestWriteReadPanes(sfExcel5, 1, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_Panes_Hor;
begin
  TestWriteReadPanes(sfExcel5, 1, 0);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_Panes_Vert;
begin
  TestWriteReadPanes(sfExcel5, 0, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF5_Panes_None;
begin
  TestWriteReadPanes(sfExcel5, 0, 0);
end;

{ Tests for BIFF8 frozen panes }
procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_Panes_HorVert;
begin
  TestWriteReadPanes(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_Panes_Hor;
begin
  TestWriteReadPanes(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_Panes_Vert;
begin
  TestWriteReadPanes(sfExcel8, 0, 2);
end;

procedure TSpreadWriteReadOptionsTests.TestWriteReadBIFF8_Panes_None;
begin
  TestWriteReadPanes(sfExcel8, 0, 0);
end;

initialization
  RegisterTest(TSpreadWriteReadOptionsTests);

end.

