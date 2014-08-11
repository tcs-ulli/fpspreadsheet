unit virtualmodetests;
{ Tests for VirtualMode }

{$mode objfpc}{$H+}

interface

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  fpsutils, testsutility;

type
  { TSpreadVirtualModeTests }

  TSpreadVirtualModeTests= class(TTestCase)
  private
    procedure WriteVirtualCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
      var AValue:Variant; var AStyleCell: PCell);
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteVirtualMode(AFormat: TsSpreadsheetFormat; ABufStreamMode: Boolean);

  published
    // Virtual mode tests for all file formats
    procedure TestWriteVirtualMode_BIFF2;
    procedure TestWriteVirtualMode_BIFF5;
    procedure TestWriteVirtualMode_BIFF8;
    procedure TestWriteVirtualMode_ODS;
    procedure TestWriteVirtualMode_OOXML;

    procedure TestWriteVirtualMode_BIFF2_BufStream;
    procedure TestWriteVirtualMode_BIFF5_BufStream;
    procedure TestWriteVirtualMode_BIFF8_BufStream;
    procedure TestWriteVirtualMode_ODS_BufStream;
    procedure TestWriteVirtualMode_OOXML_BufStream;
  end;

implementation

uses
   numberstests, stringtests;

const
  VIRTUALMODE_SHEET = 'VirtualMode'; //worksheet name

procedure TSpreadVirtualModeTests.SetUp;
begin
end;

procedure TSpreadVirtualModeTests.TearDown;
begin
end;

procedure TSpreadVirtualModeTests.WriteVirtualCellDataHandler(Sender: TObject;
  ARow, ACol: Cardinal; var AValue:Variant; var AStyleCell: PCell);
begin
  Unused(ACol);
  Unused(AStyleCell);
  // First read the SollNumbers, then the first 4 SollStrings
  // See comment in TestVirtualMode().
  if ARow < Length(SollNumbers) then
    AValue := SollNumbers[ARow]
  else
    AValue := SollStrings[ARow - Length(SollNumbers)];
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode(AFormat: TsSpreadsheetFormat;
  ABufStreamMode: Boolean);
var
  tempFile: String;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  row, col: Integer;
  value: Double;
  s: String;
begin
  try
    workbook := TsWorkbook.Create;
    try
      worksheet := workbook.AddWorksheet(VIRTUALMODE_SHEET);
      workbook.Options := workbook.Options + [boVirtualMode];
      if ABufStreamMode then
        workbook.Options := workbook.Options + [boBufStream];
      workbook.VirtualColCount := 1;
      workbook.VirtualRowCount := Length(SollNumbers) + 4;
      // We'll use only the first 4 SollStrings, the others cause trouble due to utf8 and formatting.
      workbook.OnWriteCellData := @WriteVirtualCellDataHandler;
      tempFile:=NewTempFile;
      workbook.WriteToFile(tempfile, AFormat, true);
    finally
      workbook.Free;
    end;

    workbook := TsWorkbook.Create;
    try
      workbook.ReadFromFile(tempFile, AFormat);
      worksheet := workbook.GetWorksheetByIndex(0);
      col := 0;
      CheckEquals(Length(SollNumbers) + 4, worksheet.GetLastRowIndex+1,
        'Row count mismatch');
      for row := 0 to Length(SollNumbers)-1 do
      begin
        value := worksheet.ReadAsNumber(row, col);
        CheckEquals(SollNumbers[row], value,
          'Test number value mismatch, cell '+CellNotation(workSheet, row, col))
      end;
      for row := Length(SollNumbers) to worksheet.GetLastRowIndex do
      begin
        s := worksheet.ReadAsUTF8Text(row, col);
        CheckEquals(SollStrings[row - Length(SollNumbers)], s,
          'Test string value mismatch, cell '+CellNotation(workSheet, row, col));
      end;
    finally
      workbook.Free;
    end;

  finally
    DeleteFile(tempFile);
  end;
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_BIFF2;
begin
  TestWriteVirtualMode(sfExcel2, false);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_BIFF5;
begin
  TestWriteVirtualMode(sfExcel5, false);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_BIFF8;
begin
  TestWriteVirtualMode(sfExcel8, false);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_ODS;
begin
  TestWriteVirtualMode(sfOpenDocument, false);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_OOXML;
begin
  TestWriteVirtualMode(sfOOXML, false);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_BIFF2_BufStream;
begin
  TestWriteVirtualMode(sfExcel2, True);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_BIFF5_BufStream;
begin
  TestWriteVirtualMode(sfExcel5, true);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_BIFF8_BufStream;
begin
  TestWriteVirtualMode(sfExcel8, true);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_ODS_BufStream;
begin
  TestWriteVirtualMode(sfOpenDocument, true);
end;

procedure TSpreadVirtualModeTests.TestWriteVirtualMode_OOXML_BufStream;
begin
  TestWriteVirtualMode(sfOOXML, true);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadVirtualModeTests);

end.

