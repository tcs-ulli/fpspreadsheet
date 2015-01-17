unit dbexporttests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpstypes, fpsallformats, fpspreadsheet,
  testsutility, db, bufdataset, fpsexport;

type
  TExportTestData=record
    id: integer;
    Name: string;
    DOB: TDateTime;
  end;

var
  ExportTestData: array[0..4] of TExportTestData;

procedure InitExportTestData;

type
{ TSpreadExportTests }

  TSpreadExportTests = class(TTestCase)
  private
    FDataset: TBufDataset;
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure TestExport;
  end;

implementation

procedure InitExportTestData;
begin
  with ExportTestData[0] do
  begin
    id:=1;
    name:='Elvis Wesley';
    dob:=encodedate(1912,12,31);
  end;

  with ExportTestData[1] do
  begin
    id:=2;
    name:='Kingsley Dill';
    dob:=encodedate(1918,11,11);
  end;

  with ExportTestData[2] do
  begin
    id:=3;
    name:='Joe Snort';
    dob:=encodedate(1988,8,4);
  end;

  with ExportTestData[3] do
  begin
    id:=4;
    //> may give problems with character encoding
    //http://forum.lazarus.freepascal.org/index.php/topic,26471.0.html
    name:='Hagen > Dit';
    dob:=encodedate(1944,2,24);
  end;

  with ExportTestData[4] do
  begin
    id:=5;
    name:='';
    dob:=encodedate(2112,4,12);
  end;
end;

{ TSpreadExportTests }

procedure TSpreadExportTests.SetUp;
var
  i:integer;
begin
  inherited SetUp;
  InitExportTestData;

  FDataset:=TBufDataset.Create(nil);
  with FDataset.FieldDefs do
  begin
    Add('id',ftAutoinc);
    Add('name',ftString,40);
    Add('dob',ftDateTime);
  end;
  FDataset.CreateDataset;

  for i:=low(ExportTestData) to high(ExportTestData) do
  begin
    FDataset.Append;
    //autoinc field should be filled by bufdataset
    FDataSet.Fields.FieldByName('name').AsString:=ExportTestData[i].Name;
    FDataSet.Fields.FieldByName('dob').AsDateTime:=ExportTestData[i].dob;
    FDataSet.Post;
  end;
end;

procedure TSpreadExportTests.TearDown;
begin
  FDataset.Free;
  inherited TearDown;
end;

procedure TSpreadExportTests.TestExport;
var
  Exp: TFPSExport;
  ExpSettings: TFPSExportFormatSettings;
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: cardinal;
  TempFile: string;
  TheDate: TDateTime;
begin
  Exp := TFPSExport.Create(nil);
  ExpSettings := TFPSExportFormatSettings.Create(true);
  try
    ExpSettings.ExportFormat := efXLS;
    ExpSettings.HeaderRow := true;
    Exp.FormatSettings := ExpSettings;
    Exp.Dataset:=FDataset;
    Exp.FromCurrent:=false; //export from beginning
    TempFile := NewTempFile;
    Exp.FileName := TempFile;
    CheckEquals(length(ExportTestData),Exp.Execute,'Number of exported records');
    CheckTrue(FileExists(TempFile),'Export file must exist');

    // Open the workbook for verification
    MyWorkbook := TsWorkbook.Create;
    try
      // Format must match ExpSettings.ExportFormat above
      MyWorkbook.ReadFromFile(TempFile, sfExcel8);
      MyWorksheet := MyWorkbook.GetFirstWorksheet;
      // ignore header row for now
      for Row := 1 to length(ExportTestData) do
      begin
        // cell 0 is id
        CheckEquals(ExportTestData[Row-1].id,MyWorkSheet.ReadAsNumber(Row,0),'Cell data: id');
        CheckEquals(ExportTestData[Row-1].name,MyWorkSheet.ReadAsUTF8Text(Row,1),'Cell data: name');
        MyWorkSheet.ReadAsDateTime(Row,2,TheDate);
        CheckEquals(ExportTestData[Row-1].dob,TheDate,'Cell data: dob');
      end;
    finally
      MyWorkBook.Free;
    end;
  finally
    Exp.Free;
    ExpSettings.Free;
    DeleteFile(TempFile);
  end;
end;

initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadExportTests);
  InitExportTestData; //useful to have norm data if other code want to use this unit
end.

end.

