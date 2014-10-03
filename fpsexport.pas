unit fpsexport;

{ 
  Exports dataset to spreadsheet/tabular format 
  either XLS (Excel), XLSX (Excel), ODS (OpenOffice/LibreOffice)
  or wikitable
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, db, fpsallformats, fpspreadsheet, fpdbexport;
  
Type

  { TFPSExportFieldItem }

  TFPSExportFieldItem = Class(TExportFieldItem)
  private
    FDestField: TField;
  protected
    property DestField : TField read FDestField;
  end;

  { TFPSExportFormatSettings }

  TExportFormat = (efXLS {BIFF8},efXLSX,efODS,efWikiTable);
  
  TFPSExportFormatSettings = class(TExportFormatSettings)
  private
    FExportFormat: TExportFormat;
    FHeaderRow: boolean;
  public
    procedure Assign(Source : TPersistent); override;
    procedure InitSettings; override;
  published
    // File format for the export
    property ExportFormat: TExportFormat read FExportFormat write FExportFormat;
    // Write the field list to the first row of the spreadsheet or not
    property HeaderRow: boolean read FHeaderRow write FHeaderRow default false;
  end;

  { TCustomFPSExport }
  TCustomFPSExport = Class(TCustomDatasetExporter)
  private
    FRow: cardinal; //current row in exported spreadsheet
    FSpreadsheet: TsWorkbook;
    FSheet: TsWorksheet;
    FFileName: string;
    function GetSettings: TFPSExportFormatSettings;
    procedure SetSettings(const AValue: TFPSExportFormatSettings);
  protected
    function CreateFormatSettings: TCustomExportFormatSettings; override;

    function CreateExportFields: TExportFields; override;
    procedure DoBeforeExecute; override;
    procedure DoAfterExecute; override;
    procedure DoDataHeader; override;
    procedure DoDataRowEnd; override;
    procedure ExportField(EF : TExportFieldItem); override;
    property FileName: String read FFileName write FFileName;
    property Workbook: TsWorkbook read FSpreadsheet;
  public
    property FormatSettings: TFPSExportFormatSettings read GetSettings write SetSettings;
  end;

  TFPSExport = Class(TCustomFPSExport)
  published
    property FileName;
    property Dataset;
    property ExportFields;
    property FromCurrent;
    property RestorePosition;
    property FormatSettings;
    property OnExportRow;
  end;
  
procedure RegisterFPSExportFormat;
procedure UnRegisterFPSExportFormat;

Const
  SFPSExport = 'xls';
  SPFSFilter = '*.xls'; //todo: add others?
  
ResourceString
  SFPSDescription = 'Spreadsheet files';

implementation


{ TCustomFPSExport }

function TCustomFPSExport.GetSettings: TFPSExportFormatSettings;
begin
  result:=TFPSExportFormatSettings(Inherited FormatSettings);
end;

procedure TCustomFPSExport.SetSettings
  (const AValue: TFPSExportFormatSettings);
begin
  Inherited FormatSettings.Assign(AValue);
end;

function TCustomFPSExport.CreateFormatSettings: TCustomExportFormatSettings;
begin
  result:=TFPSExportFormatSettings.Create(True);
end;

function TCustomFPSExport.CreateExportFields: TExportFields;
begin
  result:=TExportFields.Create(TFPSExportFieldItem);
end;

procedure TCustomFPSExport.DoBeforeExecute;
begin
  Inherited;
  FSpreadsheet:=TsWorkbook.Create;
  // For extra performance. Note that virtual mode is not an option
  // due to the data export determining flow of the program.
  FSpreadsheet.Options:=FSpreadsheet.Options+[boBufStream];
  FSheet:=FSpreadsheet.AddWorksheet('1');
  FRow:=0;
end;

procedure TCustomFPSExport.DoDataHeader;
var
  i: integer;
begin
  if FormatSettings.FHeaderRow then
  begin
    for i:=0 to ExportFields.Count-1 do
    begin
      FSheet.WriteUTF8Text(FRow,i,ExportFields[i].ExportedName);
    end;
    inc(FRow);
  end;
  inherited DoDataHeader;
end;

procedure TCustomFPSExport.DoAfterExecute;
begin
  FRow:=0;
  // Overwrite existing file similar to how dbf export does it
  case Formatsettings.ExportFormat of
    efXLS: FSpreadSheet.WriteToFile(FFileName,sfExcel8,true);
    efXLSX: FSpreadsheet.WriteToFile(FFilename,sfOOXML,true);
    efODS: FSpreadSheet.WriteToFile(FFileName,sfOpenDocument,true);
    efWikiTable: FSpreadSheet.WriteToFile(FFileName,sfWikitable_wikimedia,true);
    else
      ;// raise error?
  end;

  // Don't free FSheet; done by FSpreadsheet
  try
    FreeAndNil(FSpreadsheet);
  finally
    Inherited;
  end;
end;

procedure TCustomFPSExport.DoDataRowEnd;
begin
  FRow:=FRow+1;
end;

procedure TCustomFPSExport.ExportField(EF: TExportFieldItem);
var
  F : TFPSExportFieldItem;
begin
  F:=EF as TFPSExportFieldItem;
  with F do
  begin
    // Export depending on field datatype;
    // Fall back to string if unknown datatype
    If Field.IsNull then
      FSheet.WriteUTF8Text(FRow,EF.Index,'')
    else if Field.Datatype in (IntFieldTypes+[ftAutoInc,ftLargeInt]) then
      FSheet.WriteNumber(FRow,EF.Index,Field.AsInteger)
    else if Field.Datatype in [ftBCD,ftCurrency,ftFloat,ftFMTBcd] then
      FSheet.WriteNumber(FRow,EF.Index,Field.AsFloat)
    else if Field.DataType in [ftString,ftFixedChar] then
      FSheet.WriteUTF8Text(FRow,EF.Index,Field.AsString)
    else if (Field.DataType in ([ftWideMemo,ftWideString,ftFixedWideChar]+BlobFieldTypes)) then
      FSheet.WriteUTF8Text(FRow,EF.Index,UTF8Encode(Field.AsWideString))
      { Note: we test for the wide text fields before the MemoFieldTypes, in order to
      let ftWideMemo end up at the right place }
    else if Field.DataType in MemoFieldTypes then
      FSheet.WriteUTF8Text(FRow,EF.Index,Field.AsString)
    else if Field.DataType=ftBoolean then
      FSheet.WriteBoolValue(FRow,EF.Index,Field.AsBoolean)
    else if field.DataType in DateFieldTypes then
      FSheet.WriteDateTime(FRow,EF.Index,Field.AsDateTime)
    else //fallback to string
      FSheet.WriteUTF8Text(FRow,EF.Index,Field.AsString);
  end;
end;

procedure RegisterFPSExportFormat;
begin
  RegisterExportFormat(SFPSExport,SFPSDescription,SPFSFilter,TFPSExport);
end;

procedure UnRegisterFPSExportFormat;
begin
  UnregisterExportFormat(SFPSExport);
end;

{ TFPSExportFormatSettings }

procedure TFPSExportFormatSettings.Assign(Source: TPersistent);
var
  FS : TFPSExportFormatSettings;
begin
  If Source is TFPSExportFormatSettings then
  begin
    FS:=Source as TFPSExportFormatSettings;
    HeaderRow:=FS.HeaderRow;
    ExportFormat:=FS.ExportFormat;
  end;
  inherited Assign(Source);
end;

procedure TFPSExportFormatSettings.InitSettings;
begin
  inherited InitSettings;
  FExportFormat:=efXLS; //often used
end;

end.

