unit fpsexport;

{ 
  Exports dataset to spreadsheet/tabular format 
  either XLS (Excel), XLSX (Excel), ODS (OpenOffice/LibreOffice)
  or wikitable
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, db,
  {%H-}fpsallformats, fpspreadsheet, fpsstrings, fpdbexport;
  
Type

  { TFPSExportFieldItem }

  TFPSExportFieldItem = Class(TExportFieldItem)
  private
    FDestField: TField;
  protected
    property DestField : TField read FDestField;
  end;

  TExportFormat = (efXLS {BIFF8},efXLSX,efODS,efWikiTable);

  { TFPSExportFormatSettings }
  {@@ Specific export settings that apply to spreadsheet export}
  TFPSExportFormatSettings = class(TExportFormatSettings)
  private
    FExportFormat: TExportFormat;
    FHeaderRow: boolean;
    FSheetName: String;
  public
    procedure Assign(Source : TPersistent); override;
    procedure InitSettings; override;
  published
    {@@ File format for the export }
    property ExportFormat: TExportFormat read FExportFormat write FExportFormat;
    {@@ Flag that determines whether to write the field list to the first
        row of the spreadsheet }
    property HeaderRow: boolean read FHeaderRow write FHeaderRow default false;
    {@@ Sheet name }
    property SheetName: String read FSheetName write FSheetName;
  end;

  { TGetSheetNameEvent }
  TsGetSheetNameEvent = procedure (Sender: TObject; ASheetIndex: Integer;
    var ASheetName: String) of object;

  { TCustomFPSExport }
  TCustomFPSExport = Class(TCustomDatasetExporter)
  private
    FRow: cardinal; //current row in exported spreadsheet
    FSpreadsheet: TsWorkbook;
    FSheet: TsWorksheet;
    FFileName: string;
    FMultipleSheets: Boolean;
    FOnGetSheetName: TsGetSheetNameEvent;
    function CalcSheetNameMask(const AMask: String): String;
    function CalcUniqueSheetName(const AMask: String): String;
    function GetSettings: TFPSExportFormatSettings;
    procedure SaveWorkbook;
    procedure SetSettings(const AValue: TFPSExportFormatSettings);
  protected
    function CreateExportFields: TExportFields; override;
    function CreateFormatSettings: TCustomExportFormatSettings; override;
    procedure DoBeforeExecute; override;
    procedure DoAfterExecute; override;
    procedure DoDataHeader; override;
    procedure DoDataRowEnd; override;
    function  DoGetSheetName: String; virtual;
    procedure ExportField(EF : TExportFieldItem); override;
    property FileName: String read FFileName write FFileName;
    property Workbook: TsWorkbook read FSpreadsheet;
    property RestorePosition default true;
    property OnGetSheetName: TsGetSheetNameEvent read FOnGetSheetName write FOnGetSheetName;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure WriteExportFile;
    {@@ Settings for the export. Note: a lot of generic settings are preent
        that are not relevant for this export, e.g. decimal point settings }
    property FormatSettings: TFPSExportFormatSettings read GetSettings write SetSettings;
    {@@ MultipleSheets: export several datasets to multiple sheets in
        the sasme file. Otherwise a single-sheet workbook is created. }
    property MultipleSheets: Boolean read FMultipleSheets write FMultipleSheets default false;
  end;

  { TFPSExport }
  {@@ Export class allowing dataset export to spreadsheet(like) file }
  TFPSExport = Class(TCustomFPSExport)
  published
    {@@ Destination filename }
    property FileName;
    {@@ Source dataset }
    property Dataset;
    {@@ Fields to be exported }
    property ExportFields;
    {@@ Settings - e.g. export format - to be used }
    property FormatSettings;
    {@@ Export starting from current record or beginning. }
    property FromCurrent;
    {@@ Flag indicating whether to return to current dataset position after export }
    property RestorePosition;
    {@@ Procedure to run when exporting a row }
    property OnExportRow;
    {@@ Determines the name of the worksheet }
    property OnGetSheetName;
  end;

{@@ Register export format with fpsdbexport so it can be dynamically used }
procedure RegisterFPSExportFormat;
{@@ Remove registration. Opposite to RegisterFPSExportFormat }
procedure UnRegisterFPSExportFormat;

const
  SFPSExport = 'xls';
  SPFSExtension = '.xls'; //Add others? Doesn't seem to fit other dxport units
  
implementation


{ TCustomFPSExport }

constructor TCustomFPSExport.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  RestorePosition := true;
end;

destructor TCustomFPSExport.Destroy;
begin
  // Last chance to save file if calling WriteExportFile has been forgotten
  // in case of multiple sheets.
  if FMultipleSheets and (FSpreadsheet <> nil) then
  begin
    SaveWorkbook;
    FreeAndNil(FSpreadsheet);
  end;
  inherited;
end;

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
  if FFileName='' then
    Raise EDataExporter.Create(rsExportFileIsRequired);
  if (not RestorePosition) and FMultipleSheets then
    Raise EDataExporter.Create(rsMultipleSheetsOnlyWithRestorePosition);

  if (not FMultipleSheets) or (FSpreadsheet = nil) then
  begin
    FSpreadsheet:=TsWorkbook.Create;
    FSpreadsheet.Options:=FSpreadsheet.Options+[boBufStream];
    // For extra performance. Note that virtual mode is not an option
    // due to the data export determining flow of the program.
  end;
  FSheet:=FSpreadsheet.AddWorksheet(DoGetSheetName);
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

{ Writes the workbook populated during the export process to file }
procedure TCustomFPSExport.SaveWorkbook;
begin
  FRow:=0;
  // Overwrite existing file similar to how dbf export does it
  case Formatsettings.ExportFormat of
    efXLS:
      FSpreadSheet.WriteToFile(FFileName,sfExcel8,true);
    efXLSX:
      FSpreadsheet.WriteToFile(FFilename,sfOOXML,true);
    efODS:
      FSpreadSheet.WriteToFile(FFileName,sfOpenDocument,true);
    efWikiTable:
      FSpreadSheet.WriteToFile(FFileName,sfWikitable_wikimedia,true);
    else
      raise Exception.Create('[TCustomFPSExport.SaveWorkbook] ExportFormat unknown');
  end;
end;

procedure TCustomFPSExport.DoAfterExecute;
begin
  if not FMultipleSheets then
  begin
    SaveWorkbook;
    FreeAndNil(FSpreadsheet);  // Don't free FSheet; done by FSpreadsheet
  end;
  // Multi-sheet workbooks are written when WriteExportFile is called.
  inherited;
end;

procedure TCustomFPSExport.DoDataRowEnd;
begin
  FRow:=FRow+1;
end;

function TCustomFPSExport.CalcSheetNameMask(const AMask: String): String;
begin
  Result := AMask;
  // No %d in the mask string
  if pos('%d', Result) = 0 then
  begin
    // If the mask string is already used we'll add a number to the sheet name
    if not FSpreadsheet.ValidWorksheetName(Result) then
    begin
      Result := AMask + '%d';
      exit;
    end;
  end;
end;

function TCustomFPSExport.CalcUniqueSheetName(const AMask: String): String;
var
  i: Integer;
begin
  if pos('%d', AMask) > 0 then
  begin
    i := 0;
    repeat
      inc(i);
      Result := Format(AMask, [i]);
    until (FSpreadsheet.GetWorksheetByName(Result) = nil);
  end else
    Result := AMask;
  if not FSpreadsheet.ValidWorksheetName(Result) then
    Raise EDataExporter.CreateFmt(rsInvalidWorksheetName, [Result]);
end;

{ Method which provides the name of the worksheet into which the dataset is to
  be exported. There are several cases:
  (1) Use the name defined in the FormatSettings.
  (2) Provide the name in an event handler for OnGetSheetname.
  The name provided from these sources can contain a %d placeholder which will
  be replaced by a number such that the sheet name is unique.
  If it does not contain a %d then a %d may be added if needed to get a unique
  sheet name. }
function TCustomFPSExport.DoGetSheetName: String;
var
  mask: String;
begin
  mask := CalcSheetNameMask(FormatSettings.SheetName);
  Result := CalcUniqueSheetName(mask);
  if Assigned(FOnGetSheetName) then
  begin
    FOnGetSheetName(Self, FSpreadsheet.GetWorksheetCount, mask);
    Result := CalcUniqueSheetName(mask);
  end;
end;

procedure TCustomFPSExport.ExportField(EF: TExportFieldItem);
var
  F : TFPSExportFieldItem;
  dt: TDateTime;
begin
  F := EF as TFPSExportFieldItem;
  with F do
  begin
    // Export depending on field datatype;
    // Fall back to string if unknown datatype
    If Field.IsNull then
      FSheet.WriteBlank(FRow, EF.Index)
    else if Field.Datatype in (IntFieldTypes+[ftAutoInc,ftLargeInt]) then
      FSheet.WriteNumber(FRow, EF.Index,Field.AsInteger)
    else if Field.Datatype in [ftBCD,ftCurrency,ftFloat,ftFMTBcd] then
      FSheet.WriteCurrency(FRow, EF.Index, Field.AsFloat)
    else if Field.DataType in [ftString,ftFixedChar] then
      FSheet.WriteUTF8Text(FRow, EF.Index, Field.AsString)
    else if (Field.DataType in ([ftWideMemo,ftWideString,ftFixedWideChar]+BlobFieldTypes)) then
      FSheet.WriteUTF8Text(FRow, EF.Index, UTF8Encode(Field.AsWideString))
      { Note: we test for the wide text fields before the MemoFieldTypes, in order to
      let ftWideMemo end up at the right place }
    else if Field.DataType in MemoFieldTypes then
      FSheet.WriteUTF8Text(FRow, EF.Index, Field.AsString)
    else if Field.DataType=ftBoolean then
      FSheet.WriteBoolValue(FRow, EF.Index, Field.AsBoolean)
    else if Field.DataType in DateFieldTypes then
      case Field.DataType of
        ftDate: FSheet.WriteDateTime(FRow, EF.Index, Field.AsDateTime, nfShortDate);
        ftTime: FSheet.WriteDateTime(FRow, EF.Index, Field.AsDatetime, nfLongTime);
        else    // try to guess best format if Field.DataType is ftDateTime
                dt := Field.AsDateTime;
                if dt < 1.0 then
                  FSheet.WriteDateTime(FRow, EF.Index, Field.AsDateTime, nfLongTime)
                else if frac(dt) = 0 then
                  FSheet.WriteDateTime(FRow, EF.Index, Field.AsDateTime, nfShortDate)
                else
                  FSheet.WriteDateTime(FRow, EF.Index, Field.AsDateTime, nfShortDateTime);
      end
    else //fallback to string
      FSheet.WriteUTF8Text(FRow, EF.Index, Field.AsString);
  end;
end;

procedure TCustomFPSExport.WriteExportFile;
begin
  if FMultipleSheets then begin
    SaveWorkbook;
    FreeAndNil(FSpreadsheet);
    // Don't free FSheet; done by FSpreadsheet
  end;
end;


procedure RegisterFPSExportFormat;
begin
  ExportFormats.RegisterExportFormat(SFPSExport,rsFPSExportDescription,SPFSExtension,TFPSExport);
end;

procedure UnRegisterFPSExportFormat;
begin
  ExportFormats.UnregisterExportFormat(SFPSExport);
end;

{ TFPSExportFormatSettings }

procedure TFPSExportFormatSettings.Assign(Source: TPersistent);
var
  FS : TFPSExportFormatSettings;
begin
  If Source is TFPSExportFormatSettings then
  begin
    FS:=Source as TFPSExportFormatSettings;
    HeaderRow := FS.HeaderRow;
    ExportFormat := FS.ExportFormat;
    SheetName := FS.SheetName;
  end;
  inherited Assign(Source);
end;

procedure TFPSExportFormatSettings.InitSettings;
begin
  inherited InitSettings;
  FExportFormat := efXLS; //often used format
  FSheetName := 'Sheet';
end;

end.

