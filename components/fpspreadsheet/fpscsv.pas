unit fpscsv;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  fpspreadsheet;

type
  TsCSVReader = class(TsCustomSpreadReader)
  private
    FFormatSettings: TFormatSettings;
    FRow, FCol: Cardinal;
    FCellValue: String;
    FWorksheetName: String;
  protected
    procedure ProcessCellValue(AStream: TStream);
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    procedure ReadFromFile(AFileName: String; AData: TsWorkbook); override;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); override;
    procedure ReadFromStrings(AStrings: TStrings; AData: TsWorkbook); override;
  end;

  TsCSVWriter = class(TsCustomSpreadWriter)
  private
    FFormatSettings: TFormatSettings;

  protected
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); override;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;

    procedure WriteSheet(AStream: TStream; AWorksheet: TsWorksheet);

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    procedure WriteToStream(AStream: TStream); override;
    procedure WriteToStrings(AStrings: TStrings); override;
  end;

  TsCSVParams = record
    LineDelimiter: String;        // LineEnding
    ColDelimiter: Char;           // ';', ',', TAB (#9)
    QuoteChar: Char;              // use #0 if strings are not quoted
    NumberFormat: String;         // if empty, numbers are formatted as in sheet
    DateTimeFormat: String;       // if empty, date/times are formatted as in sheet
    DecimalSeparator: Char;       // '.', ',', #0 if using workbook's formatsetting
    SheetIndex: Integer;          // -1 for all sheets
  end;

var
  CSVParams: TsCSVParams = (
    LineDelimiter: '';            // is replaced by LineEnding at runtime
    ColDelimiter: ';';
    QuoteChar: '"';
    NumberFormat: '';             // Use number format of worksheet
    DateTimeFormat: '';           // Use DateTime format of worksheet
    DecimalSeparator: '.';
    SheetIndex: 0;                // Store sheet #0
  );

implementation

uses
  StrUtils, DateUtils, fpsutils;

{ -----------------------------------------------------------------------------}
{                              TsCSVReader                                     }
{------------------------------------------------------------------------------}
constructor TsCSVReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FFormatSettings := AWorkbook.FormatSettings;
  FWorksheetName := 'Sheet1';
end;

procedure TsCSVReader.ProcessCellValue(AStream: TStream);
begin
  if FCellValue = '' then
    ReadBlank(AStream)
  else
  if (Length(FCellValue) > 1) and (
     ((FCellValue[1] = '"') and (FCellValue[Length(FCellValue)] = '"'))
       or
     (not (CSVParams.QuoteChar in [#0, '"']) and (FCellValue[1] = CSVParams.QuoteChar)
       and (FCellValue[Length(FCellValue)] = CSVParams.QuoteChar))
     ) then
  begin
    Delete(FCellValue, Length(FCellValue), 1);
    Delete(FCellValue, 1, 1);
    ReadLabel(AStream);
  end else
    ReadNumber(AStream);
end;

procedure TsCSVReader.ReadBlank(AStream: TStream);
begin
  // We could write a blank cell, but since CSV does not support formatting
  // this would be a waste of memory. --> Just do nothing
end;

procedure TsCSVReader.ReadFormula(AStream: TStream);
begin
  // Nothing to do - CSV does not support formulas
end;

procedure TsCSVReader.ReadFromFile(AFileName: String; AData: TsWorkbook);
begin
  FWorksheetName := ChangeFileExt(ExtractFileName(AFileName), '');
  inherited;
end;

procedure TsCSVReader.ReadFromStream(AStream: TStream; AData: TsWorkbook);
var
  n: Int64;
  ch: Char;
  nextch: Char;
begin
  FWorkbook := AData;
  FWorksheet := AData.AddWorksheet(FWorksheetName);
  n := AStream.Size;
  FCellValue := '';
  FRow := 0;
  FCol := 0;
  while AStream.Position < n do begin
    ch := char(AStream.ReadByte);
    if ch = CSVParams.ColDelimiter then begin
      // End of column reached
      ProcessCellValue(AStream);
      inc(FCol);
      FCellValue := '';
    end else
    if (ch = #13) or (ch = #10) then begin
      // End of row reached
      ProcessCellValue(AStream);
      inc(FRow);
      FCol := 0;
      FCellValue := '';

      // look for CR+LF: if true, skip next byte
      if AStream.Position+1 < n then begin
        nextch := char(AStream.ReadByte);
        if ((ch = #13) and (nextch <> #10)) then
          AStream.Position := AStream.Position - 1;  // re-read nextchar in next loop
      end;
    end else
      FCellValue := FCellValue + ch;
  end;
end;

procedure TsCSVReader.ReadFromStrings(AStrings: TStrings; AData: TsWorkbook);
var
  stream: TStringStream;
begin
  stream := TStringStream.Create(AStrings.Text);
  try
    ReadFromStream(stream, AData);
  finally
    stream.Free;
  end;
end;

procedure TsCSVReader.ReadLabel(AStream: TStream);
begin
  Unused(AStream);
  FWorksheet.WriteUTF8Text(FRow, FCol, FCellValue);
end;

procedure TsCSVReader.ReadNumber(AStream: TStream);
var
  dbl: Double;
  dt: TDateTime;
  fs: TFormatSettings;
begin
  Unused(AStream);

  // Try as float
  fs := FFormatSettings;
  if CSVParams.DecimalSeparator <> #0 then
    fs.DecimalSeparator := CSVParams.DecimalSeparator;
  if TryStrToFloat(FCellValue, dbl, fs) then
  begin
    FWorksheet.WriteNumber(FRow, FCol, dbl);
    FWorkbook.FormatSettings.DecimalSeparator := fs.DecimalSeparator;
    exit;
  end;
  if fs.DecimalSeparator = '.'
    then fs.DecimalSeparator := ','
    else fs.DecimalSeparator := '.';
  if TryStrToFloat(FCellValue, dbl, fs) then
  begin
    FWorksheet.WriteNumber(FRow, FCol, dbl);
    FWorkbook.FormatSettings.DecimalSeparator := fs.DecimalSeparator;
    exit;
  end;

  // Try as date/time
  fs := FFormatSettings;
  if TryStrToDateTime(FCellValue, dt, fs) then
  begin
    FWorksheet.WriteDateTime(FRow, FCol, dt);
    exit;
  end;

  // Could not convert to float or date/time. Show at least as label.
  FWorksheet.WriteUTF8Text(FRow, FCol, FCellValue);
end;


{ -----------------------------------------------------------------------------}
{                              TsCSVWriter                                     }
{------------------------------------------------------------------------------}
constructor TsCSVWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FFormatSettings := AWorkbook.FormatSettings;
  if CSVParams.DecimalSeparator <> #0 then
    FFormatSettings.DecimalSeparator := CSVParams.DecimalSeparator;
  if CSVParams.LineDelimiter = '' then
    CSVParams.LineDelimiter := LineEnding;
end;

procedure TsCSVWriter.WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  // nothing to do
end;

procedure TsCSVWriter.WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: TDateTime; ACell: PCell);
var
  s: String;
begin
  Unused(ARow, ACol);
  if CSVParams.DateTimeFormat <> '' then
    s := FormatDateTime(CSVParams.DateTimeFormat, AValue, FFormatSettings)
  else
    s := FWorksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream, s);
end;

procedure TsCSVWriter.WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  // no formulas in CSV
  Unused(AStream);
  Unused(ARow, ACol, AStream);
end;

procedure TsCSVWriter.WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: string; ACell: PCell);
var
  s: String;
begin
  Unused(ARow, ACol);
  if ACell = nil then
    exit;
  s := ACell^.UTF8StringValue;
  if CSVParams.QuoteChar <> #0 then
    s := CSVParams.QuoteChar + s + CSVParams.QuoteChar;
  AppendToStream(AStream, s);
end;

procedure TsCSVWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
var
  s: String;
  mask: String;
begin
  Unused(ARow, ACol);
  if ACell = nil then
    exit;
  if CSVParams.NumberFormat <> '' then
    s := Format(CSVParams.NumberFormat, [AValue], FFormatSettings)
  else
    s := FWorksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream, s);
end;

procedure TsCSVWriter.WriteSheet(AStream: TStream; AWorksheet: TsWorksheet);
var
  r, c: Cardinal;
  lastRow, lastCol: Cardinal;
  cell: PCell;
begin
  FWorksheet := AWorksheet;
  lastRow := FWorksheet.GetLastOccupiedRowIndex;
  lastCol := FWorksheet.GetLastOccupiedColIndex;
  for r := 0 to lastRow do
    for c := 0 to lastCol do begin
      cell := FWorksheet.FindCell(r, c);
      if cell <> nil then
        WriteCellCallback(cell, AStream);
      if c = lastCol then
        AppendToStream(AStream, CSVParams.LineDelimiter)
      else
        AppendToStream(AStream, CSVParams.ColDelimiter);
    end;
end;

procedure TsCSVWriter.WriteToStream(AStream: TStream);
var
  n: Integer;
begin
  if (CSVParams.SheetIndex >= 0) and (CSVParams.SheetIndex < FWorkbook.GetWorksheetCount)
    then n := CSVParams.SheetIndex
    else n := 0;
  WriteSheet(AStream, FWorkbook.GetWorksheetByIndex(n));
end;

procedure TsCSVWriter.WriteToStrings(AStrings: TStrings);
var
  stream: TStream;
begin
  stream := TStringStream.Create('');
  try
    WriteToStream(stream);
    stream.Position := 0;
    AStrings.LoadFromStream(stream);
  finally
    stream.Free;
  end;
end;


initialization
  RegisterSpreadFormat(TsCSVReader, TsCSVWriter, sfCSV);

end.

