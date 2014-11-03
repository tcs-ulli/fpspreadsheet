unit fpscsv;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  fpspreadsheet, fpsCsvDocument;

type
  TsCSVReader = class(TsCustomSpreadReader)
  private
    FWorksheetName: String;
    FFormatSettings: TFormatSettings;
    function IsBool(AText: String; out AValue: Boolean): Boolean;
    function IsDateTime(AText: String; out ADateTime: TDateTime;
      out ANumFormat: TsNumberFormat): Boolean;
    function IsNumber(AText: String; out ANumber: Double; out ANumFormat: TsNumberFormat;
      out ADecimals: Integer; out ACurrencySymbol, AWarning: String): Boolean;
    function IsQuotedText(var AText: String): Boolean;
    procedure ReadCellValue(ARow, ACol: Cardinal; AText: String);
  protected
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
    FCSVBuilder: TCSVBuilder;
    FEncoding: String;
    FFormatSettings: TFormatSettings;
  protected
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
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

  TsCSVLineEnding = (leSystem, leCRLF, leCR, leLF);

  TsCSVParams = record   // W = writing, R = reading, RW = reading/writing
    SheetIndex: Integer;             // W: Index of the sheet to be written
    LineEnding: TsCSVLineEnding;     // W: Specification for line ending to be written
    Delimiter: Char;                 // RW: Column delimiter
    QuoteChar: Char;                 // RW: Character for quoting texts
    Encoding: String;                // RW: Encoding of file
    DetectContentType: Boolean;      // R: try to convert strings to content types
    NumberFormat: String;            // W: if empty write numbers like in sheet, otherwise use this format
    AutoDetectNumberFormat: Boolean; // R: automatically detects decimal/thousand separator used in numbers
    TrueText: String;                // RW: String for boolean TRUE
    FalseText: String;               // RW: String for boolean FALSE
    FormatSettings: TFormatSettings; // RW: add'l parameters for conversion
  end;

var
  CSVParams: TsCSVParams = (
    SheetIndex: 0;
    LineEnding: leSystem;
    Delimiter: ';';
    QuoteChar: '"';
    Encoding: '';    // '' = auto-detect when reading, UTF8 when writing
    DetectContentType: true;
    NumberFormat: '';
    AutoDetectNumberFormat: true;
    TrueText: 'TRUE';
    FalseText: 'FALSE';
  {%H-});

function LineEndingAsString(ALineEnding: TsCSVLineEnding): String;


implementation

uses
  //StrUtils,
  DateUtils, LConvEncoding, Math, fpsutils, fpscurrency;

{ Initializes the FormatSettings of the CSVParams to default values which
  can be replaced by the FormatSettings of the workbook's FormatSettings }
procedure InitCSVFormatSettings;
var
  i: Integer;
begin
  with CSVParams.FormatSettings do
  begin
    CurrencyFormat := Byte(-1);
    NegCurrFormat := Byte(-1);
    ThousandSeparator := #0;
    DecimalSeparator := #0;
    CurrencyDecimals := Byte(-1);
    DateSeparator := #0;
    TimeSeparator := #0;
    ListSeparator := #0;
    CurrencyString := '';
    ShortDateFormat := '';
    LongDateFormat := '';
    TimeAMString := '';
    TimePMString := '';
    ShortTimeFormat := '';
    LongTimeFormat := '';
    for i:=1 to 12 do
    begin
      ShortMonthNames[i] := '';
      LongMonthNames[i] := '';
    end;
    for i:=1 to 7 do
    begin
      ShortDayNames[i] := '';
      LongDayNames[i] := '';
    end;
    TwoDigitYearCenturyWindow := Word(-1);
  end;
end;

procedure ReplaceFormatSettings(var AFormatSettings: TFormatSettings;
  const ADefaultFormats: TFormatSettings);
var
  i: Integer;
begin
  if AFormatSettings.CurrencyFormat = Byte(-1) then
    AFormatSettings.CurrencyFormat := ADefaultFormats.CurrencyFormat;
  if AFormatSettings.NegCurrFormat = Byte(-1) then
    AFormatSettings.NegCurrFormat := ADefaultFormats.NegCurrFormat;
  if AFormatSettings.ThousandSeparator = #0 then
    AFormatSettings.ThousandSeparator := ADefaultFormats.ThousandSeparator;
  if AFormatSettings.DecimalSeparator = #0 then
    AFormatSettings.DecimalSeparator := ADefaultFormats.DecimalSeparator;
  if AFormatSettings.CurrencyDecimals = Byte(-1) then
    AFormatSettings.CurrencyDecimals := ADefaultFormats.CurrencyDecimals;
  if AFormatSettings.DateSeparator = #0 then
    AFormatSettings.DateSeparator := ADefaultFormats.DateSeparator;
  if AFormatSettings.TimeSeparator = #0 then
    AFormatSettings.TimeSeparator := ADefaultFormats.TimeSeparator;
  if AFormatSettings.ListSeparator = #0 then
    AFormatSettings.ListSeparator := ADefaultFormats.ListSeparator;
  if AFormatSettings.CurrencyString = '' then
    AFormatSettings.CurrencyString := ADefaultFormats.CurrencyString;
  if AFormatSettings.ShortDateFormat = '' then
    AFormatSettings.ShortDateFormat := ADefaultFormats.ShortDateFormat;
  if AFormatSettings.LongDateFormat = '' then
    AFormatSettings.LongDateFormat := ADefaultFormats.LongDateFormat;
  if AFormatSettings.ShortTimeFormat = '' then
    AFormatSettings.ShortTimeFormat := ADefaultFormats.ShortTimeFormat;
  if AFormatSettings.LongTimeFormat = '' then
    AFormatSettings.LongTimeFormat := ADefaultFormats.LongTimeFormat;
  for i:=1 to 12 do
  begin
    if AFormatSettings.ShortMonthNames[i] = '' then
      AFormatSettings.ShortMonthNames[i] := ADefaultFormats.ShortMonthNames[i];
    if AFormatSettings.LongMonthNames[i] = '' then
      AFormatSettings.LongMonthNames[i] := ADefaultFormats.LongMonthNames[i];
  end;
  for i:=1 to 7 do
  begin
    if AFormatSettings.ShortDayNames[i] = '' then
      AFormatSettings.ShortDayNames[i] := ADefaultFormats.ShortDayNames[i];
    if AFormatSettings.LongDayNames[i] = '' then
      AFormatSettings.LongDayNames[i] := ADefaultFormats.LongDayNames[i];
  end;
  if AFormatSettings.TwoDigitYearCenturyWindow = Word(-1) then
    AFormatSettings.TwoDigitYearCenturyWindow := ADefaultFormats.TwoDigitYearCenturyWindow;
end;

function LineEndingAsString(ALineEnding: TsCSVLineEnding): String;
begin
  case ALineEnding of
    leSystem: Result := LineEnding;
    leCR    : Result := #13;
    leLF    : Result := #10;
    leCRLF  : Result := #13#10;
  end;
end;


{ -----------------------------------------------------------------------------}
{                              TsCSVReader                                     }
{------------------------------------------------------------------------------}

constructor TsCSVReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FFormatSettings := CSVParams.FormatSettings;
  ReplaceFormatSettings(FFormatSettings, AWorkbook.FormatSettings);
  FWorksheetName := 'Sheet1';  // will be replaced by filename
end;

function TsCSVReader.IsBool(AText: String; out AValue: Boolean): Boolean;
begin
  if SameText(AText, CSVParams.TrueText) then
  begin
    AValue := true;
    Result := true;
  end else
  if SameText(AText, CSVParams.FalseText) then
  begin
    AValue := false;
    Result := true;
  end else
    Result := false;
end;

function TsCSVReader.IsDateTime(AText: String; out ADateTime: TDateTime;
  out ANumFormat: TsNumberFormat): Boolean;

  { Test whether the text is formatted according to a built-in date/time format.
    Converts the obtained date/time value back to a string and compares. }
  function TestFormat(lNumFmt: TsNumberFormat): Boolean;
  var
    fmt: string;
  begin
    fmt := BuildDateTimeFormatString(lNumFmt, FFormatSettings);
    Result := FormatDateTime(fmt, ADateTime, FFormatSettings) = AText;
    if Result then ANumFormat := lNumFmt;
  end;

begin
  Result := TryStrToDateTime(AText, ADateTime, FFormatSettings);
  if Result then
  begin
    ANumFormat := nfCustom;
    if abs(ADateTime) > 1 then       // this is most probably a date
    begin
      if TestFormat(nfShortDateTime) then
        exit;
      if TestFormat(nfLongDate) then
        exit;
      if TestFormat(nfShortDate) then
        exit;
    end else
    begin          // this case is time-only
      if TestFormat(nfLongTimeAM) then
        exit;
      if TestFormat(nfLongTime) then
        exit;
      if TestFormat(nfShortTimeAM) then
        exit;
      if TestFormat(nfShortTime) then
        exit;
    end;
  end;
end;

function TsCSVReader.IsNumber(AText: String; out ANumber: Double;
  out ANumFormat: TsNumberFormat; out ADecimals: Integer;
  out ACurrencySymbol, AWarning: String): Boolean;
var
  p: Integer;
  DecSep, ThousSep: Char;
begin
  Result := false;
  AWarning := '';

  // To detect whether the text is a currency value we look for the currency
  // string. If we find it, we delete it and convert the remaining string to
  // a number.
  ACurrencySymbol := FFormatSettings.CurrencyString;
  if RemoveCurrencySymbol(ACurrencySymbol, AText) then
  begin
    if IsNegative(AText) then
    begin
      if AText = '' then
        exit;
      AText := '-' + AText;
    end;
  end else
    ACurrencySymbol := '';

  if CSVParams.AutoDetectNumberFormat then
    Result := TryStrToFloatAuto(AText, ANumber, DecSep, ThousSep, AWarning)
  else begin
    Result := TryStrToFloat(AText, ANumber, FFormatSettings);
    if Result then
    begin
      if pos(FFormatSettings.DecimalSeparator, AText) = 0
        then DecSep := #0
        else DecSep := FFormatSettings.DecimalSeparator;
      if pos(CSVParams.FormatSettings.ThousandSeparator, AText) = 0
        then ThousSep := #0
        else ThousSep := FFormatSettings.ThousandSeparator;
    end;
  end;

  // Try to determine the number format
  if Result then
  begin
    if ThousSep <> #0 then
      ANumFormat := nfFixedTh
    else
      ANumFormat := nfGeneral;
    // count number of decimal places and try to catch special formats
    ADecimals := 0;
    if DecSep <> #0 then
    begin
      // Go to the decimal separator and search towards the end of the string
      p := pos(DecSep, AText) + 1;
      while (p <= Length(AText)) do begin
        // exponential format
        if AText[p] in ['+', '-', 'E', 'e'] then
        begin
          ANumFormat := nfExp;
          break;
        end else
        // percent format
        if AText[p] = '%' then
        begin
          ANumFormat := nfPercentage;
          break;
        end else
        begin
          inc(p);
          inc(ADecimals);
        end;
      end;
      if (ADecimals > 0) and (ADecimals < 9) and (ANumFormat = nfGeneral) then
        // "no formatting" assumed if there are "many" decimals
        ANumFormat := nfFixed;
    end else
    begin
      p := Length(AText);
      while (p > 0) do begin
        case AText[p] of
          '%'     : ANumFormat := nfPercentage;
          'e', 'E': ANumFormat := nfExp;
          else      dec(p);
        end;
        break;
      end;
    end;
  end else
    ACurrencySymbol := '';
end;

{ Checks if text is quoted; strips any starting and ending quotes }
function TsCSVReader.IsQuotedText(var AText: String): Boolean;
begin
  if (Length(AText) > 1) and (CSVParams.QuoteChar <> #0) and
   (AText[1] = CSVParams.QuoteChar) and
   (AText[Length(AText)] = CSVParams.QuoteChar) then
  begin
    Delete(AText, 1, 1);
    Delete(AText, Length(AText), 1);
    Result := true;
  end else
    Result := false;
end;

procedure TsCSVReader.ReadBlank(AStream: TStream);
begin
  Unused(AStream);
end;

{ Determines content types from/for the text read from the csv file and writes
  the corresponding data to the worksheet. }
procedure TsCSVReader.ReadCellValue(ARow, ACol: Cardinal; AText: String);
var
  dblValue: Double;
  dtValue: TDateTime;
  boolValue: Boolean;
  currSym: string;
  warning: String;
  nf: TsNumberFormat;
  decs: Integer;
begin
  // Empty strings are blank cells -- nothing to do
  if AText = '' then
    exit;

  // Do not try to interpret the strings. --> everything is a LABEL cell.
  if not CSVParams.DetectContentType then
  begin
    FWorksheet.WriteUTF8Text(ARow, aCol, AText);
    exit;
  end;

  // Check for a NUMBER or CURRENCY cell
  if IsNumber(AText, dblValue, nf, decs, currSym, warning) then
  begin
    if currSym <> '' then
      FWorksheet.WriteCurrency(ARow, ACol, dblValue, nfCurrency, decs, currSym)
    else
      FWorksheet.WriteNumber(ARow, ACol, dblValue, nf, decs);
    if warning <> '' then
      FWorkbook.AddErrorMsg('Cell %s: %s', [GetCellString(ARow, ACol), warning]);
    exit;
  end;

  // Check for a DATE/TIME cell
  // No idea how to apply the date/time formatsettings here...
  if IsDateTime(AText, dtValue, nf) then
  begin
    FWorksheet.WriteDateTime(ARow, ACol, dtValue, nf);
    exit;
  end;

  // Check for a BOOLEAN cell
  if IsBool(AText, boolValue) then
  begin
    FWorksheet.WriteBoolValue(ARow, aCol, boolValue);
    exit;
  end;

  // What is left is handled as a TEXT cell
  FWorksheet.WriteUTF8Text(ARow, ACol, AText);
end;

procedure TsCSVReader.ReadFormula(AStream: TStream);
begin
  Unused(AStream);
end;

procedure TsCSVReader.ReadFromFile(AFileName: String; AData: TsWorkbook);
begin
  FWorksheetName := ChangeFileExt(ExtractFileName(AFileName), '');
  inherited;
end;

procedure TsCSVReader.ReadFromStream(AStream: TStream; AData: TsWorkbook);
var
  Parser: TCSVParser;
  s: String;
  encoding: String;
begin
  // Try to determine encoding of the input file
  SetLength(s, Min(1000, AStream.Size));
  AStream.ReadBuffer(s[1], Length(s));
  if CSVParams.Encoding = '' then
    encoding := GuessEncoding(s)
  else
    encoding := CSVParams.Encoding;

  // Store workbook for internal use, and create worksheet
  FWorkbook := AData;
  FWorksheet := AData.AddWorksheet(FWorksheetName, true);

  // Create csv parser, read file and store in worksheet
  Parser := TCSVParser.Create;
  try
    Parser.Delimiter := CSVParams.Delimiter;
    Parser.LineEnding := LineEndingAsString(CSVParams.LineEnding);
    Parser.QuoteChar := CSVParams.QuoteChar;
    // Indicate column counts between rows may differ:
    Parser.EqualColCountPerRow := false;
    Parser.SetSource(AStream);
    while Parser.ParseNextCell do begin
      // Convert string to UTF8
      s := Parser.CurrentCellText;
      s := ConvertEncoding(s, encoding, EncodingUTF8);
      ReadCellValue(Parser.CurrentRow, Parser.CurrentCol, s);
    end;
  finally
    Parser.Free;
  end;
end;

procedure TsCSVReader.ReadFromStrings(AStrings: TStrings; AData: TsWorkbook);
var
  Stream: TStringStream;
begin
  Stream := TStringStream.Create(AStrings.Text);
  try
    ReadFromStream(Stream, AData);
  finally
    Stream.Free;
  end;
end;

procedure TsCSVReader.ReadLabel(AStream: TStream);
begin
  Unused(AStream);
end;

procedure TsCSVReader.ReadNumber(AStream: TStream);
begin
  Unused(AStream);
end;


{ -----------------------------------------------------------------------------}
{                              TsCSVWriter                                     }
{------------------------------------------------------------------------------}

constructor TsCSVWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FFormatSettings := CSVParams.FormatSettings;
  ReplaceFormatSettings(FFormatSettings, FWorkbook.FormatSettings);
  if CSVParams.Encoding = '' then
    FEncoding := 'utf8'
  else
    FEncoding := CSVParams.Encoding;
end;

procedure TsCSVWriter.WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  // nothing to do
end;

{ Write boolean cell to stream formatted as string }
procedure TsCSVWriter.WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: Boolean; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  if AValue then
    s := CSVParams.TrueText
  else
    s := CSVParams.FalseText;
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  FCSVBuilder.AppendCell(s);
end;

{ Write date/time values in the same way they are displayed in the sheet }
procedure TsCSVWriter.WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: TDateTime; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol, AValue);
  s := FWorksheet.ReadAsUTF8Text(ACell);
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  FCSVBuilder.AppendCell(s);
end;

{ CSV does not support formulas, but we can write the formula results to
  to stream. }
procedure TsCSVWriter.WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  if ACell = nil then
    exit;
  case ACell^.ContentType of
    cctBool      : WriteBool(AStream, ARow, ACol, ACell^.BoolValue, ACell);
    cctEmpty     : ;
    cctDateTime  : WriteDateTime(AStream, ARow, ACol, ACell^.DateTimeValue, ACell);
    cctNumber    : WriteNumber(AStream, ARow, ACol, ACell^.NumberValue, ACell);
    cctUTF8String: WriteLabel(AStream, ARow, ACol, ACell^.UTF8StringValue, ACell);
    cctError     : ;
  end;
end;

{ Writes a LABEL cell to the stream. }
procedure TsCSVWriter.WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: string; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol, AValue);
  if ACell = nil then
    exit;
  s := ACell^.UTF8StringValue;
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  // No need to quote; csvdocument will do that for us...
  FCSVBuilder.AppendCell(s);
end;

{ Writes a number cell to the stream. }
procedure TsCSVWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol);
  if ACell = nil then
    exit;
  if CSVParams.NumberFormat <> '' then
    s := Format(CSVParams.NumberFormat, [AValue], FFormatSettings)
  else
    s := FWorksheet.ReadAsUTF8Text(ACell, FFormatSettings);
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  FCSVBuilder.AppendCell(s);
end;

procedure TsCSVWriter.WriteSheet(AStream: TStream; AWorksheet: TsWorksheet);
var
  r, c: Cardinal;
  LastRow, LastCol: Cardinal;
  Cell: PCell;
begin
  FWorksheet := AWorksheet;

  FCSVBuilder := TCSVBuilder.Create;
  try
    FCSVBuilder.Delimiter := CSVParams.Delimiter;
    FCSVBuilder.LineEnding := LineEndingAsString(CSVParams.LineEnding);
    FCSVBuilder.QuoteChar := CSVParams.QuoteChar;
    FCSVBuilder.SetOutput(AStream);

    LastRow := FWorksheet.GetLastOccupiedRowIndex;
    LastCol := FWorksheet.GetLastOccupiedColIndex;
    for r := 0 to LastRow do
      for c := 0 to LastCol do
      begin
        Cell := FWorksheet.FindCell(r, c);
        if Cell <> nil then
          WriteCellCallback(Cell, AStream);
        if c = LastCol then
          FCSVBuilder.AppendRow;
      end;
  finally
    FreeAndNil(FCSVBuilder);
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
  Stream: TStream;
begin
  Stream := TStringStream.Create('');
  try
    WriteToStream(Stream);
    Stream.Position := 0;
    AStrings.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;


initialization
  InitCSVFormatSettings;
  RegisterSpreadFormat(TsCSVReader, TsCSVWriter, sfCSV);

end.

