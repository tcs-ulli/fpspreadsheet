unit fpsNumFormatParser;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  SysUtils, fpspreadsheet;


const
  psOK = 0;
  psErrNoValidColorIndex = 1;
  psErrNoValidCompareNumber = 2;
  psErrUnknownInfoInBrackets = 3;
  psErrConditionalFormattingNotSupported = 4;
  psErrNoUsableFormat = 5;
  psErrNoValidNumberFormat = 6;
  psErrNoValidDateTimeFormat = 7;

{ TsNumFormatParser }

type
  TsCompareOperation = (coNotUsed,
    coEqual, coNotEqual, coLess, coGreater, coLessEqual, coGreaterEqual
  );

  TsConversionDirection = (cdToFPSpreadsheet, cdFromFPSpreadsheet);

  TsNumFormatSection = record
    FormatString: String;
    CompareOperation: TsCompareOperation;
    CompareValue: Double;
    Color: TsColor;
    CountryCode: String;
    CurrencySymbol: String;
    Decimals: Byte;
    NumFormat: TsNumberFormat;
  end;

  TsNumFormatSections = array of TsNumFormatSection;

  TsNumFormatParser = class
  private
    FCreateMethod: Byte;
    FWorkbook: TsWorkbook;
    FCurrent: PChar;
    FStart: PChar;
    FEnd: PChar;
    FCurrSection: Integer;
    FSections: TsNumFormatSections;
    FFormatSettings: TFormatSettings;
    FFormatString: String;
    FNumFormat: TsNumberFormat;
    FConversionDirection: TsConversionDirection;
    FStatus: Integer;
    function GetFormatString: String;
    function GetParsedSectionCount: Integer;
    function GetParsedSections(AIndex: Integer): TsNumFormatSection;

  protected
    procedure AddChar(AChar: Char);
    procedure AddSection;
    procedure AnalyzeBracket(const AValue: String);
    procedure AnalyzeText(const AValue: String);
    procedure CheckSections;
    function CreateFormatStringFromSection(ASection: Integer): String; virtual;
    function CreateFormatStringFromSections: String;
    procedure Parse(const AFormatString: String);
    procedure ScanAMPM(var s: String);
    procedure ScanBrackets;
    procedure ScanDateTime;
    procedure ScanDateTimeParts(TestToken, Replacement: Char; var s: String);
    procedure ScanFormat;
    procedure ScanNumber;
    procedure ScanText;

  public
    constructor Create(AWorkbook: TsWorkbook; const AFormatString: String;
      AConversionDirection: TsConversionDirection = cdToFPSpreadsheet); overload;
    constructor Create(AWorkbook: TsWorkbook; const AFormatSections: TsNumFormatSections;
      AConversionDirection: TsConversionDirection = cdFromFPSpreadsheet); overload;
    destructor Destroy; override;
    procedure CopySections(const FromSections: TsNumFormatSections;
      var ToSections: TsNumFormatSections);
    procedure CopySectionsTo(var ADestination: TsNumFormatSections);
    property Builtin_NumFormat: TsNumberFormat read FNumFormat;
    property FormatString: String read GetFormatString;
    property ParsedSectionCount: Integer read GetParsedSectionCount;
    property ParsedSections[AIndex: Integer]: TsNumFormatSection read GetParsedSections;
    property Status: Integer read FStatus;
  end;

implementation

uses
  fpsutils;

const
  COMPARE_STR: array[TsCompareOperation] of string = (
    '', '=', '<>', '<', '>', '<=', '>'
  );

{ TsNumFormatParser }

{ Creates a number format parser for analyzing a formatstring that has been read
  from a spreadsheet file. The conversion, by default, will go FROM the file TO
  the fpspreadsheet procedures. }
constructor TsNumFormatParser.Create(AWorkbook: TsWorkbook;
  const AFormatString: String; AConversionDirection: TsConversionDirection = cdToFPSpreadsheet);
begin
  inherited Create;
  FCreateMethod := 0;
  FConversionDirection := AConversionDirection;
  FWorkbook := AWorkbook;
  FFormatSettings := DefaultFormatSettings;
  FFormatSettings.DecimalSeparator := '.';
  FFormatSettings.ThousandSeparator := ',';
  Parse(AFormatString);
end;

{ Creates a number format parser to create a format string from the individual
  format sections given in "AFormatSections". It is assumed by default that the
  format string will be written to file. Therefore, it can contain features of
  the destination file format and, in general, will not work if called by
  fpspreadsheet. }
constructor TsNumFormatParser.Create(AWorkbook: TsWorkbook;
  const AFormatSections: TsNumFormatSections;
  AConversionDirection: TsConversionDirection = cdFromFPSpreadsheet);
begin
  inherited Create;
  FCreateMethod := 1;
  FConversionDirection := AConversionDirection;
  FWorkbook := AWorkbook;
  CopySections(AFormatSections, FSections);
end;

destructor TsNumFormatParser.Destroy;
begin
  FSections := nil;
  inherited Destroy;
end;

procedure TsNumFormatParser.AddChar(AChar: Char);
begin
  with FSections[FCurrSection] do
    FormatString := FormatString + AChar;
end;

procedure TsNumFormatParser.AddSection;
begin
  FCurrSection := Length(FSections);
  SetLength(FSections, FCurrSection + 1);
  with FSections[FCurrSection] do begin
    FormatString := '';
    CompareOperation := coNotUsed;
    CompareValue := 0.0;
    Color := scNotDefined;
    CountryCode := '';
    CurrencySymbol := '';
    Decimals := 0;
    NumFormat := nfGeneral;
  end;
end;

procedure TsNumFormatParser.AnalyzeBracket(const AValue: String);
var
  lValue: String;
  n: Integer;
begin
  lValue := lowercase(AValue);
  // Colors
  if lValue = 'red' then
    FSections[FCurrSection].Color := scRed
  else
  if lValue = 'black' then
    FSections[FCurrSection].Color := scBlack
  else
  if lValue = 'blue' then
    FSections[FCurrSection].Color := scBlue
  else
  if lValue = 'white' then
    FSections[FCurrSection].Color := scWhite
  else
  if lValue = 'green' then
    FSections[FCurrSection].Color := scGreen
  else
  if lValue = 'cyan' then
    FSections[FCurrSection].Color := scCyan
  else
  if lValue = 'magenta' then
    FSections[FCurrSection].Color := scMagenta
  else
  if copy(lValue, 1, 5) = 'color' then begin
    lValue := copy(lValue, 6, Length(lValue));
    if not TryStrToInt(trim(lValue), n) then begin
      FStatus := psErrNoValidColorIndex;
      exit;
    end;
    FSections[FCurrSection].Color := n;
  end
  else
  // Conditions
  if lValue[1] in ['=', '<', '>'] then begin
    n := 1;
    case lValue[1] of
      '=': FSections[FCurrSection].CompareOperation := coEqual;
      '<': case lValue[2] of
             '>': begin FSections[FCurrSection].CompareOperation := coNotEqual; inc(n); end;
             '=': begin FSections[FCurrSection].CompareOperation := coLessEqual; inc(n); end;
             else FSections[FCurrSection].CompareOperation := coLess;
           end;
      '>': case lValue[2] of
             '=': begin FSections[FCurrSection].CompareOperation := coGreaterEqual; inc(n); end;
             else FSections[FCurrSection].CompareOperation := coGreater;
           end;
    end;
    Delete(lValue, 1, n);
    if not TryStrToFloat(trim(lValue), FSections[FCurrSection].CompareValue) then
      FStatus := psErrNoValidCompareNumber;
  end else
  // Locale information
  if lValue[1] = '$' then begin
    FSections[FCurrSection].CountryCode := Copy(AValue, 2, Length(AValue));
  end else
    FStatus := psErrUnknownInfoInBrackets;
end;

procedure TsNumFormatParser.AnalyzeText(const AValue: String);
var
  uValue: String;
begin
  uValue := Uppercase(AValue);
  if (uValue = '$') or (uValue = 'USD') or (uValue = '€') or (uValue = 'EUR') or
     (uValue = '£') or (uValue = 'GBP') or (uValue = '¥') or (uValue = 'JPY')
  then
    FSections[FCurrSection].CurrencySymbol := AValue;
end;

procedure TsNumFormatParser.CheckSections;
var
  i: Integer;
  ns: Integer;
  s: String;
begin
  ns := Length(FSections);

  for i:=0 to ns-1 do begin
    if FSections[i].FormatString = '' then
      FSections[i].NumFormat := nfGeneral;

    if (FSections[i].CurrencySymbol <> '') and (FSections[i].NumFormat = nfFixedTh) then
      FSections[i].NumFormat := nfCurrency;

    if FSections[i].CompareOperation <> coNotUsed then begin
      FStatus := psErrConditionalFormattingNotSupported;
      exit;
    end;

    // Check format strings
    case FSections[i].NumFormat of
      nfGeneral, nfFixed, nfFixedTh, nfPercentage, nfExp, nfSci, nfCurrency:
        try
          s := FormatFloat(FSections[i].FormatString, 1.0, FWorkBook.FormatSettings);
        except
          FStatus := psErrNoValidNumberFormat;
          exit;
        end;

      nfShortDateTime, nfShortDate, nfShortTime, nfShortTimeAM,
      nfLongDate, nfLongTime, nfLongTimeAM, nfFmtDateTime:
        try
          s := FormatDateTimeEx(FSections[i].FormatString, now(), FWorkbook.FormatSettings);
        except
          FStatus := psErrNoValidDateTimeFormat;
          exit;
        end;
    end;
  end;

  // Extract built-in NumFormat identifier for currency (needs several entries in
  // three sections).
  if (ns = 3) and
     (FSections[0].NumFormat = nfCurrency) and
     (FSections[1].NumFormat = nfCurrency) and
     (FSections[2].NumFormat = nfCurrency)
  then begin
    if ((FSections[2].FormatString = '-') or (FSections[2].FormatString = '"-"')) then begin
      if (FSections[1].Color = scRed) then
        FNumFormat := nfCurrencyDashRed
      else
        FNumFormat := nfCurrencyDash;
    end else begin
      if (FSections[1].Color = scRed) then
        FNumFormat := nfCurrencyRed;
    end;
  end else
  // If there are other multi-section formatstrings they must be a custom format
  if (ns > 1) then
    FNumFormat := nfCustom
  else
    FNumFormat := FSections[0].NumFormat;

  if ns = 2 then
    FFormatString := Format('%s;%s;%s', [
      FSections[0].FormatString,
      FSections[1].FormatString,
      FSections[0].FormatString  // make sure that fpc understands the "zero"
    ])
  else
  if ns > 0 then begin
    FFormatString := FSections[0].FormatString;
    for i:=1 to ns-1 do
      FFormatString := Format('%s;%s', [FFormatString, FSections[i].FormatString]);
  end else
    FStatus := psErrNoUsableFormat;
end;

procedure TsNumFormatParser.CopySections(
  const FromSections: TsNumFormatSections; var ToSections: TsNumformatSections);
var
  i: Integer;
begin
  SetLength(ToSections, Length(FromSections));
  for i:= 0 to High(FromSections) do begin
    ToSections[i].FormatString := FromSections[i].FormatString;
    ToSections[i].CompareOperation := FromSections[i].CompareOperation;
    ToSections[i].CompareValue := FromSections[i].CompareValue;
    ToSections[i].Color := FromSections[i].Color;
    ToSections[i].CurrencySymbol := FromSections[i].CurrencySymbol;
    ToSections[i].Decimals := FromSections[i].Decimals;
    ToSections[i].NumFormat := FromSections[i].NumFormat;
  end;
end;

procedure TsNumFormatParser.CopySectionsTo(var ADestination: TsNumFormatSections);
begin
  CopySections(FSections, ADestination);
end;

function TsNumFormatParser.CreateFormatStringFromSections: String;
var
  i: Integer;
begin
  if Length(FSections) = 0 then
    Result := ''
  else begin
    Result := CreateFormatStringFromSection(0);
    for i:=1 to High(FSections) do
      Result := Result + ';' + CreateFormatStringFromSection(i);
  end;
end;

function TsNumFormatParser.CreateFormatStringFromSection(ASection: Integer): String;
begin
  with FSections[ASection] do
    if (NumFormat = nfFmtDateTime) or (NumFormat = nfCustom) then begin
      Result := FormatString;
      exit;
    end;

  Result := BuildNumberFormatString(FSections[ASection].NumFormat,
    FWorkbook.FormatSettings,
    FSections[ASection].Decimals,
    FSections[ASection].CurrencySymbol
  );
  if FConversionDirection = cdFromFPSpreadsheet then begin
    // This is typical of Excel, but is valid for all others as well.
    // Override if you need to change
    if FSections[ASection].Color < 8 then
      Result := Format('[%s]%s', [FWorkbook.GetColorName(FSections[ASection].Color), Result])
    else
    if FSections[ASection].Color < scNotDefined then
      Result := Format('[Color%d]%s', [FSections[ASection].Color, Result]);

    if FSections[ASection].CompareOperation <> coNotUsed then
      Result := Format('[%s%g]%s', [
        COMPARE_STR[FSections[ASection].CompareOperation],
        FSections[ASection].CompareValue,
        Result
      ]);
  end;
end;

function TsNumFormatParser.GetFormatString: String;
begin
  case FCreateMethod of
    0: Result := FFormatString;
    1: Result := CreateFormatStringFromSections;
  end;
end;

function TsNumFormatParser.GetParsedSectionCount: Integer;
begin
  Result := Length(FSections);
end;

function TsNumFormatParser.GetParsedSections(AIndex: Integer): TsNumFormatSection;
begin
  Result := FSections[AIndex];
end;

procedure TsNumFormatParser.Parse(const AFormatString: String);
var
  token: Char;
begin
  FStatus := psOK;
  AddSection;
  FStart := @AFormatString[1];
  FEnd := FStart + Length(AFormatString) - 1;
  FCurrent := FStart;
  while (FCurrent <= FEnd) and (FStatus = psOK) do begin
    token := FCurrent^;
    case token of
      '[': ScanBrackets;
      ';': AddSection;
      else ScanFormat;
    end;
    inc(FCurrent);
  end;
  CheckSections;
end;

{ Extracts the text between square brackets --> AnalyzeBracket }
procedure TsNumFormatParser.ScanBrackets;
var
  s: String;
  token: Char;
begin
  inc(FCurrent);  // cursor stands at '['
  while (FCurrent <= FEnd) and (FStatus = psOK) do begin
    token := FCurrent^;
    case token of
      ']': begin
             AnalyzeBracket(s);
             break;
           end;
      else
           s := s + token;
    end;
    inc(FCurrent);
  end;
end;

procedure TsNumFormatParser.ScanDateTime;
var
  token: Char;
  done: Boolean;
  s: String;
  i: Integer;
  nf: TsNumberFormat;
  partStr: String;
  isTime: Boolean;
  isAMPM: Boolean;
begin
  done := false;
  s := '';
  isTime := false;
  isAMPM := false;

  while (FCurrent <= FEnd) and (FStatus = psOK) and (not done) do begin
    token := FCurrent^;
    case token of
      '\':
        begin
          inc(FCurrent);
          token := FCurrent^;
          s := s + token;
        end;
      'Y', 'y':
        begin
          ScanDateTimeParts(token, token, s);
          isTime := false;
        end;
      'M', 'm':
        ScanDateTimeParts(token, token, s);
        {if isTime then    // help fpc to separate "month" and "minute"
           ScanDateTimeParts(token, 'n', s)
         else   // both "month" and "minute" work in fpc to some degree
           ScanDateTimeParts(token, token, s);}
      'N', 'n':
        ScanDateTimeParts(token, 'n', s);  // fpc dialect for "minutes"
      'D', 'd':
        begin
          ScanDateTimeParts(token, token, s);
          isTime := false;
        end;
      'H', 'h':
        begin
          ScanDateTimeParts(token, token, s);
          isTime := true;
        end;
      'S', 's':
        begin
          ScanDateTimeParts(token, token, s);
          isTime := true;
        end;
      '/', ':', '.', ']', '[', ' ':
        s := s + token;
      '0', 'z', 'Z':
        ScanDateTimeParts(token, token, s);
      'A', 'a':
        begin
          ScanAMPM(s);
          isAMPM := true;
        end;
      else
        begin
          done := true;
          dec(FCurrent);
          // char pointer must be at end of date/time mask.
        end;
    end;
    if not done then inc(FCurrent);
  end;

  FSections[FCurrSection].FormatString := FSections[FCurrSection].FormatString + s;
  s := FSections[FCurrSection].FormatString;
  if s <> '' then begin
    if s = FWorkbook.FormatSettings.LongDateFormat then
      nf := nfLongDate
    else
    if s = FWorkbook.FormatSettings.ShortDateFormat then
      nf := nfShortDate
    else
    if s = StripAMPM(FWorkbook.FormatSettings.LongTimeFormat) then
      nf := IfThen(isAMPM, nfLongTimeAM, nfLongTime)
    else
    if s = StripAMPM(FWorkbook.FormatSettings.ShortTimeFormat) then
      nf := IfThen(isAMPM, nfShortTimeAM, nfShortTime)
    else
      nf := nfFmtDateTime;

    FSections[FCurrSection].NumFormat := nf;
  end;
end;

procedure TsNumFormatParser.ScanAMPM(var s: String);
var
  token: Char;
begin
  while (FCurrent <= FEnd) do begin
    token := FCurrent^;
    if token in ['A', 'a', 'P', 'p', 'm', 'M', '/'] then
      s := s + token
    else begin
      dec(FCurrent);
      exit;
    end;
    inc(FCurrent);
  end;
end;

procedure TsNumFormatParser.ScanDateTimeParts(TestToken, Replacement: Char;
  var s: String);
var
  token: Char;
begin
  s := s + Replacement;
  while (FCurrent <= FEnd) do begin
    inc(FCurrent);
    token := FCurrent^;
    if token = TestToken then
      s := s + Replacement
    else begin
      dec(FCurrent);
      break;
    end;
  end;
end;

procedure TsNumFormatParser.ScanFormat;
var
  token: Char;
  done: Boolean;
begin
  done := false;
  while (FCurrent <= FEnd) and (FStatus = psOK) and (not done) do begin
    token := FCurrent^;
    case token of
      // Strip Excel's formatting symbols
      '\', '*':
        ;
      '_':
        inc(FCurrent);
      '"':
        begin
          inc(FCurrent);
          ScanText;
        end;
      '0', '#', '.', ',', '-':
        ScanNumber;
      'y', 'Y', 'm', 'M',  'd', 'D', 'h', 'N', 'n', 's', '[':
        ScanDateTime;
      ' ':
        AddChar(token);
      ';':
        begin
          done := true;
          dec(FCurrent);
          // Cursor must stay on the ";"
        end;
    end;
    if not done then inc(FCurrent);
  end;
end;

procedure TsNumFormatParser.ScanNumber;
var
  token: Char;
  done: Boolean;
  countdecs: Boolean;
  s: String;
  hasThSep: Boolean;
  isExp: Boolean;
  isSci: Boolean;
  hasHash: Boolean;
  hasPerc: Boolean;
  nf: TsNumberFormat;
begin
  countdecs := false;
  done := false;
  hasThSep := false;
  hasHash := false;
  hasPerc := false;
  isExp := false;
  isSci := false;
  s := '';
  while (FCurrent <= FEnd) and (FStatus = psOK) and (not done) do begin
    token := FCurrent^;
    case token of
      ',': begin
             hasThSep := true;
             s := s + token;
           end;
      '.': begin
             countdecs := true;
             FSections[FCurrSection].Decimals := 0;
             s := s + token;
           end;
      '0': begin
             if countdecs then inc(FSections[FCurrSection].Decimals);
             s := s + token;
           end;
      'E', 'e':
           begin
             if hasHash and countdecs then isSci := true else isExp := true;
             countdecs := false;
             s := s + token;
           end;
      '+', '-':
           s := s + token;
      '#': begin
             hasHash := true;
             countdecs := false;
             s := s + token;
           end;
      '%': begin
             hasPerc := true;
             s := s + token;
           end;
      else begin
             done := true;
             dec(FCurrent);
           end;
    end;
    if not done then
      inc(FCurrent);
  end;

  if s <> '' then begin
    if isExp then
      nf := nfExp
    else if isSci then
      nf := nfSci
    else if hasPerc then
      nf := nfPercentage
    else if hasThSep then
      nf := nfFixedTh
    else
      nf := nfFixed;
  end else
    nf := nfGeneral;

  FSections[FCurrSection].NumFormat := nf;
  FSections[FCurrSection].FormatString := FSections[FCurrSection].FormatString + s;
end;

{ Scans a text in quotation marks. Tries to interpret the text as a currency
  symbol (--> AnalyzeText) }
procedure TsNumFormatParser.ScanText;
var
  token: Char;
  done: Boolean;
  s: String;
begin
  done := false;
  s := '';
  while (FCurrent <= FEnd) and (FStatus = psOK) and not done do begin
    token := FCurrent^;
    if token = '"' then begin
      done := true;
      AnalyzeText(s);
    end else begin
      s := s + token;
      inc(FCurrent);
    end;
  end;
  FSections[FCurrSection].FormatString := Format('%s"%s"',
    [FSections[FCurrSection].FormatString, s]);
end;

end.
