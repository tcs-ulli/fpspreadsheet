unit fpsNumFormatParser;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

uses
  Classes, SysUtils, fpstypes, fpspreadsheet;


const
  psOK = 0;
  psErrNoValidColorIndex = 1;
  psErrNoValidCompareNumber = 2;
  psErrUnknownInfoInBrackets = 3;
  psErrConditionalFormattingNotSupported = 4;
  psErrNoUsableFormat = 5;
  psErrNoValidNumberFormat = 6;
  psErrNoValidDateTimeFormat = 7;
  psErrQuoteExpected = 8;
  psErrMultipleCurrSymbols = 9;
  psErrMultipleFracSymbols = 10;
  psErrMultipleExpChars = 11;
  psAmbiguousSymbol = 12;

type

  { TsNumFormatParser }

  TsNumFormatParser = class
  private
    FToken: Char;
    FCurrent: PChar;
    FStart: PChar;
    FEnd: PChar;
    FCurrSection: Integer;
    FStatus: Integer;
    function GetCurrencySymbol: String;
    function GetDecimals: byte;
    function GetFracDenominator: Integer;
    function GetFracInt: Integer;
    function GetFracNumerator: Integer;
    function GetFormatString: String;
    function GetNumFormat: TsNumberFormat;
    function GetParsedSectionCount: Integer;
    function GetParsedSections(AIndex: Integer): TsNumFormatSection;
    procedure SetDecimals(AValue: Byte);

  protected
    FWorkbook: TsWorkbook;
    FSections: TsNumFormatSections;

    { Administration while scanning }
    procedure AddElement(AToken: TsNumFormatToken; AText: String); overload;
    procedure AddElement(AToken: TsNumFormatToken; AIntValue: Integer); overload;
    procedure AddElement(AToken: TsNumFormatToken; AFloatValue: Double); overload;
    procedure AddSection;
    procedure DeleteElement(ASection, AIndex: Integer);
    procedure InsertElement(ASection, AIndex: Integer; AToken: TsNumFormatToken; AText: String); overload;
    procedure InsertElement(ASection, AIndex: Integer; AToken: TsNumFormatToken; AIntValue: Integer); overload;
    procedure InsertElement(ASection, AIndex: Integer; AToken: TsNumFormatToken; AFloatValue: Double); overload;
    function NextToken: Char;
    function PrevToken: Char;

    { Scanning/parsing }
    procedure ScanAMPM;
    procedure ScanAndCount(ATestChar: Char; out ACount: Integer);
    procedure ScanBrackets;
    procedure ScanCondition(AFirstChar: Char);
    procedure ScanCurrSymbol;
    procedure ScanDateTime;
    procedure ScanFormat;
    procedure ScanNumber;
    procedure ScanQuotedText;
    // Main scanner
    procedure Parse(const AFormatString: String);

    { Analysis while scanning }
    procedure AnalyzeColor(AValue: String);
    function AnalyzeCurrency(const AValue: String): Boolean;

    { Analysis after scanning }
    // General
    procedure CheckSections;
    procedure CheckSection(ASection: Integer);
    procedure FixMonthMinuteToken(ASection: Integer);
    // Format string
    function BuildFormatString: String; virtual;
    // Token analysis
    function GetTokenIntValueAt(AToken: TsNumFormatToken;
      ASection,AIndex: Integer): Integer;
    function IsNumberAt(ASection,AIndex: Integer; out ANumFormat: TsNumberFormat;
      out ADecimals: Byte; out ANextIndex: Integer): Boolean;
    function IsTextAt(AText: string; ASection, AIndex: Integer): Boolean;
    function IsTokenAt(AToken: TsNumFormatToken; ASection,AIndex: Integer): Boolean;

  public
    constructor Create(AWorkbook: TsWorkbook; const AFormatString: String);
    destructor Destroy; override;
    procedure ClearAll;
    function GetDateTimeCode(ASection: Integer): String;
    function IsDateTimeFormat: Boolean;
    function IsTimeFormat: Boolean;
    procedure LimitDecimals;

    property CurrencySymbol: String read GetCurrencySymbol;
    property Decimals: Byte read GetDecimals write SetDecimals;
    property FormatString: String read GetFormatString;
    property FracDenominator: Integer read GetFracDenominator;
    property FracInt: Integer read GetFracInt;
    property FracNumerator: Integer read GetFracNumerator;
    property NumFormat: TsNumberFormat read GetNumFormat;
    property ParsedSectionCount: Integer read GetParsedSectionCount;
    property ParsedSections[AIndex: Integer]: TsNumFormatSection read GetParsedSections;
    property Status: Integer read FStatus;
  end;


implementation

uses
  TypInfo, LazUTF8, fpsutils, fpsCurrency;


{ TsNumFormatParser }

{@@ Creates a number format parser for analyzing a formatstring that has been
  read from a spreadsheet file. }
constructor TsNumFormatParser.Create(AWorkbook: TsWorkbook;
  const AFormatString: String);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  Parse(AFormatString);
  CheckSections;
end;

destructor TsNumFormatParser.Destroy;
begin
  FSections := nil;
  inherited Destroy;
end;

procedure TsNumFormatParser.AddElement(AToken: TsNumFormatToken; AText: String);
var
  n: Integer;
begin
  n := Length(FSections[FCurrSection].Elements);
  SetLength(FSections[FCurrSection].Elements, n+1);
  FSections[FCurrSection].Elements[n].Token := AToken;
  FSections[FCurrSection].Elements[n].TextValue := AText;
end;

procedure TsNumFormatParser.AddElement(AToken: TsNumFormatToken; AIntValue: Integer);
var
  n: Integer;
begin
  n := Length(FSections[FCurrSection].Elements);
  SetLength(FSections[FCurrSection].Elements, n+1);
  FSections[FCurrSection].Elements[n].Token := AToken;
  FSections[FCurrSection].Elements[n].IntValue := AIntValue;
end;

procedure TsNumFormatParser.AddElement(AToken: TsNumFormatToken; AFloatValue: Double); overload;
var
  n: Integer;
begin
  n := Length(FSections[FCurrSection].Elements);
  SetLength(FSections[FCurrSection].Elements, n+1);
  FSections[FCurrSection].Elements[n].Token := AToken;
  FSections[FCurrSection].Elements[n].FloatValue := AFloatValue;
end;

procedure TsNumFormatParser.AddSection;
begin
  FCurrSection := Length(FSections);
  SetLength(FSections, FCurrSection + 1);
  with FSections[FCurrSection] do
    SetLength(Elements, 0);
end;

procedure TsNumFormatParser.AnalyzeColor(AValue: String);
var
  n: Integer;
begin
  AValue := lowercase(AValue);
  // Colors
  if AValue = 'red' then
    AddElement(nftColor, ord(scRed))
  else
  if AValue = 'black' then
    AddElement(nftColor, ord(scBlack))
  else
  if AValue = 'blue' then
    AddElement(nftColor, ord(scBlue))
  else
  if AValue = 'white' then
    AddElement(nftColor, ord(scWhite))
  else
  if AValue = 'green' then
    AddElement(nftColor, ord(scGreen))
  else
  if AValue = 'cyan' then
    AddElement(nftColor, ord(scCyan))
  else
  if AValue = 'magenta' then
    AddElement(nftColor, ord(scMagenta))
  else
  if copy(AValue, 1, 5) = 'color' then begin
    AValue := copy(AValue, 6, Length(AValue));
    if not TryStrToInt(trim(AValue), n) then begin
      FStatus := psErrNoValidColorIndex;
      exit;
    end;
    AddElement(nftColor, n);
  end else
    FStatus := psErrUnknownInfoInBrackets;
end;

function TsNumFormatParser.AnalyzeCurrency(const AValue: String): Boolean;
begin
  if (FWorkbook = nil) or (FWorkbook.FormatSettings.CurrencyString = '') then
    Result := false
  else
    Result := CurrencyRegistered(AValue);
end;

{ Creates a formatstring for all sections.
  Note: this implementation is only valid for the fpc and Excel dialects of
  format string. }
function TsNumFormatParser.BuildFormatString: String;
var
  i: Integer;
begin
  if Length(FSections) > 0 then begin
    Result := BuildFormatStringFromSection(FSections[0]);
    for i:=1 to High(FSections) do
      Result := Result + ';' + BuildFormatStringFromSection(FSections[i]);
  end;
end;

procedure TsNumFormatParser.CheckSections;
var
  i: Integer;
begin
  for i:=0 to High(FSections) do
    CheckSection(i);

  if (Length(FSections) > 1) and (FSections[1].NumFormat = nfCurrencyRed) then
    for i:=0 to High(FSections) do
      if FSections[i].NumFormat = nfCurrency then
        FSections[i].NumFormat := nfCurrencyRed;
end;

procedure TsNumFormatParser.CheckSection(ASection: Integer);
var
  el, next, i: Integer;
  section: PsNumFormatSection;
  nfs, nfsTest: String;
  nf: TsNumberFormat;
  datetimeFormats: set of TsNumberformat;
  f1,f2: Integer;
  decs: Byte;
begin
  if FStatus <> psOK then
    exit;

  section := @FSections[ASection];
  section^.Kind := [];

  for el := 0 to High(section^.Elements) do
    case section^.Elements[el].Token of
      nftPercent:
        section^.Kind := section^.Kind + [nfkPercent];
      nftExpChar:
        if (nfkExp in section^.Kind) then
          FStatus := psErrMultipleExpChars
        else
          section^.Kind := section^.Kind + [nfkExp];
      nftFracSymbol:
        if (nfkFraction in section^.Kind) then
          FStatus := psErrMultipleFracSymbols
        else
          section^.Kind := section^.Kind + [nfkFraction];
      nftCurrSymbol:
        begin
          if (nfkCurrency in section^.Kind) then
            FStatus := psErrMultipleCurrSymbols
          else begin
            section^.Kind := section^.Kind + [nfkCurrency];
            section^.CurrencySymbol := section^.Elements[el].TextValue;
          end;
        end;
      nftYear, nftMonth, nftDay:
        section^.Kind := section^.Kind + [nfkDate];
      nftHour, nftMinute, nftSecond, nftMilliseconds:
        begin
          section^.Kind := section^.Kind + [nfkTime];
          if section^.Elements[el].IntValue < 0 then
            section^.Kind := section^.Kind + [nfkTimeInterval];
        end;
    end;

  if FStatus <> psOK then
    exit;

  if (section^.Kind * [nfkDate, nfkTime] <> []) and
     (section^.Kind * [nfkPercent, nfkExp, nfkCurrency, nfkFraction] <> []) then
  begin
    FStatus := psErrNoValidDateTimeFormat;
    exit;
  end;

  section^.NumFormat := nfCustom;

  if (section^.Kind * [nfkDate, nfkTime] <> []) then
  begin
    FixMonthMinuteToken(ASection);
    nfs := GetFormatString;
    if (nfkTimeInterval in section^.Kind) then
      section^.NumFormat := nfTimeInterval
    else
    begin
      datetimeFormats := [nfShortDateTime, nfLongDate, nfShortDate, nfLongTime,
        nfShortTime, nfLongTimeAM, nfShortTimeAM, nfDayMonth, nfMonthYear];
      for nf in datetimeFormats do
      begin
        nfsTest := BuildDateTimeFormatString(nf, FWorkbook.FormatSettings);
        if Length(nfsTest) = Length(nfs) then
        begin
          for i := 1 to Length(nfsTest) do
            case nfsTest[i] of
              '/': if not (nf in [nfLongTimeAM, nfShortTimeAM]) then
                     nfsTest[i] := FWorkbook.FormatSettings.DateSeparator;
              ':': nfsTest[i] := FWorkbook.FormatSettings.TimeSeparator;
              'n': nfsTest[i] := 'm';
            end;
          if SameText(nfs, nfsTest) then
          begin
            section^.NumFormat := nf;
            break;
          end;
        end;
{
        if SameText(nfs, BuildDateTimeFormatString(nf, FWorkbook.FormatSettings)) then
        begin
          section^.NumFormat := nf;
          break;
        end;
        }
      end;
    end;
  end else
  begin
    el := 0;
    while el < Length(section^.Elements) do
    begin
      if IsNumberAt(ASection, el, nf, decs, next) then begin
        section^.Decimals := decs;
        if nf = nfFixedTh then begin
          if (nfkCurrency in section^.Kind) then
            section^.NumFormat := nfCurrency
          else
            section^.NumFormat := nfFixedTh
        end else
        begin
          section^.NumFormat := nf;
          if (nfkPercent in section^.Kind) then
            section^.NumFormat := nfPercentage
          else
          if (nfkExp in section^.Kind) then
            section^.NumFormat := nfExp
          else
          if (nfkCurrency in section^.Kind) then
            section^.NumFormat := nfCurrency
          else
          if (nfkFraction in section^.Kind) and (decs = 0) then begin
            f1 := section^.Elements[el].IntValue;  // int part or numerator
            el := next;
            while IsTokenAt(nftSpace, ASection, el) or IsTextAt(' ', ASection, el) do
              inc(el);
            if IsTokenAt(nftFracSymbol, ASection, el) then begin
              inc(el);
              while IsTokenAt(nftSpace, ASection, el) or IsTextAt(' ', aSection, el) do
                inc(el);
              if IsNumberAt(ASection, el, nf, decs, next) and (nf in [nfFixed, nfFraction]) and (decs = 0) then
              begin
                section^.FracInt := 0;
                section^.FracNumerator := f1;
                section^.FracDenominator := section^.Elements[el].IntValue;
                section^.NumFormat := nfFraction;
              end;
            end else
            if IsNumberAt(ASection, el, nf, decs, next) and (nf in [nfFixed, nfFraction]) and (decs = 0) then
            begin
              f2 := section^.Elements[el].IntValue;
              el := next;
              while IsTokenAt(nftSpace, ASection, el) or IsTextAt(' ', ASection, el) do
                inc(el);
              if IsTokenAt(nftFracSymbol, ASection, el) then
              begin
                inc(el);
                while IsTokenAt(nftSpace, ASection, el) or IsTextAt(' ', ASection, el) do
                  inc(el);
                if IsNumberAt(ASection, el, nf, decs, next) and (nf in [nfFixed, nfFraction]) and (decs=0) then
                begin
                  section^.FracInt := f1;
                  section^.FracNumerator := f2;
                  section^.FracDenominator := section^.Elements[el].IntValue;
                  section^.NumFormat := nfFraction;
                end;
              end;
            end;
          end;
        end;
        break;
      end else
      if IsTokenAt(nftColor, ASection, el) then
        section^.Color := section^.Elements[el].IntValue;
      inc(el);
    end;
    if (section^.NumFormat = nfCurrency) and (section^.Color = scRed) then
      section^.NumFormat := nfCurrencyRed;
  end;
end;

procedure TsNumFormatParser.ClearAll;
var
  i, j: Integer;
begin
  for i:=0 to Length(FSections)-1 do begin
    for j:=0 to Length(FSections[i].Elements) do
      if FSections[i].Elements <> nil then
        FSections[i].Elements[j].TextValue := '';
    FSections[i].Elements := nil;
    FSections[i].CurrencySymbol := '';
  end;
  FSections := nil;
end;

procedure TsNumFormatParser.DeleteElement(ASection, AIndex: Integer);
var
  i, n: Integer;
begin
  n := Length(FSections[ASection].Elements);
  for i:= AIndex+1 to n-1 do
    FSections[ASection].Elements[i-1] := FSections[ASection].Elements[i];
  SetLength(FSections[ASection].Elements, n-1);
end;

{ Identify the ambiguous "m" token ("month" or "minute") }
procedure TsNumFormatParser.FixMonthMinuteToken(ASection: Integer);
var
  i, j: Integer;

  // Finds the previous date/time element skipping spaces, date/time sep etc.
  function PrevDateTimeElement(j: Integer): Integer;
  begin
    Result := -1;
    dec(j);
    while (j >= 0) do begin
      with FSections[ASection].Elements[j] do
        if Token in [nftYear, nftMonth, nftDay, nftHour, nftMinute, nftSecond] then
        begin
          Result := j;
          exit;
        end;
      dec(j);
    end;
  end;

  // Finds the next date/time element skipping spaces, date/time sep etc.
  function NextDateTimeElement(j: Integer): Integer;
  begin
    Result := -1;
    inc(j);
    while (j < Length(FSections[ASection].Elements)) do begin
      with FSections[ASection].Elements[j] do
        if Token in [nftYear, nftMonth, nftDay, nftHour, nftMinute, nftSecond] then
        begin
          Result := j;
          exit;
        end;
      inc(j);
    end;
  end;

begin
  for i:=0 to High(FSections[ASection].Elements) do
    // Find index of nftMonthMinute token...
    if FSections[ASection].Elements[i].Token = nftMonthMinute then begin
      // ... and, using its neighbors, decide whether it is a month or a minute.
      j := NextDateTimeElement(i);
      if j <> -1 then
        case FSections[ASection].Elements[j].Token of
          nftDay, nftYear:
            begin
              FSections[ASection].Elements[i].Token := nftMonth;
              Continue;
            end;
          nftSecond:
            begin
              FSections[ASection].Elements[i].Token := nftMinute;
              Continue;
            end;
        end;
      j := PrevDateTimeElement(i);
      if j <> -1 then
        case FSections[ASection].Elements[j].Token of
          nftDay, nftYear:
            begin
              FSections[ASection].Elements[i].Token := nftMonth;
              Continue;
            end;
          nftHour:
            begin
              FSections[ASection].Elements[i].Token := nftMinute;
              Continue;
            end;
        end;
    end;
end;

procedure TsNumFormatParser.InsertElement(ASection, AIndex: Integer;
  AToken: TsNumFormatToken; AText: String);
var
  i, n: Integer;
begin
  n := Length(FSections[ASection].Elements);
  SetLength(FSections[ASection].Elements, n+1);
  for i:= n-1 downto AIndex+1 do
    FSections[ASection].Elements[i+1] := FSections[ASection].Elements[i];
  FSections[ASection].Elements[AIndex+1].Token := AToken;
  FSections[ASection].Elements[AIndex+1].TextValue := AText;
end;

procedure TsNumFormatParser.InsertElement(ASection, AIndex: Integer;
  AToken: TsNumFormatToken; AIntValue: Integer);
var
  i, n: Integer;
begin
  n := Length(FSections[ASection].Elements);
  SetLength(FSections[ASection].Elements, n+1);
  for i:= n-1 downto AIndex+1 do
    FSections[ASection].Elements[i+1] := FSections[ASection].Elements[i];
  FSections[ASection].Elements[AIndex+1].Token := AToken;
  FSections[ASection].Elements[AIndex+1].IntValue := AIntValue;
end;

procedure TsNumFormatParser.InsertElement(ASection, AIndex: Integer;
  AToken: TsNumFormatToken; AFloatValue: Double);
var
  i, n: Integer;
begin
  n := Length(FSections[ASection].Elements);
  SetLength(FSections[ASection].Elements, n+1);
  for i:= n-1 downto AIndex+1 do
    FSections[ASection].Elements[i+1] := FSections[ASection].Elements[i];
  FSections[ASection].Elements[AIndex+1].Token := AToken;
  FSections[ASection].Elements[AIndex+1].FloatValue := AFloatValue;
end;

function TsNumFormatParser.GetFormatString: String;
begin
  Result := BuildFormatString;
end;

{ Extracts the currency symbol form the formatting sections. It is assumed that
  all two or three sections of the currency/accounting format use the same
  currency symbol, otherwise it would be custom format anyway which ignores
  the currencysymbol value. }
function TsNumFormatParser.GetCurrencySymbol: String;
begin
  if Length(FSections) > 0 then
    Result := FSections[0].CurrencySymbol
  else
    Result := '';
end;

{ Creates a string which summarizes the date/time formats in the given section.
  The string contains a 'y' for a nftYear, a 'm' for a nftMonth, a
  'd' for a nftDay, a 'h' for a nftHour, a 'n' for a nftMinute, a 's' for a
  nftSeconds, and a 'z' for a nftMilliseconds token. The order is retained.
  Needed for biff2 }
function TsNumFormatParser.GetDateTimeCode(ASection: Integer): String;
var
  i: Integer;
begin
  Result := '';
  if ASection < Length(FSections) then
    with FSections[ASection] do begin
      i := 0;
      while i < Length(Elements) do begin
        case Elements[i].Token of
          nftYear        : Result := Result + 'y';
          nftMonth       : Result := Result + 'm';
          nftDay         : Result := Result + 'd';
          nftHour        : Result := Result + 'h';
          nftMinute      : Result := Result + 'n';
          nftSecond      : Result := Result + 's';
          nftMilliSeconds: Result := Result + 'z';
        end;
        inc(i);
      end;
    end;
end;

{ Extracts the number of decimals from the sections. Since they are needed only
  for default formats having only a single section, only the first section is
  considered. In case of currency/accounting having two or three sections, it is
  assumed that all sections have the same decimals count, otherwise it would not
  be a standard format. }
function TsNumFormatParser.GetDecimals: Byte;
begin
  if Length(FSections) > 0 then
    Result := FSections[0].Decimals
  else
    Result := 0;
end;

function TsNumFormatParser.GetFracDenominator: Integer;
begin
  if Length(FSections) > 0 then
    Result := FSections[0].FracDenominator
  else
    Result := 0;
end;

function TsNumFormatParser.GetFracInt: Integer;
begin
  if Length(FSections) > 0 then
    Result := FSections[0].FracInt
  else
    Result := 0;
end;

function TsNumFormatParser.GetFracNumerator: Integer;
begin
  if Length(FSections) > 0 then
    Result := FSections[0].FracNumerator
  else
    Result := 0;
end;

{ Tries to extract a common builtin number format from the sections. If there
  are multiple sections, it is always a custom format, except for Currency and
  Accounting. }
function TsNumFormatParser.GetNumFormat: TsNumberFormat;
begin
  if Length(FSections) = 0 then
    result := nfGeneral
  else begin
    Result := FSections[0].NumFormat;
    if (Result = nfCurrency) then begin
      if Length(FSections) = 2 then begin
        Result := FSections[1].NumFormat;
        if FSections[1].CurrencySymbol <> FSections[0].CurrencySymbol then begin
          Result := nfCustom;
          exit;
        end;
        if (FSections[0].NumFormat in [nfCurrency, nfCurrencyRed]) and
           (FSections[1].NumFormat in [nfCurrency, nfCurrencyRed])
        then
          exit;
      end else
      if Length(FSections) = 3 then begin
        Result := FSections[1].NumFormat;
        if (FSections[0].CurrencySymbol <> FSections[1].CurrencySymbol) or
           (FSections[1].CurrencySymbol <> FSections[2].CurrencySymbol)
        then begin
          Result := nfCustom;
          exit;
        end;
        if (FSections[0].NumFormat in [nfCurrency, nfCurrencyRed]) and
           (FSections[1].NumFormat in [nfCurrency, nfCurrencyRed]) and
           (FSections[2].NumFormat in [nfCurrency, nfCurrencyRed])
        then
          exit;
      end;
      Result := nfCustom;
      exit;
    end;
    if Length(FSections) > 1 then
      Result := nfCustom;
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

function TsNumFormatParser.GetTokenIntValueAt(AToken: TsNumFormatToken;
  ASection, AIndex: Integer): Integer;
begin
  if IsTokenAt(AToken, ASection, AIndex) then
    Result := FSections[ASection].Elements[AIndex].IntValue
  else
    Result := -1;
end;

{ Returns true if the format elements contain at least one date/time token }
function TsNumFormatParser.IsDateTimeFormat: Boolean;
var
  section: TsNumFormatSection;
begin
  for section in FSections do
    if section.Kind * [nfkDate, nfkTime] <> [] then
    begin
      Result := true;
      exit;
    end;
  Result := false;
end;

function TsNumFormatParser.IsNumberAt(ASection, AIndex: Integer;
  out ANumFormat: TsNumberFormat; out ADecimals: Byte;
  out ANextIndex: Integer): Boolean;
var
  token: TsNumFormatToken;
begin
  if (ASection > High(FSections)) or (AIndex > High(FSections[ASection].Elements))
  then begin
    Result := false;
    ANextIndex := AIndex;
    exit;
  end;

  Result := true;
  ANumFormat := nfCustom;
  ADecimals := 0;
  token := FSections[ASection].Elements[AIndex].Token;

  if token in [nftFracNumOptDigit, nftFracNumZeroDigit, nftFracNumSpaceDigit,
    nftFracDenomOptDigit, nftFracDenomZeroDigit, nftFracDenomSpaceDigit] then
  begin
    ANumFormat := nfFraction;
    ANextIndex := AIndex + 1;
    exit;
  end;

  if (token = nftIntTh) and (FSections[ASection].Elements[AIndex].IntValue = 1) then   // '#,##0'
    ANumFormat := nfFixedTh
  else
  if (token = nftIntZeroDigit) and (FSections[ASection].Elements[AIndex].IntValue = 1) then // '0'
    ANumFormat := nfFixed;

  if (token in [nftIntTh, nftIntZeroDigit, nftIntOptDigit, nftIntSpaceDigit]) then
  begin
    if IsTokenAt(nftDecSep, ASection, AIndex+1) then
    begin
      if AIndex + 2 < Length(FSections[ASection].Elements) then
      begin
        token := FSections[ASection].Elements[AIndex+2].Token;
        if (token in [nftZeroDecs, nftOptDecs, nftSpaceDecs]) then
        begin
          ANextIndex := AIndex + 3;
          ADecimals := FSections[ASection].Elements[AIndex+2].IntValue;
          if (token <> nftZeroDecs) then
            ANumFormat := nfCustom;
          exit;
        end;
      end;
    end else
    if IsTokenAt(nftSpace, ASection, AIndex+1) then
    begin
      ANumFormat := nfFraction;
      ANextIndex := AIndex + 1;
      exit;
    end else
    begin
      ANextIndex := AIndex + 1;
      exit;
    end;
  end;

  ANextIndex := AIndex;
  Result := false;
end;

function TsNumFormatParser.IsTextAt(AText: String; ASection, AIndex: Integer): Boolean;
begin
  Result := IsTokenAt(nftText, ASection, AIndex) and
    (FSections[ASection].Elements[AIndex].TextValue = AText);
end;

{ Returns true if the format elements contain only time, no date tokens. }
function TsNumFormatParser.IsTimeFormat: Boolean;
var
  section: TsNumFormatSection;
begin
  for section in FSections do
    if (nfkTime in section.Kind) then
    begin
      Result := true;
      exit;
    end;
  Result := false;
end;

function TsNumFormatParser.IsTokenAt(AToken: TsNumFormatToken;
  ASection, AIndex: Integer): Boolean;
begin
  Result := (ASection < Length(FSections)) and
            (AIndex < Length(FSections[ASection].Elements)) and
            (FSections[ASection].Elements[AIndex].Token = AToken);
end;

{ Limits the decimals to 0 or 2, as required by Excel2. }
procedure TsNumFormatParser.LimitDecimals;
var
  i, j: Integer;
begin
  for j:=0 to High(FSections) do
    for i:=0 to High(FSections[j].Elements) do
      if FSections[j].Elements[i].Token = nftZeroDecs then
        if FSections[j].Elements[i].IntValue > 0 then
          FSections[j].Elements[i].IntValue := 2;
end;

function TsNumFormatParser.NextToken: Char;
begin
  if FCurrent < FEnd then begin
    inc(FCurrent);
    Result := FCurrent^;
  end else
    Result := #0;
end;

function TsNumFormatParser.PrevToken: Char;
begin
  if FCurrent > nil then begin
    dec(FCurrent);
    Result := FCurrent^;
  end else
    Result := #0;
end;

procedure TsNumFormatParser.Parse(const AFormatString: String);
begin
  FStatus := psOK;

  AddSection;
  if (AFormatString = '') or SameText(AFormatString, 'General') then
    exit;

  FStart := @AFormatString[1];
  FEnd := FStart + Length(AFormatString);
  FCurrent := FStart;
  FToken := FCurrent^;
  while (FCurrent < FEnd) and (FStatus = psOK) do begin
    case FToken of
      '[': ScanBrackets;
      '"': ScanQuotedText;
      ':': AddElement(nftDateTimeSep, ':');
      ';': AddSection;
      else ScanFormat;
    end;
    FToken := NextToken;
  end;
end;

{ Scans an AM/PM sequence (or AMPM or A/P).
  At exit, cursor is a next character }
procedure TsNumFormatParser.ScanAMPM;
var
  s: String;
  el: Integer;
begin
  s := '';
  while (FCurrent < FEnd) do begin
    if (FToken in ['A', 'a', 'P', 'p', 'm', 'M', '/']) then
      s := s + FToken
    else
      break;
    FToken := NextToken;
  end;
  if s <> '' then
  begin
    AddElement(nftAMPM, s);
    // Tag the hour element for AM/PM format needed
    el := High(FSections[FCurrSection].Elements)-1;
    for el := High(FSections[FCurrSection].Elements)-1 downto 0 do
      if FSections[FCurrSection].Elements[el].Token = nftHour then
      begin
        FSections[FCurrSection].Elements[el].TextValue := 'AM';
        break;
      end;
  end;
end;

{ Counts the number of characters equal to ATestChar. Stops at the next
  different character. This is also where the cursor is at exit. }
procedure TsNumFormatParser.ScanAndCount(ATestChar: Char; out ACount: Integer);
begin
  ACount := 0;
  repeat
    inc(ACount);
    FToken := NextToken;
  until (FToken <> ATestChar) or (FCurrent >= FEnd);
end;

{ Extracts the text between square brackets. This can be
  - a time duration like [hh]
  - a condition, like [>= 2.0]
  - a currency symbol like [$€-409]
  - a color like [red] or [color25]
  The procedure is left with the cursor at ']' }
procedure TsNumFormatParser.ScanBrackets;
var
  s: String;
  n: Integer;
  prevtoken: Char;
  isText: Boolean;
begin
  s := '';
  isText := false;
  FToken := NextToken;   // Cursor was at '['
  while (FCurrent < FEnd) and (FStatus = psOK) do begin
    case FToken of
      'h', 'H', 'm', 'M', 'n', 'N', 's', 'S':
        if isText then
          s := s + FToken
        else
        begin
          prevtoken := FToken;
          ScanAndCount(FToken, n);
          if (FToken in [']', #0]) then begin
            case prevtoken of
              'h', 'H'          : AddElement(nftHour, -n);
              'm', 'M', 'n', 'N': AddElement(nftMinute, -n);
              's', 'S'          : AddElement(nftSecond, -n);
            end;
            break;
          end else
           FStatus := psErrUnknownInfoInBrackets;
        end;

      '<', '>', '=':
        begin
          ScanCondition(FToken);
          if FToken = ']' then
            break
          else
           FStatus := psErrUnknownInfoInBrackets;
        end;

      '$':
        begin
          ScanCurrSymbol;
          if FToken = ']' then
            break
          else
           FStatus := psErrUnknownInfoInBrackets;
        end;

      ']':
        begin
          AnalyzeColor(s);
          break;
        end;

      else
        s := s + FToken;
        isText := true;
    end;
    FToken := NextToken;
  end;
end;

{ Scans a condition like [>=2.0]. Starts after the "[" and ends before at "]".
  Returns first character after the number (spaces allowed). }
procedure TsNumFormatParser.ScanCondition(AFirstChar: Char);
var
  s: String;
//  op: TsCompareOperation;
  value: Double;
  res: Integer;
begin
  s := AFirstChar;
  FToken := NextToken;
  if FToken in ['>', '<', '='] then s := s + FToken else FToken := PrevToken;
  {
  if s = '=' then op := coEqual else
  if s = '<>' then op := coNotEqual else
  if s = '<' then op := coLess else
  if s = '>' then op := coGreater else
  if s = '<=' then op := coLessEqual else
  if s = '>=' then op := coGreaterEqual
  else begin
    FStatus := psErrUnknownInfoInBrackets;
    FToken := #0;
    exit;
  end;
    }
  while (FToken = ' ') and (FCurrent < FEnd) do
    FToken := NextToken;

  if FCurrent >= FEnd then begin
    FStatus := psErrUnknownInfoInBrackets;
    FToken := #0;
    exit;
  end;

  s := FToken;
  while (FCurrent < FEnd) and (FToken in ['+', '-', '.', '0'..'9']) do begin
    FToken := NextToken;
    s := s + FToken;
  end;
  val(s, value, res);
  if res <> 0 then begin
    FStatus := psErrUnknownInfoInBrackets;
    FToken := #0;
    exit;
  end;

  while (FCurrent < FEnd) and (FToken = ' ') do
    FToken := NextToken;
  if FToken = ']' then
    AddElement(nftCompareOp, value)
  else begin
    FStatus := psErrUnknownInfoInBrackets;
    FToken := #0;
  end;
end;

{ Scans to end of a symbol like [$EUR-409], starting after the $ and ending at
  the "]".
  After the "$" follows the currency symbol, after the "-" country information }
procedure TsNumFormatParser.ScanCurrSymbol;
var
  s: String;
begin
  s := '';
  FToken := NextToken;
  while (FCurrent < FEnd) and not (FToken in ['-', ']']) do begin
    s := s + FToken;
    FToken := NextToken;
  end;
  if s <> '' then
    AddElement(nftCurrSymbol, s);
  if FToken <> ']' then begin
    FToken := NextToken;
    while (FCurrent < FEnd) and (FToken <> ']') do begin
      s := s + FToken;
      FToken := NextToken;
    end;
    if s <> '' then
      AddElement(nftCountry, s);
  end;
end;

{ Scans a date/time format. Procedure is left with the cursor at the last char
  of the date/time format. }
procedure TsNumFormatParser.ScanDateTime;
var
  n: Integer;
  token: Char;
begin
  while (FCurrent < FEnd) and (FStatus = psOK) do begin
    case FToken of
      '\':  // means that the next character is taken literally
        begin
          FToken := NextToken;     // skip the "\"...
          AddElement(nftEscaped, FToken);
          FToken := NextToken;
        end;
      'Y', 'y':
        begin
          ScanAndCount(FToken, n);
          AddElement(nftYear, n);
        end;
      'm', 'M', 'n', 'N':
        begin
          ScanAndCount(FToken, n);
          AddElement(nftMonthMinute, n);  // Decide on minute or month later
        end;
      'D', 'd':
        begin
          ScanAndCount(FToken, n);
          AddElement(nftDay, n);
        end;
      'H', 'h':
        begin
          ScanAndCount(FToken, n);
          AddElement(nftHour, n);
        end;
      'S', 's':
        begin
          ScanAndCount(FToken, n);
          AddElement(nftSecond, n);
        end;
      '/', ':':
        begin
          AddElement(nftDateTimeSep, FToken);
          FToken := NextToken;
        end;
      '.':
        begin
          token := NextToken;
          if token in ['z', '0'] then begin
            AddElement(nftDecSep, FToken);
            FToken := NextToken;
            ScanAndCount(FToken, n);
            AddElement(nftMilliseconds, n);
          end else begin
            AddElement(nftDateTimeSep, FToken);
            FToken := token;
          end;
        end;
      '[':
        begin
          ScanBrackets;
          FToken := NextToken;
        end;
      'A', 'a':
        ScanAMPM;
      ',', '-':
        begin
          AddElement(nftText, FToken);
          FToken := NextToken;
        end
      else
        // char pointer must be at end of date/time mask.
        FToken := PrevToken;
        Exit;
    end;
  end;
end;

procedure TsNumFormatParser.ScanFormat;
var
  done: Boolean;
begin
  done := false;
  while (FCurrent < FEnd) and (FStatus = psOK) and (not done) do begin
    case FToken of
      '\': // Excel: add next character literally
        begin
          FToken := NextToken;
          AddElement(nftText, FToken);
        end;
      '*':  // Excel: repeat next character to fill cell. For accounting format.
        begin
          FToken := NextToken;
          AddElement(nftRepeat, FToken);
        end;
      '_':  // Excel: Leave width of next character empty
        begin
          FToken := NextToken;
          AddElement(nftEmptyCharWidth, FToken);
        end;
      '@':  // Excel: Indicates text format
        begin
          AddElement(nftTextFormat, FToken);
        end;
      '"':
        ScanQuotedText;
      '(', ')':
        AddElement(nftSignBracket, FToken);
      '0', '#', '?', '.', ',', '-':
        ScanNumber;
      'y', 'Y', 'm', 'M',  'd', 'D', 'h', 'N', 'n', 's':
        ScanDateTime;
      '[':
        ScanBrackets;
      ' ':
        AddElement(nftSpace, FToken);
      'A', 'a':
        begin
          ScanAMPM;
          FToken := PrevToken;
        end;
      ';':  // End of the section. Important: Cursor must stay on ';'
        begin
          AddSection;
          Exit;
        end;
      else
        AddElement(nftText, FToken);
    end;
    FToken := NextToken;
  end;
end;

{ Scans a floating point format. Procedure is left with the cursor at the last
  character of the format. }
procedure TsNumFormatParser.ScanNumber;
var
  hasDecSep: Boolean;
  isFrac: Boolean;
  n: Integer;
  el: Integer;
  savedCurrent: PChar;
begin
  hasDecSep := false;
  isFrac := false;
  while (FCurrent < FEnd) and (FStatus = psOK) do begin
    case FToken of
      ',': AddElement(nftThSep, ',');
      '.': begin
             AddElement(nftDecSep, '.');
             hasDecSep := true;
           end;
      '#': begin
             ScanAndCount('#', n);
             savedCurrent := FCurrent;
             if not (hasDecSep or isFrac) and (n = 1) and (FToken = ',') then
             begin
               FToken := NextToken;
               ScanAndCount('#', n);
               case n of
                 0: begin
                      FToken := PrevToken;
                      ScanAndCount('0', n);
                      FToken := prevToken;
                      if n = 3 then
                        AddElement(nftIntTh, 3)
                      else
                        FCurrent := savedCurrent;
                    end;
                 1: begin
                      ScanAndCount('0', n);
                      FToken := prevToken;
                      if n = 2 then
                        AddElement(nftIntTh, 2)
                      else
                        FCurrent := savedCurrent;
                    end;
                 2: begin
                      ScanAndCount('0', n);
                      FToken := prevToken;
                      if (n = 1) then
                        AddElement(nftIntTh, 1)
                      else
                        FCurrent := savedCurrent;
                    end;
               end;
             end else
             begin
               FToken := PrevToken;
               if isFrac then
                 AddElement(nftFracDenomOptDigit, n)
               else
               if hasDecSep then
                 AddElement(nftOptDecs, n)
               else
                 AddElement(nftIntOptDigit, n);
             end;
           end;
      '0': begin
             ScanAndCount('0', n);
             FToken := PrevToken;
             if hasDecSep then
               AddElement(nftZeroDecs, n)
             else
             if isFrac then
               AddElement(nftFracDenomZeroDigit, n)
             else
               AddElement(nftIntZeroDigit, n);
           end;
      '?': begin
             ScanAndCount('?', n);
             FToken := PrevToken;
             if hasDecSep then
               AddElement(nftSpaceDecs, n)
             else
             if isFrac then
               AddElement(nftFracDenomSpaceDigit, n)
             else
               AddElement(nftIntSpaceDigit, n);
           end;
      'E', 'e':
           begin
             AddElement(nftExpChar, FToken);
             FToken := NextToken;
             if FToken in ['+', '-'] then
               AddElement(nftExpSign, FToken);
             FToken := NextToken;
             if FToken = '0' then begin
               ScanAndCount('0', n);
               FToken := PrevToken;
               AddElement(nftExpDigits, n);
             end;
           end;
      '+', '-':
           AddElement(nftSign, FToken);
      '%': AddElement(nftPercent, FToken);
      '/': begin
             isFrac := true;
             AddElement(nftFracSymbol, FToken);
             // go back and replace correct token for numerator
             el := High(FSections[FCurrSection].Elements);
             while el > 0 do begin
               dec(el);
               case FSections[FCurrSection].Elements[el].Token of
                 nftIntOptDigit:
                   begin
                     FSections[FCurrSection].Elements[el].Token := nftFracNumOptDigit;
                     break;
                   end;
                 nftIntSpaceDigit:
                   begin
                     FSections[FCurrSection].Elements[el].Token := nftFracNumSpaceDigit;
                     break;
                   end;
                 nftIntZeroDigit:
                   begin
                     FSections[FCurrSection].Elements[el].Token := nftFracNumZeroDigit;
                     break;
                   end;
               end;
             end;
           end;

      else
           FToken := PrevToken;
           Exit;
    end;
    FToken := NextToken;
  end;
end;

{ Scans a text in quotation marks. Tries to interpret the text as a currency
  symbol (--> AnalyzeText).
  The procedure is entered and left with the cursor at a quotation mark. }
procedure TsNumFormatParser.ScanQuotedText;
var
  s: String;
begin
  s := '';
  FToken := NextToken;   // Cursor war at '"'
  while (FCurrent < FEnd) and (FStatus = psOK) do begin
    if FToken = '"' then begin
      if AnalyzeCurrency(s) then
        AddElement(nftCurrSymbol, s)
      else
        AddElement(nftText, s);
      exit;
    end else begin
      s := s + FToken;
      FToken := NextToken;
    end;
  end;
  // When the procedure gets here the final quotation mark is missing
  FStatus := psErrQuoteExpected;
end;

procedure TsNumFormatParser.SetDecimals(AValue: Byte);
var
  i, j, n: Integer;
begin
  for j := 0 to High(FSections) do begin
    n := Length(FSections[j].Elements);
    i := n-1;
    while (i > -1) do begin
      case FSections[j].Elements[i].Token of
        nftIntOptDigit, nftIntZeroDigit, nftIntSpaceDigit, nftIntTh:
          // no decimals so far --> add decimal separator and decimals element
          if (AValue > 0) then begin
            // Don't use "AddElements" because nfCurrency etc have elements after the number.
            InsertElement(j, i, nftDecSep, '.');
            InsertElement(j, i+1, nftZeroDecs, AValue);
            break;
          end;
        nftZeroDecs, nftOptDecs, nftSpaceDecs:
          if AValue > 0 then begin
            // decimals are already used, just replace value of decimal places
            FSections[j].Elements[i].IntValue := AValue;
            FSections[j].Elements[i].Token := nftZeroDecs;
            break;
          end else begin
            // No decimals any more: delete decs and decsep elements
            DeleteElement(j, i);
            DeleteElement(j, i-1);
            break;
          end;
      end;
      dec(i);
    end;
  end;
end;

end.
