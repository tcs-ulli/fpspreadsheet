unit fpsNumFormatParser;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

uses
  SysUtils, fpstypes, fpspreadsheet;


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
  psAmbiguousSymbol = 9;

{ TsNumFormatParser }

type
  TsNumFormatToken = (nftText, nftThSep, nftDecSep,
    nftYear, nftMonth, nftDay, nftHour, nftMinute, nftSecond, nftMilliseconds,
    nftAMPM, nftMonthMinute, nftDateTimeSep,
    nftSign, nftSignBracket,
    nftDigit, nftOptDigit, nftOptSpaceDigit, nftDecs, nftOptDec,
    nftExpChar, nftExpSign, nftExpDigits,
    nftPercent,
    nftFraction,
    nftCurrSymbol, nftCountry,
    nftColor, nftCompareOp, nftCompareValue,
    nftSpace, nftEscaped,
    nftRepeat, nftEmptyCharWidth,
    nftTextFormat);

  TsNumFormatElement = record
    Token: TsNumFormatToken;
    IntValue: Integer;
    FloatValue: Double;
    TextValue: String;
  end;

  TsNumFormatElements = array of TsNumFormatElement;

  TsNumFormatSection = record
    Elements: TsNumFormatElements;
    NumFormat: TsNumberFormat;
    Decimals: Byte;
    FracInt: Integer;
    FracNumerator: Integer;
    FracDenominator: Integer;
    CurrencySymbol: String;
    Color: TsColor;
  end;

  TsNumFormatSections = array of TsNumFormatSection;

  TsNumFormatParser = class
  private
    FCreateMethod: Byte;
    FToken: Char;
    FCurrent: PChar;
    FStart: PChar;
    FEnd: PChar;
    FCurrSection: Integer;
    FHasRedSection: Boolean;
    FStatus: Integer;
    function GetCurrencySymbol: String;
    function GetDecimals: byte;
    function GetFracDenominator: Integer;
    function GetFracInt: Integer;
    function GetFracNumerator: Integer;
    function GetFormatString(ADialect: TsNumFormatDialect): String;
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
    // Format string
    function BuildFormatString(ADialect: TsNumFormatDialect): String; virtual;
    function BuildFormatStringFromSection(ASection: Integer;
      ADialect: TsNumFormatDialect): String; virtual;
    // NumberFormat
    procedure EvalNumFormatOfSection(ASection: Integer);
    function IsCurrencyAt(ASection: Integer; out ANumFormat: TsNumberFormat;
      out ADecimals: byte; out ACurrencySymbol: String; out AColor: TsColor): Boolean;
    function IsDateAt(ASection,AIndex: Integer; out ANumberFormat: TsNumberFormat;
      var ANextIndex: Integer): Boolean;
    function IsNumberAt(ASection,AIndex: Integer; out ANumberFormat: TsNumberFormat;
      out ADecimals: Byte; out ANextIndex: Integer): Boolean;
    function IsTextAt(AText: string; ASection, AIndex: Integer): Boolean;
    function IsTimeAt(ASection,AIndex: Integer; out ANumberFormat: TsNumberFormat;
      out ANextIndex: Integer): Boolean;
    function IsTokenAt(AToken: TsNumFormatToken; ASection,AIndex: Integer): Boolean;

  public
    constructor Create(AWorkbook: TsWorkbook; const AFormatString: String;
      const ANumFormat: TsNumberFormat = nfGeneral);
    destructor Destroy; override;
    procedure ClearAll;
    function GetDateTimeCode(ASection: Integer): String;
    function IsDateTimeFormat: Boolean;
    function IsTimeFormat: Boolean;
    procedure LimitDecimals;
    procedure Localize;

    property CurrencySymbol: String read GetCurrencySymbol;
    property Decimals: Byte read GetDecimals write SetDecimals;
    property FormatString[ADialect: TsNumFormatDialect]: String read GetFormatString;
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
  TypInfo, StrUtils, LazUTF8, fpsutils, fpsCurrency;


{ TsNumFormatParser }

{@@ Creates a number format parser for analyzing a formatstring that has been
  read from a spreadsheet file.
  In case of "red" number formats we also have to specify the number format
  because the format string might not contain the color information, and we
  extract it from the NumFormat in this case. }
constructor TsNumFormatParser.Create(AWorkbook: TsWorkbook;
  const AFormatString: String; const ANumFormat: TsNumberFormat = nfGeneral);
begin
  inherited Create;
  FCreateMethod := 0;
  FWorkbook := AWorkbook;
  FHasRedSection := (ANumFormat = nfCurrencyRed);
  Parse(AFormatString);
end;

destructor TsNumFormatParser.Destroy;
begin
  FSections := nil;
//  ClearAll;
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
function TsNumFormatParser.BuildFormatString(ADialect: TsNumFormatDialect): String;
var
  i: Integer;
begin
  if Length(FSections) > 0 then begin
    Result := BuildFormatStringFromSection(0, ADialect);
    for i := 1 to High(FSections) do
      Result := Result + ';' + BuildFormatStringFromSection(i, ADialect);
  end else
    Result := '';
end;

{ Creates a format string for the given section. This implementation covers
  the formatstring dialects of fpc (nfdDefault) and Excel (nfdExcel). }
function TsNumFormatParser.BuildFormatStringFromSection(ASection: Integer;
  ADialect: TsNumFormatDialect): String;
var
  element: TsNumFormatElement;
  i: Integer;
begin
  Result := '';

  if (ASection < 0) and (ASection >= GetParsedSectionCount) then
    exit;
  for i := 0 to High(FSections[ASection].Elements)  do begin
    element := FSections[ASection].Elements[i];
    case element.Token of
      nftText:
        if element.TextValue <> '' then result := Result + '"' + element.TextValue + '"';
      nftThSep, nftDecSep:
        Result := Result + element.TextValue;
      nftDigit:
        Result := Result + '0';
      nftOptDigit, nftOptDec:
        Result := Result + '#';
      nftOptSpaceDigit:
        Result := Result + '#';  /// !!!!!!!!!!! TO BE CHANGED !!!!!!!!!!!!!!!!!
      nftYear:
        Result := Result + DupeString(IfThen(ADialect = nfdExcel, 'Y', 'y'), element.IntValue);
      nftMonth:
        Result := Result + DupeString(IfThen(ADialect = nfdExcel, 'M', 'm'), element.IntValue);
      nftDay:
        Result := Result + DupeString(IfThen(ADialect = nfdExcel, 'D', 'd'), element.IntValue);
      nftHour:
        if element.IntValue < 0
          then Result := Result + '[' + DupeString('h', -element.IntValue) + ']'
          else Result := Result + DupeString('h', element.IntValue);
      nftMinute:
        if element.IntValue < 0
          then Result := result + '[' + DupeString(IfThen(ADialect = nfdExcel, 'm', 'n'), -element.IntValue) + ']'
          else Result := Result + DupeString(IfThen(ADialect = nfdExcel, 'm', 'n'), element.IntValue);
      nftSecond:
        if element.IntValue < 0
          then Result := Result + '[' + DupeString('s', -element.IntValue) + ']'
          else Result := Result + DupeString('s', element.IntValue);
      nftDecs, nftExpDigits, nftMilliseconds:
        Result := Result + Dupestring('0', element.IntValue);
      nftSpace, nftSign, nftSignBracket, nftExpChar, nftExpSign, nftPercent,
      nftFraction, nftAMPM, nftDateTimeSep:
        if element.TextValue <> '' then Result := Result + element.TextValue;
      nftCurrSymbol:
        if element.TextValue <> '' then begin
          if ADialect = nfdExcel then
            Result := Result + '[$' + element.TextValue + ']'
          else
            Result := Result + '"' + element.TextValue + '"';
        end;
      nftEscaped:
        if element.TextValue <> '' then begin
          if ADialect = nfdExcel then
            Result := Result + '\' + element.TextValue
          else
            Result := Result + element.TextValue;
        end;
      nftTextFormat:
        if element.TextValue <> '' then
          if ADialect = nfdExcel then Result := Result + element.TextValue;
      nftRepeat:
        if element.TextValue <> '' then Result := Result + '*' + element.TextValue;
      nftColor:
        if ADialect = nfdExcel then begin
          case element.IntValue of
            scBlack  : Result := '[black]';
            scWhite  : Result := '[white]';
            scRed    : Result := '[red]';
            scBlue   : Result := '[blue]';
            scGreen  : Result := '[green]';
            scYellow : Result := '[yellow]';
            scMagenta: Result := '[magenta]';
            scCyan   : Result := '[cyan]';
            else       Result := Format('[Color%d]', [element.IntValue]);
          end;
        end;
    end;
  end;
end;

procedure TsNumFormatParser.CheckSections;
var
  i: Integer;
begin
  for i:=0 to Length(FSections)-1 do
    CheckSection(i);
end;

procedure TsNumFormatParser.CheckSection(ASection: Integer);
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
  // Fix the ambiguous "m":
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

  EvalNumFormatOfSection(ASection);
  {
    FSections[ASection].NumFormat,
    FSections[ASection].Decimals,
    FSections[ASection].Numerator,
    FSections[ASection].Denominator,
    FSections[ASection].CurrencySymbol,
    FSections[ASection].Color
  );
  }
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

function TsNumFormatParser.GetFormatString(ADialect: TsNumFormatDialect): String;
var
  i: Integer;
begin
  Result := '';
  if Length(FSections) > 0 then begin
    Result := BuildFormatStringFromSection(0, ADialect);
    for i:=1 to High(FSections) do
      Result := Result + ';' + BuildFormatStringFromSection(i, ADialect);
  end;
end;

procedure TsNumFormatParser.EvalNumFormatOfSection(ASection: Integer);
{
  out ANumFormat: TsNumberFormat; out ADecimals: byte;
  out ANumerator, ADenominator: Integer; out ACurrencySymbol: String;
  out AColor: TsColor);
  }
var
  nf, nf1: TsNumberFormat;
  next: Integer = 0;
  decs, num, denom: Byte;
  cs: String;
  clr: TsColor;
begin
  nf := nfCustom;
  decs := 0;
  num := 0;
  denom := 0;
  cs := '';
  clr := scNotDefined;

  with FSections[ASection] do begin
    if Length(Elements) = 0 then begin
      FSections[ASection].NumFormat := nfGeneral;
      exit;
    end;

    // Look for number formats
    if IsNumberAt(ASection, 0, nf, decs, next) then begin
      // nfFixed, nfFixedTh
      if next = Length(Elements) then
      begin
        FSections[ASection].NumFormat := nf;
        FSections[ASection].Decimals := decs;
        exit;
      end;
      // nfPercentage
      if IsTokenAt(nftPercent, ASection, next) and (next+1 = Length(Elements))
      then begin
        FSections[ASection].NumFormat := nfPercentage;
        FSections[ASection].Decimals := decs;
        exit;
      end;
      // nfFraction
      if (IsTokenAt(nftSpace, ASection, next) or IsTextAt(' ', ASection, next)) and
         IsNumberAt(ASection, next+1, nf, num, next) and
         IsTokenAt(nftFraction, ASection, next) and
         IsNumberAt(ASection, next+1, nf, denom, next) and
         (next = Length(Elements))
      then begin
        FSections[ASection].NumFormat := nfFraction;
        FSections[ASection].FracInt := integer(decs);  // "decs" means "number of integer digits", here
        FSections[ASection].FracNumerator := integer(num);
        FSections[ASection].FracDenominator := integer(denom);
        exit;
      end;
      if IsTokenAt(nftFraction, ASection, next) and
         IsNumberAt(ASection, next+1, nf, denom, next) and
         (next = Length(Elements))
      then begin
        FSections[ASection].NumFormat := nfFraction;
        FSections[ASection].FracNumerator := integer(decs);
        FSections[ASection].FracDenominator := integer(denom);
        exit;
      end;
      // nfExp
      if IsTokenAt(nftExpChar, ASection, next) then begin
        if IsTokenAt(nftExpSign, ASection, next+1) and IsTokenAt(nftExpDigits, ASection, next+2) and
          (next+3 = Length(Elements))
        then begin
          if nf = nfFixed then
          begin
            FSections[ASection].NumFormat := nfExp;
            FSections[ASection].Decimals := decs;
            exit;
          end;
        end;
      end;
    end;

    // Currency?
    if IsCurrencyAt(ASection, nf, decs, cs, clr) then
    begin
      FSections[ASection].NumFormat := nf;
      FSections[ASection].Decimals := decs;
      FSections[ASection].CurrencySymbol := cs;
      FSections[ASection].Color := clr;
      exit;
    end;

    // Look for date formats
    if IsDateAt(ASection, 0, nf, next) then begin
      if (next = Length(Elements)) then
      begin
        FSections[ASection].NumFormat := nf;
        exit;
      end;
      if IsTokenAt(nftSpace, ASection, next) and IsTimeAt(ASection, next+1, nf1, next) and
         (next = Length(Elements))
      then begin
        if (nf = nfShortDate) and (nf1 = nfShortTime) then
          FSections[ASection].NumFormat := nfShortDateTime;
      end;
      exit;
    end;

    // Look for time formats
    if IsTimeAt(ASection, 0, nf, next) then
      if next = Length(Elements) then
      begin
        FSections[ASection].NumFormat := nf;
        exit;
      end;
  end;

  // If we get here it must be a custom format.
  FSections[ASection].NumFormat := nfCustom;
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

{ Checks if a currency-type of format string begins at index AIndex, and returns
  the numberformat code, the count of decimals, the currency sambol, and the
  color.
  Note that the check is not very exact, but should cover most cases. }
function TsNumFormatParser.IsCurrencyAt(ASection: Integer;
  out ANumFormat: TsNumberFormat; out ADecimals: byte;
  out ACurrencySymbol: String; out AColor: TsColor): Boolean;
var
  hasCurrSymbol: Boolean;
  hasColor: Boolean;
  el: Integer;
begin
  Result := false;

  ANumFormat := nfCustom;
  ACurrencySymbol := '';
  ADecimals := 0;
  AColor := scNotDefined;
  hasColor := false;
  hasCurrSymbol := false;

  // Looking for the currency symbol: it is the unique identifier of the
  // currency format.
  for el := 0 to High(FSections[ASection].Elements) do
    if FSections[ASection].Elements[el].Token = nftCurrSymbol then begin
      Result := true;
      break;
    end;

  if not Result then
    exit;

  { When the format string comes from fpc it does not contain a color token.
    Color would be lost when saving. Therefore, we take the color from the
    knowledge of the NumFormat passed on creation: nfCurrencyRed has color red
    in the second section! }
  if (ASection = 1) and FHasRedSection then
    AColor := scRed;

  // Now that we know that it is a currency format analyze the elements again
  // and determine color, decimals and currency symbol.
  el := 0;
  while (el < Length(FSections[ASection].Elements)) do begin
    case FSections[ASection].Elements[el].Token of
      nftColor:
        begin
          AColor := FSections[ASection].Elements[el].IntValue;
          hasColor := true;
        end;
      nftRepeat:
        ;
      nftCurrSymbol:
        begin
          ACurrencySymbol := FSections[ASection].Elements[el].TextValue;
          hasCurrSymbol := true;
        end;
      nftOptDigit:
        if IsNumberAt(ASection, el, ANumFormat, ADecimals, el) then
          dec(el)
        else begin
          Result := false;
          exit;
        end;
      nftDigit:
        if IsNumberAt(ASection, el, ANumFormat, ADecimals, el) then
          dec(el)
        else begin
          Result := false;
          exit;
        end;
    end;
    inc(el);
  end;

  if (ASection = 1) and FHasRedSection and not hasColor then
    InsertElement(ASection, 0, nftColor, scRed);

  Result := hasCurrSymbol and ((ANumFormat = nfFixedTh) or (ASection = 2));
  if Result then begin
    if AColor = scNotDefined then ANumFormat := nfCurrency else
    if AColor = scRed then ANumFormat := nfCurrencyRed;
  end else
    ANumFormat := nfCustom;
end;

function TsNumFormatParser.IsDateAt(ASection,AIndex: Integer;
  out ANumberFormat: TsNumberFormat; var ANextIndex: Integer): Boolean;

  function CheckFormat(AFmtStr: String; var idx: Integer): Boolean;
  var
    i: Integer;
    s: String;
  begin
    Result := false;
    idx := AIndex;
    i := 1;
    while (i < Length(AFmtStr)) and (idx < Length(FSections[ASection].Elements)) do begin
      case AFmtStr[i] of
        'y', 'Y':
          begin
            if not IsTokenAt(nftYear, ASection, idx) then Exit;
            inc(idx);
            inc(i);
            while (i < Length(AFmtStr)) and (AFmtStr[i] in ['y', 'Y']) do inc(i);
          end;
        'm', 'M':
          begin
            if not IsTokenAt(nftMonth, ASection, idx) then Exit;
            inc(idx);
            inc(i);
            while (i < Length(AFmtStr)) and (AFmtStr[i] in ['m', 'M']) do inc(i);
          end;
        'd', 'D':
          begin
            if not IsTokenAt(nftDay, ASection, idx) then exit;
            inc(idx);
            inc(i);
            while (i < Length(AFmtStr)) and (AFmtStr[i] in ['d', 'D']) do inc(i);
          end;
        '/':
          begin
            if not IsTokenAt(nftDateTimeSep, ASection, idx) then exit;
            s := FSections[ASection].Elements[idx].TextValue;
            if not ((s = '/') or (s = FWorkbook.FormatSettings.DateSeparator)) then
              exit;
            inc(idx);
            inc(i);
          end;
        else
          begin
            if not (IsTokenAt(nftDateTimeSep, ASection, idx) and
                   (FSections[ASection].Elements[idx].textValue = AFmtStr[i]))
            then
              exit;
            inc(idx);
            inc(i);
          end;
      end;
    end;  // while ...
    Result := true;
    ANextIndex := idx;
  end;

begin
  if FWorkbook = nil then begin
    Result := false;
    exit;
  end;

  // The default format nfShortDate is defined by the ShortDateFormat of the
  // Workbook's FormatSettings. Check whether the current format string matches.
  // But watch out for different date separators!
  if CheckFormat(FWorkbook.FormatSettings.ShortDateFormat, ANextIndex) then begin
    Result := true;
    ANumberFormat := nfShortDate;
  end else
  // dto. with the LongDateFormat
  if CheckFormat(FWorkbook.FormatSettings.LongDateFormat, ANextIndex) then begin
    Result := true;
    ANumberFormat := nfLongDate;
  end else
    Result := false;
end;

{ Returns true if the format elements contain at least one date/time token }
function TsNumFormatParser.IsDateTimeFormat: Boolean;
var
  section: Integer;
  elem: Integer;
begin
  Result := true;
  for section := 0 to High(FSections) do
    for elem := 0 to High(FSections[section].Elements) do
      if FSections[section].Elements[elem].Token in [nftYear, nftMonth, nftDay,
        nftHour, nftMinute, nftSecond]
      then
        exit;
  Result := false;
end;

{ Checks whether the format tokens beginning at AIndex for ASection represent
  at standard number format, like nfFixed, nfPercentage etc.
  Returns TRUE if it does.
  NOTE: ADecimals can have various meanings -- see EvalNumFormatOfSection}
function TsNumFormatParser.IsNumberAt(ASection,AIndex: Integer;
  out ANumberFormat: TsNumberFormat; out ADecimals: Byte;
  out ANextIndex: Integer): Boolean;
var
  i: Integer;
begin
  Result := false;
  ANumberFormat := nfGeneral;
  ADecimals := 0;
  ANextIndex := MaxInt;
  // Let's look for digit tokens ('0') first
  if IsTokenAt(nftDigit, ASection, AIndex) then begin      // '0'
    if IsTokenAt(nftDecSep, ASection, AIndex+1) and        // '.'
       IsTokenAt(nftDecs, ASection, AIndex+2)              // count of decimals
    then begin
      // This is the case with decimal separator, like "0.000"
      Result := true;
      ANumberFormat := nfFixed;
      ADecimals := FSections[ASection].Elements[AIndex+2].IntValue;
      ANextIndex := AIndex+3;
    end else
    if not IsTokenAt(nftDecSep, ASection, AIndex+1) then begin
      // and this is the (only) case without decimal separator ("0")
      Result := true;
      ANumberFormat := nfFixed;
      ADecimals := 0;
      ANextIndex := AIndex+1;
    end;
  end else
  // Now look also for optional digits ('#')
  if IsTokenAt(nftOptDigit, ASection, AIndex) then begin   // '#'
    if IsTokenAt(nftThSep, ASection, AIndex+1) and         // ','
       IsTokenAt(nftOptDigit, ASection, AIndex+2) and      // '#'
       IsTokenAt(nftOptDigit, ASection, Aindex+3) and      // '#'
       IsTokenAt(nftDigit, ASection, AIndex+4)             // '0'
    then begin
      if IsTokenAt(nftDecSep, ASection, AIndex+5) and      // '.'
         IsTokenAt(nftDecs, ASection, AIndex+6)            // count of decimals
      then begin
        // This is the case with decimal separator, like "#,##0.000"
        Result := true;
        ANumberFormat := nfFixedTh;
        ADecimals := FSections[ASection].Elements[AIndex+6].IntValue;
        ANextIndex := AIndex+7;
      end else
      if not IsTokenAt(nftDecSep, ASection, AIndex+5) then begin
        // and this is without decimal separator, "#,##0"
        result := true;
        ANumberFormat := nfFixedTh;
        ADecimals := 0;
        ANextIndex := AIndex + 5;
      end;
    end else
    begin      // Isolated '#'
      result := true;
      inc(AIndex);
      ANextIndex := AIndex;
      ADecimals := 1;
      while IsTokenAt(nftOptDigit, ASection, AIndex) do
      begin
        inc(AIndex);
        inc(ANextIndex);
        inc(ADecimals);
      end;
      ANumberFormat := nfFraction;
    end;
  end else
  if IsTokenAt(nftOptSpaceDigit, ASection, AIndex) then // '?'
  begin
    Result := true;
    inc(AIndex);
    ANextIndex := AIndex;
    ADecimals := 1;
    while IsTokenAt(nftOptSpaceDigit, ASection, AIndex) do
    begin
      inc(AIndex);
      inc(ANextIndex);
      inc(ADecimals);
    end;
    ANumberFormat := nfFraction;
  end;
end;

function TsNumFormatParser.IsTextAt(AText: String; ASection, AIndex: Integer): Boolean;
begin
  Result := IsTokenAt(nftText, ASection, AIndex) and
    (FSections[ASection].Elements[AIndex].TextValue = AText);
end;

function TsNumFormatParser.IsTimeAt(ASection,AIndex: Integer;
  out ANumberFormat: TsNumberFormat; out ANextIndex: Integer): Boolean;

  function CheckFormat(AFmtStr: String; out idx: Integer;
    out AMPM, IsInterval: boolean): Boolean;
  var
    i: Integer;
    s: String;
  begin
    Result := false;
    AMPM := false;
    IsInterval := false;
    idx := AIndex;
    i := 1;
    while (i < Length(AFmtStr)) and (idx < Length(FSections[ASection].Elements)) do begin
      case AFmtStr[i] of
        'h', 'H':
          begin
            if not IsTokenAt(nftHour, ASection, idx) then Exit;
            if FSections[ASection].Elements[idx].IntValue < 0 then isInterval := true;
            inc(idx);
            inc(i);
            while (i < Length(AFmtStr)) and (AFmtStr[i] in ['h', 'H']) do inc(i);
          end;
        'm', 'M', 'n', 'N':
          begin
            if not IsTokenAt(nftMinute, ASection, idx) then Exit;
            if FSections[ASection].Elements[idx].IntValue < 0 then isInterval := true;
            inc(idx);
            inc(i);
            while (i < Length(AFmtStr)) and (AFmtStr[i] in ['m', 'M', 'n', 'N']) do inc(i);
          end;
        's', 'S':
          begin
            if not IsTokenAt(nftSecond, ASection, idx) then exit;
            if FSections[ASection].Elements[idx].IntValue < 0 then isInterval := true;
            inc(idx);
            inc(i);
            while (i < Length(AFmtStr)) and (AFmtStr[i] in ['s', 'S']) do inc(i);
          end;
        ':':
          begin
            if not IsTokenAt(nftDateTimeSep, ASection, idx) then exit;
            s := FSections[ASection].Elements[idx].TextValue;
            if not ((s = ':') or (s = FWorkbook.FormatSettings.DateSeparator)) then
              exit;
            inc(idx);
            inc(i);
          end;
        ' ':
          if (i+1 <= Length(AFmtStr)) and (AFmtStr[i+1] in ['a', 'A']) then begin
            inc(idx);
            inc(i);
          end else
            exit;
        'a', 'A':
          begin
            if not IsTokenAt(nftAMPM, ASection, idx) then exit;
            inc(idx);
            inc(i);
            while (i < Length(AFmtStr)) and (AFmtStr[i] in ['m','M','/','p','P']) do inc(i);
            AMPM := true;
          end;
        '[':
           begin
             if not IsTokenAt(nftHour, ASection, idx+1) then exit;
             if IsTextAt(']', ASection, idx+2) then begin
               inc(i, 3);
               inc(idx, 3);
               IsInterval := true;
             end else
             if IsTokenAt(nftHour, ASection, idx+2) and IsTextAt(']', ASection, idx+3) then begin
               inc(i, 4);
               inc(idx, 4);
               isInterval := true;
             end else
               exit;
           end
        else
          exit;
      end;
    end;
    Result := i >= Length(AFmtStr); //true;
  end;

var
  AMPM, isInterval: Boolean;
  i: Integer;
  fmt: String;
begin
  if FWorkbook = nil then begin
    Result := false;
    exit;
  end;

  Result := true;
  fmt := AddAMPM(FWorkbook.FormatSettings.LongTimeFormat, FWorkbook.FormatSettings);
  if CheckFormat(fmt, ANextIndex, AMPM, isInterval) then begin
    ANumberFormat := IfThen(AMPM, nfLongTimeAM, IfThen(isInterval, nfTimeInterval, nfLongTime));
    exit;
  end;
  fmt := FWorkbook.FormatSettings.LongTimeFormat;
  if CheckFormat(fmt, ANextIndex, AMPM, isInterval) then begin
    ANumberFormat := IfThen(AMPM, nfLongTimeAM, IfThen(isInterval, nfTimeInterval, nfLongTime));
    exit;
  end;
  fmt := AddAMPM(FWorkbook.FormatSettings.ShortTimeFormat, FWorkbook.FormatSettings);
  if CheckFormat(fmt, ANextIndex, AMPM, isInterval) then begin
    ANumberFormat := IfThen(AMPM, nfShortTimeAM, nfShortTime);
    exit;
  end;
  fmt := FWorkbook.FormatSettings.ShortTimeFormat;
  if CheckFormat(fmt, ANextIndex, AMPM, isInterval) then begin
    ANumberFormat := IfThen(AMPM, nfShortTimeAM, nfShortTime);
    exit;
  end;

  for i:=0 to High(FSections[ASection].Elements) do
    if (FSections[ASection].Elements[i].Token in [nftHour, nftMinute, nftSecond]) and
       (FSections[ASection].Elements[i].IntValue < 0)
    then begin
      ANumberFormat := nfTimeInterval;
      exit;
    end;

  Result := false;
end;

{ Returns true if the format elements contain only time, no date tokens. }
function TsNumFormatParser.IsTimeFormat: Boolean;
var
  section: Integer;
  elem: Integer;
begin
  Result := false;
  for section := 0 to High(FSections) do
    for elem := 0 to High(FSections[section].Elements) do
      if FSections[section].Elements[elem].Token in [nftHour, nftMinute, nftSecond]
      then begin
        Result := true;
      end else
      if FSections[section].Elements[elem].Token in
         [nftYear, nftMonth, nftDay, nftExpChar, nftCurrSymbol]
      then begin
        Result := false;
        exit;
      end;
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
      if FSections[j].Elements[i].Token = nftDecs then
        if FSections[j].Elements[i].IntValue > 0 then
          FSections[j].Elements[i].IntValue := 2;
end;

{ Localizes the thousand- and decimal separator symbols by replacing them with
  the FormatSettings value of the workbook. A recreated format string will be
  localized as required by Excel2. }
procedure TsNumFormatParser.Localize;
var
  i, j: Integer;
  fs: TFormatSettings;
  txt: String;
begin
  fs := FWorkbook.FormatSettings;
  for j:=0 to High(FSections) do
    for i:=0 to High(FSections[j].Elements) do begin
      txt := FSections[j].Elements[i].TextValue;
      case FSections[j].Elements[i].Token of
        nftThSep     : txt := fs.ThousandSeparator;
        nftDecSep    : txt := fs.DecimalSeparator;
        nftCurrSymbol: txt := UTF8ToAnsi(txt);
      end;
      FSections[j].Elements[i].TextValue := txt;
    end;
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
  if (AFormatString = '') or (lowercase(AFormatString) = 'general') then
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

  CheckSections;
end;

{ Scans an AM/PM sequence (or AMPM or A/P).
  At exit, cursor is a next character }
procedure TsNumFormatParser.ScanAMPM;
var
  s: String;
begin
  s := '';
  while (FCurrent < FEnd) do begin
    if (FToken in ['A', 'a', 'P', 'p', 'm', 'M', '/']) then
      s := s + FToken
    else
      break;
    FToken := NextToken;
  end;
  AddElement(nftAMPM, s);
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
begin
  s := '';
  FToken := NextToken;   // Cursor was at '['
  while (FCurrent < FEnd) and (FStatus = psOK) do begin
    case FToken of
      'h', 'H', 'm', 'M', 'n', 'N', 's', 'S':
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
  AddElement(nftCurrSymbol, s);
  if FToken <> ']' then begin
    FToken := NextToken;
    while (FCurrent < FEnd) and (FToken <> ']') do begin
      s := s + FToken;
      FToken := NextToken;
    end;
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
          Addelement(nftText, FToken);
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
    end;
    FToken := NextToken;
  end;
end;

{ Scans a floating point format. Procedure is left with the cursor at the last
  character of the format. }
procedure TsNumFormatParser.ScanNumber;
var
  hasDecSep: Boolean;
  n: Integer;
begin
  hasDecSep := false;
  while (FCurrent < FEnd) and (FStatus = psOK) do begin
    case FToken of
      ',': AddElement(nftThSep, ',');
      '.': begin
             AddElement(nftDecSep, '.');
             hasDecSep := true;
           end;
      '0': if hasDecSep then begin
             ScanAndCount('0', n);
             FToken := PrevToken;
             AddElement(nftDecs, n);
           end else
             AddElement(nftDigit, '0');
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
      '#': AddElement(nftOptDigit, FToken);
      '?': AddElement(nftOptSpaceDigit, FToken);
      '%': AddElement(nftPercent, FToken);
      '/': AddElement(nftFraction, FToken);
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
        nftDigit:
          // no decimals so far --> add decimal separator and decimals element
          if (AValue > 0) then begin
            // Don't use "AddElements" because nfCurrency etc have elements after the number.
            InsertElement(j, i, nftDecSep, '.');
            InsertElement(j, i+1, nftDecs, AValue);
            break;
          end;
        nftDecs:
          if AValue > 0 then begin
            // decimals are already used, just replace value of decimal places
            FSections[j].Elements[i].IntValue := AValue;
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
