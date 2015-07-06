{@@ ----------------------------------------------------------------------------
  Unit fpsNumFormat contains classes and procedures to analyze and process
  <b>number formats</b>.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
            distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpsNumFormat;

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

interface

uses
  Classes, SysUtils,
  fpstypes;

type
  {@@ Set of characters }
  TsDecsChars = set of char;

  {@@ Tokens used by the elements of the number format parser. If, e.g. a
    format string is "0.000" then the number format parser detects the following
    three tokens

      - nftIntZeroDigit with integer value 1  (i.e. 1 zero-digit for the integer part)
      - nftDecSep (i.e. decimal separator)
      - ntZeroDecs with integer value 2 (i.e. 3 decimal places. }
  TsNumFormatToken = (
    nftGeneral,            // token for "general" number format
    nftText,               // must be quoted, stored in TextValue
    nftThSep,              // ',', replaced by FormatSettings.ThousandSeparator
    nftDecSep,             // '.', replaced by FormatSettings.DecimalSeparator
    nftYear,               // 'y' or 'Y', count stored in IntValue
    nftMonth,              // 'm' or 'M', count stored in IntValue
    nftDay,                // 'd' or 'D', count stored in IntValue
    nftHour,               // 'h' or 'H', count stored in IntValue
    nftMinute,             // 'n' or 'N' (or 'm'/'M'), count stored in IntValue
    nftSecond,             // 's' or 'S', count stored in IntValue
    nftMilliseconds,       // 'z', 'Z', '0', count stored in IntValue
    nftAMPM,               //
    nftMonthMinute,        // 'm'/'M' or 'n'/'N', meaning depending on context
    nftDateTimeSep,        // '/' or ':', replaced by value from FormatSettings, stored in TextValue
    nftSign,               // '+' or '-', stored in TextValue
    nftSignBracket,        // '(' or ')' for negative values, stored in TextValue
    nftIntOptDigit,        // '#', count stored in IntValue
    nftIntZeroDigit,       // '0', count stored in IntValue
    nftIntSpaceDigit,      // '?', count stored in IntValue
    nftIntTh,              // '#,##0' sequence for nfFixed, count of 0 stored in IntValue
    nftZeroDecs,           // '0' after dec sep, count stored in IntValue
    nftOptDecs,            // '#' after dec sep, count stored in IntValue
    nftSpaceDecs,          // '?' after dec sep, count stored in IntValue
    nftExpChar,            // 'e' or 'E', stored in TextValue
    nftExpSign,            // '+' or '-' in exponent
    nftExpDigits,          // '0' digits in exponent, count stored in IntValue
    nftPercent,            // '%' percent symbol
    nftFactor,             // thousand separators at end of format string, each one divides value by 1000
    nftFracSymbol,         // '/' fraction symbol
    nftFracNumOptDigit,    // '#' in numerator, count stored in IntValue
    nftFracNumSpaceDigit,  // '?' in numerator, count stored in IntValue
    nftFracNumZeroDigit,   // '0' in numerator, count stored in IntValue
    nftFracDenomOptDigit,  // '#' in denominator, count stored in IntValue
    nftFracDenomSpaceDigit,// '?' in denominator, count stored in IntValue
    nftFracDenomZeroDigit, // '0' in denominator, count stored in IntValue
    nftFracDenom,          // specified denominator, value stored in IntValue
    nftCurrSymbol,         // e.g., '"€"' or '[$€]', stored in TextValue
    nftCountry,
    nftColor,              // e.g., '[red]', Color in IntValue
    nftCompareOp,
    nftCompareValue,
    nftSpace,
    nftEscaped,            // '\'
    nftRepeat,
    nftEmptyCharWidth,
    nftTextFormat);

  {@@ Element of the parsed number format sequence. Each element is identified
    by a token and has optional parameters stored as integer, float, and/or text. }
  TsNumFormatElement = record
    {@@ Token identifying the number format element }
    Token: TsNumFormatToken;
    {@@ Integer value associated with the number format element }
    IntValue: Integer;
    {@@ Floating point value associated with the number format element }
    FloatValue: Double;
    {@@ String value associated with the number format element }
    TextValue: String;
  end;

  {@@ Array of parsed number format elements }
  TsNumFormatElements = array of TsNumFormatElement;

  {@@ Summary information classifying a number format section }
  TsNumFormatKind = (nfkPercent, nfkExp, nfkCurrency, nfkFraction,
    nfkDate, nfkTime, nfkTimeInterval, nfkHasColor, nfkHasThSep, nfkHasFactor);

  {@@ Set of summary elements classifying and describing a number format section }
  TsNumFormatKinds = set of TsNumFormatKind;

  {@@ Number format string can be composed of several parts separated by a
    semicolon. The number format parser extracts the format information into
    individual sections for each part }
  TsNumFormatSection = record
    {@@ Parser number format elements used by this section }
    Elements: TsNumFormatElements;
    {@@ Summary information describing the section }
    Kind: TsNumFormatKinds;
    {@@ Reconstructed number format identifier for the built-in fps formats }
    NumFormat: TsNumberFormat;
    {@@ Number of decimal places used by the format string }
    Decimals: Byte;
    {@@ Factor by which a number will be multiplied before converting to string }
    Factor: Double;
    {@@ Digits to be used for the integer part of a fraction }
    FracInt: Integer;
    {@@ Digits to be used for the numerator part of a fraction }
    FracNumerator: Integer;
    {@@ Digits to be used for the denominator part of a fraction }
    FracDenominator: Integer;
    {@@ Currency string to be used in case of currency/accounting formats }
    CurrencySymbol: String;
    {@@ Color to be used when displaying the converted string }
    Color: TsColor;
  end;

  {@@ Pointer to a parsed number format section }
  PsNumFormatSection = ^TsNumFormatSection;

  {@@ Array of parsed number format sections }
  TsNumFormatSections = array of TsNumFormatSection;

  { TsNumFormatParams }

  {@@ Describes a parsed number format and provides all the information to
    convert a number to a number or date/time string. These data are created
    by the number format parser from a format string. }
  TsNumFormatParams = class(TObject)
  private
  protected
    function GetNumFormat: TsNumberFormat; virtual;
    function GetNumFormatStr: String; virtual;
  public
    {@@ Array of the format sections }
    Sections: TsNumFormatSections;
    procedure DeleteElement(ASectionIndex, AElementIndex: Integer);
    procedure InsertElement(ASectionIndex, AElementIndex: Integer;
      AToken: TsNumFormatToken);
    function SectionsEqualTo(ASections: TsNumFormatSections): Boolean;
    procedure SetCurrSymbol(AValue: String);
    procedure SetDecimals(AValue: Byte);
    procedure SetNegativeRed(AEnable: Boolean);
    procedure SetThousandSep(AEnable: Boolean);
    property NumFormat: TsNumberFormat read GetNumFormat;
    property NumFormatStr: String read GetNumFormatStr;
  end;


  { TsNumFormatList }

  {@@ Class of number format parameters }
  TsNumFormatParamsClass = class of TsNumFormatParams;

  {@@ List containing parsed number format parameters }
  TsNumFormatList = class(TFPList)
  private
    FOwnsData: Boolean;
    function GetItem(AIndex: Integer): TsNumFormatParams;
    procedure SetItem(AIndex: Integer; const AValue: TsNumFormatParams);
  protected
    FFormatSettings: TFormatSettings;
    FClass: TsNumFormatParamsClass;
    procedure AddBuiltinFormats; virtual;
  public
    constructor Create(AFormatSettings: TFormatSettings; AOwnsData: Boolean);
    destructor Destroy; override;
    function AddFormat(ASections: TsNumFormatSections): Integer; overload;
    function AddFormat(AFormatStr: String): Integer; overload;
    procedure Clear;
    procedure Delete(AIndex: Integer);
    function Find(ASections: TsNumFormatSections): Integer; overload;
    function Find(AFormatstr: String): Integer; overload;
    property Items[AIndex: Integer]: TsNumFormatParams read GetItem write SetItem; default;
  end;


{ Utility functions }

function AddAMPM(const ATimeFormatString: String;
  const AFormatSettings: TFormatSettings): String;
function AddIntervalBrackets(AFormatString: String): String;

procedure BuildCurrencyFormatList(AList: TStrings;
  APositive: Boolean; AValue: Double; const ACurrencySymbol: String);

function BuildCurrencyFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; ADecimals, APosCurrFmt, ANegCurrFmt: Integer;
  ACurrencySymbol: String; Accounting: Boolean = false): String;
function BuildDateTimeFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; AFormatString: String = ''): String;
function BuildFractionFormatString(AMixedFraction: Boolean;
  ANumeratorDigits, ADenominatorDigits: Integer): String;
function BuildNumberFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; ADecimals: Integer = -1): String;

function BuildFormatStringFromSection(const ASection: TsNumFormatSection): String;

function ConvertFloatToStr(AValue: Double; AParams: TsNumFormatParams;
  AFormatSettings: TFormatSettings): String;
function CountDecs(AFormatString: String; ADecChars: TsDecsChars = ['0']): Byte;

function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsCurrencyFormat(ANumFormat: TsNumFormatParams): Boolean; overload;

function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsDateTimeFormat(AFormatStr: String): Boolean; overload;
function IsDateTimeFormat(ANumFormat: TsNumFormatParams): Boolean; overload;

function IsDateFormat(ANumFormat: TsNumFormatParams): Boolean;

function IsTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsTimeFormat(AFormatStr: String): Boolean; overload;
function IsTimeFormat(ANumFormat: TsNumFormatParams): Boolean; overload;
function IsLongTimeFormat(AFormatStr: String; ATimeSeparator: char): Boolean; overload;

function IsTimeIntervalFormat(ANumFormat: TsNumFormatParams): Boolean;

function MakeLongDateFormat(ADateFormat: String): String;
function MakeShortDateFormat(ADateFormat: String): String;
procedure MakeTimeIntervalMask(Src: String; var Dest: String);
function StripAMPM(const ATimeFormatString: String): String;


implementation

uses
  StrUtils, Math,
  fpsUtils, fpsNumFormatParser;

const
  {@@ Array of format strings identifying the order of number and
    currency symbol of a positive currency value. The number is expected at
    index 0, the currency symbol at index 1 of the parameter array used by the
    fpc Format() function. }
  POS_CURR_FMT: array[0..3] of string = (
    ('%1:s%0:s'),        // 0: $1
    ('%0:s%1:s'),        // 1: 1$
    ('%1:s %0:s'),       // 2: $ 1
    ('%0:s %1:s')        // 3: 1 $
  );
  {@@ Array of format strings identifying the order of number and
    currency symbol of a negative currency value. The sign is shown
    as a dash character ("-") or by means of brackets. The number
    is expected at index 0, the currency symbol at index 1 of the
    parameter array for the fpc Format() function. }
  NEG_CURR_FMT: array[0..15] of string = (
    ('(%1:s%0:s)'),      //  0: ($1)
    ('-%1:s%0:s'),       //  1: -$1
    ('%1:s-%0:s'),       //  2: $-1
    ('%1:s%0:s-'),       //  3: $1-
    ('(%0:s%1:s)'),      //  4: (1$)
    ('-%0:s%1:s'),       //  5: -1$
    ('%0:s-%1:s'),       //  6: 1-$
    ('%0:s%1:s-'),       //  7: 1$-
    ('-%0:s %1:s'),      //  8: -1 $
    ('-%1:s %0:s'),      //  9: -$ 1
    ('%0:s %1:s-'),      // 10: 1 $-
    ('%1:s %0:s-'),      // 11: $ 1-
    ('%1:s -%0:s'),      // 12: $ -1
    ('%0:s- %1:s'),      // 13: 1- $
    ('(%1:s %0:s)'),     // 14: ($ 1)
    ('(%0:s %1:s)')      // 15: (1 $)
  );

{==============================================================================}
{                         Float-to-string conversion                           }
{==============================================================================}

type
  {@@ Set of parsed number format tokens }
  TsNumFormatTokenSet = set of TsNumFormatToken;

const
  {@@ Set of tokens which terminate number information in a format string }
  TERMINATING_TOKENS: TsNumFormatTokenSet =
    [nftSpace, nftText, nftEscaped, nftPercent, nftCurrSymbol, nftSign, nftSignBracket];
  {@@ Set of tokens which describe the integer part of a number format }
  INT_TOKENS: TsNumFormatTokenSet =
    [nftIntOptDigit, nftIntZeroDigit, nftIntSpaceDigit];
  {@@ Set of tokens which describe the decimals of a number format }
  DECS_TOKENS: TsNumFormatTokenSet =
    [nftZeroDecs, nftOptDecs, nftSpaceDecs];
  {@@ Set of tokens which describe the numerator of a fraction format }
  FRACNUM_TOKENS: TsNumFormatTokenSet =
    [nftFracNumOptDigit, nftFracNumZeroDigit, nftFracNumSpaceDigit];
  {@@ Set of tokens which describe the denominator of a fraction format }
  FRACDENOM_TOKENS: TsNumFormatTokenSet =
    [nftFracDenomOptDigit, nftFracDenomZeroDigit, nftFracDenomSpaceDigit, nftFracDenom];
  {@@ Set of tokens which describe the exponent in exponential formatting of a number }
  EXP_TOKENS: TsNumFormatTokenSet =
    [nftExpDigits];   // todo: expand by optional digits (0.00E+#)

{ Helper function which checks whether a sequence of format tokens for
  exponential formatting begins at the specified index in the format elements }
function CheckExp(const AElements: TsNumFormatElements; AIndex: Integer): Boolean;
var
  numEl: Integer;
  i: Integer;
begin
  numEl := Length(AElements);

  Result := (AIndex < numEl) and (AElements[AIndex].Token in INT_TOKENS);
  if not Result then
    exit;

  numEl := Length(AElements);
  i := AIndex + 1;
  while (i < numEl) and (AElements[i].Token in INT_TOKENS) do inc(i);

  // no decimal places
  if (i+2 < numEl) and
     (AElements[i].Token = nftExpChar) and
     (AElements[i+1].Token = nftExpSign) and
     (AElements[i+2].Token in EXP_TOKENS)
  then begin
    Result := true;
    exit;
  end;

  // with decimal places
  if (i < numEl) and (AElements[i].Token = nftDecSep) //and (AElements[i+1].Token in DECS_TOKENS)
  then begin
    inc(i);
    while (i < numEl) and (AElements[i].Token in DECS_TOKENS) do inc(i);
    if (i + 2 < numEl) and
       (AElements[i].Token = nftExpChar) and
       (AElements[i+1].Token = nftExpSign) and
       (AElements[i+2].Token in EXP_TOKENS)
    then begin
      Result := true;
      exit;
    end;
  end;

  Result := false;
end;

{ Helper function which checks whether a sequence of format tokens for
  fraction formatting begins at the specified index in the format elements }
function CheckFraction(const AElements: TsNumFormatElements; AIndex: Integer;
  out digits: Integer): Boolean;
var
  numEl: Integer;
  i: Integer;
begin
  digits := 0;
  numEl := Length(AElements);

  Result := (AIndex < numEl);
  if not Result then
    exit;

  i := AIndex;
  // Check for mixed fraction (integer split off, sample format "# ??/??"
  if (AElements[i].Token in (INT_TOKENS + [nftIntTh])) then
  begin
    inc(i);
    while (i < numEl) and (AElements[i].Token in (INT_TOKENS + [nftIntTh])) do inc(i);
    while (i < numEl) and (AElements[i].Token in TERMINATING_TOKENS) do inc(i);
  end;

  if (i = numEl) or not (AElements[i].Token in FRACNUM_TOKENS) then
    exit(false);

  // Here follows the ordinary fraction (no integer split off); sample format "??/??"
  while (i < numEl) and (AElements[i].Token in FRACNUM_TOKENS) do inc(i);
  while (i < numEl) and (AElements[i].Token in TERMINATING_TOKENS) do inc(i);
  if (i = numEl) or (AElements[i].Token <> nftFracSymbol) then
    exit(False);

  inc(i);
  while (i < numEl) and (AElements[i].Token in TERMINATING_TOKENS) do inc(i);
  if (i = numEl) or (not (AElements[i].Token in FRACDENOM_TOKENS)) then
    exit(false);

  while (i < numEL) and (AElements[i].Token in FRACDENOM_TOKENS) do
  begin
    case AElements[i].Token of
      nftFracDenomZeroDigit : inc(digits, AElements[i].IntValue);
      nftFracDenomSpaceDigit: inc(digits, AElements[i].IntValue);
      nftFracDenomOptDigit  : inc(digits, AElements[i].IntValue);
      nftFracDenom          : digits := -AElements[i].IntValue;  // "-" indicates a literal denominator value!
    end;
    inc(i);
  end;
  Result := true;
end;

{ Processes a sequence of #, 0, and ? tokens.
  Adds leading (GrowRight=false) or trailing (GrowRight=true) zeros and/or
  spaces as specified by the format elements to the number value string.
  On exit AIndex points to the first non-integer token. }
function ProcessIntegerFormat(AValue: String; AFormatSettings: TFormatSettings;
  const AElements: TsNumFormatElements; var AIndex: Integer;
  ATokens: TsNumFormatTokenSet; GrowRight, UseThSep: Boolean): String;
const
  OptTokens = [nftIntOptDigit, nftFracNumOptDigit, nftFracDenomOptDigit, nftOptDecs];
  ZeroTokens = [nftIntZeroDigit, nftFracNumZeroDigit, nftFracDenomZeroDigit, nftZeroDecs, nftIntTh];
  SpaceTokens = [nftIntSpaceDigit, nftFracNumSpaceDigit, nftFracDenomSpaceDigit, nftSpaceDecs];
  AllOptTokens = OptTokens + SpaceTokens;
var
  fs: TFormatSettings absolute AFormatSettings;
  i, j, L: Integer;
  numEl: Integer;
begin
  Result := AValue;
  numEl := Length(AElements);
  if GrowRight then
  begin
    // This branch is intended for decimal places, i.e. there may be trailing zeros.
    i := AIndex;
    if (AValue = '0') and (AElements[i].Token in AllOptTokens) then
      Result := '';
    // Remove trailing zeros
    while (Result <> '') and (Result[Length(Result)] = '0')
      do Delete(Result, Length(Result), 1);
    // Add trailing zeros or spaces as required by the elements.
    i := AIndex;
    L := 0;
    while (i < numEl) and (AElements[i].Token in ATokens) do
    begin
      if AElements[i].Token in ZeroTokens then
      begin
        inc(L, AElements[i].IntValue);
        while Length(Result) < L do Result := Result + '0'
      end else
      if AElements[i].Token in SpaceTokens then
      begin
        inc(L, AElements[i].IntValue);
        while Length(Result) < L do Result := Result + ' ';
      end;
      inc(i);
    end;
    if UseThSep then begin
      j := 2;
      while (j < Length(Result)) and (Result[j-1] <> ' ') and (Result[j] <> ' ') do
      begin
        Insert(fs.ThousandSeparator, Result, 1);
        inc(j, 3);
      end;
    end;
    AIndex := i;
  end else
  begin
    // This branch is intended for digits (or integer and numerator parts of fractions)
    // --> There are no leading zeros.
    // Find last digit token of the sequence
    i := AIndex;
    while (i < numEl) and (AElements[i].Token in ATokens) do
      inc(i);
    j := i;
    if i > 0 then dec(i);
    if (AValue = '0') and (AElements[i].Token in AllOptTokens) and (i = AIndex) then
      Result := '';
    // From the end of the sequence, going backward, add leading zeros or spaces
    // as required by the elements of the format.
    L := 0;
    while (i >= AIndex) do begin
      if AElements[i].Token in ZeroTokens then
      begin
        inc(L, AElements[i].IntValue);
        while Length(Result) < L do Result := '0' + Result;
      end else
      if AElements[i].Token in SpaceTokens then
      begin
        inc(L, AElements[i].IntValue);
        while Length(Result) < L do Result := ' ' + Result;
      end;
      dec(i);
    end;
    AIndex := j;
    if UseThSep then
    begin
     // AIndex := j + 1;
      j := Length(Result) - 2;
      while (j > 1) and (Result[j-1] <> ' ') and (Result[j] <> ' ') do
      begin
        Insert(fs.ThousandSeparator, Result, j);
        dec(j, 3);
      end;
    end;
  end;
end;

{ Converts the floating point number to an exponential number string according
  to the format specification in AElements.
  It must have been verified before, that the elements in fact are valid for
  an exponential format. }
function ProcessExpFormat(AValue: Double; AFormatSettings: TFormatSettings;
  const AElements: TsNumFormatElements; var AIndex: Integer): String;
var
  fs: TFormatSettings absolute AFormatSettings;
  expchar: String;
  expSign: String;
  se, si, sd: String;
  decs, expDigits: Integer;
  intZeroDigits, intOptDigits, intSpaceDigits: Integer;
  numStr: String;
  i, id, p: Integer;
  numEl: Integer;
begin
  Result := '';
  numEl := Length(AElements);

  // Determine digits of integer part of mantissa
  intZeroDigits := 0;
  intOptDigits := 0;
  intSpaceDigits := 0;
  i := AIndex;
  while (AElements[i].Token in INT_TOKENS) do begin
    case AElements[i].Token of
      nftIntZeroDigit : inc(intZeroDigits, AElements[i].IntValue);
      nftIntSpaceDigit: inc(intSpaceDigits, AElements[i].IntValue);
      nftIntOptDigit  : inc(intOptDigits, AElements[i].IntValue);
    end;
    inc(i);
  end;

  // No decimal places
  if (i + 2 < numEl) and (AElements[i].Token = nftExpChar) then
  begin
    expChar := AElements[i].TextValue;
    expSign := AElements[i+1].TextValue;
    expDigits := 0;
    i := i+2;
    while (i < numEl) and (AElements[i].Token in EXP_TOKENS) do
    begin
      inc(expDigits, AElements[i].IntValue);  // not exactly what Excel does... Rather exotic case...
      inc(i);
    end;
    numstr := FormatFloat('0'+expChar+expSign+DupeString('0', expDigits), AValue, fs);
    p := pos('e', Lowercase(numStr));
    se := copy(numStr, p, Length(numStr));   // exp part of the number string, incl "E"
    numStr := copy(numstr, 1, p-1);          // mantissa of the number string
    numStr := ProcessIntegerFormat(numStr, fs, AElements, AIndex, INT_TOKENS, false, false);
    Result := numStr + se;
    AIndex := i;
  end
  else
  // With decimal places
  if (i + 1 < numEl) and (AElements[i].Token = nftDecSep) then
  begin
    inc(i);
    id := i;     // index of decimal elements
    decs := 0;
    while (i < numEl) and (AElements[i].Token in DECS_TOKENS) do
    begin
      case AElements[i].Token of
        nftZeroDecs,
        nftSpaceDecs: inc(decs, AElements[i].IntValue);
      end;
      inc(i);
    end;
    expChar := AElements[i].TextValue;
    expSign := AElements[i+1].TextValue;
    expDigits := 0;
    inc(i, 2);
    while (i < numEl) and (AElements[i].Token in EXP_TOKENS) do
    begin
      inc(expDigits, AElements[i].IntValue);
      inc(i);
    end;
    if decs=0 then
      numstr := FormatFloat('0'+expChar+expSign+DupeString('0', expDigits), AValue, fs)
    else
      numStr := FloatToStrF(AValue, ffExponent, decs+1, expDigits, fs);
    if (abs(AValue) >= 1.0) and (expSign = '-') then
      Delete(numStr, pos('+', numStr), 1);
    p := pos('e', Lowercase(numStr));
    se := copy(numStr, p, Length(numStr));    // exp part of the number string, incl "E"
    numStr := copy(numStr, 1, p-1);           // mantissa of the number string
    p := pos(fs.DecimalSeparator, numStr);
    if p = 0 then
    begin
      si := numstr;
      sd := '';
    end else
    begin
      si := ProcessIntegerFormat(copy(numStr, 1, p-1), fs, AElements, AIndex, INT_TOKENS, false, false);  // integer part of the mantissa
      sd := ProcessIntegerFormat(copy(numStr, p+1, Length(numStr)), fs, AElements, id, DECS_TOKENS, true, false);  // fractional part of the mantissa
    end;
    // Put all parts together...
    Result := si + fs.DecimalSeparator + sd + se;
    AIndex := i;
  end;
end;

function ProcessFracFormat(AValue: Double; const AFormatSettings: TFormatSettings;
  ADigits: Integer; const AElements: TsNumFormatElements;
  var AIndex: Integer): String;
var
  fs: TFormatSettings absolute AFormatSettings;
  frint, frnum, frdenom, maxdenom: Int64;
  sfrint, sfrnum, sfrdenom: String;
  sfrsym, sintnumspace, snumsymspace, ssymdenomspace: String;
  i, numEl: Integer;
begin
  sintnumspace := '';
  snumsymspace := '';
  ssymdenomspace := '';
  sfrsym := '/';
  if ADigits >= 0 then
    maxDenom := Round(IntPower(10, ADigits));
  numEl := Length(AElements);

  i := AIndex;
  if AElements[i].Token in (INT_TOKENS + [nftIntTh]) then begin
    // Split-off integer
    if (AValue >= 1) then
    begin
      frint := trunc(AValue);
      AValue := frac(AValue);
    end else
      frint := 0;
    if ADigits >= 0 then
      FloatToFraction(AValue, maxdenom, frnum, frdenom)
    else
    begin
      frdenom := -ADigits;
      frnum := round(AValue*frdenom);
    end;
    sfrint := ProcessIntegerFormat(IntToStr(frint), fs, AElements, i,
      INT_TOKENS + [nftIntTh], false, (AElements[i].Token = nftIntTh));
    while (i < numEl) and (AElements[i].Token in TERMINATING_TOKENS) do
    begin
      sintnumspace := sintnumspace + AElements[i].TextValue;
      inc(i);
    end;
  end else
  begin
    // "normal" fraction
    sfrint := '';
    if ADigits > 0 then
      FloatToFraction(AValue, maxdenom, frnum, frdenom)
    else
    begin
      frdenom := -ADigits;
      frnum := round(AValue*frdenom);
    end;
    sintnumspace := '';
  end;

  // numerator and denominator
  sfrnum := ProcessIntegerFormat(IntToStr(frnum), fs, AElements, i,
    FRACNUM_TOKENS, false, false);
  while (i < numEl) and (AElements[i].Token in TERMINATING_TOKENS) do
  begin
    snumsymspace := snumsymspace + AElements[i].TextValue;
    inc(i);
  end;
  inc(i);  // fraction symbol
  while (i < numEl) and (AElements[i].Token in TERMINATING_TOKENS) do
  begin
    ssymdenomspace := ssymdenomspace + AElements[i].TextValue;
    inc(i);
  end;

  sfrdenom := ProcessIntegerFormat(IntToStr(frdenom), fs, AElements, i,
    FRACDENOM_TOKENS, false, false);
  AIndex := i+1;

  // Special cases
  if (frnum = 0) then
  begin
    if sfrnum = '' then begin
      sintnumspace := '';
      snumsymspace := '';
      ssymdenomspace := '';
      sfrdenom := '';
      sfrsym := '';
    end else
    if trim(sfrnum) = '' then begin
      sfrdenom := DupeString(' ', Length(sfrdenom));
      sfrsym := ' ';
    end;
  end;
  if sfrint = '' then sintnumspace := '';

  // Compose result string
  Result := sfrnum + snumsymspace + sfrsym + ssymdenomspace + sfrdenom;
  if (Trim(Result) = '') and (sfrint = '') then
    sfrint := '0';
  if sfrint <> '' then
    Result := sfrint + sintnumSpace + result;
end;

function ProcessFloatFormat(AValue: Double; AFormatSettings: TFormatSettings;
  const AElements: TsNumFormatElements; var AIndex: Integer): String;
var
  fs: TFormatSettings absolute AFormatSettings;
  numEl: Integer;
  numStr, s: String;
  p, i: Integer;
  decs: Integer;
  useThSep: Boolean;
begin
  Result := '';
  numEl := Length(AElements);

  // Extract integer part
  Result := IntToStr(trunc(AValue));
  useThSep := AElements[AIndex].Token = nftIntTh;
  Result := ProcessIntegerFormat(Result, fs, AElements, AIndex,
    (INT_TOKENS + [nftIntTh]), false, UseThSep);

  // Decimals
  if (AIndex < numEl) and (AElements[AIndex].Token = nftDecSep) then
  begin
    inc(AIndex);
    i := AIndex;
    // Count decimal digits in format elements
    decs := 0;
    while (AIndex < numEl) and (AElements[AIndex].Token in DECS_TOKENS) do begin
      inc(decs, AElements[AIndex].IntValue);
      inc(AIndex);
    end;
    // Convert value to string
    numstr := FloatToStrF(AValue, ffFixed, MaxInt, decs, fs);
    p := Pos(fs.DecimalSeparator, numstr);
    s := Copy(numstr, p+1, Length(numstr));
    s := ProcessIntegerFormat(s, fs, AElements, i, DECS_TOKENS, true, false);
    if s <> '' then
      Result := Result + fs.DecimalSeparator + s;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Converts a floating point number to a string as determined by the specified
  number format parameters

  @param AValue           Value to be converted to a string
  @param AParams          Number format parameters which will be applied in the
                          conversion. The number format params are obtained
                          by the number format parser from the number format
                          string.
  @param AFormatSettings  Format settings needed by the number format parser for
                          the conversion
  @return Converted string
-------------------------------------------------------------------------------}
function ConvertFloatToStr(AValue: Double; AParams: TsNumFormatParams;
  AFormatSettings: TFormatSettings): String;
var
  fs: TFormatSettings absolute AFormatSettings;
  sidx: Integer;
  section: TsNumFormatSection;
  i, el, numEl: Integer;
  isNeg: Boolean;
  yr, mon, day, hr, min, sec, ms: Word;
  s: String;
  digits: Integer;
begin
  Result := '';
  if IsNaN(AValue) then
    exit;

  if AParams = nil then
  begin
    Result := FloatToStrF(AValue, ffGeneral, 20, 20, fs);
    exit;
  end;

  sidx := 0;
  if (AValue < 0) and (Length(AParams.Sections) > 1) then
    sidx := 1;
  if (AValue = 0) and (Length(AParams.Sections) > 2) then
    sidx := 2;
  isNeg := (AValue < 0);
  AValue := abs(AValue);   // section 0 adds the sign back, section 1 has the sign in the elements
  section := AParams.Sections[sidx];
  numEl := Length(section.Elements);

  if nfkPercent in section.Kind then
    AValue := AValue * 100.0;
  if nfkHasFactor in section.Kind then
    AValue := AValue * section.Factor;
  if nfkTime in section.Kind then
    DecodeTime(AValue, hr, min, sec, ms);
  if nfkDate in section.Kind then
    DecodeDate(AValue, yr, mon, day);

  el := 0;
  while (el < numEl) do begin
    if section.Elements[el].Token = nftGeneral then
    begin
      s := FloatToStrF(AValue, ffGeneral, 20, 20, fs);
      if (sidx=0) and isNeg then s := '-' + s;
      Result := Result + s;
    end
    else
    // Integer token: can be the start of a number, exp, or mixed fraction format
    // Cases with thousand separator are handled here as well.
    if section.Elements[el].Token in (INT_TOKENS + [nftIntTh]) then begin
      // Check for exponential format
      if CheckExp(section.Elements, el) then
        s := ProcessExpFormat(AValue, fs, section.Elements, el)
      else
      // Check for fraction format
      if CheckFraction(section.Elements, el, digits) then
        s := ProcessFracFormat(AValue, fs, digits, section.Elements, el)
      else
      // Floating-point or integer
        s := ProcessFloatFormat(AValue, fs, section.Elements, el);
      if (sidx = 0) and isNeg then s := '-' + s;
      Result := Result + s;
      Continue;
    end
    else
    // Regular fraction (without integer being split off)
    if (section.Elements[el].Token in FRACNUM_TOKENS) and
       CheckFraction(section.Elements, el, digits) then
    begin
      s := ProcessFracFormat(AValue, fs, digits, section.Elements, el);
      if (sidx = 0) and isNeg then s := '-' + s;
      Result := Result + s;
      Continue;
    end
    else
      case section.Elements[el].Token of
        nftSpace, nftText, nftEscaped, nftCurrSymbol,
        nftSign, nftSignBracket, nftPercent:
          Result := Result + section.Elements[el].TextValue;

        nftEmptyCharWidth:
          Result := Result + ' ';

        nftDateTimeSep:
          case section.Elements[el].TextValue of
            '/': Result := Result + fs.DateSeparator;
            ':': Result := Result + fs.TimeSeparator;
            else Result := Result + section.Elements[el].TextValue;
          end;

        nftDecSep:
          Result := Result + fs.DecimalSeparator;

        nftThSep:
          Result := Result + fs.ThousandSeparator;

        nftYear:
          case section.Elements[el].IntValue of
            1,
            2: Result := Result + IfThen(yr mod 100 < 10, '0'+IntToStr(yr mod 100), IntToStr(yr mod 100));
            4: Result := Result + IntToStr(yr);
          end;

        nftMonth:
          case section.Elements[el].IntValue of
            1: Result := Result + IntToStr(mon);
            2: Result := Result + IfThen(mon < 10, '0'+IntToStr(mon), IntToStr(mon));
            3: Result := Result + fs.ShortMonthNames[mon];
            4: Result := Result + fs.LongMonthNames[mon];
          end;

        nftDay:
          case section.Elements[el].IntValue of
            1: result := result + IntToStr(day);
            2: result := Result + IfThen(day < 10, '0'+IntToStr(day), IntToStr(day));
            3: Result := Result + fs.ShortDayNames[DayOfWeek(AValue)];
            4: Result := Result + fs.LongDayNames[DayOfWeek(AValue)];
          end;

        nftHour:
          begin
            if section.Elements[el].IntValue < 0 then  // This case is for nfTimeInterval
              s := IntToStr(Int64(hr) + trunc(AValue) * 24)
            else
            if section.Elements[el].TextValue = 'AM' then  // This tag is set in case of AM/FM format
            begin
              hr := hr mod 12;
              if hr = 0 then hr := 12;
              s := IntToStr(hr)
            end else
              s := IntToStr(hr);
            if (abs(section.Elements[el].IntValue) = 2) and (Length(s) = 1) then
              s := '0' + s;
            Result := Result + s;
          end;

        nftMinute:
          begin
            if section.Elements[el].IntValue < 0 then  // case for nfTimeInterval
              s := IntToStr(int64(min) + trunc(AValue) * 24 * 60)
            else
              s := IntToStr(min);
            if (abs(section.Elements[el].IntValue) = 2) and (Length(s) = 1) then
              s := '0' + s;
            Result := Result + s;
          end;

       nftSecond:
          begin
            if section.Elements[el].IntValue < 0 then  // case for nfTimeInterval
              s := IntToStr(Int64(sec) + trunc(AValue) * 24 * 60 * 60)
            else
              s := IntToStr(sec);
            if (abs(section.Elements[el].IntValue) = 2) and (Length(s) = 1) then
              s := '0' + s;
            Result := Result + s;
          end;

        nftMilliseconds:
          case section.Elements[el].IntValue of
            1: Result := Result + IntToStr(ms div 100);
            2: Result := Result + Format('%02d', [ms div 10]);
            3: Result := Result + Format('%03d', [ms]);
          end;

        nftAMPM:
          begin
            s := section.Elements[el].TextValue;
            if lowercase(s) = 'ampm' then
              s := IfThen(frac(AValue) < 0.5, fs.TimeAMString, fs.TimePMString)
            else
            begin
              i := pos('/', s);
              if i > 0 then
                s := IfThen(frac(AValue) < 0.5, copy(s, 1, i-1), copy(s, i+1, Length(s)))
              else
                s := IfThen(frac(AValue) < 0.5, 'AM', 'PM');
            end;
            Result := Result + s;
          end;
      end;  // case
    inc(el);
  end;  // while
end;


{==============================================================================}
{                           Utility functions                                  }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Adds an AM/PM format code to a pre-built time formatting string. The strings
  replacing "AM" or "PM" in the final formatted number are taken from the
  TimeAMString or TimePMString of the specified FormatSettings.

  @param   ATimeFormatString  String of time formatting codes (such as 'hh:nn')
  @param   AFormatSettings    FormatSettings for locale-dependent information
  @result  Formatting string with AM/PM option activated.

  Example:  ATimeFormatString = 'hh:nn' ==> 'hh:nn AM/PM'
-------------------------------------------------------------------------------}
function AddAMPM(const ATimeFormatString: String;
  const AFormatSettings: TFormatSettings): String;
var
  am, pm: String;
  fs: TFormatSettings absolute AFormatSettings;
begin
  am := IfThen(fs.TimeAMString <> '', fs.TimeAMString, 'AM');
  pm := IfThen(fs.TimePMString <> '', fs.TimePMString, 'PM');
  Result := Format('%s %s/%s', [StripAMPM(ATimeFormatString), am, pm]);
end;

{@@ ----------------------------------------------------------------------------
  The given format string is assumed to represent a time interval, i.e. its
  first time symbol must be enclosed by square brackets. Checks if this is true,
  and adds the brackes if not.

  @param   AFormatString   String with time formatting codes
  @return  Unchanged format string if its first time code is in square brackets
           (as in '[h]:nn:ss'). If not, the first time code is enclosed in
           square brackets.
-------------------------------------------------------------------------------}
function AddIntervalBrackets(AFormatString: String): String;
var
  p: Integer;
  s1, s2: String;
begin
  if AFormatString[1] = '[' then
    Result := AFormatString
  else begin
    p := pos(':', AFormatString);
    if p <> 0 then begin
      s1 := copy(AFormatString, 1, p-1);
      s2 := copy(AFormatString, p, Length(AFormatString));
      Result := Format('[%s]%s', [s1, s2]);
    end else
      Result := Format('[%s]', [AFormatString]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Builds a string list with samples of the predefined currency formats

  @param  AList            String list in which the format samples are stored
  @param  APositive        If true, samples are built for positive currency
                           values, otherwise for negative values
  @param  AValue           Currency value to be used when calculating the sample
                           strings
  @param  ACurrencySymbol  Currency symbol string to be used in the samples
-------------------------------------------------------------------------------}
procedure BuildCurrencyFormatList(AList: TStrings;
  APositive: Boolean; AValue: Double; const ACurrencySymbol: String);
var
  valueStr: String;
  i: Integer;
begin
  valueStr := Format('%.0n', [AValue]);
  AList.BeginUpdate;
  try
    if AList.Count = 0 then
    begin
      if APositive then
        for i:=0 to High(POS_CURR_FMT) do
          AList.Add(Format(POS_CURR_FMT[i], [valueStr, ACurrencySymbol]))
      else
        for i:=0 to High(NEG_CURR_FMT) do
          AList.Add(Format(NEG_CURR_FMT[i], [valueStr, ACurrencySymbol]));
    end else
    begin
      if APositive then
        for i:=0 to High(POS_CURR_FMT) do
          AList[i] := Format(POS_CURR_FMT[i], [valueStr, ACurrencySymbol])
      else
        for i:=0 to High(NEG_CURR_FMT) do
          AList[i] := Format(NEG_CURR_FMT[i], [valueStr, ACurrencySymbol]);
    end;
  finally
    AList.EndUpdate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Builds a currency format string. The presentation of negative values (brackets,
  or minus signs) is taken from the provided format settings. The format string
  consists of three sections, separated by semicolons.

  @param  ANumberFormat   Identifier of the built-in number format for which the
                          format string is to be generated.
  @param  AFormatSettings FormatSettings to be applied (used to extract default
                          values for the parameters following)
  @param  ADecimals       number of decimal places. If < 0, the CurrencyDecimals
                          of the FormatSettings is used.
  @param  APosCurrFmt     Identifier for the order of currency symbol, value and
                          spaces of positive values
                          - see pcfXXXX constants in fpsTypes.pas.
                          If < 0, the CurrencyFormat of the FormatSettings is used.
  @param  ANegCurrFmt     Identifier for the order of currency symbol, value and
                          spaces of negative values. Specifies also usage of ().
                          - see ncfXXXX constants in fpsTypes.pas.
                          If < 0, the NegCurrFormat of the FormatSettings is used.
  @param  ACurrencySymbol String to identify the currency, like $ or USD.
                          If ? the CurrencyString of the FormatSettings is used.
  @param  Accounting      If true, adds spaces for alignment of decimals

  @return                 String of formatting codes

  @example                '"$"#,##0.00;("$"#,##0.00);"$"0.00'
-------------------------------------------------------------------------------}
function BuildCurrencyFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings;
  ADecimals, APosCurrFmt, ANegCurrFmt: Integer; ACurrencySymbol: String;
  Accounting: Boolean = false): String;
var
  decs: String;
  pcf, ncf: Byte;
  p, n: String;
  negRed: Boolean;
begin
  pcf := IfThen(APosCurrFmt < 0, AFormatSettings.CurrencyFormat, APosCurrFmt);
  ncf := IfThen(ANegCurrFmt < 0, AFormatSettings.NegCurrFormat, ANegCurrFmt);
  if (ADecimals < 0) then
    ADecimals := AFormatSettings.CurrencyDecimals;
  if ACurrencySymbol = '?' then
    ACurrencySymbol := AFormatSettings.CurrencyString;
  if ACurrencySymbol <> '' then
    ACurrencySymbol := '"' + ACurrencySymbol + '"';
  decs := DupeString('0', ADecimals);
  if ADecimals > 0 then decs := '.' + decs;

  negRed := (ANumberFormat = nfCurrencyRed);
  p := POS_CURR_FMT[pcf];   // Format mask for positive values
  n := NEG_CURR_FMT[ncf];   // Format mask for negative values

  // add extra space for the sign of the number for perfect alignment in Excel
  if Accounting then
    case ncf of
      0, 14: p := p + '_)';
      3, 11: p := p + '_-';
       4, 15: p := '_(' + p;
      5, 8 : p := '_-' + p;
    end;

  if ACurrencySymbol <> '' then begin
    Result := Format(p, ['#,##0' + decs, ACurrencySymbol]) + ';'
            + IfThen(negRed, '[red]', '')
            + Format(n, ['#,##0' + decs, ACurrencySymbol]) + ';'
            + Format(p, ['0'+decs, ACurrencySymbol]);
  end
  else begin
    Result := '#,##0' + decs;
    if negRed then
      Result := Result +';[red]'
    else
      Result := Result +';';
    case ncf of
      0, 14, 15           : Result := Result + '(#,##0' + decs + ')';
      1, 2, 5, 6, 8, 9, 12: Result := Result + '-#,##0' + decs;
      else                  Result := Result + '#,##0' + decs + '-';
    end;
    Result := Result + ';0' + decs;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Builds a date/time format string from the number format code.

  @param   ANumberFormat    built-in number format identifier
  @param   AFormatSettings  Format settings from which locale-dependent
                            information like day-month-year order is taken.
  @param   AFormatString    Optional pre-built formatting string. It is used
                            only for the format nfInterval where square brackets
                            are added to the first time code field.
  @return  String of date/time formatting code constructed from the built-in
           format information.
-------------------------------------------------------------------------------}
function BuildDateTimeFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; AFormatString: String = '') : string;
var
  i, j: Integer;
  Unwanted: set of ansichar;
begin
  case ANumberFormat of
    nfShortDateTime:
      Result := AFormatSettings.ShortDateFormat + ' ' + AFormatSettings.ShortTimeFormat;
      // In the DefaultFormatSettings this is: d/m/y hh:nn
    nfShortDate:
      Result := AFormatSettings.ShortDateFormat;   // --> d/m/y
    nfLongDate:
      Result := AFormatSettings.LongDateFormat;    // --> dd mm yyyy
    nfShortTime:
      Result := StripAMPM(AFormatSettings.ShortTimeFormat);    // --> hh:nn
    nfLongTime:
      Result := StripAMPM(AFormatSettings.LongTimeFormat);     // --> hh:nn:ss
    nfShortTimeAM:
      begin                                       // --> hh:nn AM/PM
        Result := AFormatSettings.ShortTimeFormat;
        if (pos('a', lowercase(AFormatSettings.ShortTimeFormat)) = 0) then
          Result := AddAMPM(Result, AFormatSettings);
      end;
    nfLongTimeAM:                                 // --> hh:nn:ss AM/PM
      begin
        Result := AFormatSettings.LongTimeFormat;
        if pos('a', lowercase(AFormatSettings.LongTimeFormat)) = 0 then
          Result := AddAMPM(Result, AFormatSettings);
      end;
    nfDayMonth,                                  // --> dd/mmm
    nfMonthYear:                                 // --> mmm/yy
      begin
        Result := AFormatSettings.ShortDateFormat;
        case ANumberFormat of
          nfDayMonth:
            unwanted := ['y', 'Y'];
          nfMonthYear:
            unwanted := ['d', 'D'];
        end;
        for i:=Length(Result) downto 1 do
          if Result[i] in unwanted then Delete(Result, i, 1);
        while not (Result[1] in (['m', 'M', 'd', 'D', 'y', 'Y'] - unwanted)) do
          Delete(Result, 1, 1);
        while not (Result[Length(Result)] in (['m', 'M', 'd', 'D', 'y', 'Y'] - unwanted)) do
          Delete(Result, Length(Result), 1);
        i := 1;
        while not (Result[i] in ['m', 'M']) do inc(i);
        j := i;
        while (j <= Length(Result)) and (Result[j] in ['m', 'M']) do inc(j);
        while (j - i < 3) do begin
          Insert(Result[i], Result, j);
          inc(j);
        end;
      end;
    nfTimeInterval:                               // --> [h]:nn:ss
      if AFormatString = '' then
        Result := '[h]:nn:ss'
      else
        Result := AddIntervalBrackets(AFormatString);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Builds a number format string for fraction formatting from the number format
  code and the count of numerator and denominator digits.

  @param  AMixedFraction     If TRUE, fraction is presented as mixed fraction
  @param  ANumeratorDigits   Count of numerator digits
  @param  ADenominatorDigits Count of denominator digits. If the value is negative
                             then its absolute value is inserted literally as
                             as denominator.

  @return String of formatting code, here something like: '##/##' or '# ##/##'
-------------------------------------------------------------------------------}
function BuildFractionFormatString(AMixedFraction: Boolean;
  ANumeratorDigits, ADenominatorDigits: Integer): String;
begin
  if ADenominatorDigits < 0 then  // a negative value indicates a fixed denominator value
    Result := Format('%s/%d', [
      DupeString('?', ANumeratorDigits), -ADenominatorDigits
    ])
  else
    Result := Format('%s/%s', [
      DupeString('?', ANumeratorDigits), DupeString('?', ADenominatorDigits)
    ]);
  if AMixedFraction then
    Result := '# ' + Result;
end;

{@@ ----------------------------------------------------------------------------
  Builds a number format string from the number format code and the count of
  decimal places.

  @param  ANumberFormat   Identifier of the built-in numberformat for which a
                          format string is to be generated
  @param  AFormatSettings FormatSettings for default parameters
  @param  ADecimals       Number of decimal places. If < 0 the CurrencyDecimals
                          value of the FormatSettings is used. In case of a
                          fraction format "ADecimals" refers to the maximum count
                          digits of the denominator.

  @return String of formatting codes

  @example  ANumberFormat = nfFixedTh, ADecimals = 2 --> '#,##0.00'
-------------------------------------------------------------------------------}
function BuildNumberFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; ADecimals: Integer = -1): String;
var
  decs: String;
begin
  Result := '';
  if ADecimals = -1 then
    ADecimals := AFormatSettings.CurrencyDecimals;
  decs := DupeString('0', ADecimals);
  if ADecimals > 0 then decs := '.' + decs;
  case ANumberFormat of
    nfFixed:
      Result := '0' + decs;
    nfFixedTh:
      Result := '#,##0' + decs;
    nfExp:
      Result := '0' + decs + 'E+00';
    nfPercentage:
      Result := '0' + decs + '%';
    nfFraction:
      if ADecimals = 0 then    // "ADecimals" has a different meaning here...
        Result := '# ??/??'    // This is the default fraction format
      else
      begin
        decs := DupeString('?', ADecimals);
        Result := '# ' + decs + '/' + decs;
      end;
    nfCurrency, nfCurrencyRed:
      Result := BuildCurrencyFormatString(ANumberFormat, AFormatSettings,
        ADecimals, AFormatSettings.CurrencyFormat, AFormatSettings.NegCurrFormat,
        AFormatSettings.CurrencyString);
    nfShortDateTime, nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfDayMonth, nfMonthYear, nfTimeInterval:
      raise Exception.Create('BuildNumberFormatString: Use BuildDateTimeFormatSstring '+
        'to create a format string for date/time values.');
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a format string for the specified parsed number format section.
  The format string is created according to Excel convention (which is understood
  by ODS as well).

  @param  ASection  Parsed section of number format elements as created by the
                    number format parser
  @return Excel-compatible format string
-------------------------------------------------------------------------------}
function BuildFormatStringFromSection(const ASection: TsNumFormatSection): String;
var
  element: TsNumFormatElement;
  i, n: Integer;
begin
  Result := '';

  for i := 0 to High(ASection.Elements)  do begin
    element := ASection.Elements[i];
    case element.Token of
      nftGeneral:
        Result := Result + 'General';
      nftIntOptDigit, nftOptDecs, nftFracNumOptDigit, nftFracDenomOptDigit:
        if element.IntValue > 0 then
          Result := Result + DupeString('#', element.IntValue);
      nftIntZeroDigit, nftZeroDecs, nftFracNumZeroDigit, nftFracDenomZeroDigit, nftExpDigits:
        if element.IntValue > 0 then
          Result := result + DupeString('0', element.IntValue);
      nftIntSpaceDigit, nftSpaceDecs, nftFracNumSpaceDigit, nftFracDenomSpaceDigit:
        if element.Intvalue > 0 then
          Result := result + DupeString('?', element.IntValue);
      nftFracDenom:
        Result := Result + IntToStr(element.IntValue);
      nftIntTh:
        case element.Intvalue of
          0: Result := Result + '#,###';
          1: Result := Result + '#,##0';
          2: Result := Result + '#,#00';
          3: Result := Result + '#,000';
        end;
      nftDecSep, nftThSep:
        Result := Result + element.TextValue;
      nftFracSymbol:
        Result := Result + '/';
      nftPercent:
        Result := Result + '%';
      nftFactor:
        if element.IntValue <> 0 then
        begin
          n := element.IntValue;
          while (n > 0) do
          begin
            Result := Result + element.TextValue;
            dec(n);
          end;
        end;
      nftSpace:
        Result := Result + ' ';
      nftText:
        if element.TextValue <> '' then result := Result + '"' + element.TextValue + '"';
      nftYear:
        Result := Result + DupeString('Y', element.IntValue);
      nftMonth:
        Result := Result + DupeString('M', element.IntValue);
      nftDay:
        Result := Result + DupeString('D', element.IntValue);
      nftHour:
        if element.IntValue < 0
          then Result := Result + '[' + DupeString('h', -element.IntValue) + ']'
          else Result := Result + DupeString('h', element.IntValue);
      nftMinute:
        if element.IntValue < 0
          then Result := result + '[' + DupeString('m', -element.IntValue) + ']'
          else Result := Result + DupeString('m', element.IntValue);
      nftSecond:
        if element.IntValue < 0
          then Result := Result + '[' + DupeString('s', -element.IntValue) + ']'
          else Result := Result + DupeString('s', element.IntValue);
      nftMilliseconds:
        Result := Result + DupeString('0', element.IntValue);
      nftSign, nftSignBracket, nftExpChar, nftExpSign, nftAMPM, nftDateTimeSep:
        if element.TextValue <> '' then Result := Result + element.TextValue;
      nftCurrSymbol:
        if element.TextValue <> '' then
          Result := Result + '[$' + element.TextValue + ']';
      nftEscaped:
        if element.TextValue <> '' then
          Result := Result + '\' + element.TextValue;
      nftTextFormat:
        if element.TextValue <> '' then
          Result := Result + element.TextValue;
      nftRepeat:
        if element.TextValue <> '' then Result := Result + '*' + element.TextValue;
      nftColor:
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

{@@ ----------------------------------------------------------------------------
  Counts how many decimal places are coded into a given number format string.

  @param   AFormatString   String with number format codes, such as '0.000'
  @param   ADecChars       Characters which are considered as symbols for decimals.
                           For the fixed decimals, this is the '0'. Optional
                           decimals are encoced as '#'.
  @return  Count of decimal places found (3 in above example).
-------------------------------------------------------------------------------}
function CountDecs(AFormatString: String; ADecChars: TsDecsChars = ['0']): Byte;
var
  i: Integer;
begin
  Result := 0;
  i := 1;
  while (i <= Length(AFormatString)) do begin
    if AFormatString[i] = '.' then begin
      inc(i);
      while (i <= Length(AFormatString)) and (AFormatString[i] in ADecChars) do begin
        inc(i);
        inc(Result);
      end;
      exit;
    end else
      inc(i);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given number format code is for currency,
  i.e. requires a currency symbol.

  @param  AFormat   Built-in number format identifier to be checked
  @return True if AFormat is nfCurrency or nfCurrencyRed, false otherwise.
-------------------------------------------------------------------------------}
function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [nfCurrency, nfCurrencyRed];
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to currency values.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkCurrency elements; false otherwise
-------------------------------------------------------------------------------}
function IsCurrencyFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkCurrency] <> []);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given number format code is for date/time values.

  @param   AFormat  Built-in number format identifier to be checked
  @return  True if AFormat is a date/time format (such as nfShortTime),
           false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [nfShortDateTime, nfShortDate, nfLongDate,
    nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM,
    nfDayMonth, nfMonthYear, nfTimeInterval];
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given string with formatting codes is for date/time values.

  @param   AFormatStr   String with formatting codes to be checked.
  @return  True if AFormatStr is a date/time format string (such as 'hh:nn'),
           false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(AFormatStr: string): Boolean;
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(AFormatStr, DefaultFormatSettings);
  try
    Result := parser.IsDateTimeFormat;
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to date/time values.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkDate or nfkTime elements; false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkDate, nfkTime] <> []);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to a date value.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkDate, but no nfkTime tags; false otherwise
-------------------------------------------------------------------------------}
function IsDateFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkDate, nfkTime] = [nfkDate]);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given built-in number format code is for time values.

  @param   AFormat  Built-in number format identifier to be checked
  @return  True if AFormat represents to a time-format, false otherwise
-------------------------------------------------------------------------------}
function IsTimeFormat(AFormat: TsNumberFormat): boolean;
begin
  Result := AFormat in [nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM,
    nfTimeInterval];
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given string with formatting codes is for time values.

  @param   AFormatStr   String with formatting codes to be checked
  @return  True if AFormatStr represents a time-format, false otherwise
-------------------------------------------------------------------------------}
function IsTimeFormat(AFormatStr: String): Boolean;
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(AFormatStr, DefaultFormatSettings);
  try
    Result := parser.IsTimeFormat;
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to time values.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkTime, but no nfkDate elements; false otherwise
-------------------------------------------------------------------------------}
function IsTimeFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkTime, nfkDate] = [nfkTime]);
end;

{@@ ----------------------------------------------------------------------------
  Returns TRUE if the specified format string represents a long time format, i.e.
  it contains two TimeSeparators.
-------------------------------------------------------------------------------}
function IsLongTimeFormat(AFormatStr: String; ATimeSeparator: Char): Boolean;
var
  i, n: Integer;
begin
  n := 0;
  for i:=1 to Length(AFormatStr) do
    if AFormatStr[i] = ATimeSeparator then inc(n);
  Result := (n=2);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters is a time interval
  format.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkTimeInterval elements; false otherwise
-------------------------------------------------------------------------------}
function IsTimeIntervalFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkTimeInterval] <> []);
end;

{@@ ----------------------------------------------------------------------------
  Creates a long date format string out of a short date format string.
  Retains the order of year-month-day and the separators, but uses 4 digits
  for year and 3 digits of month.

  @param  ADateFormat   String with date formatting code representing a
                        "short" date, such as 'dd/mm/yy'
  @return Format string modified to represent a "long" date, such as 'dd/mmm/yyyy'
-------------------------------------------------------------------------------}
function MakeLongDateFormat(ADateFormat: String): String;
var
  i: Integer;
begin
  Result := '';
  i := 1;
  while i < Length(ADateFormat) do begin
    case ADateFormat[i] of
      'y', 'Y':
        begin
          Result := Result + DupeString(ADateFormat[i], 4);
          while (i < Length(ADateFormat)) and (ADateFormat[i] in ['y','Y']) do
            inc(i);
        end;
      'm', 'M':
        begin
          result := Result + DupeString(ADateFormat[i], 3);
          while (i < Length(ADateFormat)) and (ADateFormat[i] in ['m','M']) do
            inc(i);
        end;
      else
        Result := Result + ADateFormat[i];
        inc(i);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Modifies the short date format such that it has a two-digit year and a two-digit
  month. Retains the order of year-month-day and the separators.

  @param   ADateFormat   String with date formatting codes representing a
                         "long" date, such as 'dd/mmm/yyyy'
  @return  Format string modified to represent a "short" date, such as 'dd/mm/yy'
-------------------------------------------------------------------------------}
function MakeShortDateFormat(ADateFormat: String): String;
var
  i: Integer;
begin
  Result := '';
  i := 1;
  while i < Length(ADateFormat) do begin
    case ADateFormat[i] of
      'y', 'Y':
        begin
          Result := Result + DupeString(ADateFormat[i], 2);
          while (i < Length(ADateFormat)) and (ADateFormat[i] in ['y','Y']) do
            inc(i);
        end;
      'm', 'M':
        begin
          result := Result + DupeString(ADateFormat[i], 2);
          while (i < Length(ADateFormat)) and (ADateFormat[i] in ['m','M']) do
            inc(i);
        end;
      else
        Result := Result + ADateFormat[i];
        inc(i);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a "time interval" format string having the first time code identifier
  in square brackets.

  @param  Src   Source format string, must be a time format string, like 'hh:nn'
  @param  Dest  Destination format string, will have the first time code element
                of the src format string in square brackets, like '[hh]:nn'.
-------------------------------------------------------------------------------}
procedure MakeTimeIntervalMask(Src: String; var Dest: String);
var
  L: TStrings;
begin
  L := TStringList.Create;
  try
    L.StrictDelimiter := true;
    L.Delimiter := ':';
    L.DelimitedText := Src;
    if L[0][1] <> '[' then L[0] := '[' + L[0];
    if L[0][Length(L[0])] <> ']' then L[0] := L[0] + ']';
    Dest := L.DelimitedText;
  finally
    L.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes an AM/PM formatting code from a given time formatting string. Variants
  of "AM/PM" are considered as well. The string is left unchanged if it does not
  contain AM/PM codes.

  @param   ATimeFormatString  String of time formatting codes (such as 'hh:nn AM/PM')
  @return  Formatting string with AM/PM being removed (--> 'hh:nn')
-------------------------------------------------------------------------------}
function StripAMPM(const ATimeFormatString: String): String;
var
  i: Integer;
begin
  Result := '';
  i := 1;
  while i <= Length(ATimeFormatString) do begin
    if ATimeFormatString[i] in ['a', 'A'] then begin
      inc(i);
      while (i <= Length(ATimeFormatString)) and (ATimeFormatString[i] in ['p', 'P', 'm', 'M', '/'])  do
        inc(i);
    end else
      Result := Result + ATimeFormatString[i];
    inc(i);
  end;
end;


{==============================================================================}
{                             TsNumFormatParams                                }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Deletes a parsed number format element from the specified format section.

  @param  ASectionIndex   Index of the format section containing the element to
                          be deleted
  @param  AElementIndex   Index of the format element to be deleted
-------------------------------------------------------------------------------}
procedure TsNumFormatParams.DeleteElement(ASectionIndex, AElementIndex: Integer);
var
  i, n: Integer;
begin
  with Sections[ASectionIndex] do
  begin
    n := Length(Elements);
    for i := AElementIndex+1 to n-1 do
      Elements[i-1] := Elements[i];
    SetLength(Elements, n-1);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates the built-in number format identifier from the parsed number format
  sections and elements

  @return  Built-in number format identifer if the format is built into
           fpspreadsheet, or nfCustom otherwise

  @see     TsNumFormat
-------------------------------------------------------------------------------}
function TsNumFormatParams.GetNumFormat: TsNumberFormat;
begin
  Result := nfCustom;
  case Length(Sections) of
    0: Result := nfGeneral;
    1: Result := Sections[0].NumFormat;
    2: if (Sections[0].NumFormat = Sections[1].NumFormat) and
          (Sections[0].NumFormat in [nfCurrency, nfCurrencyRed])
       then
         Result := Sections[0].NumFormat;
    3: if (Sections[0].NumFormat = Sections[1].NumFormat) and
          (Sections[1].NumFormat = Sections[2].NumFormat) and
          (Sections[0].NumFormat in [nfCurrency, nfCurrencyRed])
       then
         Result := Sections[0].NumFormat;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Constructs the number format string from the parsed sections and elements.
  The format symbols are selected according to Excel syntax.

  @return  Excel-compatible number format string.
-------------------------------------------------------------------------------}
function TsNumFormatParams.GetNumFormatStr: String;
var
  i: Integer;
begin
  if Length(Sections) > 0 then begin
    Result := BuildFormatStringFromSection(Sections[0]);
    for i := 1 to High(Sections) do
      Result := Result + ';' + BuildFormatStringFromSection(Sections[i]);
  end else
    Result := '';
end;

{@@ ----------------------------------------------------------------------------
  Inserts a parsed format token into the specified format section before the
  specified element.

  @param  ASectionIndex   Index of the parsed format section into which the
                          token is to be inserted
  @param  AElementIndex   Index of the format element before which the token
                          is to be inserted
  @param  AToken          Parsed format token to be inserted

  @see    TsNumFormatToken
-------------------------------------------------------------------------------}
procedure TsNumFormatParams.InsertElement(ASectionIndex, AElementIndex: Integer;
  AToken: TsNumFormatToken);
var
  i, n: Integer;
begin
  with Sections[ASectionIndex] do
  begin
    n := Length(Elements);
    SetLength(Elements, n+1);
    for i:=n-1 downto AElementIndex do
      Elements[i+1] := Elements[i];
    Elements[AElementIndex].Token := AToken;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the parsed format sections passed as a parameter are identical
  to the interal section array.

  @param  ASections  Array of parsed format sections to be compared with the
                     internal format sections
-------------------------------------------------------------------------------}
function TsNumFormatParams.SectionsEqualTo(ASections: TsNumFormatSections): Boolean;
var
  i, j: Integer;
begin
  Result := false;
  if Length(ASections) <> Length(Sections) then
    exit;
  for i := 0 to High(Sections) do begin
    if Length(Sections[i].Elements) <> Length(ASections[i].Elements) then
      exit;

    for j:=0 to High(Sections[i].Elements) do
    begin
      if Sections[i].Elements[j].Token <> ASections[i].Elements[j].Token then
        exit;

      if Sections[i].NumFormat <> ASections[i].NumFormat then
        exit;
      if Sections[i].Decimals <> ASections[i].Decimals then
        exit;
      {
      if Sections[i].Factor <> ASections[i].Factor then
        exit;
        }
      if Sections[i].FracInt <> ASections[i].FracInt then
        exit;
      if Sections[i].FracNumerator <> ASections[i].FracNumerator then
        exit;
      if Sections[i].FracDenominator <> ASections[i].FracDenominator then
        exit;
      if Sections[i].CurrencySymbol <> ASections[i].CurrencySymbol then
        exit;
      if Sections[i].Color <> ASections[i].Color then
        exit;

      case Sections[i].Elements[j].Token of
        nftText, nftThSep, nftDecSep, nftDateTimeSep,
        nftAMPM, nftSign, nftSignBracket,
        nftExpChar, nftExpSign, nftPercent, nftFracSymbol, nftCurrSymbol,
        nftCountry, nftSpace, nftEscaped, nftRepeat, nftEmptyCharWidth,
        nftTextFormat:
          if Sections[i].Elements[j].TextValue <> ASections[i].Elements[j].TextValue
            then exit;

        nftYear, nftMonth, nftDay,
        nftHour, nftMinute, nftSecond, nftMilliseconds,
        nftMonthMinute,
        nftIntOptDigit, nftIntZeroDigit, nftIntSpaceDigit, nftIntTh,
        nftZeroDecs, nftOptDecs, nftSpaceDecs, nftExpDigits, nftFactor,
        nftFracNumOptDigit, nftFracNumSpaceDigit, nftFracNumZeroDigit,
        nftFracDenomOptDigit, nftFracDenomSpaceDigit, nftFracDenomZeroDigit,
        nftColor:
          if Sections[i].Elements[j].IntValue <> ASections[i].Elements[j].IntValue
            then exit;

        nftCompareOp, nftCompareValue:
          if Sections[i].Elements[j].FloatValue <> ASections[i].Elements[j].FloatValue
            then exit;
      end;
    end;
  end;
  Result := true;
end;

{@@ ----------------------------------------------------------------------------
  Defines the currency symbol used in the format params sequence

  @param  AValue  String containing the currency symbol to be used in the
                  converted numbers
-------------------------------------------------------------------------------}
procedure TsNumFormatParams.SetCurrSymbol(AValue: String);
var
  section: TsNumFormatSection;
  s, el: Integer;
begin
  for s:=0 to High(Sections) do
  begin
    section := Sections[s];
    if (nfkCurrency in section.Kind) then
    begin
      section.CurrencySymbol := AValue;
      for el := 0 to High(section.Elements) do
        if section.Elements[el].Token = nftCurrSymbol then
          section.Elements[el].Textvalue := AValue;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds or modifies parsed format tokens such that the specified number of
  decimal places is displayed

  @param  AValue  Number of decimal places to be shown
-------------------------------------------------------------------------------}
procedure TsNumFormatParams.SetDecimals(AValue: byte);
var
  section: TsNumFormatSection;
  s, el: Integer;
begin
  for s := 0 to High(Sections) do
  begin
    section := Sections[s];
    if section.Kind * [nfkFraction, nfkDate, nfkTime] <> [] then
      Continue;
    section.Decimals := AValue;
    for el := High(section.Elements) downto 0 do
      case section.Elements[el].Token of
        nftZeroDecs:
          section.Elements[el].Intvalue := AValue;
        nftOptDecs, nftSpaceDecs:
          DeleteElement(s, el);
      end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  If AEnable is true a format section for negative numbers is added (or an
  existing one is modified) such that negative numbers are displayed in red.
  If AEnable is false the format tokens are modified such that negative values
  are displayed in default color.

  @param  AEnable  The format tokens are modified such as to display negative
                   values in red if AEnable is true.
-------------------------------------------------------------------------------}
procedure TsNumFormatParams.SetNegativeRed(AEnable: Boolean);
var
  el: Integer;
begin
  // Enable negative-value color
  if AEnable then
  begin
    if Length(Sections) = 1 then begin
      SetLength(Sections, 2);
      Sections[1] := Sections[0];
      InsertElement(1, 0, nftColor);
      Sections[1].Elements[0].Intvalue := scRed;
      InsertElement(1, 1, nftSign);
      Sections[1].Elements[1].TextValue := '-';
    end else
    begin
      if not (nfkHasColor in Sections[1].Kind) then
        InsertElement(1, 0, nftColor);
      for el := 0 to High(Sections[1].Elements) do
        if Sections[1].Elements[el].Token = nftColor then
          Sections[1].Elements[el].IntValue := scRed;
    end;
    Sections[1].Kind := Sections[1].Kind + [nfkHasColor];
    Sections[1].Color := scRed;
  end else
  // Disable negative-value color
  if Length(Sections) >= 2 then
  begin
    Sections[1].Kind := Sections[1].Kind - [nfkHasColor];
    Sections[1].Color := scBlack;
    for el := High(Sections[1].Elements) downto 0 do
      if Sections[1].Elements[el].Token = nftColor then
        DeleteElement(1, el);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Inserts a thousand separator token into the format elements at the
  appropriate position, or removes it

  @param  AEnable   A thousand separator is inserted if AEnable is true, or else
                    deleted.
-------------------------------------------------------------------------------}
procedure TsNumFormatParams.SetThousandSep(AEnable: Boolean);
var
  section: TsNumFormatSection;
  s, el: Integer;
  replaced: Boolean;
begin
  for s := 0 to High(Sections) do
  begin
    section := Sections[s];
    replaced := false;
    for el := High(section.Elements) downto 0 do
    begin
      if AEnable then
      begin
        if section.Elements[el].Token in [nftIntOptDigit, nftIntSpaceDigit, nftIntZeroDigit] then
        begin
          if replaced then
            DeleteElement(s, el)
          else begin
            section.Elements[el].Token := nftIntTh;
            Include(section.Kind, nfkHasThSep);
            replaced := true;
          end;
        end;
      end else
      begin
        if section.Elements[el].Token = nftIntTh then begin
          section.Elements[el].Token := nftIntZeroDigit;
          Exclude(section.Kind, nfkHasThSep);
          break;
        end;
      end;
    end;
  end;
end;


{==============================================================================}
{                           TsNumFormatList                                    }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Constructor of the number format list class.

  @param  AFormatSettings   Format settings needed internally by the number
                            format parser (currency symbol, etc.)
  @param  AOwnsData         If true then the list is responsible to destroy
                            the list items
-------------------------------------------------------------------------------}
constructor TsNumFormatList.Create(AFormatSettings: TFormatSettings;
  AOwnsData: Boolean);
begin
  inherited Create;
  FClass := TsNumFormatParams;
  FFormatSettings := AFormatSettings;
  FOwnsData := AOwnsData;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the number format list class.

  Clears the list items if the list "owns" the data.
-------------------------------------------------------------------------------}
destructor TsNumFormatList.Destroy;
begin
  Clear;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Adds the specified sections of a parsed number format to the list.
  Duplicates are not checked before adding the format item.

  @param  ASections   Array of number format sections as obtained by the
                      number format parser for a given format string
  @return Index of the format item in the list.
-------------------------------------------------------------------------------}
function TsNumFormatList.AddFormat(ASections: TsNumFormatSections): Integer;
var
  nfp: TsNumFormatParams;
begin
  Result := Find(ASections);
  if Result = -1 then begin
    nfp := FClass.Create;
    nfp.Sections := ASections;
    Result := inherited Add(nfp);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format as specified by a format string to the list
  Uses the number format parser to convert the format string to format sections
  and elements.

  Duplicates are not checked before adding the format item.

  @param  AFormatStr  Excel-like format string describing the format to be added
  @return Index of the format item in the list
-------------------------------------------------------------------------------}
function TsNumFormatList.AddFormat(AFormatStr: String): Integer;
var
  parser: TsNumFormatParser;
  newSections: TsNumFormatSections;
  i: Integer;
begin
  parser := TsNumFormatParser.Create(AFormatStr, FFormatSettings);
  try
    SetLength(newSections, parser.ParsedSectionCount);
    for i:=0 to High(newSections) do
      newSections[i] := parser.ParsedSections[i];
    Result := AddFormat(newSections);
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds the number formats to the list which are built into the file format.

  Does nothing here. Must be overridden by derived classes for each file format.
-------------------------------------------------------------------------------}
procedure TsNumFormatList.AddBuiltinFormats;
begin
end;

{@@ ----------------------------------------------------------------------------
  Clears the list.
  If the list "owns" the format items they are destroyed.

  @see  TsNumFormatList.Create
-------------------------------------------------------------------------------}
procedure TsNumFormatList.Clear;
var
  i: Integer;
begin
  for i := Count-1 downto 0 do Delete(i);
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the number format item having the specified index in the list.
  If the list "owns" the format items, the item is destroyed.

  @param  AIndex  Index of the format item to be deleted
  @see TsNumformatList.Create
-------------------------------------------------------------------------------}
procedure TsNumFormatList.Delete(AIndex: Integer);
var
  p: TsNumFormatParams;
begin
  if FOwnsData then
  begin
    p := GetItem(AIndex);
    if p <> nil then p.Free;
  end;
  inherited Delete(AIndex);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a parsed format item having the specified format sections is
  contained in the list and returns its index if found, or -1 if not found.

  @param  ASections   Array of number format sections as obtained by the
                      number format parser for a given format string
  @return Index of the found format item, or -1 if not found
-------------------------------------------------------------------------------}
function TsNumFormatList.Find(ASections: TsNumFormatSections): Integer;
var
  nfp: TsNumFormatParams;
begin
  for Result := 0 to Count-1 do begin
    nfp := GetItem(Result);
    if nfp.SectionsEqualTo(ASections) then
      exit;
  end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a format item corresponding to the specified format string is
  contained in the list and returns its index if found, or -1 if not.

  Should be called before adding a format to the list to avoid duplicates.

  @param   AFormatStr  Number format string of the format item which is seeked
  @return  Index of the found format item, or -1 if not found
  @see     TsNumFormatList.Add
-------------------------------------------------------------------------------}
function TsNumFormatList.Find(AFormatStr: String): Integer;
var
  nfp: TsNumFormatParams;
begin
  nfp := CreateNumFormatParams(AFormatStr, FFormatSettings);
  if nfp = nil then
    Result := -1
  else
    Result := Find(nfp.Sections);
end;

{@@ ----------------------------------------------------------------------------
  Getter function returning the correct type of the list items
  (i.e., TsNumFormatParams which are parsed format descriptions).

  @param  AIndex   Index of the format item
  @return Pointer to the list item at the specified index, cast to the type
          TsNumFormatParams
-------------------------------------------------------------------------------}
function TsNumFormatList.GetItem(AIndex: Integer): TsNumFormatParams;
begin
  Result := TsNumFormatParams(inherited Items[AIndex]);
end;

{@@ ----------------------------------------------------------------------------
  Setter function for the list items

  @param  AIndex  Index of the format item
  @param  AValue  Pointer to the parsed format description to be stored in the
                  list at the specified index.
-------------------------------------------------------------------------------}
procedure TsNumFormatList.SetItem(AIndex: Integer;
  const AValue: TsNumFormatParams);
begin
  inherited Items[AIndex] := AValue;
end;


end.
