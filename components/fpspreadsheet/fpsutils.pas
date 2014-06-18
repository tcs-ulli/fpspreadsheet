{
  Utility functions and constants from FPSpreadsheet
}

// to do: Remove the patched FormatDateTime when the feature of square brackets
//        in time format codes is in the rtl

unit fpsutils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, StrUtils, fpspreadsheet, fpsNumFormatParser;

// Exported types
type
  TsSelectionDirection = (fpsVerticalSelection, fpsHorizontalSelection);
  TsDecsChars = set of char;

  // to be removed when fpc trunk is stable
  TFormatDateTimeOption = (fdoInterval);
  TFormatDateTimeOptions =  set of TFormatDateTimeOption;

const
  // Date formatting string for unambiguous date/time display as strings
  // Can be used for text output when date/time cell support is not available
  ISO8601Format='yyyymmdd"T"hhmmss';
  // Extended ISO 8601 date/time format, used in e.g. ODF/opendocument
  ISO8601FormatExtended='yyyy"-"mm"-"dd"T"hh":"mm":"ss';
  // ISO 8601 time-only format, used in ODF/opendocument
  ISO8601FormatTimeOnly='"PT"hh"H"nn"M"ss"S"';

// Endianess helper functions
function WordToLE(AValue: Word): Word;
function DWordToLE(AValue: Cardinal): Cardinal;
function IntegerToLE(AValue: Integer): Integer;
function WideStringToLE(const AValue: WideString): WideString;

function WordLEtoN(AValue: Word): Word;
function DWordLEtoN(AValue: Cardinal): Cardinal;
function WideStringLEToN(const AValue: WideString): WideString;

function LongRGBToExcelPhysical(const RGB: DWord): DWord;

// Other routines
function ParseIntervalString(const AStr: string;
  var AFirstCellRow, AFirstCellCol, ACount: Integer;
  var ADirection: TsSelectionDirection): Boolean;
function ParseCellRangeString(const AStr: string;
  var AFirstCellRow, AFirstCellCol, ALastCellRow, ALastCellCol: Integer;
  var AFlags: TsRelFlags): Boolean;
function ParseCellString(const AStr: string;
  var ACellRow, ACellCol: Integer; var AFlags: TsRelFlags): Boolean; overload;
function ParseCellString(const AStr: string;
  var ACellRow, ACellCol: Integer): Boolean; overload;
function ParseCellRowString(const AStr: string;
  var AResult: Integer): Boolean;
function ParseCellColString(const AStr: string;
  var AResult: Integer): Boolean;

function GetColString(AColIndex: Integer): String;
function GetCellString(ARow,ACol: Cardinal; AFlags: TsRelFlags): String;
function GetCellRangeString(ARow1, ACol1, ARow2, ACol2: Cardinal; AFlags: TsRelFlags): String;

function GetErrorValueStr(AErrorValue: TsErrorValue): String;

function UTF8TextToXMLText(AText: ansistring): ansistring;

function IfThen(ACondition: Boolean; AValue1,AValue2: TsNumberFormat): TsNumberFormat; overload;

function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean;
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsDateTimeFormat(AFormatStr: String): Boolean; overload;
function IsTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsTimeFormat(AFormatStr: String): Boolean; overload;

function BuildCurrencyFormatString(ADialect: TsNumFormatDialect;
  ANumberFormat: TsNumberFormat; const AFormatSettings: TFormatSettings;
  ADecimals, APosCurrFormat, ANegCurrFormat: Integer; ACurrencySymbol: String): String;
function BuildDateTimeFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; AFormatString: String = ''): String;
function BuildNumberFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; ADecimals: Integer = -1): String;

function AddAMPM(const ATimeFormatString: String;
  const AFormatSettings: TFormatSettings): String;
function StripAMPM(const ATimeFormatString: String): String;
function CountDecs(AFormatString: String; ADecChars: TsDecsChars = ['0']): Byte;
function AddIntervalBrackets(AFormatString: String): String;
function MakeLongDateFormat(AShortDateFormat: String): String;
function MakeShortDateFormat(AShortDateFormat: String): String;
function SpecialDateTimeFormat(ACode: String;
  const AFormatSettings: TFormatSettings; ForWriting: Boolean): String;
function SplitAccountingFormatString(const AFormatString: String; ASection: ShortInt;
  var ALeft, ARight: String): Byte;
procedure SplitFormatString(const AFormatString: String; out APositivePart,
  ANegativePart, AZeroPart: String);

function SciFloat(AValue: Double; ADecimals: Byte): String; overload;
function SciFloat(AValue: Double; ADecimals: Byte; AFormatSettings: TFormatSettings): String; overload;
procedure MakeTimeIntervalMask(Src: String; var Dest: String);

// These two functions are copies of fpc trunk until they are available in stable fpc.
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  Options : TFormatDateTimeOptions = []): string;
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []): string;

function TwipsToPts(AValue: Integer): Single;
function PtsToTwips(AValue: Single): Integer;
function cmToPts(AValue: Double): Double;
function PtsToCm(AValue: Double): Double;
function InToPts(AValue: Double): Double;
function mmToPts(AValue: Double): Double;
function PtsToMM(AValue: Double): Double;
function pxToPts(AValue, AScreenPixelsPerInch: Integer): Double;
function PtsToPx(AValue: Double; AScreenPixelsPerInch: Integer): Integer;
function HTMLLengthStrToPts(AValue: String): Double;

function HTMLColorStrToColor(AValue: String): TsColorValue;
function ColorToHTMLColorStr(AValue: TsColorValue): String;

var
  ScreenPixelsPerInch: Integer = 96;

implementation

uses
  Math;

{
  Endianess helper functions

  Excel files are all written with Little Endian byte order,
  so it's necessary to swap the data to be able to build a
  correct file on big endian systems.

  These routines are preferable to System unit routines because they
  ensure that the correct overloaded version of the conversion routines
  will be used, avoiding typecasts which are less readable.

  They also guarantee delphi compatibility. For Delphi we just support
  big-endian isn't support, because Delphi doesn't support it.
}

function WordToLE(AValue: Word): Word;
begin
  {$IFDEF FPC}
    Result := NtoLE(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

function DWordToLE(AValue: Cardinal): Cardinal;
begin
  {$IFDEF FPC}
    Result := NtoLE(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

function IntegerToLE(AValue: Integer): Integer;
begin
  {$IFDEF FPC}
    Result := NtoLE(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

function WordLEtoN(AValue: Word): Word;
begin
  {$IFDEF FPC}
    Result := LEtoN(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

function DWordLEtoN(AValue: Cardinal): Cardinal;
begin
  {$IFDEF FPC}
    Result := LEtoN(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

function WideStringToLE(const AValue: WideString): WideString;
var
  j: integer;
begin
  {$IFDEF FPC}
    {$IFDEF FPC_LITTLE_ENDIAN}
      Result:=AValue;
    {$ELSE}
      Result:=AValue;
      for j := 1 to Length(AValue) do begin
        PWORD(@Result[j])^:=NToLE(PWORD(@Result[j])^);
      end;
    {$ENDIF}
  {$ELSE}
    Result:=AValue;
  {$ENDIF}
end;

function WideStringLEToN(const AValue: WideString): WideString;
var
  j: integer;
begin
  {$IFDEF FPC}
    {$IFDEF FPC_LITTLE_ENDIAN}
      Result:=AValue;
    {$ELSE}
      Result:=AValue;
      for j := 1 to Length(AValue) do begin
        PWORD(@Result[j])^:=LEToN(PWORD(@Result[j])^);
      end;
    {$ENDIF}
  {$ELSE}
    Result:=AValue;
  {$ENDIF}
end;

{ Converts RGB part of a LongRGB logical structure to its physical representation
  IOW: RGBA (where A is 0 and omitted in the function call) => ABGR
  Needed for conversion of palette colors. }
function LongRGBToExcelPhysical(const RGB: DWord): DWord;
begin
  {$IFDEF FPC}
  {$IFDEF ENDIAN_LITTLE}
  result := RGB shl 8; //tags $00 at end for the A byte
  result := SwapEndian(result); //flip byte order
  {$ELSE}
  //Big endian
  result := RGB; //leave value as is //todo: verify if this turns out ok
  {$ENDIF}
  {$ELSE}
  // messed up result
  {$ENDIF}
end;

{@@
  Parses strings like A5:A10 into an selection interval information
}
function ParseIntervalString(const AStr: string;
  var AFirstCellRow, AFirstCellCol, ACount: Integer;
  var ADirection: TsSelectionDirection): Boolean;
var
  //Cells: TStringList;
  LastCellRow, LastCellCol: Integer;
  p: Integer;
  s1, s2: String;
begin
  Result := True;

  { Simpler:
  use "pos" instead of the TStringList overhead.
  And: the StringList is not free'ed here

  // First get the cells
  Cells := TStringList.Create;
  ExtractStrings([':'],[], PChar(AStr), Cells);

  // Then parse each of them
  Result := ParseCellString(Cells[0], AFirstCellRow, AFirstCellCol);
  if not Result then Exit;
  Result := ParseCellString(Cells[1], LastCellRow, LastCellCol);
  if not Result then Exit;
  }

  // First find the position of the colon and split into parts
  p := pos(':', AStr);
  if p = 0 then exit(false);
  s1 := copy(AStr, 1, p-1);
  s2 := copy(AStr, p+1, Length(AStr));

  // Then parse each of them
  Result := ParseCellString(s1, AFirstCellRow, AFirstCellCol);
  if not Result then Exit;
  Result := ParseCellString(s2, LastCellRow, LastCellCol);
  if not Result then Exit;

  if AFirstCellRow = LastCellRow then
  begin
    ADirection := fpsHorizontalSelection;
    ACount := LastCellCol - AFirstCellCol + 1;
  end
  else if AFirstCellCol = LastCellCol then
  begin
    ADirection := fpsVerticalSelection;
    ACount := LastCellRow - AFirstCellRow + 1;
  end
  else Exit(False);
end;

{@@
  Parses strings like A5:C10 into a range selection information.
  Returns in AFlags also information on relative/absolute cells.
}
function ParseCellRangeString(const AStr: string;
  var AFirstCellRow, AFirstCellCol, ALastCellRow, ALastCellCol: Integer;
  var AFlags: TsRelFlags): Boolean;
var
  p: Integer;
  s: String;
  f: TsRelFlags;
begin
  Result := True;

  // First find the colon
  p := pos(':', AStr);
  if p = 0 then exit(false);

  // Analyze part after the colon
  s := copy(AStr, p+1, Length(AStr));
  Result := ParseCellString(s, ALastCellRow, ALastCellCol, f);
  if not Result then exit;

  // Analyze part before the colon
  s := copy(AStr, 1, p-1);
  Result := ParseCellString(s, AFirstCellRow, AFirstCellCol, AFlags);

  // Add flags of 2nd part
  if rfRelRow in f then Include(AFlags, rfRelRow2);
  if rfRelCol in f then Include(AFlags, rfRelCol2);
end;


{@@
  Parses a cell string, like 'A1' into zero-based column and row numbers
  Note that there can be several letters to address for more than 26 columns.
  'AFlags' indicates relative addresses.

  Example "AMP$200" --> (rel) column 1029 (= 26*26*1 + 26*16 + 26 - 1)
                        (abs) row = 199 (abs)
}
function ParseCellString(const AStr: String; var ACellRow, ACellCol: Integer;
  var AFlags: TsRelFlags): Boolean;

  function Scan(AStartPos: Integer): Boolean;
  const
    LETTERS = ['A'..'Z'];
    DIGITS  = ['0'..'9'];
  var
    i: Integer;
    isAbs: Boolean;
  begin
    Result := false;

    i := AStartPos;
    // Scan letters
    while (i <= Length(AStr)) do begin
      if (UpCase(AStr[i]) in LETTERS) then begin
        ACellCol := ord(UpCase(AStr[i])) - ord('A') + 1 + ACellCol * 26;
        inc(i);
      end
      else
      if (AStr[i] in DIGITS) or (AStr[i] = '$') then
        break
      else begin
        ACellCol := 0;
        exit;      // Only letters or $ allowed
      end;
    end;
    if AStartPos = 1 then Include(AFlags, rfRelCol);

    isAbs := (AStr[i] = '$');
    if isAbs then inc(i);

    // Scan digits
    while (i <= Length(AStr)) do begin
      if (AStr[i] in DIGITS) then begin
        ACellRow := (ord(AStr[i]) - ord('0')) + ACellRow * 10;
        inc(i);
      end
      else begin
        ACellCol := 0;
        ACellRow := 0;
        AFlags := [];
        exit;
      end;
    end;

    dec(ACellCol);
    dec(ACellRow);
    if not isAbs then Include(AFlags, rfRelRow);

    Result := true;
  end;

begin
  ACellCol := 0;
  ACellRow := 0;
  AFlags := [];

  if AStr = '' then
    Exit(false);

  if (AStr[1] = '$') then
    Result := Scan(2)
  else
    Result := Scan(1);
end;

{ for compatibility with old version which does not return flags for relative
  cell addresses }
function ParseCellString(const AStr: string;
  var ACellRow, ACellCol: Integer): Boolean;
var
  flags: TsRelFlags;
begin
  ParseCellString(AStr, ACellRow, ACellCol, flags);
end;

function ParseCellRowString(const AStr: string; var AResult: Integer): Boolean;
begin
  try
    AResult := StrToInt(AStr) - 1;
  except
    Result := False;
  end;
  Result := True;
end;

function ParseCellColString(const AStr: string; var AResult: Integer): Boolean;
const
  INT_NUM_LETTERS = 26;
begin
  Result := False;
  AResult := 0;

  if Length(AStr) = 1 then AResult := Ord(AStr[1]) - Ord('A')
  else if Length(AStr) = 2 then
  begin
    AResult := (Ord(AStr[1]) - Ord('A') + 1) * INT_NUM_LETTERS
     + Ord(AStr[2]) - Ord('A');
  end
  else if Length(AStr) = 3 then
  begin
    AResult := (Ord(AStr[1]) - Ord('A') + 1) * INT_NUM_LETTERS * INT_NUM_LETTERS
     + (Ord(AStr[2]) - Ord('A') + 1) * INT_NUM_LETTERS
     +  Ord(AStr[3]) - Ord('A');
  end
  else Exit(False);

  Result := True;
end;

function Letter(AValue: Integer): char;
begin
  Result := Char(AValue + ord('A'));
end;

{ Calculates an Excel column name ('A', 'B' etc) from the zero-based column index }
function GetColString(AColIndex: Integer): String;
{ Code adapted from: http://stackoverflow.com/questions/12796973/vba-function-to-convert-column-number-to-letter }
var
  n: Integer;
  c: byte;
begin
  Result := '';
  n := AColIndex + 1;
  while (n > 0) do begin
    c := (n - 1) mod 26;
    Result := char(c + ord('A')) + Result;
    n := (n - c) div 26;
  end;
end;

const
  RELCHAR: Array[boolean] of String = ('$', '');

function GetCellString(ARow, ACol: Cardinal; AFlags: TsRelFlags): String;
begin
  Result := Format('%s%s%s%d', [
    RELCHAR[rfRelCol in AFlags], GetColString(ACol),
    RELCHAR[rfRelRow in AFlags], ARow+1
  ]);
end;

function GetCellRangeString(ARow1, ACol1, ARow2, ACol2: Cardinal; AFlags: TsRelFlags): String;
begin
  Result := Format('%s%s%s%d:%s%s%s%d', [
    RELCHAR[rfRelCol in AFlags], GetColString(ACol1),
    RELCHAR[rfRelRow in AFlags], ARow1 + 1,
    RELCHAR[rfRelCol2 in AFlags], GetColString(ACol2),
    RELCHAR[rfRelRow2 in AFlags], ARow2 + 1
  ]);
//  Result := GetCellString(ARow1, ACol1, AFlags) + ':' + GetCellString(ARow2, ACol2, [rfRelRow2, rfRelCol2]);
end;


{ Returns the message text assigned to an error value }
function GetErrorValueStr(AErrorValue: TsErrorValue): String;
begin
  case AErrorValue of
    errOK                   : Result := '';
    errEmptyIntersection    : Result := '#NULL!';
    errDivideByZero         : Result := '#DIV/0!';
    errWrongType            : Result := '#VALUE!';
    errIllegalRef           : Result := '#REF!';
    errWrongName            : Result := '#NAME?';
    errOverflow             : Result := '#NUM!';
    errArgError             : Result := '#N/A';
    // --- no Excel errors --
    errFormulaNotSupported  : Result := '#FORMULA?';
    else                      Result := '#UNKNOWN ERROR';
  end;
end;

{In XML files some chars must be translated}
function UTF8TextToXMLText(AText: ansistring): ansistring;
var
  Idx:Integer;
  WrkStr, AppoSt:ansistring;
begin
  WrkStr:='';

  for Idx:=1 to Length(AText) do
  begin
    case AText[Idx] of
      '&': begin
        AppoSt:=Copy(AText, Idx, 6);

        if (Pos('&amp;',  AppoSt) = 1) or
           (Pos('&lt;',   AppoSt) = 1) or
           (Pos('&gt;',   AppoSt) = 1) or
           (Pos('&quot;', AppoSt) = 1) or
           (Pos('&apos;', AppoSt) = 1) then begin
          //'&' is the first char of a special chat, it must not be converted
          WrkStr:=WrkStr + AText[Idx];
        end else begin
          WrkStr:=WrkStr + '&amp;';
        end;
      end;
      '<': WrkStr:=WrkStr + '&lt;';
      '>': WrkStr:=WrkStr + '&gt;';
      '"': WrkStr:=WrkStr + '&quot;';
      '''':WrkStr:=WrkStr + '&apos;';
    else
      WrkStr:=WrkStr + AText[Idx];
    end;
  end;

  Result:=WrkStr;
end;

{ Returns either AValue1 or AValue2, depending on the condition.
  For reduciton of typing... }
function IfThen(ACondition: Boolean; AValue1, AValue2: TsNumberFormat): TsNumberFormat;
begin
  if ACondition then Result := AValue1 else Result := AValue2;
end;

{ Checks whether the given number format code is for currency or accounting
  i.e. requires currency symbol. }
function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [nfCurrency, nfCurrencyRed, nfAccounting, nfAccountingRed];
end;

{ Checks whether the given number format code is for date/times. }
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [{nfFmtDateTime, }nfShortDateTime, nfShortDate, nfLongDate,
    nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM, nfTimeInterval];
end;

function IsDateTimeFormat(AFormatStr: string): Boolean;
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(nil, AFormatStr);
  try
    Result := parser.IsDateTimeFormat;
  finally
    parser.Free;
  end;
end;

function IsTimeFormat(AFormat: TsNumberFormat): boolean;
begin
  Result := AFormat in [nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM,
    nfTimeInterval];
end;

function IsTimeFormat(AFormatStr: String): Boolean;
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(nil, AFormatStr);
  try
    Result := parser.IsTimeFormat;
  finally
    parser.Free;
  end;
end;

{ Builds a date/time format string from the numberformat code. If the format code
  is nfFmtDateTime the given AFormatString is used. AFormatString can use the
  abbreviations "dm" (for "d/mmm"), "my" (for "mmm/yy"), "ms" (for "mm:ss")
  and "msz" (for "mm:ss.z"). }
function BuildDateTimeFormatString(ANumberFormat: TsNumberFormat;
  const AFormatSettings: TFormatSettings; AFormatString: String = '') : string;
var
  fmt: String;
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
    nfTimeInterval:                               // --> [h]:nn:ss
      if AFormatString = '' then
        Result := '[h]:mm:ss'
      else
        Result := AddIntervalBrackets(AFormatString);
  end;
end;

{ Builds a currency format string. The presentation of negative values (brackets,
  or minus signs) is taken from the provided format settings. The format string
  consists of three sections, separated by semicolons.
  Additional code is inserted for the destination file format:
  - AAccountingStyle = true adds code to align the currency symbols below each
    other.
  - ANegativeValuesRed adds code to the second section of the format code (for
    negative values) to apply a red font color.
  This code has to be removed by StripAccountingSymbols before applying to
  FormatFloat. }
function BuildCurrencyFormatString(ADialect: TsNumFormatDialect;
  ANumberFormat: TsNumberFormat; const AFormatSettings: TFormatSettings;
  ADecimals, APosCurrFormat, ANegCurrFormat: Integer; ACurrencySymbol: String): String;
const
  POS_FMT: array[0..3, boolean] of string = (
    // Parameter 0 is "value", parameter 1 is "currency symbol"
    // AccountingStyle = false --> 1st column, true --> 2nd column
    ('"%1:s"%0:s',  '"%1:s"* %0:s'),      // 0: $1
    ('%0:s"%1:s"',  '%0:s "%1:s"'),       // 1: 1$
    ('"%1:s" %0:s', '"%1:s"* %0:s'),      // 2: $ 1
    ('%0:s "%1:s"', '%0:s "%1:s"')        // 3: 1 $
  );
  NEG_FMT: array[0..15, boolean] of string = (
    ('("%1:s"%0:s)',  '"%1:s"* (%0:s)'),  //  0: ($1)
    ('-"%1:s"%0:s',   '-* "%1:s" %0:s'),  //  1: -$1
    ('"%1:s"-%0:s',   '"%1:s"* -%0:s'),   //  2: $-1
    ('"%1:s"%0:s-',   '"%1:s"%0:s-'),     //  3: $1-
    ('(%0:s"%1:s")',  '(%0:s)%1:s"'),     //  4: (1$)
    ('-%0:s"%1:s"',   '-* %0:s"%1:s"'),   //  5: -1$
    ('%0:s-"%1:s"',   '%0:s-"%1:s"'),     //  6: 1-$
    ('%0:s"%1:s"-',   '%0:s-"%1:s"'),     //  7: 1$-
    ('-%0:s "%1:s"',  '-* %0:s"%1:s"'),   //  8: -1 $
    ('-"%1:s" %0:s',  '-* "%1:s" %0:s'), //  9: -$ 1
    ('%0:s "%1:s"-',  '%0:s- "%1:s"'),    // 10: 1 $-
    ('"%1:s" %0:s-',  '"%1:s"* %0:s-'),   // 11: $ 1-
    ('"%1:s" -%0:s',  '"%1:s"* -%0:s'),   // 12: $ -1
    ('%0:s- "%1:s"',  '%0:s- "%1:s"'),    // 13: 1- $
    ('("%1:s" %0:s)', '"%1:s"* (%0:s)'),  // 14: ($ 1)
    ('(%0:s "%1:s")', '(%0:s "%1:s")')    // 15: (1 $)
  );
var
  decs: String;
  pcf, ncf: Byte;
  p, n: String;
  accStyle: Boolean;
  negRed: Boolean;
begin
  pcf := IfThen(APosCurrFormat < 0, AFormatSettings.CurrencyFormat, APosCurrFormat);
  ncf := IfThen(ANegCurrFormat < 0, AFormatSettings.NegCurrFormat, ANegCurrFormat);
  if ADecimals < 0 then
    ADecimals := AFormatSettings.CurrencyDecimals;
  if ACurrencySymbol = '?' then
    ACurrencySymbol := AnsiToUTF8(AFormatSettings.CurrencyString);
  decs := DupeString('0', ADecimals);
  if ADecimals > 0 then decs := '.' + decs;

  accStyle := ANumberFormat in [nfAccounting, nfAccountingRed];
  negRed := ANumberFormat in [nfCurrencyRed, nfAccountingRed];

  p := POS_FMT[pcf, accStyle];
  n := NEG_FMT[ncf, accStyle];
  // add extra space for the sign of the number for perfect alignment in Excel
  if accStyle then
    case ncf of
      0, 14: p := p + '_)';
      3, 11: p := p + '_-';
      4, 15: p := '_(' + p;
      5, 8 : p := '_-' + p;
    end;

  if ACurrencySymbol <> '' then begin
    Result := Format(p, ['#,##0' + decs, ACurrencySymbol]) + ';'
            + IfThen(negRed and (ADialect = nfdExcel), '[red]', '')
            + Format(n, ['#,##0' + decs, ACurrencySymbol]) + ';'
            + Format(p, [IfThen(accStyle, '-', '0'+decs), ACurrencySymbol]);
  end
  else begin
    Result := '#,##0' + decs;
    if negRed and (ADialect = nfdExcel) then
      Result := Result +';[red]'
    else
      Result := Result +';';
    case ncf of
      0, 14, 15           : Result := Result + '(#,##0' + decs + ')';
      1, 2, 5, 6, 8, 9, 12: Result := Result + '-#,##0' + decs;
      else                  Result := Result + '#,##0' + decs + '-';
    end;
    Result := Result + ';' + IfThen(accStyle, '-', '0'+decs);
  end;
end;

{ Builds a number format string from the number format code, the count of
  decimals, and the currencysymbol (if not empty). }
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
    nfSci:
      Result := '##0' + decs + 'E+0';
    nfPercentage:
      Result := '0' + decs + '%';
    nfCurrency, nfCurrencyRed, nfAccounting, nfAccountingRed:
      raise Exception.Create('BuildNumberFormatString: Use BuildCurrencyFormatString '+
        'to create a format string for currency values.');
    nfShortDateTime, nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfTimeInterval:
      raise Exception.Create('BuildNumberFormatString: Use BuildDateTimeFormatSstring '+
        'to create a format string for date/time values.');
  end;
end;

function AddAMPM(const ATimeFormatString: String;
  const AFormatSettings: TFormatSettings): String;
var
  am, pm: String;
begin
  am := IfThen(AFormatSettings.TimeAMString <> '', AFormatSettings.TimeAMString, 'AM');
  pm := IfThen(AFormatSettings.TimePMString <> '', AFormatSettings.TimePMString, 'PM');
  Result := Format('%s %s/%s', [StripAMPM(ATimeFormatString), am, pm]);
end;

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

{ The given format string is assumed to be for time intervals, i.e. its first
  time symbol must be enclosed by square brackets. Checks if this is true, and
  adds the brackes if not. }
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

{ Creates a long date format string out of a short one. Retains the order of
  year-month-day and the separators, but uses 4 digits for year and 3 digits of m }
function MakeLongDateFormat(AShortDateFormat: String): String;
var
  i: Integer;
begin
  Result := '';
  i := 1;
  while i < Length(AShortDateFormat) do begin
    case AShortDateFormat[i] of
      'y', 'Y':
        begin
          Result := Result + DupeString(AShortDateFormat[i], 4);
          while (i < Length(AShortDateFormat)) and (AShortDateFormat[i] in ['y','Y']) do
            inc(i);
        end;
      'm', 'M':
        begin
          result := Result + DupeString(AShortDateFormat[i], 3);
          while (i < Length(AShortDateFormat)) and (AShortDateFormat[i] in ['m','M']) do
            inc(i);
        end;
      else
        Result := Result + AShortDateFormat[i];
        inc(i);
    end;
  end;
end;

{ Modifies the short date format such that it has a two-digit year and a two-digit
  month. Retains the order of year-month-day and the separators. }
function MakeShortDateFormat(AShortDateFormat: String): String;
var
  i: Integer;
begin
  Result := '';
  i := 1;
  while i < Length(AShortDateFormat) do begin
    case AShortDateFormat[i] of
      'y', 'Y':
        begin
          Result := Result + DupeString(AShortDateFormat[i], 2);
          while (i < Length(AShortDateFormat)) and (AShortDateFormat[i] in ['y','Y']) do
            inc(i);
        end;
      'm', 'M':
        begin
          result := Result + DupeString(AShortDateFormat[i], 2);
          while (i < Length(AShortDateFormat)) and (AShortDateFormat[i] in ['m','M']) do
            inc(i);
        end;
      else
        Result := Result + AShortDateFormat[i];
        inc(i);
    end;
  end;
end;

{ Creates the formatstrings for the date/time codes "dm", "my", "ms" and "msz"
  out of the formatsettings. }
function SpecialDateTimeFormat(ACode: String;
  const AFormatSettings: TFormatSettings; ForWriting: Boolean): String;
var
  pd, pm, py: Integer;
  sdf: String;
  MonthChar, MinuteChar, MillisecChar: Char;
begin
  if ForWriting then begin
    MonthChar := 'M'; MinuteChar := 'm'; MillisecChar := '0';
  end else begin
    MonthChar := 'm'; MinuteChar := 'n'; MillisecChar := 'z';
  end;
  ACode := lowercase(ACode);
  sdf := lowercase(AFormatSettings.ShortDateFormat);
  pd := pos('d', sdf);
  pm := pos('m', sdf);
  py := pos('y', sdf);
  if ACode = 'dm' then begin
    Result := DupeString(MonthChar, 3);
    Result := IfThen(pd < py, 'd/'+Result, Result+'/d');            // d/mmm
  end else
  if ACode = 'my' then begin
    Result := DupeString(MonthChar, 3);
    Result := IfThen(pm < py, Result+'/yy', 'yy/'+Result);          // mmm/yy
  end else
  if ACode = 'ms' then begin
    Result := DupeString(MinuteChar, 2) + ':ss';                    // mm:ss
  end
  else if ACode = 'msz' then
    Result := DupeString(MinuteChar, 2) + ':ss.' + MillisecChar     // mm:ss.z
  else
    Result := ACode;
end;

{ Splits the sections +1 (positive) or -1 (negative values) or 0 (zero values)
  of the accounting format string at the position of the '*' into a left
  and right part and returns 1 if the format string is in the left, and 2 if
  it is in the right part. Additionally removes Excel format codes '_' }
function SplitAccountingFormatString(const AFormatString: String; ASection: ShortInt;
  var ALeft, ARight: String): Byte;
var
  P: PChar;
  PStart, PEnd: PChar;
  token: Char;
  done: Boolean;
  i: Integer;
begin
  Result := 0;
  PStart := PChar(@AFormatString[1]);
  PEnd := PStart + Length(AFormatString);
  P := PStart;

  done := false;
  case ASection of
    -1 : while (P < PEnd) and not done do begin
           token := P^;
           if token = ';' then done := true;
           inc(P);
         end;
     0 : for i := 1 to 2 do begin
           done := false;
           while (P < PEnd) and not done do begin
             token := P^;
             if token = ';' then done := true;
             inc(P);
           end;
         end;
    +1: ;
  end;

  ALeft := '';
  done := false;

  while (P < PEnd) and not done do begin
    token := P^;
    case token of
      '_': inc(P);
      ';': done := true;
      '"': ;
      '*': begin
             inc(P);
             done := true;
           end;
      '0',
      '#': begin
             ALeft := ALeft + token;
             Result := 1;
           end;
      else ALeft := ALeft + token;
    end;
    inc(P);
  end;

  ARight := '';
  done := false;
  while (P < PEnd) and not done do begin
    token := P^;
    case token of
      '_': inc(P);
      ';': done := true;
      '"': ;
      '0',
      '#': begin
             ARight := ARight + token;
             Result := 2;
           end;
      else ARight := ARight + token;
    end;
    inc(P);
  end;
end;

procedure SplitFormatString(const AFormatString: String; out APositivePart,
  ANegativePart, AZeroPart: String);
var
  P, PStart, PEnd: PChar;
  token: Char;
  where: Byte;  // 0 = positive part, 1 = negative part, 2 = zero part
begin
  APositivePart := '';
  ANegativePart := '';
  AZeroPart := '';
  if AFormatString = '' then
    exit;

  PStart := PChar(@AFormatString[1]);
  PEnd := PStart + Length(AFormatString);
  P := PStart;
  where := 0;
  while P < PEnd do begin
    token := P^;
    case token of
      '"': begin   // Skip quoted strings
             inc(P);
             token := P^;
             while (P < PEnd) and (token <> '"') do begin
               inc(P);
               token := P^;
             end;
           end;
      ';': begin  // Separator between parts
             inc(where);
             if where = 3 then
               exit;
           end
      else case where of
             0: APositivePart := APositivePart + token;
             1: ANegativePart := ANegativePart + token;
             2: AZeroPart := AZeroPart + token;
           end;
    end;
    inc(P);
  end;
end;

{ Formats the number AValue in "scientific" format with the given number of
  decimals. "Scientific" is the same as "exponential", but with exponents rounded
  to multiples of 3 (like for "kilo" - "Mega" - "Giga" etc.). }
function SciFloat(AValue: Double; ADecimals: Byte;
  AFormatSettings: TFormatSettings): String;
var
  m: Double;
  ex: Integer;
begin
  if AValue = 0 then
    Result := Format('%0.*fE+0', [ADecimals, 0.0], AFormatSettings)
    // Excel shows "000.0E+0", but I think the "0.0E+0" shown here is better.
  else begin
    ex := floor(log10(abs(AValue)));  // exponent
    // round exponent to multiples of 3
    ex := (ex div 3) * 3;
    if ex < 0 then dec(ex, 3);
    m := AValue * Power(10, -ex);     // mantisse
    Result := Format('%.*fE+%d', [ADecimals, m, ex], AFormatSettings);
  end;
end;

function SciFloat(AValue: Double; ADecimals: Byte): String;
begin
  Result := SciFloat(AValue, ADecimals, DefaultFormatSettings);
end;

{ Creates a "time interval" format string having the first code identifier
  in square brackets. }
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

{ Excel's unit of row heights is "twips", i.e. 1/20 point.
  Converts Twips to points. }
function TwipsToPts(AValue: Integer): Single;
begin
  Result := AValue / 20;
end;

{ Converts points to twips (1 twip = 1/20 point) }
function PtsToTwips(AValue: Single): Integer;
begin
  Result := round(AValue * 20);
end;

{ Converts centimeters to points (72 pts = 1 inch) }
function cmToPts(AValue: Double): Double;
begin
  Result := AValue * 72 / 2.54;
end;

{ Converts points to centimeters }
function PtsToCm(AValue: Double): Double;
begin
  Result := AValue / 72 * 2.54;
end;

{ Converts inches to points (72 pts = 1 inch) }
function InToPts(AValue: Double): Double;
begin
  Result := AValue * 72;
end;

{ Converts millimeters to points (72 pts = 1 inch) }
function mmToPts(AValue: Double): Double;
begin
  Result := AValue * 72 / 25.4;
end;

{ Converts points to millimeters }
function PtsToMM(AValue: Double): Double;
begin
  Result := AValue / 72 * 25.4;
end;

{ Converts pixels to points. }
function pxToPts(AValue, AScreenPixelsPerInch: Integer): Double;
begin
  Result := (AValue / AScreenPixelsPerInch) * 72;
end;

{ Converts points to pixels }
function PtsToPx(AValue: Double; AScreenPixelsPerInch: Integer): Integer;
begin
  Result := Round(AValue / 72 * AScreenPixelsPerInch);
end;

{ converts a HTML length string to points. The units are assumed to be the last
  two digits of the string }
function HTMLLengthStrToPts(AValue: String): Double;
var
  units: String;
  x: Double;
  res: Word;
begin
  if (Length(AValue) > 1) and (AValue[Length(AValue)] in ['a'..'z', 'A'..'Z']) then begin
    units := lowercase(Copy(AValue, Length(AValue)-1, 2));
    val(copy(AValue, 1, Length(AValue)-2), x, res);
    // No hasseling with the decimal point...
  end else begin
    units := '';
    val(AValue, x, res);
  end;
  if res <> 0 then
    raise Exception.CreateFmt('No valid number or units (%s)', [AValue]);

  if (units = 'pt') or (units = '') then
    Result := x
  else
  if units = 'in' then
    Result := InToPts(x)
  else if units = 'cm' then
    Result := cmToPts(x)
  else if units = 'mm' then
    Result := mmToPts(x)
  else if units = 'px' then
    Result := pxToPts(Round(x), ScreenPixelsPerInch)
  else
    raise Exception.Create('Unknown length units');
end;

{ converts a HTML color string to a TsColorValue. For ods }
function HTMLColorStrToColor(AValue: String): TsColorValue;
begin
  if AValue = '' then
    Result := scNotDefined
  else
  if AValue[1] = '#' then begin
    AValue[1] := '$';
    Result := LongRGBToExcelPhysical(StrToInt(AValue));
  end else begin
    AValue := lowercase(AValue);
    if AValue = 'red' then
      Result := $0000FF
    else if AValue = 'cyan' then
      Result := $FFFF00
    else if AValue = 'blue' then
      Result := $FF0000
    else if AValue = 'purple' then
      Result := $800080
    else if AValue = 'yellow' then
      Result := $00FFFF
    else if AValue = 'lime' then
      Result := $00FF00
    else if AValue = 'white' then
      Result := $FFFFFF
    else if AValue = 'black' then
      Result := $000000
    else if (AValue = 'gray') or (AValue = 'grey') then
      Result := $808080
    else if AValue = 'silver' then
      Result := $C0C0C0
    else if AValue = 'maroon' then
      Result := $000080
    else if AValue = 'green' then
      Result := $008000
    else if AValue = 'olive' then
      Result := $008080;
  end;
end;

{ converts an rgb color value to a string as used in HTML code (for ods) }
function ColorToHTMLColorStr(AValue: TsColorValue): String;
type
  TRGB = record r,g,b,a: Byte end;
var
  rgb: TRGB;
begin
  rgb := TRGB(AValue);
  Result := Format('#%.2x%.2x%.2x', [rgb.r, rgb.g, rgb.b]);
end;



{******************************************************************************}
{******************************************************************************}
{                   Patch for SysUtils.FormatDateTime                          }
{  Remove when the feature of square brackets in time format masks is in rtl   }
{******************************************************************************}
{******************************************************************************}

// Copied from "fpc/rtl/objpas/sysutils/datei.inc"

procedure DateTimeToString(out Result: string; const FormatStr: string; const DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []);
var
  ResultLen: integer;
  ResultBuffer: array[0..255] of char;
  ResultCurrent: pchar;
{$IFDEF MSWindows}
  isEnable_E_Format : Boolean;
  isEnable_G_Format : Boolean;
  eastasiainited : boolean;
{$ENDIF MSWindows}
             (*    ---- not needed here ---
{$IFDEF MSWindows}
  procedure InitEastAsia;
  var     ALCID : LCID;
         PriLangID , SubLangID : Word;

  begin
    ALCID := GetThreadLocale;
    PriLangID := ALCID and $3FF;
    if (PriLangID>0) then
       SubLangID := (ALCID and $FFFF) shr 10
      else
        begin
          PriLangID := SysLocale.PriLangID;
          SubLangID := SysLocale.SubLangID;
        end;
    isEnable_E_Format := (PriLangID = LANG_JAPANESE)
                  or
                  (PriLangID = LANG_KOREAN)
                  or
                  ((PriLangID = LANG_CHINESE)
                   and
                   (SubLangID = SUBLANG_CHINESE_TRADITIONAL)
                  );
    isEnable_G_Format := (PriLangID = LANG_JAPANESE)
                  or
                  ((PriLangID = LANG_CHINESE)
                   and
                   (SubLangID = SUBLANG_CHINESE_TRADITIONAL)
                  );
    eastasiainited :=true;
  end;
{$ENDIF MSWindows}
                     *)
  procedure StoreStr(Str: PChar; Len: Integer);
  begin
    if ResultLen + Len < SizeOf(ResultBuffer) then
    begin
      StrMove(ResultCurrent, Str, Len);
      ResultCurrent := ResultCurrent + Len;
      ResultLen := ResultLen + Len;
    end;
  end;

  procedure StoreString(const Str: string);
  var Len: integer;
  begin
   Len := Length(Str);
   if ResultLen + Len < SizeOf(ResultBuffer) then
     begin
       StrMove(ResultCurrent, pchar(Str), Len);
       ResultCurrent := ResultCurrent + Len;
       ResultLen := ResultLen + Len;
     end;
  end;

  procedure StoreInt(Value, Digits: Integer);
  var
    S: string[16];
    Len: integer;
  begin
    System.Str(Value:Digits, S);
    for Len := 1 to Length(S) do
    begin
      if S[Len] = ' ' then
        S[Len] := '0'
      else
        Break;
    end;
    StoreStr(pchar(@S[1]), Length(S));
  end ;

var
  Year, Month, Day, DayOfWeek, Hour, Minute, Second, MilliSecond: word;
  DT : TDateTime;

  procedure StoreFormat(const FormatStr: string; Nesting: Integer; TimeFlag: Boolean);
  var
    Token, lastformattoken, prevlasttoken: char;
    FormatCurrent: pchar;
    FormatEnd: pchar;
    Count: integer;
    Clock12: boolean;
    P: pchar;
    tmp: integer;
    isInterval: Boolean;

  begin
    if Nesting > 1 then  // 0 is original string, 1 is included FormatString
      Exit;

    FormatCurrent := PChar(FormatStr);
    FormatEnd := FormatCurrent + Length(FormatStr);
    Clock12 := false;
    isInterval := false;
    P := FormatCurrent;
    // look for unquoted 12-hour clock token
    while P < FormatEnd do
    begin
      Token := P^;
      case Token of
        '''', '"':
        begin
          Inc(P);
          while (P < FormatEnd) and (P^ <> Token) do
            Inc(P);
        end;
        'A', 'a':
        begin
          if (StrLIComp(P, 'A/P', 3) = 0) or
             (StrLIComp(P, 'AMPM', 4) = 0) or
             (StrLIComp(P, 'AM/PM', 5) = 0) then
          begin
            Clock12 := true;
            break;
          end;
        end;
      end;  // case
      Inc(P);
    end ;
    token := #255;
    lastformattoken := ' ';
    prevlasttoken := 'H';
    while FormatCurrent < FormatEnd do
    begin
      Token := UpCase(FormatCurrent^);
      Count := 1;
      P := FormatCurrent + 1;
      case Token of
        '''', '"':
        begin
          while (P < FormatEnd) and (p^ <> Token) do
            Inc(P);
          Inc(P);
          Count := P - FormatCurrent;
          StoreStr(FormatCurrent + 1, Count - 2);
        end ;
        'A':
        begin
          if StrLIComp(FormatCurrent, 'AMPM', 4) = 0 then
          begin
            Count := 4;
            if Hour < 12 then
              StoreString(FormatSettings.TimeAMString)
            else
              StoreString(FormatSettings.TimePMString);
          end
          else if StrLIComp(FormatCurrent, 'AM/PM', 5) = 0 then
          begin
            Count := 5;
            if Hour < 12 then StoreStr(FormatCurrent, 2)
                         else StoreStr(FormatCurrent+3, 2);
          end
          else if StrLIComp(FormatCurrent, 'A/P', 3) = 0 then
          begin
            Count := 3;
            if Hour < 12 then StoreStr(FormatCurrent, 1)
                         else StoreStr(FormatCurrent+2, 1);
          end
          else
            raise EConvertError.Create('Illegal character in format string');
        end ;
        '/': StoreStr(@FormatSettings.DateSeparator, 1);
        ':': StoreStr(@FormatSettings.TimeSeparator, 1);
	'[': if (fdoInterval in Options) then isInterval := true else StoreStr(FormatCurrent, 1);
	']': if (fdoInterval in Options) then isInterval := false else StoreStr(FormatCurrent, 1);
        ' ', 'C', 'D', 'H', 'M', 'N', 'S', 'T', 'Y', 'Z', 'F' :
        begin
          while (P < FormatEnd) and (UpCase(P^) = Token) do
            Inc(P);
          Count := P - FormatCurrent;
          case Token of
            ' ': StoreStr(FormatCurrent, Count);
            'Y': begin
              if Count > 2 then
                StoreInt(Year, 4)
              else
                StoreInt(Year mod 100, 2);
            end;
            'M': begin
	      if isInterval and ((prevlasttoken = 'H') or TimeFlag) then
	        StoreInt(Minute + (Hour + trunc(abs(DateTime))*24)*60, 0)
	      else
              if (lastformattoken = 'H') or TimeFlag then
              begin
                if Count = 1 then
                  StoreInt(Minute, 0)
                else
                  StoreInt(Minute, 2);
              end
              else
              begin
                case Count of
                  1: StoreInt(Month, 0);
                  2: StoreInt(Month, 2);
                  3: StoreString(FormatSettings.ShortMonthNames[Month]);
                else
                  StoreString(FormatSettings.LongMonthNames[Month]);
                end;
              end;
            end;
            'D': begin
              case Count of
                1: StoreInt(Day, 0);
                2: StoreInt(Day, 2);
                3: StoreString(FormatSettings.ShortDayNames[DayOfWeek]);
                4: StoreString(FormatSettings.LongDayNames[DayOfWeek]);
                5: StoreFormat(FormatSettings.ShortDateFormat, Nesting+1, False);
              else
                StoreFormat(FormatSettings.LongDateFormat, Nesting+1, False);
              end ;
            end ;
            'H':
	      if isInterval then
	        StoreInt(Hour + trunc(abs(DateTime))*24, 0)
	      else
	      if Clock12 then
              begin
                tmp := hour mod 12;
                if tmp=0 then tmp:=12;
                if Count = 1 then
                  StoreInt(tmp, 0)
                else
                  StoreInt(tmp, 2);
              end
              else begin
                if Count = 1 then
		  StoreInt(Hour, 0)
                else
                  StoreInt(Hour, 2);
              end;
            'N': if isInterval then
	           StoreInt(Minute + (Hour + trunc(abs(DateTime))*24)*60, 0)
		 else
		 if Count = 1 then
                   StoreInt(Minute, 0)
                 else
                   StoreInt(Minute, 2);
            'S': if isInterval then
	           StoreInt(Second + (Minute + (Hour + trunc(abs(DateTime))*24)*60)*60, 0)
	         else
	         if Count = 1 then
                   StoreInt(Second, 0)
                 else
                   StoreInt(Second, 2);
            'Z': if Count = 1 then
                   StoreInt(MilliSecond, 0)
                 else
		   StoreInt(MilliSecond, 3);
            'T': if Count = 1 then
		   StoreFormat(FormatSettings.ShortTimeFormat, Nesting+1, True)
                 else
	           StoreFormat(FormatSettings.LongTimeFormat, Nesting+1, True);
            'C': begin
                   StoreFormat(FormatSettings.ShortDateFormat, Nesting+1, False);
                   if (Hour<>0) or (Minute<>0) or (Second<>0) then
                     begin
                      StoreString(' ');
                      StoreFormat(FormatSettings.LongTimeFormat, Nesting+1, True);
                     end;
                 end;
            'F': begin
                   StoreFormat(FormatSettings.ShortDateFormat, Nesting+1, False);
                   StoreString(' ');
                   StoreFormat(FormatSettings.LongTimeFormat, Nesting+1, True);
                 end;
            (* ------------ not needed here...
{$IFDEF MSWindows}
            'E':
               begin
                 if not Eastasiainited then InitEastAsia;
                 if Not(isEnable_E_Format) then StoreStr(@FormatCurrent^, 1)
                  else
                   begin
                     while (P < FormatEnd) and (UpCase(P^) = Token) do
                     P := P + 1;
                     Count := P - FormatCurrent;
                     StoreString(ConvertEraYearString(Count,Year,Month,Day));
                   end;
		 prevlasttoken := lastformattoken;
                 lastformattoken:=token;
               end;
             'G':
               begin
                 if not Eastasiainited then InitEastAsia;
                 if Not(isEnable_G_Format) then StoreStr(@FormatCurrent^, 1)
                  else
                   begin
                     while (P < FormatEnd) and (UpCase(P^) = Token) do
                     P := P + 1;
                     Count := P - FormatCurrent;
                     StoreString(ConvertEraString(Count,Year,Month,Day));
                   end;
		 prevlasttoken := lastformattoken;
                 lastformattoken:=token;
               end;
{$ENDIF MSWindows}
*)
          end;
	  prevlasttoken := lastformattoken;
          lastformattoken := token;
        end;
        else
          StoreStr(@Token, 1);
      end ;
      Inc(FormatCurrent, Count);
    end;
  end;

begin
{$ifdef MSWindows}
  eastasiainited:=false;
{$endif MSWindows}
  DecodeDateFully(DateTime, Year, Month, Day, DayOfWeek);
  DecodeTime(DateTime, Hour, Minute, Second, MilliSecond);
  ResultLen := 0;
  ResultCurrent := @ResultBuffer[0];
  if FormatStr <> '' then
    StoreFormat(FormatStr, 0, False)
  else
    StoreFormat('C', 0, False);
  ResultBuffer[ResultLen] := #0;
  result := StrPas(@ResultBuffer[0]);
end ;

procedure DateTimeToString(out Result: string; const FormatStr: string;
  const DateTime: TDateTime; Options : TFormatDateTimeOptions = []);
begin
  DateTimeToString(Result, FormatStr, DateTime, DefaultFormatSettings, Options);
end;

function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  Options : TFormatDateTimeOptions = []): string;
begin
  DateTimeToString(Result, FormatStr, DateTime, DefaultFormatSettings,Options);
end;

function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []): string;
begin
  DateTimeToString(Result, FormatStr, DateTime, FormatSettings,Options);
end;

end.

