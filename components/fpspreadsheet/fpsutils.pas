{
  Utility functions and constants from FPSpreadsheet
}

// to do: Remove the patched FormatDateTime when the feature of square brackets
//        in time format codes is in the rtl

unit fpsutils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, StrUtils;

// Exported types
type
  TsSelectionDirection = (fpsVerticalSelection, fpsHorizontalSelection);

  TsRelFlag = (rfRelRow, rfRelCol, rfRelRow2, rfRelCol2);
  TsRelFlags = set of TsRelFlag;

const
  // Date formatting string for unambiguous date/time display as strings
  // Can be used for text output when date/time cell support is not available
  ISO8601Format='yyyymmdd"T"hhmmss';
  // Extended ISO 8601 date/time format, used in e.g. ODF/opendocument
  ISO8601FormatExtended='yyyy"-"mm"-"dd"T"hh":"mm":"ss';

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

function UTF8TextToXMLText(AText: ansistring): ansistring;

function TwipsToMillimeters(AValue: Integer): Single;
function MillimetersToTwips(AValue: Single): Integer;

function IsExpNumberFormat(s: String; out Decimals: Word; out IsSci: Boolean): Boolean;
function IsFixedNumberFormat(s: String; out Decimals: Word): Boolean;
function IsPercentNumberFormat(s: String; out Decimals: Word): Boolean;
function IsThousandSepNumberFormat(s: String; out Decimals: Word): Boolean;
function IsDateFormat(s: String; out IsLong: Boolean): Boolean;
function IsTimeFormat(s: String; out isLong, isAMPM, isInterval: Boolean;
  out SecDecimals: Word): Boolean;

function SciFloat(AValue: Double; ADecimals: Word): String;
//function TimeIntervalToString(AValue: TDateTime; AFormatStr: String): String;
procedure MakeTimeIntervalMask(Src: String; var Dest: String);

function FormatDateTime(const FormatStr: string; DateTime: TDateTime): string;

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
  Return also information on relative/absolute cells.
}
function ParseCellRangeString(const AStr: string;
  var AFirstCellRow, AFirstCellCol, ALastCellRow, ALastCellCol: Integer;
  var AFlags: TsRelFlags): Boolean;
var
  p: Integer;
  s: String;
begin
  Result := True;

  // First find the colon
  p := pos(':', AStr);
  if p = 0 then exit(false);

  // Analyze part after the colon
  s := copy(AStr, p+1, Length(AStr));
  Result := ParseCellString(s, ALastCellRow, ALastCellCol, AFlags);
  if not Result then exit;
  if (rfRelRow in AFlags) then begin
    Include(AFlags, rfRelRow2);
    Exclude(AFlags, rfRelRow);
  end;
  if (rfRelCol in AFlags) then begin
    Include(AFlags, rfRelCol2);
    Exclude(AFlags, rfRelCol);
  end;

  // Analyze part before the colon
  s := copy(AStr, 1, p-1);
  Result := ParseCellString(s, AFirstCellRow, AFirstCellCol, AFlags);
end;


{@@
  Parses a cell string, like 'A1' into zero-based column and row numbers

  The parser is a simple state machine, with the following states:

  0 - Reading Column part 1 (necesserely needs a letter)
  1 - Reading Column part 2, but could be the first number as well
  2 - Reading Row

  'AFlags' indicates relative addresses.
}
function ParseCellString(const AStr: string; var ACellRow, ACellCol: Integer;
  var AFlags: TsRelFlags): Boolean;
var
  i: Integer;
  state: Integer;
  Col, Row: string;
  lChar: Char;
  isAbs: Boolean;
const
  cLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
   'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'W', 'X', 'Y', 'Z'];
  cDigits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
begin
  // Starting state
  Result := True;
  state := 0;
  Col := '';
  Row := '';
  AFlags := [rfRelCol, rfRelRow];
  isAbs := false;

  // Separates the string into a row and a col
  for i := 1 to Length(AStr) do
  begin
    lChar := AStr[i];

    if lChar = '$' then begin
      if isAbs then
        exit(false);
      isAbs := true;
      continue;
    end;

    case state of

    0:
    begin
      if lChar in cLetters then
      begin
        Col := lChar;
        if isAbs then
          Exclude(AFlags, rfRelCol);
        isAbs := false;
        state := 1;
      end
      else Exit(False);
    end;

    1:
    begin
      if lChar in cLetters then
        Col := Col + lChar
      else if lChar in cDigits then
      begin
        Row := lChar;
        if isAbs then
          Exclude(AFlags, rfRelRow);
        isAbs := false;
        state := 2;
      end
      else Exit(False);
    end;

    2:
    begin
      if lChar in cDigits then Row := Row + lChar
      else Exit(False);
    end;

    end;
  end;

  // Now parses each separetely
  ParseCellRowString(Row, ACellRow);
  ParseCellColString(Col, ACellCol);
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

function GetColString(AColIndex: Integer): String;
begin
  if AColIndex < 26 then
    Result := Letter(AColIndex)
  else
  if AColIndex < 26*26 then
    Result := Letter(AColIndex div 26) + Letter(AColIndex mod 26)
  else
  if AColIndex < 26*26*26 then
    Result := Letter(AColIndex div (26*26)) + Letter((AColIndex mod (26*26)) div 26)
      + Letter(AColIndex mod (26*26*26))
  else
    Result := 'too big';
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

{ Excel's unit of row heights is "twips", i.e. 1/20 point. 72 pts = 1 inch = 25.4 mm
  The procedure TwipsToMillimeters performs the conversion to millimeters. }
function TwipsToMillimeters(AValue: Integer): Single;
begin
  Result := 25.4 * AValue / (20 * 72);
end;

{ Converts Millimeters to Twips, i.e. 1/20 pt }
function MillimetersToTwips(AValue: Single): Integer;
begin
  Result := Round((AValue * 20 * 72) / 25.4);
end;


{ Format checking procedures }

{ This simple parsing procedure of the Excel format string checks for a fixed
  float format s, i.e. s can be '0', '0.00', '000', '0,000', and returns the
  number of decimals, i.e. number of zeros behind the decimal point }
function IsFixedNumberFormat(s: String; out Decimals: Word): Boolean;
var
  i: Integer;
  p: Integer;
  decs: String;
begin
  Decimals := 0;

  // Excel time formats with milliseconds ("mm:ss.000") can be incorrectly
  // detected as fixed number formats. Check this case at first.
  if pos('s.0', s) > 0 then begin
    Result := false;
    exit;
  end;

  // Check if s is a valid format mask.
  try
    FormatFloat(s, 1.0);
  except
    on EConvertError do begin
      Result := false;
      exit;
    end;
  end;

  // If it is count the zeros - each one is a decimal.
  if s = '0' then
    Result := true
  else begin
    p := pos('.', s);  // position of decimal point;
    if p = 0 then begin
      Result := false;
    end else begin
      Result := true;
      for i:= p+1 to Length(s) do
        if s[i] = '0' then begin
          inc(Decimals)
        end
        else
          exit;     // ignore characters after the last 0
    end;
  end;
end;

{ This function checks whether the format string corresponds to a thousand
  separator format like "#,##0.000' and returns the number of fixed decimals
  (i.e. zeros after the decimal point) }
function IsThousandSepNumberFormat(s: String; out Decimals: Word): Boolean;
var
  i, p: Integer;
begin
  Decimals := 0;

  // Check if s is a valid format string
  try
    FormatFloat(s, 1.0);
  except
    on EConvertError do begin
      Result := false;
      exit;
    end;
  end;

  // If it is look for the thousand separator. If found count decimals.
  Result := (Pos(',', s) > 0);
  if Result then begin
    p := pos('.', s);
    if p > 0 then
      for i := p+1 to Length(s) do
        if s[i] = '0' then
          inc(Decimals)
        else
          exit;  // ignore format characters after the last 0
  end;
end;

{ This function checks whether the format string corresponds to percent
  formatting and determines the number of decimals }
function IsPercentNumberFormat(s: String; out Decimals: Word): Boolean;
var
  i, p: Integer;
begin
  Decimals := 0;
  // The signature of the percent format is a percent sign at the end of the
  // format string.
  Result := (s <> '') and (s[Length(s)] = '%');
  if Result then begin
    // Check for a valid format string
    try
      FormatDateTime(s, 1.0);
    except
      on EConvertError do begin
        Result := false;
        exit;
      end;
    end;
    // Count decimals
    p := pos('.', s);
    if p > 0 then
      for i := p+1 to Length(s)-1 do
        if s[i] = '0' then
          inc(Decimals)
        else
          exit;  // ignore characters after last 0
  end;
end;

{ This function checks whether the format string corresponds to exponential
  formatting and determines the number of decimals. If it contains a # character
  the function assumes a "scientific" format rounding the exponent to multiples
  of 2. }
function IsExpNumberFormat(s: String; out Decimals: Word;
  out IsSci: Boolean): Boolean;
var
  i, pdp, pe, ph: Integer;
begin
  Result := false;
  Decimals := 0;
  IsSci := false;

  if SameText(s, 'General') then
    exit;

  // Check for a valid format string
  try
    FormatDateTime(s, 1.0);
  except
    on EConvertError do begin
      exit;
    end;
  end;

  pe := pos('e', lowercase(s));
  result := pe > 0;
  if Result then begin
    // The next character must be a "+", "-", or "0"
    if (pe = Length(s)) or not (s[pe+1] in ['+', '-', '0']) then begin
      Result := false;
      exit;
    end;
    // Count decimals
    pdp := pos('.', s);
    if (pdp > 0) then begin
      if pdp < pe then
        for i:=pdp+1 to pe-1 do
          if s[i] = '0' then
            inc(Decimals)
          else
            break;   // ignore characters after last 0
    end;
    // Look for hash signs # as indicator of the "scientific" format
    ph := pos('#', s);
    if ph > 0 then IsSci := true;
  end;
end;

{ IsDateFormat checks if the format string s corresponds to a date format }
function IsDateFormat(s: String; out IsLong: Boolean): Boolean;
begin
  // Day, month, year are separated by a slash
  Result := (pos('/', s) > 0);
  if Result then
    // Check validity of format string
    try
      FormatDateTime(s, now);
      s := Lowercase(s);
      isLong := (pos('mmm', s) <> 0) or (pos('mmmm', s) <> 0);
    except on EConvertError do
      Result := false;
    end;
end;

{ IsTimeFormat checks if the format string s is a time format. isLong is
  true if the string contains hours, minutes and seconds (two colons).
  isAMPM is true if the string contains "AM/PM", "A/P" or "AMPM".
  isInterval is true if the string contains square bracket codes for time intervals.
  SecDecimals is the number of decimals for the seconds. }
function IsTimeFormat(s: String; out isLong, isAMPM, isInterval: Boolean;
  out SecDecimals: Word): Boolean;
var
  p, pdp, i, count: Integer;
begin
  isLong := false;
  isAMPM := false;
  SecDecimals := 0;

  // Time parts are separated by a colon
  p := pos(':', s);
  result := p > 0;

  if Result then begin
    count := 1;
    s := Uppercase(s);

    // If there are is a second colon s is a "long" time format
    for i:=p+1 to Length(s) do
      if s[i] = ':' then begin
        isLong := true;
        break;
      end;

    // Seek for "AM/PM" etc to detect that specific format
    isAMPM := (pos('AM/PM', s) > 0) or (pos('A/P', s) > 0) or (pos('AMPM', s) > 0);

    // Look for square brackets indicating the interval format.
    p := pos('[', s);
    if p > 0 then isInterval := (pos(']', s) > 0) else isInterval := false;

    // Count decimals
    pdp := pos('.', s);
    if (pdp > 0) then
      for i:=pdp+1 to Length(s) do
        if (s[i] in ['0', 'z', 'Z']) then
          inc(SecDecimals)
        else
          break;   // ignore characters after last 0

    // Check validity of format string
    try
      FormatDateTime(s, now);
    except on EConvertError do
      Result := false;
    end;
  end;
end;

{ Formats the number AValue in "scientific" format with the given number of
  decimals. "Scientific" is the same as "exponential", but with exponents rounded
  to multiples of 3 (like for "kilo" - "Mega" - "Giga" etc.). }
function SciFloat(AValue: Double; ADecimals: Word): String;
var
  m: Double;
  ex: Integer;
begin
  if AValue = 0 then
    Result := '0.0'
  else begin
    ex := floor(log10(abs(AValue)));  // exponent
    // round exponent to multiples of 3
    ex := (ex div 3) * 3;
    if ex < 0 then dec(ex, 3);
    m := AValue * Power(10, -ex);     // mantisse
    Result := Format('%.*fE%d', [ADecimals, m, ex]);
  end;
end;
               (*
{ Formats the number AValue as a time string according to the format string.
  If the hour part is between square brackets it can be greater than 24 hours.
  Dto for the minutes or seconds part, with the higher-value part being added
  and no longer being shown explicitly.
  Example:
    AValue = 1:30:02, FormatStr = "[mm]:ss]" --> "90:02" }
function TimeIntervalToString(AValue: TDateTime; AFormatStr: String): String;
var
  hrs, mins, secs: Integer;
  diff: Double;
  h,m,s,z: Word;
  ts: String;
  fmt: String;
  p: Integer;
begin                         {
  fmt := Lowercase(AFormatStr);
  p := pos('h]', fmt);
  if p > 0 then begin
    System.Delete(fmt, 1, p+2);
    Result := FormatDateTime(fmt, AValue);
    DecodeTime(frac(abs(AValue)), h, m, s, z);
    hrs := h + trunc(abs(AValue))*24;
    Result := FormatDateTime(fmt, AValue);
  end;
  for i
  p := pos('h
  }
  ts := DefaultFormatSettings.TimeSeparator;
  DecodeTime(frac(abs(AValue)), h, m, s, z);
  hrs := h + trunc(abs(AValue))*24;
  if z > 499 then inc(s);
  if hrs > 0 then
    Result := Format('%d%s%.2d%s%.2d', [hrs, ts, m, ts, s])
  else
    Result := Format('%d%s%.2d', [m, ts, s]);
  if AValue < 0.0 then Result := '-' + Result;
end;
          *)
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


{******************************************************************************}
{******************************************************************************}
{                   Patch for SysUtils.FormatDateTime                          }
{ Remove when the feature of square brackets in time format masks is in rtl    }
{******************************************************************************}
{******************************************************************************}

// Copied from "fpc/rtl/objpas/sysutils/datei.inc"
procedure DateTimeToString(out Result: string; const FormatStr: string; const DateTime: TDateTime; const FormatSettings: TFormatSettings);
var
  ResultLen: integer;
  ResultBuffer: array[0..255] of char;
  ResultCurrent: pchar;

{$IFDEF MSWindows}
  isEnable_E_Format : Boolean;
  isEnable_G_Format : Boolean;
  eastasiainited : boolean;
{$ENDIF MSWindows}

(* This part is in the original code. It is not needed here and avoids a
   dependency on the unit Windows.

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
        '[': isInterval := true;
        ']': isInterval := false;
        ' ', 'C', 'D', 'H', 'M', 'N', 'S', 'T', 'Y','Z' :
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
                StoreInt(Minute + Hour*60 + trunc(DateTime)*24*60, 0)
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
                StoreInt(Hour + trunc(DateTime)*24, 0)
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
                   StoreInt(Minute + 60*Hour + 60*24*trunc(DateTime), 0)
                 else
                 if Count = 1 then
                   StoreInt(Minute, 0)
                 else
                   StoreInt(Minute, 2);
            'S': if isInterval then
                   StoreInt(Second + Minute*60 + Hour*60*60 + trunc(DateTime)*24*60*60, 0)
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

(* This part is in the original code. It is not needed here and avoids a
   dependency on the unit Windows.

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

function FormatDateTime(const FormatStr: string; DateTime: TDateTime): string;
begin
  DateTimeToString(Result, FormatStr, DateTime, DefaultFormatSettings);
end;

end.

