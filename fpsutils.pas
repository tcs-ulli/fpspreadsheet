{@@ ----------------------------------------------------------------------------
  Utility functions and declarations for FPSpreadsheet

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpsutils;

// to do: Remove the patched FormatDateTime when the feature of square brackets
//        in time format codes is in the rtl
// to do: Remove the declaration UTF8FormatSettings and InitUTF8FormatSettings
//        when this same modification is in LazUtils of Laz stable


{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, StrUtils,
  fpspreadsheet, fpsNumFormatParser;

// Exported types
type
  {@@ Selection direction along column or along row }
  TsSelectionDirection = (fpsVerticalSelection, fpsHorizontalSelection);

  {@@ Set of characters }
  TsDecsChars = set of char;

  {@@ Options for the FormatDateTime function to activate time interval strings
      with more than 24 hours.
      Will be removed when this feature is in the stable release of FPC }
  TFormatDateTimeOption = (fdoInterval);

  {@@ Options for the FormatDateTime function to activate time interval strings
      with more than 24 hours.
      Will be removed when this feature is in the stable release of FPC }
  TFormatDateTimeOptions =  set of TFormatDateTimeOption;

const
  {@@ Date formatting string for unambiguous date/time display as strings
      Can be used for text output when date/time cell support is not available }
  ISO8601Format='yyyymmdd"T"hhmmss';
  {@@ Extended ISO 8601 date/time format, used in e.g. ODF/opendocument }
  ISO8601FormatExtended='yyyy"-"mm"-"dd"T"hh":"mm":"ss';
  {@@  ISO 8601 date-only format, used in ODF/opendocument }
  ISO8601FormatDateOnly='yyyy"-"mm"-"dd';
  {@@  ISO 8601 time-only format, used in ODF/opendocument }
  ISO8601FormatTimeOnly='"PT"hh"H"nn"M"ss"S"';
  {@@ ISO 8601 time-only format, with hours overflow }
  ISO8601FormatHoursOverflow='"PT"[hh]"H"nn"M"ss.zz"S"';

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
  out AFirstCellRow, AFirstCellCol, ACount: Cardinal;
  out ADirection: TsSelectionDirection): Boolean;
function ParseCellRangeString(const AStr: string;
  out AFirstCellRow, AFirstCellCol, ALastCellRow, ALastCellCol: Cardinal;
  out AFlags: TsRelFlags): Boolean; overload;
function ParseCellRangeString(const AStr: string;
  out AFirstCellRow, AFirstCellCol, ALastCellRow, ALastCellCol: Cardinal): Boolean; overload;
function ParseCellString(const AStr: string;
  out ACellRow, ACellCol: Cardinal; out AFlags: TsRelFlags): Boolean; overload;
function ParseCellString(const AStr: string;
  out ACellRow, ACellCol: Cardinal): Boolean; overload;
function ParseCellRowString(const AStr: string;
  out AResult: Cardinal): Boolean;
function ParseCellColString(const AStr: string;
  out AResult: Cardinal): Boolean;

function GetColString(AColIndex: Integer): String;
function GetCellString(ARow,ACol: Cardinal;
  AFlags: TsRelFlags = [rfRelRow, rfRelCol]): String;
function GetCellRangeString(ARow1, ACol1, ARow2, ACol2: Cardinal;
  AFlags: TsRelFlags = [rfRelRow, rfRelCol, rfRelRow2, rfRelCol2];
  Compact: Boolean = false): String;

function GetErrorValueStr(AErrorValue: TsErrorValue): String;

function IfThen(ACondition: Boolean; AValue1,AValue2: TsNumberFormat): TsNumberFormat; overload;

function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean;
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsDateTimeFormat(AFormatStr: String): Boolean; overload;
function IsTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsTimeFormat(AFormatStr: String): Boolean; overload;

procedure BuildCurrencyFormatList(AList: TStrings;
  APositive: Boolean; AValue: Double; const ACurrencySymbol: String);
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
function DayNamesToString(const ADayNames: TWeekNameArray;
  const AEmptyStr: String): String;
function MakeLongDateFormat(ADateFormat: String): String;
function MakeShortDateFormat(ADateFormat: String): String;
function MonthNamesToString(const AMonthNames: TMonthNameArray;
  const AEmptyStr: String): String;
function SpecialDateTimeFormat(ACode: String;
  const AFormatSettings: TFormatSettings; ForWriting: Boolean): String;
procedure SplitFormatString(const AFormatString: String; out APositivePart,
  ANegativePart, AZeroPart: String);

procedure MakeTimeIntervalMask(Src: String; var Dest: String);

// These two functions are copies of fpc trunk until they are available in stable fpc.
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  Options : TFormatDateTimeOptions = []): string;
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []): string;

function TryStrToFloatAuto(AText: String; out ANumber: Double;
  out ADecimalSeparator, AThousandSeparator: Char; out AWarning: String): Boolean;

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
function ColorToHTMLColorStr(AValue: TsColorValue; AExcelDialect: Boolean = false): String;
function UTF8TextToXMLText(AText: ansistring): ansistring;
function ValidXMLText(var AText: ansistring; ReplaceSpecialChars: Boolean = true): Boolean;

function TintedColor(AColor: TsColorValue; tint: Double): TsColorValue;
function HighContrastColor(AColorValue: TsColorValue): TsColor;

function AnalyzeCompareStr(AString: String; out ACompareOp: TsCompareOperation): String;

function InitSortParams(ASortByCols: Boolean = true; ANumSortKeys: Integer = 1;
  ASortPriority: TsSortPriority = spNumAlpha): TsSortParams;

procedure AppendToStream(AStream: TStream; const AString: String); inline; overload;
procedure AppendToStream(AStream: TStream; const AString1, AString2: String); inline; overload;
procedure AppendToStream(AStream: TStream; const AString1, AString2, AString3: String); inline; overload;

{ For silencing the compiler... }
procedure Unused(const A1);
procedure Unused(const A1, A2);
procedure Unused(const A1, A2, A3);

{ For debugging purposes }
procedure DumpFontsToFile(AWorkbook: TsWorkbook; AFileName: String);

var
  {@@ Default value for the screen pixel density (pixels per inch). Is needed
  for conversion of distances to pixels}
  ScreenPixelsPerInch: Integer = 96;
  {@@ FPC format settings for which all strings have been converted to UTF8 }
  UTF8FormatSettings: TFormatSettings;

implementation

uses
  Math, lazutf8, fpsStrings;

type
  TRGBA = record r, g, b, a: byte end;

const
  POS_CURR_FMT: array[0..3] of string = (
    // Format parameter 0 is "value", parameter 1 is "currency symbol"
    ('%1:s%0:s'),        // 0: $1
    ('%0:s%1:s'),        // 1: 1$
    ('%1:s %0:s'),       // 2: $ 1
    ('%0:s %1:s')        // 3: 1 $
  );
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

{******************************************************************************}
{                       Endianess helper functions                             }
{******************************************************************************}

{ Excel files are all written with little endian byte order,
  so it's necessary to swap the data to be able to build a
  correct file on big endian systems.

  The routines WordToLE, DWordToLE, IntegerToLE etc are preferable to
  System unit routines because they ensure that the correct overloaded version
  of the conversion routines will be used, avoiding typecasts which are less readable.

  They also guarantee delphi compatibility. For Delphi we just support
  big-endian isn't support, because Delphi doesn't support it.
}

{@@ ----------------------------------------------------------------------------
  WordLEToLE converts a word value from big-endian to little-endian byte order.

  @param   AValue  Big-endian word value
  @return          Little-endian word value
-------------------------------------------------------------------------------}
function WordToLE(AValue: Word): Word;
begin
  {$IFDEF FPC}
    Result := NtoLE(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

{@@ ----------------------------------------------------------------------------
  DWordLEToLE converts a DWord value from big-endian to little-endian byte-order.

  @param   AValue  Big-endian DWord value
  @return          Little-endian DWord value
-------------------------------------------------------------------------------}
function DWordToLE(AValue: Cardinal): Cardinal;
begin
  {$IFDEF FPC}
    Result := NtoLE(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

{@@ ----------------------------------------------------------------------------
  Converts an integer value from big-endian to little-endian byte-order.

  @param   AValue  Big-endian integer value
  @return          Little-endian integer value
-------------------------------------------------------------------------------}
function IntegerToLE(AValue: Integer): Integer;
begin
  {$IFDEF FPC}
    Result := NtoLE(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

{@@ ----------------------------------------------------------------------------
  Converts a word value from little-endian to big-endian byte-order.

  @param   AValue  Little-endian word value
  @return          Big-endian word value
-------------------------------------------------------------------------------}
function WordLEtoN(AValue: Word): Word;
begin
  {$IFDEF FPC}
    Result := LEtoN(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

{@@ ----------------------------------------------------------------------------
  Converts a DWord value from little-endian to big-endian byte-order.

  @param   AValue  Little-endian DWord value
  @return          Big-endian DWord value
-------------------------------------------------------------------------------}
function DWordLEtoN(AValue: Cardinal): Cardinal;
begin
  {$IFDEF FPC}
    Result := LEtoN(AValue);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

{@@ ----------------------------------------------------------------------------
  Converts a widestring from big-endian to little-endian byte-order.

  @param   AValue  Big-endian widestring
  @return          Little-endian widestring
-------------------------------------------------------------------------------}
function WideStringToLE(const AValue: WideString): WideString;
{$IFNDEF FPC}
var
  j: integer;
{$ENDIF}
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

{@@ ----------------------------------------------------------------------------
  Converts a widestring from little-endian to big-endian byte-order.

  @param   AValue  Little-endian widestring
  @return          Big-endian widestring
-------------------------------------------------------------------------------}
function WideStringLEToN(const AValue: WideString): WideString;
{$IFNDEF FPC}
var
  j: integer;
{$ENDIF}
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

{@@ ----------------------------------------------------------------------------
  Converts the RGB part of a LongRGB logical structure to its physical representation.
  In other words: RGBA (where A is 0 and omitted in the function call) => ABGR
  Needed for conversion of palette colors.

  @param  RGB  DWord value containing RGBA bytes in big endian byte-order
  @return      DWord containing RGB bytes in little-endian byte-order (A = 0)
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Parses strings like A5:A10 into an selection interval information

  @param  AStr           Cell range string, such as A5:A10
  @param  AFirstCellRow  Row index of the first cell of the range (output)
  @param  AFirstCellCol  Column index of the first cell of the range (output)
  @param  ACount         Number of cells included in the range (output)
  @param  ADirection     fpsVerticalSelection if the range is along a column,
                         fpsHorizontalSelection if the range is along a row

  @return                false if the string is not a valid cell range
-------------------------------------------------------------------------------}
function ParseIntervalString(const AStr: string;
  out AFirstCellRow, AFirstCellCol, ACount: Cardinal;
  out ADirection: TsSelectionDirection): Boolean;
var
  //Cells: TStringList;
  LastCellRow, LastCellCol: Cardinal;
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

{@@ ----------------------------------------------------------------------------
  Parses strings like A5:C10 into a range selection information.
  Returns in AFlags also information on relative/absolute cells.

  @param  AStr           Cell range string, such as A5:C10
  @param  AFirstCellRow  Row index of the top/left cell of the range (output)
  @param  AFirstCellCol  Column index of the top/left cell of the range (output)
  @param  ALastCellRow   Row index of the bottom/right cell of the range (output)
  @param  ALastCellCol   Column index of the bottom/right cell of the range (output)
  @param  AFlags         a set containing an element for AFirstCellRow, AFirstCellCol,
                         ALastCellRow, ALastCellCol if they represent relative
                         cell addresses.

  @return                false if the string is not a valid cell range
-------------------------------------------------------------------------------}
function ParseCellRangeString(const AStr: string;
  out AFirstCellRow, AFirstCellCol, ALastCellRow, ALastCellCol: Cardinal;
  out AFlags: TsRelFlags): Boolean;
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


{@@ ----------------------------------------------------------------------------
  Parses strings like A5:C10 into a range selection information.
  Information on relative/absolute cells is ignored.

  @param  AStr           Cell range string, such as A5:C10
  @param  AFirstCellRow  Row index of the top/left cell of the range (output)
  @param  AFirstCellCol  Column index of the top/left cell of the range (output)
  @param  ALastCellRow   Row index of the bottom/right cell of the range (output)
  @param  ALastCellCol   Column index of the bottom/right cell of the range (output)
  @return                false if the string is not a valid cell range
--------------------------------------------------------------------------------}
function ParseCellRangeString(const AStr: string;
  out AFirstCellRow, AFirstCellCol, ALastCellRow, ALastCellCol: Cardinal): Boolean;
var
  flags: TsRelFlags;
begin
  Result := ParseCellRangeString(AStr,
    AFirstCellRow, AFirstCellCol,
    ALastCellRow, ALastCellCol,
    flags
  );
end;

{@@ ----------------------------------------------------------------------------
  Parses a cell string, like 'A1' into zero-based column and row numbers
  Note that there can be several letters to address for more than 26 columns.
  'AFlags' indicates relative addresses.

  @param  AStr      Cell range string, such as A1
  @param  ACellRow  Row index of the top/left cell of the range (output)
  @param  ACellCol  Column index of the top/left cell of the range (output)
  @param  AFlags    A set containing an element for ACellRow and/or ACellCol,
                    if they represent a relative cell address.
  @return           False if the string is not a valid cell range

  @example "AMP$200" --> (rel) column 1029 (= 26*26*1 + 26*16 + 26 - 1)
                         (abs) row = 199 (abs)
-------------------------------------------------------------------------------}
function ParseCellString(const AStr: String; out ACellRow, ACellCol: Cardinal;
  out AFlags: TsRelFlags): Boolean;

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
        ACellCol := Cardinal(ord(UpCase(AStr[i])) - ord('A')) + 1 + ACellCol * 26;
        if ACellCol >= MAX_COL_COUNT then
          // too many columns (dropping this limitation could cause overflow
          // if a too long string is passed
          exit;
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

    if i > Length(AStr) then
      exit;

    isAbs := (AStr[i] = '$');
    if isAbs then inc(i);

    if i > Length(AStr) then
      exit;

    // Scan digits
    while (i <= Length(AStr)) do begin
      if (AStr[i] in DIGITS) then begin
        ACellRow := Cardinal(ord(AStr[i]) - ord('0')) + ACellRow * 10;
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

{@@ ----------------------------------------------------------------------------
  Parses a cell string, like 'A1' into zero-based column and row numbers
  Note that there can be several letters to address for more than 26 columns.

  For compatibility with old version which does not return flags for relative
  cell addresses.

  @param  AStr      Cell range string, such as A1
  @param  ACellRow  Row index of the top/left cell of the range (output)
  @param  ACellCol  Column index of the top/left cell of the range (output)
  @return           False if the string is not a valid cell range
-------------------------------------------------------------------------------}
function ParseCellString(const AStr: string;
  out ACellRow, ACellCol: Cardinal): Boolean;
var
  flags: TsRelFlags;
begin
  Result := ParseCellString(AStr, ACellRow, ACellCol, flags);
end;

{@@ ----------------------------------------------------------------------------
  Parses a cell row string to a zero-based row number.

  @param  AStr      Cell row string, such as '1', 1-based!
  @param  AResult   Index of the row (zero-based!) (putput)
  @return           False if the string is not a valid cell row string
-------------------------------------------------------------------------------}
function ParseCellRowString(const AStr: string; out AResult: Cardinal): Boolean;
begin
  try
    AResult := StrToInt(AStr) - 1;
  except
    Result := False;
  end;
  Result := True;
end;

{@@ ----------------------------------------------------------------------------
  Parses a cell column string, like 'A' or 'CZ', into a zero-based column number.
  Note that there can be several letters to address more than 26 columns.

  @param  AStr      Cell range string, such as A1
  @param  AResult   Zero-based index of the column (output)
  @return           False if the string is not a valid cell column string
-------------------------------------------------------------------------------}
function ParseCellColString(const AStr: string; out AResult: Cardinal): Boolean;
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

{@@ ----------------------------------------------------------------------------
  Calculates an Excel column name ('A', 'B' etc) from the zero-based column index

  @param  AColIndex   Zero-based column index
  @return  Letter-based column name string. Can contain several letter in case of
           more than 26 columns
-------------------------------------------------------------------------------}
function GetColString(AColIndex: Integer): String;
{ Code adapted from:
  http://stackoverflow.com/questions/12796973/vba-function-to-convert-column-number-to-letter }
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

{@@ ----------------------------------------------------------------------------
  Calculates a cell address string from zero-based column and row indexes and
  the relative address state flags.

  @param   ARowIndex   Zero-based row index
  @param   AColIndex   Zero-based column index
  @param   AFlags      An optional set containing an entry for column and row
                       if these addresses are relative. By default, relative
                       addresses are assumed.
  @return  Excel type of cell address containing $ characters for absolute
           address parts.
  @example ARowIndex = 0, AColIndex = 0, AFlags = [rfRelRow] --> $A1
-------------------------------------------------------------------------------}
function GetCellString(ARow, ACol: Cardinal;
  AFlags: TsRelFlags = [rfRelRow, rfRelCol]): String;
begin
  Result := Format('%s%s%s%d', [
    RELCHAR[rfRelCol in AFlags], GetColString(ACol),
    RELCHAR[rfRelRow in AFlags], ARow+1
  ]);
end;

{@@ ----------------------------------------------------------------------------
  Calculates a cell range address string from zero-based column and row indexes
  and the relative address state flags.

  @param   ARow1       Zero-based index of the first row in the range
  @param   ACol1       Zero-based index of the first column in the range
  @param   ARow2       Zero-based index of the last row in the range
  @param   ACol2       Zero-based index of the last column in the range
  @param   AFlags      A set containing an entry for first and last column and
                       row if their addresses are relative.
  @param   Compact     If the range consists only of a single cell and compact
                       is true then the simple cell string is returned (e.g. A1).
                       If compact is false then the cell is repeated (e.g. A1:A1)
  @return  Excel type of cell address range containing '$' characters for absolute
           address parts and a ':' to separate the first and last cells of the
           range
  @example ARow1 = 0, ACol1 = 0, ARow = 2, ACol = 1, AFlags = [rfRelRow, rfRelRow2]
           --> $A1:$B3
-------------------------------------------------------------------------------}
function GetCellRangeString(ARow1, ACol1, ARow2, ACol2: Cardinal;
  AFlags: TsRelFlags = [rfRelRow, rfRelCol, rfRelRow2, rfRelCol2];
  Compact: Boolean = false): String;
begin
  if Compact and (ARow1 = ARow2) and (ACol1 = ACol2) then
    Result := GetCellString(ARow1, ACol1, AFlags)
  else
    Result := Format('%s%s%s%d:%s%s%s%d', [
      RELCHAR[rfRelCol in AFlags], GetColString(ACol1),
      RELCHAR[rfRelRow in AFlags], ARow1 + 1,
      RELCHAR[rfRelCol2 in AFlags], GetColString(ACol2),
      RELCHAR[rfRelRow2 in AFlags], ARow2 + 1
    ]);
end;

{@@ ----------------------------------------------------------------------------
  Returns the message text assigned to an error value

  @param   AErrorValue  Error code as defined by TsErrorvalue
  @return  Text corresponding to the error code.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Helper function to reduce typing: "if a conditions is true return the first
  number format, otherwise return the second format"

  @param   ACondition   Boolean expression
  @param   AValue1      First built-in number format code
  @param   AValue2      Second built-in number format code
  @return  AValue1 if ACondition is true, AValue2 otherwise.
-------------------------------------------------------------------------------}
function IfThen(ACondition: Boolean;
  AValue1, AValue2: TsNumberFormat): TsNumberFormat;
begin
  if ACondition then Result := AValue1 else Result := AValue2;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given number format code is for currency,
  i.e. requires currency symbol.

  @param  AFormat   Built-in number format identifier to be checked
  @return True if AFormat is nfCurrency or nfCurrencyRed, false otherwise.
-------------------------------------------------------------------------------}
function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [nfCurrency, nfCurrencyRed];
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given number format code is for date/time values.

  @param   AFormat  Built-in number format identifier to be checked
  @return  True if AFormat is a date/time format (such as nfShortTime),
           false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [{nfFmtDateTime, }nfShortDateTime, nfShortDate, nfLongDate,
    nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM, nfTimeInterval];
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
  parser := TsNumFormatParser.Create(nil, AFormatStr);
  try
    Result := parser.IsDateTimeFormat;
  finally
    parser.Free;
  end;
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
  parser := TsNumFormatParser.Create(nil, AFormatStr);
  try
    Result := parser.IsTimeFormat;
  finally
    parser.Free;
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

  @param  ADialect        Determines whether the format string is for use by
                          fpspreadsheet (nfdDefault) or by Excel (nfdExcel)
  @param  ANumberFormat   Identifier of the built-in number format for which the
                          format string is to be generated.
  @param  AFormatSettings FormatSettings to be applied (used to extract default
                          values for the next parameters)
  @param  ADecimals       number of decimal places. If < 0, the CurrencyDecimals
                          of the FormatSettings is used.
  @param  APosCurrFormat  Identifier for the order of currency symbol, value and
                          spaces of positive values
                          - see pcfXXXX constants in fpspreadsheet.pas.
                          If < 0, the CurrencyFormat of the FormatSettings is used.
  @param  ANegCurrFormat  Identifier for the order of currency symbol, value and
                          spaces of negative values. Specifies also usage of ().
                          - see ncfXXXX constants in fpspreadsheet.pas.
                          If < 0, the NegCurrFormat of the FormatSettings is used.
  @param  ACurrencySymbol Name of the currency, like $ or USD.
                          If ? the CurrencyString of the FormatSettings is used.

  @return String of formatting codes, such as '"$"#,##0.00;("$"#,##0.00);"$"0.00'
-------------------------------------------------------------------------------}
function BuildCurrencyFormatString(ADialect: TsNumFormatDialect;
  ANumberFormat: TsNumberFormat; const AFormatSettings: TFormatSettings;
  ADecimals, APosCurrFormat, ANegCurrFormat: Integer; ACurrencySymbol: String): String;
{
const
  POS_FMT: array[0..3] of string = (
    // Format parameter 0 is "value", parameter 1 is "currency symbol"
    ('"%1:s"%0:s'),        // 0: $1
    ('%0:s"%1:s"'),        // 1: 1$
    ('"%1:s" %0:s'),       // 2: $ 1
    ('%0:s "%1:s"')        // 3: 1 $
  );
  NEG_FMT: array[0..15] of string = (
    ('("%1:s"%0:s)'),      //  0: ($1)
    ('-"%1:s"%0:s'),       //  1: -$1
    ('"%1:s"-%0:s'),       //  2: $-1
    ('"%1:s"%0:s-'),       //  3: $1-
    ('(%0:s"%1:s")'),      //  4: (1$)
    ('-%0:s"%1:s"'),       //  5: -1$
    ('%0:s-"%1:s"'),       //  6: 1-$
    ('%0:s"%1:s"-'),       //  7: 1$-
    ('-%0:s "%1:s"'),      //  8: -1 $
    ('-"%1:s" %0:s'),      //  9: -$ 1
    ('%0:s "%1:s"-'),      // 10: 1 $-
    ('"%1:s" %0:s-'),      // 11: $ 1-
    ('"%1:s" -%0:s'),      // 12: $ -1
    ('%0:s- "%1:s"'),      // 13: 1- $
    ('("%1:s" %0:s)'),     // 14: ($ 1)
    ('(%0:s "%1:s")')      // 15: (1 $)
  );
  }
var
  decs: String;
  pcf, ncf: Byte;
  p, n: String;
  negRed: Boolean;
begin
  pcf := IfThen(APosCurrFormat < 0, AFormatSettings.CurrencyFormat, APosCurrFormat);
  ncf := IfThen(ANegCurrFormat < 0, AFormatSettings.NegCurrFormat, ANegCurrFormat);
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
  if ADialect = nfdExcel then
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
            + Format(p, ['0'+decs, ACurrencySymbol]);
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
    Result := Result + ';0' + decs;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Builds a number format string from the number format code and the count of
  decimal places.

  @param  ANumberFormat   Identifier of the built-in numberformat for which a
                          format string is to be generated
  @param  AFormatSettings FormatSettings for default parameters
  @param  ADecimals       Number of decimal places. If < 0 the CurrencyDecimals
                          value of the FormatSettings is used.

  @return String of formatting codes, such as '#,##0.00' for nfFixedTh and 2 decimals
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
    nfCurrency, nfCurrencyRed:
      Result := BuildCurrencyFormatString(nfdDefault, ANumberFormat, AFormatSettings,
        ADecimals, AFormatSettings.CurrencyFormat, AFormatSettings.NegCurrFormat,
        AFormatSettings.CurrencyString);
//      raise Exception.Create('BuildNumberFormatString: Use BuildCurrencyFormatString '+
//        'to create a format string for currency values.');
    nfShortDateTime, nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfTimeInterval:
      raise Exception.Create('BuildNumberFormatString: Use BuildDateTimeFormatSstring '+
        'to create a format string for date/time values.');
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds an AM/PM format code to a pre-built time formatting string. The strings
  replacing "AM" or "PM" in the final formatted number are taken from the
  TimeAMString or TimePMString of the given FormatSettings.

  @param   ATimeFormatString  String of time formatting codes (such as 'hh:nn')
  @param   AFormatSettings    FormatSettings for locale-dependent information
  @result  Formatting string with AM/PM option activated.

  Example:  ATimeFormatString = 'hh:nn' ==> 'hh:nn AM/PM'
-------------------------------------------------------------------------------}
function AddAMPM(const ATimeFormatString: String;
  const AFormatSettings: TFormatSettings): String;
var
  am, pm: String;
begin
  am := IfThen(AFormatSettings.TimeAMString <> '', AFormatSettings.TimeAMString, 'AM');
  pm := IfThen(AFormatSettings.TimePMString <> '', AFormatSettings.TimePMString, 'PM');
  Result := Format('%s %s/%s', [StripAMPM(ATimeFormatString), am, pm]);
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

{@@ ----------------------------------------------------------------------------
  Counts how many decimal places are coded into a given formatting string.

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
  The given format string is assumed to represent a time interval, i.e. its
  first time symbol must be enclosed by square brackets. Checks if this is true,
  and adds the brackes if not.

  @param   AFormatString   String with time formatting codes
  @return  Unchanged format string if its first time code is in square brackets
           (as in '[h]:nn:ss'), if not, the first time code is enclosed in
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
  Concatenates the day names specified in ADayNames to a single string. If all
  daynames are empty AEmptyStr is returned

  @param   ADayNames   Array[1..7] of day names as used in the Formatsettings
  @param   AEmptyStr   Is returned if all day names are empty
  @return  String having all day names concatenated and separated by the
           DefaultFormatSettings.ListSeparator
-------------------------------------------------------------------------------}
function DayNamesToString(const ADayNames: TWeekNameArray;
  const AEmptyStr: String): String;
var
  i: Integer;
  isEmpty: Boolean;
begin
  isEmpty := true;
  for i:=1 to 7 do
    if ADayNames[i] <> '' then
    begin
      isEmpty := false;
      break;
    end;

  if isEmpty then
    Result := AEmptyStr
  else
  begin
    Result := ADayNames[1];
    for i:=2 to 7 do
      Result := Result + DefaultFormatSettings.ListSeparator + ' ' + ADayNames[i];
  end;
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
  Concatenates the month names specified in AMonthNames to a single string.
  If all month names are empty AEmptyStr is returned

  @param   AMonthNames  Array[1..12] of month names as used in the Formatsettings
  @param   AEmptyStr    Is returned if all month names are empty
  @return  String having all month names concatenated and separated by the
           DefaultFormatSettings.ListSeparator
-------------------------------------------------------------------------------}
function MonthNamesToString(const AMonthNames: TMonthNameArray;
  const AEmptyStr: String): String;
var
  i: Integer;
  isEmpty: Boolean;
begin
  isEmpty := true;
  for i:=1 to 12 do
    if AMonthNames[i] <> '' then
    begin
      isEmpty := false;
      break;
    end;

  if isEmpty then
    Result := AEmptyStr
  else
  begin
    Result := AMonthNames[1];
    for i:=2 to 12 do
      Result := Result + DefaultFormatSettings.ListSeparator + ' ' + AMonthNames[i];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates the formatstrings for the date/time codes "dm", "my", "ms" and "msz"
  out of the formatsettings.

  @param   ACode   Quick formatting code for parts of date/time number formats;
                     "dm" = day + month
                     "my" = month + year
                     "ms" = minutes + seconds
                     "msz" = minutes + seconds + fractions of a second
  @return  String of formatting codes according to the parameter ACode
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Currency formatting strings consist of three parts, separated by
  semicolons, which are valid for positive, negative or zero values.
  Splits such a formatting string at the positions of the semicolons and
  returns the sections. If semicolons are used for other purposed within
  sections they have to be quoted by " or escaped by \. If the formatting
  string contains less sections than three the missing strings are returned
  as empty strings.

  @param   AFormatString  String of number formatting codes.
  @param   APositivePart  First section of the formatting string which is valid
                          for positive numbers (or positive and zero, if there
                          are only two sections)
  @param   ANegativePart  Second section of the formatting string which is valid
                          for negative numbers
  @param   AZeroPart      Third section of the formatting string for zero.
-------------------------------------------------------------------------------}
procedure SplitFormatString(const AFormatString: String; out APositivePart,
  ANegativePart, AZeroPart: String);

  procedure AddToken(AToken: Char; AWhere:Byte);
  begin
    case AWhere of
      0: APositivePart := APositivePart + AToken;
      1: ANegativePart := ANegativePart + AToken;
      2: AZeroPart := AZeroPart + AToken;
    end;
  end;

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
      '"': begin   // Let quoted text intact
             AddToken(token, where);
             inc(P);
             token := P^;
             while (P < PEnd) and (token <> '"') do begin
               AddToken(token, where);
               inc(P);
               token := P^;
             end;
             AddToken(token, where);
           end;
      ';': begin  // Separator between parts
             inc(where);
             if where = 3 then
               exit;
           end;
      '\': begin  // Skip "Escape" character and add next char immediately
             inc(P);
             token := P^;
             AddToken(token, where);
           end;
      else AddToken(token, where);
    end;
    inc(P);
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
  Converts a string to a floating point number. No assumption on decimal and
  thousand separator are made.
  Is needed for reading CSV files.
-------------------------------------------------------------------------------}
function TryStrToFloatAuto(AText: String; out ANumber: Double;
  out ADecimalSeparator, AThousandSeparator: Char; out AWarning: String): Boolean;
var
  i: Integer;
  testSep: Char;
  testSepPos: Integer;
  lastDigitPos: Integer;
  isPercent: Boolean;
  isExp: Boolean;
  fs: TFormatSettings;
  done: Boolean;
begin
  Result := false;
  AWarning := '';
  ADecimalSeparator := #0;
  AThousandSeparator := #0;
  if AText = '' then
    exit;

  fs := DefaultFormatSettings;

  // We scan the string starting from its end. If we find a point or a comma,
  // we have a candidate for the decimal or thousand separator. If we find
  // the same character again it was a thousand separator, if not it was
  // a decimal separator.

  // There is one amgiguity: Using a thousand separator for number < 1.000.000,
  // but no decimal separator misinterprets the thousand separator as a
  // decimal separator.

  done := false;      // Indicates that both decimal and thousand separators are found
  testSep := #0;      // Separator candidate to be tested
  testSepPos := 0;    // Position of this separator candidate in the string
  lastDigitPos := 0;  // Position of the last numerical digit
  isExp := false;     // Flag for exponential format
  isPercent := false; // Flag for percentage format

  i := Length(AText);    // Start at end...
  while i >= 1 do        // ...and search towards start
  begin
    case AText[i] of
      '0'..'9':
        if (lastDigitPos = 0) and (AText[i] in ['0'..'9']) then
          lastDigitPos := i;

      'e', 'E':
        isExp := true;

      '%':
        begin
          isPercent := true;
          // There may be spaces before the % sign which we don't want
          dec(i);
          while (i >= 1) do
            if AText[i] = ' ' then
              dec(i)
            else
            begin
              inc(i);
              break;
            end;
        end;

      '+', '-':
        ;

      '.', ',':
        begin
          if testSep = #0 then begin
            testSep := AText[i];
            testSepPos := i;
          end;
          // This is the right-most separator candidate in the text
          // It can be a decimal or a thousand separator.
          // Therefore, we continue searching from here.
          dec(i);
          while i >= 1 do
          begin
            if not (AText[i] in ['0'..'9', '+', '-', '.', ',']) then
              exit;

            // If we find the testSep character again it must be a thousand separator,
            // and there are no decimals.
            if (AText[i] = testSep) then
            begin
              // ... but only if there are 3 numerical digits in between
              if (testSepPos - i = 4) then
              begin
                fs.ThousandSeparator := testSep;
                // The decimal separator is the "other" character.
                if testSep = '.' then
                  fs.DecimalSeparator := ','
                else
                  fs.DecimalSeparator := '.';
                AThousandSeparator := fs.ThousandSeparator;
                ADecimalSeparator := #0; // this indicates that there are no decimals
                done := true;
                i := 0;
              end else
              begin
                Result := false;
                exit;
              end;
            end
            else
            // If we find the "other" separator character, then testSep was a
            // decimal separator and the current character is a thousand separator.
            // But there must be 3 digits in between.
            if AText[i] in ['.', ','] then
            begin
              if testSepPos - i <> 4 then  // no 3 digits in between --> no number, maybe a date.
                exit;
              fs.DecimalSeparator := testSep;
              fs.ThousandSeparator := AText[i];
              ADecimalSeparator := fs.DecimalSeparator;
              AThousandSeparator := fs.ThousandSeparator;
              done := true;
              i := 0;
            end;
            dec(i);
          end;
        end;

      else
        exit;  // Non-numeric character found, no need to continue

    end;
    dec(i);
  end;

  // Only one separator candicate found, we assume it is a decimal separator
  if (testSep <> #0) and not done then
  begin
    // Warning in case of ambiguous detection of separator. If only one separator
    // type is found and it is at the third position from the string's end it
    // might by a thousand separator or a decimal separator. We assume the
    // latter case, but create a warning.
    if (lastDigitPos - testSepPos = 3) and not isPercent then
      AWarning := Format(rsAmbiguousDecThouSeparator, [AText]);
    fs.DecimalSeparator := testSep;
    ADecimalSeparator := fs.DecimalSeparator;
    // Make sure that the thousand separator is different from the decimal sep.
    if testSep = '.' then fs.ThousandSeparator := ',' else fs.ThousandSeparator := '.';
  end;

  // Delete all thousand separators from the string - StrToFloat does not like them...
  AText := StringReplace(AText, fs.ThousandSeparator, '', [rfReplaceAll]);

  // Is the last character a percent sign?
  if isPercent then
    while (Length(AText) > 0) and (AText[Length(AText)] in ['%', ' ']) do
      Delete(AText, Length(AText), 1);

  // Try string-to-number conversion
  Result := TryStrToFloat(AText, ANumber, fs);

  // If successful ...
  if Result then
  begin
    // ... take care of the percentage sign
    if isPercent then
      ANumber := ANumber * 0.01;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Excel's unit of row heights is "twips", i.e. 1/20 point.
  Converts Twips to points.

  @param   AValue   Length value in twips
  @return  Value converted to points
-------------------------------------------------------------------------------}
function TwipsToPts(AValue: Integer): Single;
begin
  Result := AValue / 20;
end;

{@@ ----------------------------------------------------------------------------
  Converts points to twips (1 twip = 1/20 point)

  @param   AValue   Length value in points
  @return  Value converted to twips
-------------------------------------------------------------------------------}
function PtsToTwips(AValue: Single): Integer;
begin
  Result := round(AValue * 20);
end;

{@@ ----------------------------------------------------------------------------
  Converts centimeters to points (72 pts = 1 inch)

  @param   AValue  Length value in centimeters
  @return  Value converted to points
-------------------------------------------------------------------------------}
function cmToPts(AValue: Double): Double;
begin
  Result := AValue * 72 / 2.54;
end;

{@@ ----------------------------------------------------------------------------
  Converts points to centimeters

  @param   AValue   Length value in points
  @return  Value converted to centimeters
-------------------------------------------------------------------------------}
function PtsToCm(AValue: Double): Double;
begin
  Result := AValue / 72 * 2.54;
end;

{@@ ----------------------------------------------------------------------------
  Converts inches to points (72 pts = 1 inch)

  @param   AValue   Length value in inches
  @return  Value converted to points
-------------------------------------------------------------------------------}
function InToPts(AValue: Double): Double;
begin
  Result := AValue * 72;
end;

{@@ ----------------------------------------------------------------------------
  Converts millimeters to points (72 pts = 1 inch)

  @param   AValue   Length value in millimeters
  @return  Value converted to points
-------------------------------------------------------------------------------}
function mmToPts(AValue: Double): Double;
begin
  Result := AValue * 72 / 25.4;
end;

{@@ ----------------------------------------------------------------------------
  Converts points to millimeters

  @param    AValue   Length value in points
  @return   Value converted to millimeters
-------------------------------------------------------------------------------}
function PtsToMM(AValue: Double): Double;
begin
  Result := AValue / 72 * 25.4;
end;

{@@ ----------------------------------------------------------------------------
  Converts pixels to points.

  @param   AValue                Length value given in pixels
  @param   AScreenPixelsPerInch  Pixels per inch of the screen
  @return  Value converted to points
-------------------------------------------------------------------------------}
function pxToPts(AValue, AScreenPixelsPerInch: Integer): Double;
begin
  Result := (AValue / AScreenPixelsPerInch) * 72;
end;

{@@ ----------------------------------------------------------------------------
  Converts points to pixels
  @param   AValue                Length value given in points
  @param   AScreenPixelsPerInch  Pixels per inch of the screen
  @return  Value converted to pixels
-------------------------------------------------------------------------------}
function PtsToPx(AValue: Double; AScreenPixelsPerInch: Integer): Integer;
begin
  Result := Round(AValue / 72 * AScreenPixelsPerInch);
end;

{@@ ----------------------------------------------------------------------------
  Converts a HTML length string to points. The units are assumed to be the last
  two digits of the string, such as '1.25in'

  @param   AValue   HTML string representing a length with appended units code,
                    such as '1.25in'. These unit codes are accepted:
                    'px' (pixels), 'pt' (points), 'in' (inches), 'mm' (millimeters),
                    'cm' (centimeters).
  @return  Extracted length in points
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Converts a HTML color string to a TsColorValue. Need for the ODS file format.

  @param   AValue         HTML color string, such as '#FF0000'
  @return  rgb color value in little endian byte-sequence. This value is
           compatible with the TColor data type of the graphics unit.
-------------------------------------------------------------------------------}
function HTMLColorStrToColor(AValue: String): TsColorValue;
begin
  if AValue = '' then
    Result := scNotDefined
  else
  if AValue[1] = '#' then begin
    AValue[1] := '$';
    Result := LongRGBToExcelPhysical(DWord(StrToInt(AValue)));
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

{@@ ----------------------------------------------------------------------------
  Converts an rgb color value to a string as used in HTML code (for ods)

  @param   AValue          RGB color value (compatible with the TColor data type
                           of the graphics unit)
  @param   AExcelDialect   If TRUE, returned string is in Excels format for xlsx,
                           i.e. in AARRGGBB notation, like '00FF0000' for "red"
  @return  HTML-compatible string, like '#FF0000' (AExcelDialect = false)
-------------------------------------------------------------------------------}
function ColorToHTMLColorStr(AValue: TsColorValue; AExcelDialect: Boolean = false): String;
type
  TRGB = record r,g,b,a: Byte end;
var
  rgb: TRGB;
begin
  rgb := TRGB(AValue);
  if AExcelDialect then
    Result := Format('00%.2x%.2x%.2x', [rgb.r, rgb.g, rgb.b])
  else
    Result := Format('#%.2x%.2x%.2x', [rgb.r, rgb.g, rgb.b]);
end;

{@@ ----------------------------------------------------------------------------
  Converts a string encoded in UTF8 to a string usable in XML. For this purpose,
  some characters must be translated.

  @param   AText  input string encoded as UTF8
  @return  String usable in XML with some characters replaced by the HTML codes.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Checks a string for characters that are not permitted in XML strings.
  The function returns FALSE if a character <#32 is contained (except for
  #9, #10, #13), TRUE otherwise. Invalid characters are replaced by a box symbol.

  If ReplaceSpecialChars is TRUE, some other characters are converted
  to valid HTML codes by calling UTF8TextToXMLText

  @param  AText                String to be checked. Is replaced by valid string.
  @param  ReplaceSpecialChars  Special characters are replaced by their HTML
                               codes (e.g. '>' --> '&gt;')
  @return FALSE if characters < #32 were replaced, TRUE otherwise.
-------------------------------------------------------------------------------}
function ValidXMLText(var AText: ansistring;
  ReplaceSpecialChars: Boolean = true): Boolean;
const
  BOX = #$E2#$8E#$95;
var
  i: Integer;
begin
  Result := true;
  for i := Length(AText) downto 1 do
    if (AText[i] < #32) and not (AText[i] in [#9, #10, #13]) then begin
      // Replace invalid character by box symbol
      Delete(AText, i, 1);
      Insert(BOX, AText, i);
//      AText[i] := '?';
      Result := false;
    end;
  if ReplaceSpecialChars then
    AText := UTF8TextToXMLText(AText);
end;


{******************************************************************************}
{******************************************************************************}
{                   Patch for SysUtils.FormatDateTime                          }
{  Remove when the feature of square brackets in time format masks is in rtl   }
{******************************************************************************}
{******************************************************************************}

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
procedure DateTimeToString(out Result: string; const FormatStr: string; const DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []);
// Copied from "fpc/rtl/objpas/sysutils/datei.inc"
var
  ResultLen: integer;
  ResultBuffer: array[0..255] of char;
  ResultCurrent: pchar;
  (*                    ---- not needed here ---
{$IFDEF MSWindows}
  isEnable_E_Format : Boolean;
  isEnable_G_Format : Boolean;
  eastasiainited : boolean;

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

begin          (*
{$ifdef MSWindows}
  eastasiainited:=false;
{$endif MSWindows}
*)
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

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
procedure DateTimeToString(out Result: string; const FormatStr: string;
  const DateTime: TDateTime; Options : TFormatDateTimeOptions = []);
begin
  DateTimeToString(Result, FormatStr, DateTime, DefaultFormatSettings, Options);
end;

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  Options : TFormatDateTimeOptions = []): string;
begin
  DateTimeToString(Result, FormatStr, DateTime, DefaultFormatSettings,Options);
end;

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []): string;
begin
  DateTimeToString(Result, FormatStr, DateTime, FormatSettings,Options);
end;

{@@ ----------------------------------------------------------------------------
  Extracts compare information from an input string such as "<2.4".
  Is needed for some Excel-strings.

  @param  AString     Input string starting with "<", "<=", ">", ">=", "<>" or "="
                      If this start code is missing a "=" is assumed.
  @param  ACompareOp  Identifier for the comparing operation extracted
                      - see TsCompareOperation
  @return Input string with the comparing characters stripped.
-------------------------------------------------------------------------------}
function AnalyzeComparestr(AString: String; out ACompareOp: TsCompareOperation): String;

  procedure RemoveChars(ACount: Integer; ACompare: TsCompareOperation);
  begin
    ACompareOp := ACompare;
    if ACount = 0 then
      Result := AString
    else
      Result := Copy(AString, 1+ACount, Length(AString));
  end;

begin
  if Length(AString) > 1 then
    case AString[1] of
      '<' : case AString[2] of
              '>' : RemoveChars(2, coNotEqual);
              '=' : RemoveChars(2, coLessEqual);
              else  RemoveChars(1, coLess);
            end;
      '>' : case AString[2] of
              '=' : RemoveChars(2, coGreaterEqual);
              else  RemoveChars(1, coGreater);
            end;
      '=' : RemoveChars(1, coEqual);
      else  RemoveChars(0, coEqual);
    end
  else
    RemoveChars(0, coEqual);
end;

{@@ ----------------------------------------------------------------------------
  Initializes a Sortparams record. This record sets paramaters used when cells
  are sorted.

  @param  ASortByCols     If true sorting occurs along columns, i.e. the
                          ColRowIndex of the sorting keys refer to column indexes.
                          If False, sorting occurs along rows, and the
                          ColRowIndexes refer to row indexes
                          Default: true
  @param  ANumSortKeys    Determines how many columns or rows are used as sorting
                          keys. (Default: 1). Every sort key is initialized for
                          ascending sort direction and case-sensitive comparison.
  @param  ASortPriority   Determines the order or text and numeric data in
                          mixed content type cell ranges.
                          Default: spNumAlpha, i.e. numbers before text (in
                          ascending sort)
  @return The initializaed TsSortParams record
-------------------------------------------------------------------------------}
function InitSortParams(ASortByCols: Boolean = true; ANumSortKeys: Integer = 1;
  ASortPriority: TsSortPriority = spNumAlpha): TsSortParams;
var
  i: Integer;
begin
  Result.SortByCols := ASortByCols;
  Result.Priority := ASortPriority;
  SetLength(Result.Keys, ANumSortKeys);
  for i:=0 to High(Result.Keys) do begin
    Result.Keys[i].ColRowIndex := i;
    Result.Keys[i].Options := [];  // Ascending & case-sensitive
  end;
end;

{@@ ----------------------------------------------------------------------------
  Appends a string to a stream

  @param  AStream   Stream to which the string will be added
  @param  AString   String to be written to the stream
-------------------------------------------------------------------------------}
procedure AppendToStream(AStream: TStream; const AString: string);
begin
  if Length(AString) > 0 then
    AStream.WriteBuffer(AString[1], Length(AString));
end;

{@@ ----------------------------------------------------------------------------
  Appends two strings to a stream

  @param  AStream   Stream to which the strings will be added
  @param  AString1  First string to be written to the stream
  @param  AString2  Second string to be written to the stream
-------------------------------------------------------------------------------}
procedure AppendToStream(AStream: TStream; const AString1, AString2: String);
begin
  AppendToStream(AStream, AString1);
  AppendToStream(AStream, AString2);
end;

{@@ ----------------------------------------------------------------------------
  Appends three strings to a stream

  @param  AStream   Stream to which the strings will be added
  @param  AString1  First string to be written to the stream
  @param  AString2  Second string to be written to the stream
  @param  AString3  Third string to be written to the stream
-------------------------------------------------------------------------------}
procedure AppendToStream(AStream: TStream; const AString1, AString2, AString3: String);
begin
  AppendToStream(AStream, AString1);
  AppendToStream(AStream, AString2);
  AppendToStream(AStream, AString3);
end;


{ Modifying colors }
{ Next function are copies of GraphUtils to avoid a dependence on the Graphics unit. }

const
  HUE_000 = 0;
  HUE_060 = 43;
  HUE_120 = 85;
  HUE_180 = 128;
  HUE_240 = 170;

procedure RGBtoHLS(const R, G, B: Byte; out H, L, S: Byte);
var
  cMax, cMin: Integer;          // max and min RGB values
  Rdelta, Gdelta, Bdelta: Byte; // intermediate value: % of spread from max
  diff: Integer;
begin
  // calculate lightness
  cMax := MaxIntValue([R, G, B]);
  cMin := MinIntValue([R, G, B]);
  L := (integer(cMax) + cMin + 1) div 2;
  diff := cMax - cMin;

  if diff = 0
  then begin
    // r=g=b --> achromatic case
    S := 0;
    H := 0;
  end
  else begin
    // chromatic case
    // saturation
    if L <= 128
    then S := integer(diff * 255) div (cMax + cMin)
    else S := integer(diff * 255) div (510 - cMax - cMin);

    // hue
    Rdelta := (cMax - R);
    Gdelta := (cMax - G);
    Bdelta := (cMax - B);

    if R = cMax
    then H := (HUE_000 + integer(Bdelta - Gdelta) * HUE_060 div diff) and $ff
    else if G = cMax
    then H := HUE_120 + integer(Rdelta - Bdelta) * HUE_060 div diff
    else H := HUE_240 + integer(Gdelta - Rdelta) * HUE_060 div diff;
  end;
end;


procedure HLStoRGB(const H, L, S: Byte; out R, G, B: Byte);

  // utility routine for HLStoRGB
  function HueToRGB(const n1, n2: Byte; Hue: Integer): Byte;
  begin
    if Hue > 255
    then Dec(Hue, 255)
    else if Hue < 0
    then Inc(Hue, 255);

    // return r,g, or b value from this tridrant
    case Hue of
      HUE_000..HUE_060 - 1: Result := n1 + (n2 - n1) * Hue div HUE_060;
      HUE_060..HUE_180 - 1: Result := n2;
      HUE_180..HUE_240 - 1: Result := n1 + (n2 - n1) * (HUE_240 - Hue) div HUE_060;
    else
      Result := n1;
    end;
  end;

var
  n1, n2: Integer;
begin
  if S = 0
  then begin
    // achromatic case
    R := L;
    G := L;
    B := L;
  end
  else begin
    // chromatic case
    // set up magic numbers
    if L < 128
    then begin
      n2 := Integer(L) + Integer(L) * S div 255;
      n1 := 2 * L - n2;
    end
    else begin
      n2 := Integer(S) + L - Integer(L) * S div 255;
      n1 := 2 * L - n2 - 1;
    end;

    // get RGB
    R := HueToRGB(n1, n2, H + HUE_120);
    G := HueToRGB(n1, n2, H);
    B := HueToRGB(n1, n2, H - HUE_120);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Excel defines theme colors and applies a "tint" factor (-1...+1) to darken
  or brighten them.

  This method "tints" a given color with a factor

  The algorithm is described in
  http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.backgroundcolor.aspx

  @param   AColor   rgb color to be modified
  @param   tint     Factor (-1...+1) to be used for the operation
  @return  Modified color
-------------------------------------------------------------------------------}
function TintedColor(AColor: TsColorValue; tint: Double): TsColorValue;
const
  HLSMAX = 255;
var
  r, g, b: byte;
  h, l, s: Byte;
  lum: Double;
begin
  if tint = 0 then begin
    Result := AColor;
    exit;
  end;

  r := TRGBA(AColor).r;
  g := TRGBA(AColor).g;
  b := TRGBA(AColor).b;
  RGBToHLS(r, g, b, h, l, s);

  lum := l;
  if tint < 0 then
    lum := lum * (1.0 + tint)
  else
  if tint > 0 then
    lum := lum * (1.0-tint) + (HLSMAX - HLSMAX * (1.0-tint));
  l := Min(255, round(lum));
  HLSToRGB(h, l, s, r, g, b);

  TRGBA(Result).r := r;
  TRGBA(Result).g := g;
  TRGBA(Result).b := b;
  TRGBA(Result).a := 0;
end;

{@@ ----------------------------------------------------------------------------
  Returns the color index for black or white depending on a color being "bright"
  or "dark".

  @param   AColorValue    rgb color to be analyzed
  @return  The color index for black (scBlack) if AColorValue is a "bright" color,
           or white (scWhite) if AColorValue is a "dark" color.
-------------------------------------------------------------------------------}
function HighContrastColor(AColorValue: TsColorvalue): TsColor;
begin
  if TRGBA(AColorValue).r + TRGBA(AColorValue).g + TRGBA(AColorValue).b < 3*128 then
    Result := scWhite
  else
    Result := scBlack;
end;

{$PUSH}{$HINTS OFF}
{@@ Silence warnings due to an unused parameter }
procedure Unused(const A1);
// code "borrowed" from TAChart
begin
end;

{@@ Silence warnings due to two unused parameters }
procedure Unused(const A1, A2);
// code "borrowed" from TAChart
begin
end;

{@@ Silence warnings due to three unused parameters }
procedure Unused(const A1, A2, A3);
// code adapted from TAChart
begin
end;
{$POP}


{@@ ----------------------------------------------------------------------------
  Creates a FPC format settings record in which all strings are encoded as
  UTF8.
-------------------------------------------------------------------------------}
procedure InitUTF8FormatSettings;
// remove when available in LazUtils
var
  i: Integer;
begin
  UTF8FormatSettings := DefaultFormatSettings;
  UTF8FormatSettings.CurrencyString := AnsiToUTF8(DefaultFormatSettings.CurrencyString);
  for i:=1 to 12 do begin
    UTF8FormatSettings.LongMonthNames[i] := AnsiToUTF8(DefaultFormatSettings.LongMonthNames[i]);
    UTF8FormatSettings.ShortMonthNames[i] := AnsiToUTF8(DefaultFormatSettings.ShortMonthNames[i]);
  end;
  for i:=1 to 7 do begin
    UTF8FormatSettings.LongDayNames[i] := AnsiToUTF8(DefaultFormatSettings.LongDayNames[i]);
    UTF8FormatSettings.ShortDayNames[i] := AnsiToUTF8(DefaultFormatSettings.ShortDayNames[i]);
  end;
end;


{ For debugging only }

{@@ ----------------------------------------------------------------------------
  Write the fonts stored for a given workbook to a file.
-------------------------------------------------------------------------------}
procedure DumpFontsToFile(AWorkbook: TsWorkbook; AFileName: String);
var
  L: TStringList;
  i: Integer;
  fnt: TsFont;
begin
  L := TStringList.Create;
  try
    for i:=0 to AWorkbook.GetFontCount-1 do begin
      fnt := AWorkbook.GetFont(i);
      if fnt = nil then
        L.Add(Format('#%.3d: ---------------', [i]))
      else
        L.Add(Format('#%.3d: %-15s %4.1f %s%s%s%s %s', [
          i,
          fnt.FontName,
          fnt.Size,
          IfThen(fssBold in fnt.Style, 'b', '.'),
          IfThen(fssItalic in fnt.Style, 'i', '.'),
          IfThen(fssUnderline in fnt.Style, 'u', '.'),
          IfThen(fssStrikeOut in fnt.Style, 's', '.'),
          AWorkbook.GetPaletteColorAsHTMLStr(fnt.Color)
        ]));
    end;
    L.SaveToFile(AFileName);
  finally
    L.Free;
  end;
end;


initialization
  InitUTF8FormatSettings;

end.

