{------------------------------------------------------------------------------}
{  Standard built-in formula support                                           }
{------------------------------------------------------------------------------}

unit fpsfunc;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpspreadsheet, fpsExprParser;

var
  ExprFormatSettings: TFormatSettings;

procedure RegisterStdBuiltins(AManager : TsBuiltInExpressionManager);


implementation

uses
  Math, lazutf8, StrUtils, DateUtils, xlsconst, fpsUtils;


{------------------------------------------------------------------------------}
{   Builtin math functions                                                     }
{------------------------------------------------------------------------------}

procedure fpsABS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(abs(ArgToFloat(Args[0])));
end;

procedure fpsACOS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if InRange(x, -1, +1) then
    Result := FloatResult(arccos(x))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

procedure fpsACOSH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if x >= 1 then
    Result := FloatResult(arccosh(ArgToFloat(Args[0])))
  else
    Result := ErrorResult(errOverflow);
end;

procedure fpsASIN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if InRange(x, -1, +1) then
    Result := FloatResult(arcsin(ArgToFloat(Args[0])))
  else
    Result := ErrorResult(errOverflow);
end;

procedure fpsASINH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(arcsinh(ArgToFloat(Args[0])));
end;

procedure fpsATAN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(arctan(ArgToFloat(Args[0])));
end;

procedure fpsATANH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if (x > -1) and (x < +1) then
    Result := FloatResult(arctanh(ArgToFloat(Args[0])))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

procedure fpsCEILING(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CEILING( number, significance )
// returns a number rounded up to a multiple of significance
var
  num, sig: TsExprFloat;
begin
  num := ArgToFloat(Args[0]);
  sig := ArgToFloat(Args[1]);
  if sig = 0 then
    Result := ErrorResult(errDivideByZero)
  else
    Result := FloatResult(ceil(num/sig)*sig);
end;

procedure fpsCOS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(cos(ArgToFloat(Args[0])));
end;

procedure fpsCOSH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(cosh(ArgToFloat(Args[0])));
end;

procedure fpsDEGREES(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(RadToDeg(ArgToFloat(Args[0])));
end;

procedure fpsEVEN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// EVEN( number )
// rounds a number up to the nearest even integer.
// If the number is negative, the number is rounded away from zero.
var
  x: TsExprFloat;
  n: Integer;
begin
  if Args[0].ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtEmpty] then begin
    x := ArgToFloat(Args[0]);
    if x > 0 then
    begin
      n := Trunc(x) + 1;
      if odd(n) then inc(n);
    end else
    if x < 0 then
    begin
      n := Trunc(x) - 1;
      if odd(n) then dec(n);
    end else
      n := 0;
    Result := IntegerResult(n);
  end
  else
    Result := ErrorResult(errWrongType);
end;

procedure fpsEXP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(exp(ArgToFloat(Args[0])));
end;

procedure fpsFACT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// FACT( number )
//  returns the factorial of a number.
var
  res: TsExprFloat;
  i, n: Integer;
begin
  if Args[0].ResultType in [rtInteger, rtFloat, rtEmpty, rtDateTime] then
  begin
    res := 1.0;
    n := ArgToInt(Args[0]);
    if n < 0 then
      Result := ErrorResult(errOverflow)
    else
      try
        for i:=1 to n do
          res := res * i;
        Result := FloatResult(res);
      except on E:Exception do
        Result := ErrorResult(errOverflow);
      end;
  end else
    Result := ErrorResult(errWrongType);
end;

procedure fpsFLOOR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// FLOOR( number, significance )
// returns a number rounded down to a multiple of significance
var
  num, sig: TsExprFloat;
begin
  num := ArgToFloat(Args[0]);
  sig := ArgToFloat(Args[1]);
  if sig = 0 then
    Result := ErrorResult(errDivideByZero)
  else
    Result := FloatResult(floor(num/sig)*sig);
end;

procedure fpsINT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(floor(ArgToFloat(Args[0])));
end;

procedure fpsLN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if x > 0 then
    Result := FloatResult(ln(x))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

procedure fpsLOG(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LOG( number [, base] )  -  base is 10 if omitted.
var
  x: TsExprFloat;
  base: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if x <= 0 then begin
    Result := ErrorResult(errOverflow);  // #NUM!
    exit;
  end;

  if Length(Args) = 2 then begin
    base := ArgToFloat(Args[1]);
    if base < 0 then begin
      Result := ErrorResult(errOverflow);  // #NUM!
      exit;
    end;
  end else
    base := 10;

  Result := FloatResult(logn(base, x));
end;

procedure fpsLOG10(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if x > 0 then
    Result := FloatResult(log10(x))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

procedure fpsMOD(var Result: TsExpressionResult; const Args: TsExprParameterArray);
//  MOD( number, divisor )
// Returns the remainder after a number is divided by a divisor.
var
  n, m: Integer;
begin
  n := ArgToInt(Args[0]);
  m := ArgToInt(Args[1]);
  if m = 0 then
    Result := ErrorResult(errDivideByZero)
  else
    Result := IntegerResult(n mod m);
end;

procedure fpsODD(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ODD( number )
// rounds a number up to the nearest odd integer.
// If the number is negative, the number is rounded away from zero.
var
  x: TsExprFloat;
  n: Integer;
begin
  if Args[0].ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtEmpty] then
  begin
    x := ArgToFloat(Args[0]);
    if x >= 0 then
    begin
      n := Trunc(x) + 1;
      if not odd(n) then inc(n);
    end else
    begin
      n := Trunc(x) - 1;
      if not odd(n) then dec(n);
    end;
    Result := IntegerResult(n);
  end
  else
    Result := ErrorResult(errWrongType);
end;

procedure fpsPI(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Unused(Args);
  Result := FloatResult(pi);
end;

procedure fpsPOWER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  try
    Result := FloatResult(Power(ArgToFloat(Args[0]), ArgToFloat(Args[1])));
  except
    Result := ErrorResult(errOverflow);
  end;
end;

procedure fpsRADIANS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(DegToRad(ArgToFloat(Args[0])));
end;

procedure fpsRAND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Unused(Args);
  Result := FloatResult(random);
end;

procedure fpsROUND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  n: Integer;
begin
  if Args[1].ResultType = rtInteger then
    n := Args[1].ResInteger
  else
    n := round(Args[1].ResFloat);
  Result := FloatResult(RoundTo(ArgToFloat(Args[0]), n));
end;

procedure fpsSIGN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(sign(ArgToFloat(Args[0])));
end;

procedure fpsSIN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(sin(ArgToFloat(Args[0])));
end;

procedure fpsSINH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(sinh(ArgToFloat(Args[0])));
end;

procedure fpsSQRT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if x >= 0 then
    Result := FloatResult(sqrt(x))
  else
    Result := ErrorResult(errOverflow);
end;

procedure fpsTAN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  x := ArgToFloat(Args[0]);
  if frac(x / (pi*0.5)) = 0 then
    Result := ErrorResult(errOverflow)   // #NUM!
  else
    Result := FloatResult(tan(ArgToFloat(Args[0])));
end;

procedure fpsTANH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(tanh(ArgToFloat(Args[0])));
end;


{------------------------------------------------------------------------------}
{   Built-in date/time functions                                               }
{------------------------------------------------------------------------------}

procedure fpsDATE(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
// DATE( year, month, day )
begin
  Result := DateTimeResult(
    EncodeDate(ArgToInt(Args[0]), ArgToInt(Args[1]), ArgToInt(Args[2]))
  );
end;

procedure fpsDATEDIF(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
{ DATEDIF( start_date, end_date, interval )
    start_date <= end_date !
    interval = Y  - The number of complete years.
             = M  - The number of complete months.
             = D  - The number of days.
             = MD - The difference between the days (months and years are ignored).
             = YM - The difference between the months (days and years are ignored).
             = YD - The difference between the days (years and dates are ignored). }
var
  interval: String;
  start_date, end_date: TDate;
begin
  start_date := ArgToDateTime(Args[0]);
  end_date := ArgToDateTime(Args[1]);
  interval := ArgToString(Args[2]);

  if end_date > start_date then
    Result := ErrorResult(errOverflow)
  else if interval = 'Y' then
    Result := FloatResult(YearsBetween(end_date, start_date))
  else if interval = 'M' then
    Result := FloatResult(MonthsBetween(end_date, start_date))
  else if interval = 'D' then
    Result := FloatResult(DaysBetween(end_date, start_date))
  else
    Result := ErrorResult(errFormulaNotSupported);
end;

procedure fpsDATEVALUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the serial number of a date. Input is a string.
// DATE( date_string )
var
  d: TDateTime;
begin
  if TryStrToDate(Args[0].ResString, d) then
    Result := DateTimeResult(d)
  else
    Result := ErrorResult(errWrongType);
end;

procedure fpsDAY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DAY( date_value )
// date_value can be a serial number or a string
var
  y,m,d: Word;
  dt: TDateTime;
begin
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger]) then
    DecodeDate(ArgToFloat(Args[0]), y,m,d)
  else
  if Args[0].ResultType in [rtString] then
  begin
    if TryStrToDate(Args[0].ResString, dt) then
      DecodeDate(dt, y,m,d)
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
  end;
  Result := IntegerResult(d);
end;

procedure fpsHOUR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// HOUR( time_value )
// time_value can be a number or a string.
var
  h, m, s, ms: Word;
  t: double;
begin
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger]) then
    DecodeTime(ArgToFloat(Args[0]), h,m,s,ms)
  else
  if (Args[0].ResultType in [rtString]) then
  begin
    if TryStrToTime(Args[0].ResString, t) then
      DecodeTime(t, h,m,s,ms)
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
  end;
  Result := IntegerResult(h);
end;

procedure fpsMINUTE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MINUTE( serial_number or string )
var
  h, m, s, ms: Word;
  t: double;
begin
  if (Args[0].resultType in [rtDateTime, rtFloat, rtInteger]) then
    DecodeTime(ArgToFloat(Args[0]), h,m,s,ms)
  else
  if (Args[0].ResultType in [rtString]) then
  begin
    if TryStrToTime(Args[0].ResString, t) then
      DecodeTime(t, h,m,s,ms)
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
  end;
  Result := IntegerResult(m);
end;

procedure fpsMONTH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MONTH( date_value or string )
var
  y,m,d: Word;
  dt: TDateTime;
begin
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger]) then
    DecodeDate(ArgToFloat(Args[0]), y,m,d)
  else
  if (Args[0].ResultType in [rtString]) then
  begin
    if TryStrToDate(Args[0].ResString, dt) then
      DecodeDate(dt, y,m,d)
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
  end;
  Result := IntegerResult(m);
end;

procedure fpsNOW(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the current system date and time. Willrefresh the date/time value
// whenever the worksheet recalculates.
//   NOW()
begin
  Unused(Args);
  Result := DateTimeResult(Now);
end;

procedure fpsSECOND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// SECOND( serial_number )
var
  h, m, s, ms: Word;
  t: Double;
begin
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger]) then
    DecodeTime(ArgToFloat(Args[0]), h,m,s,ms)
  else
  if (Args[0].ResultType in [rtString]) then
  begin
    if TryStrToTime(Args[0].ResString, t) then
      DecodeTime(t, h,m,s,ms)
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
  end;
  Result := IntegerResult(s);
end;

procedure fpsTIME(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// TIME( hour, minute, second)
begin
  Result := DateTimeResult(
    EncodeTime(ArgToInt(Args[0]), ArgToInt(Args[1]), ArgToInt(Args[2]), 0)
  );
end;

procedure fpsTIMEVALUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the serial number of a time. Input must be a string.
// DATE( date_string )
var
  t: TDateTime;
begin
  if TryStrToTime(Args[0].ResString, t) then
    Result := DateTimeResult(t)
  else
    Result := ErrorResult(errWrongType);
end;

procedure fpsTODAY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the current system date. This function will refresh the date
// whenever the worksheet recalculates.
//   TODAY()
begin
  Unused(Args);
  Result := DateTimeResult(Date);
end;

procedure fpsWEEKDAY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
{ WEEKDAY( serial_number, [return_value] )
   return_value = 1 - Returns a number from 1 (Sunday) to 7 (Saturday) (default)
                = 2 - Returns a number from 1 (Monday) to 7 (Sunday).
                = 3 - Returns a number from 0 (Monday) to 6 (Sunday). }
var
  n: Integer;
  dow: Integer;
  dt: TDateTime;
begin
  if Length(Args) = 2 then
    n := ArgToInt(Args[1])
  else
    n := 1;
  if Args[0].ResultType in [rtDateTime, rtFloat, rtInteger] then
    dt := ArgToDateTime(Args[0])
  else
  if Args[0].ResultType in [rtString] then
    if not TryStrToDate(Args[0].ResString, dt) then
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
  dow := DayOfWeek(dt);   // Sunday = 1 ... Saturday = 7
  case n of
    1: ;
    2: if dow > 1 then dow := dow - 1 else dow := 7;
    3: if dow > 1 then dow := dow - 2 else dow := 6;
  end;
  Result := IntegerResult(dow);
end;

procedure fpsYEAR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// YEAR( date_value )
var
  y,m,d: Word;
  dt: TDateTime;
begin
  if Args[0].ResultType in [rtDateTime, rtFloat, rtInteger] then
    DecodeDate(ArgToFloat(Args[0]), y,m,d)
  else
  if Args[0].ResultType in [rtString] then
  begin
    if TryStrToDate(Args[0].ResString, dt) then
      DecodeDate(dt, y,m,d)
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
  end;
  Result := IntegerResult(y);
end;


{------------------------------------------------------------------------------}
{    Builtin string functions                                                  }
{------------------------------------------------------------------------------}

procedure fpsCHAR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CHAR( ascii_value )
// returns the character based on the ASCII value
var
  arg: Integer;
begin
  Result := ErrorResult(errWrongType);
  case Args[0].ResultType of
    rtInteger, rtFloat:
      if Args[0].ResultType in [rtInteger, rtFloat] then
      begin
        arg := ArgToInt(Args[0]);
        if (arg >= 0) and (arg < 256) then
          Result := StringResult(AnsiToUTF8(Char(arg)));
      end;
    rtError:
      Result := ErrorResult(Args[0].ResError);
    rtEmpty:
      Result.ResultType := rtEmpty;
  end;
end;

procedure fpsCODE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CODE( text )
// returns the ASCII value of a character or the first character in a string.
var
  s: String;
  ch: Char;
begin
  s := ArgToString(Args[0]);
  if s = '' then
    Result := ErrorResult(errWrongType)
  else
  begin
    ch := UTF8ToAnsi(s)[1];
    Result := IntegerResult(ord(ch));
  end;
end;

procedure fpsCONCATENATE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CONCATENATE( text1, text2, ... text_n )
// Joins two or more strings together
var
  s: String;
  i: Integer;
begin
  s := '';
  for i:=0 to Length(Args)-1 do
  begin
    if Args[i].ResultType = rtError then
    begin
      Result := ErrorResult(Args[i].ResError);
      exit;
    end;
    s := s + ArgToString(Args[i]);
  end;
  Result := StringResult(s);
end;

procedure fpsEXACT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// EXACT( text1, text2 )
// Compares two strings (case-sensitive) and returns TRUE if they are equal
var
  s1, s2: String;
begin
  s1 := ArgToString(Args[0]);
  s2 := ArgToString(Args[1]);
  Result := BooleanResult(s1 = s2);
end;

procedure fpsLEFT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LEFT( text, [number_of_characters] )
// extracts a substring from a string, starting from the left-most character
var
  s: String;
  count: Integer;
begin
  s := ArgToString(Args[0]);
  if s = '' then
    Result.ResultType := rtEmpty
  else
  begin
    if Length(Args) = 1 then
      count := 1
    else
    if Args[1].ResultType in [rtInteger, rtFloat] then
      count := ArgToInt(Args[1])
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    Result := StringResult(UTF8LeftStr(s, count));
  end;
end;

procedure fpsLEN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LEN( text )
//  returns the length of the specified string.
begin
  Result := IntegerResult(UTF8Length(ArgToString(Args[0])));
end;

procedure fpsLOWER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LOWER( text )
// converts all letters in the specified string to lowercase. If there are
// characters in the string that are not letters, they are not affected.
begin
  Result := StringResult(UTF8Lowercase(ArgToString(Args[0])));
end;

procedure fpsMID(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MID( text, start_position, number_of_characters )
// extracts a substring from a string (starting at any position).
begin
  Result := StringResult(UTF8Copy(ArgToString(Args[0]), ArgToInt(Args[1]), ArgToInt(Args[2])));
end;

procedure fpsREPLACE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// REPLACE( old_text, start, number_of_chars, new_text )
//  replaces a sequence of characters in a string with another set of characters
var
  sOld, sNew, s1, s2: String;
  start: Integer;
  count: Integer;
begin
  sOld := Args[0].ResString;
  start := ArgToInt(Args[1]);
  count := ArgToInt(Args[2]);
  sNew := Args[3].ResString;
  s1 := UTF8Copy(sOld, 1, start-1);
  s2 := UTF8Copy(sOld, start+count, UTF8Length(sOld));
  Result := StringResult(s1 + sNew + s2);
end;

procedure fpsREPT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// REPT( text, count )
// repeats text a specified number of times.
var
  s: String;
  count: Integer;
begin
  s := ArgToString(Args[0]);
  if s = '' then
    Result.ResultType := rtEmpty
  else
  if Args[1].ResultType in [rtInteger, rtFloat] then begin
    count := ArgToInt(Args[1]);
    Result := StringResult(DupeString(s, count));
  end;
end;

procedure fpsRIGHT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// RIGHT( text, [number_of_characters] )
// extracts a substring from a string, starting from the last character
var
  s: String;
  count: Integer;
begin
  s := ArgToString(Args[0]);
  if s = '' then
    Result.ResultType := rtEmpty
  else begin
    if Length(Args) = 1 then
      count := 1
    else
    if Args[1].ResultType in [rtInteger, rtFloat] then
      count := ArgToInt(Args[1])
    else
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    Result := StringResult(UTF8RightStr(s, count));
  end;
end;

procedure fpsSUBSTITUTE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// SUBSTITUTE( text, old_text, new_text, [nth_appearance] )
// replaces a set of characters with another.
var
  sOld: String;
  sNew: String;
  s1, s2: String;
  n: Integer;
  s: String;
  p: Integer;
begin
  s := ArgToString(Args[0]);
  sOld := ArgToString(Args[1]);
  sNew := ArgToString(Args[2]);
  if Length(Args) = 4 then
  begin
    n := ArgToInt(Args[3]);               // THIS PART NOT YET CHECKED !!!!!!
    if n <= 0 then
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    p := UTF8Pos(sOld, s);
    while (n > 1) do begin
      p := UTF8Pos(sOld, s, p+1);
      dec(n);
    end;
    if p > 0 then begin
      s1 := UTF8Copy(s, 1, p-1);
      s2 := UTF8Copy(s, p+UTF8Length(sOld), UTF8Length(s));
      s := s1 + sNew + s2;
    end;
    Result := StringResult(s);
  end else
    Result := StringResult(UTF8StringReplace(s, sOld, sNew, [rfReplaceAll]));
end;

procedure fpsTRIM(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// TRIM( text )
// returns a text value with the leading and trailing spaces removed
begin
  Result := StringResult(UTF8Trim(ArgToString(Args[0])));
end;

procedure fpsUPPER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// UPPER( text )
// converts all letters in the specified string to uppercase. If there are
// characters in the string that are not letters, they are not affected.
begin
  Result := StringResult(UTF8Uppercase(ArgToString(Args[0])));
end;

procedure fpsVALUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// VALUE( text )
// converts a text value that represents a number to a number.
var
  x: Double;
  n: Integer;
  s: String;
begin
  s := ArgToString(Args[0]);
  if TryStrToInt(s, n) then
    Result := IntegerResult(n)
  else
  if TryStrToFloat(s, x, ExprFormatSettings) then
    Result := FloatResult(x)
  else
    Result := ErrorResult(errWrongType);
end;


{------------------------------------------------------------------------------}
{    Built-in logical functions                                                }
{------------------------------------------------------------------------------}

procedure fpsAND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// AND( condition1, [condition2], ... )
// up to 30 parameters. At least 1 parameter.
var
  i: Integer;
  b: Boolean;
begin
  b := true;
  for i:=0 to High(Args) do
    b := b and Args[i].ResBoolean;
  Result.ResBoolean := b;
end;

procedure fpsFALSE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// FALSE ()
begin
  Unused(Args);
  Result.ResBoolean := false;
end;

procedure fpsIF(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// IF( condition, value_if_true, [value_if_false] )
begin
  if Length(Args) > 2 then
  begin
    if Args[0].ResBoolean then
      Result := Args[1]
    else
      Result := Args[2];
  end else
  begin
    if Args[0].ResBoolean then
      Result := Args[1]
    else
      Result.ResBoolean := false;
  end;
end;

procedure fpsNOT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// NOT( condition )
begin
  Result.ResBoolean := not Args[0].ResBoolean;
end;

procedure fpsOR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// OR( condition1, [condition2], ... )
// up to 30 parameters. At least 1 parameter.
var
  i: Integer;
  b: Boolean;
begin
  b := false;
  for i:=0 to High(Args) do
    b := b or Args[i].ResBoolean;
  Result.ResBoolean := b;
end;

procedure fpsTRUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// TRUE()
begin
  Unused(Args);
  Result.ResBoolean := true;
end;


{------------------------------------------------------------------------------}
{    Built-in statistical functions                                            }
{------------------------------------------------------------------------------}

procedure fpsAVEDEV(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Average value of absolute deviations of data from their mean.
// AVEDEV( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
  m: TsExprFloat;
  i: Integer;
begin
  ArgsToFloatArray(Args, data);
  m := Mean(data);
  for i:=0 to High(data) do      // replace data by their average deviation from the mean
    data[i] := abs(data[i] - m);
  Result.ResFloat := Mean(data);
end;

procedure fpsAVERAGE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// AVERAGE( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  Result.ResFloat := Mean(data);
end;

procedure fpsCOUNT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
{ counts the number of cells that contain numbers as well as the number of
  arguments that contain numbers.
    COUNT( value1, [value2, ... value_n] )  }
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  Result.ResInteger := Length(data);
end;

procedure fpsCOUNTA(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Counts the number of cells that are not empty as well as the number of
// arguments that contain values
//   COUNTA( value1, [value2, ... value_n] )
var
  i, n: Integer;
  r, c: Cardinal;
  cell: PCell;
  arg: TsExpressionResult;
begin
  n := 0;
  for i:=0 to High(Args) do
  begin
    arg := Args[i];
    case arg.ResultType of
      rtInteger, rtFloat, rtDateTime, rtBoolean:
        inc(n);
      rtString:
        if arg.ResString <> '' then inc(n);
      rtError:
        if arg.ResError <> errOK then inc(n);
      rtCell:
        begin
          cell := ArgToCell(arg);
          if cell <> nil then
            case cell^.ContentType of
              cctNumber, cctDateTime, cctBool: inc(n);
              cctUTF8String: if cell^.UTF8StringValue <> '' then inc(n);
              cctError: if cell^.ErrorValue <> errOK then inc(n);
            end;
        end;
      rtCellRange:
        for r := arg.ResCellRange.Row1 to arg.ResCellRange.Row2 do
          for c := arg.ResCellRange.Col1 to arg.ResCellRange.Col2 do
          begin
            cell := arg.Worksheet.FindCell(r, c);
            if (cell <> nil) then
              case cell^.ContentType of
                cctNumber, cctDateTime, cctBool : inc(n);
                cctUTF8String: if cell^.UTF8StringValue <> '' then inc(n);
                cctError: if cell^.ErrorValue <> errOK then inc(n);
              end;
          end;
    end;
  end;
  Result.ResInteger := n;
end;

procedure fpsCOUNTBLANK(var Result: TsExpressionResult; const Args: TsExprParameterArray);
{ Counts the number of empty cells in a range.
    COUNTBLANK( range )
  "range" is the range of cells to count empty cells. }
var
  n: Integer;
  r, c: Cardinal;
  cell: PCell;
begin
  n := 0;
  case Args[0].ResultType of
    rtEmpty:
      inc(n);
    rtCell:
      begin
        cell := ArgToCell(Args[0]);
        if cell = nil then
          inc(n)
        else
          case cell^.ContentType of
            cctNumber, cctDateTime, cctBool: ;
            cctUTF8String: if cell^.UTF8StringValue = '' then inc(n);
            cctError: if cell^.ErrorValue = errOK then inc(n);
          end;
      end;
    rtCellRange:
      for r := Args[0].ResCellRange.Row1 to Args[0].ResCellRange.Row2 do
        for c := Args[0].ResCellRange.Col1 to Args[0].ResCellRange.Col2 do begin
          cell := Args[0].Worksheet.FindCell(r, c);
          if cell = nil then
            inc(n)
          else
            case cell^.ContentType of
              cctNumber, cctDateTime, cctBool: ;
              cctUTF8String: if cell^.UTF8StringValue = '' then inc(n);
              cctError: if cell^.ErrorValue = errOK then inc(n);
            end;
        end;
  end;
  Result.ResInteger := n;
end;

procedure fpsMAX(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MAX( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  Result.ResFloat := MaxValue(data);
end;

procedure fpsMIN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MIN( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  Result.ResFloat := MinValue(data);
end;

procedure fpsPRODUCT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// PRODUCT( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
  i: Integer;
  p: TsExprFloat;
begin
  ArgsToFloatArray(Args, data);
  p := 1.0;
  for i := 0 to High(data) do
    p := p * data[i];
  Result.ResFloat := p;
end;

procedure fpsSTDEV(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the standard deviation of a population based on a sample of numbers
// of numbers.
//   STDEV( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  if Length(data) > 1 then
    Result.ResFloat := StdDev(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;

procedure fpsSTDEVP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the standard deviation of a population based on an entire population
// STDEVP( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  if Length(data) > 0 then
    Result.ResFloat := PopnStdDev(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;

procedure fpsSUM(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// SUM( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  Result.ResFloat := Sum(data);
end;

procedure fpsSUMSQ(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the sum of the squares of a series of values.
// SUMSQ( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  Result.ResFloat := SumOfSquares(data);
end;

procedure fpsVAR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the variance of a population based on a sample of numbers.
// VAR( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  if Length(data) > 1 then
    Result.ResFloat := Variance(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;

procedure fpsVARP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the variance of a population based on an entire population of numbers.
// VARP( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray;
begin
  ArgsToFloatArray(Args, data);
  if Length(data) > 0 then
    Result.ResFloat := PopnVariance(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;


{------------------------------------------------------------------------------}
{    Builtin info functions                                                    }
{------------------------------------------------------------------------------}

{ !!!!!!!!!!!!!!  not working !!!!!!!!!!!!!!!!!!!!!! }
{ !!!!!!!!!!!!!! needs localized strings !!!!!!!!!!! }

procedure fpsCELL(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CELL( type, [range] )

{ from http://www.techonthenet.com/excel/formulas/cell.php:

  "type" is the type of information that we retrieve for the cell and can have
  one of the following values:
    Value         Explanation
    ------------- --------------------------------------------------------------
    "address"     Address of the cell. If the cell refers to a range, it is the
                  first cell in the range.
    "col"         Column number of the cell.
    "color"       Returns 1 if the color is a negative value; Otherwise it returns 0.
    "contents"    Contents of the upper-left cell.
    "filename"    Filename of the file that contains reference.
    "format"      Number format of the cell according to:
                     "G"    General
                     "F0"   0
                     ",0"   #,##0
                     "F2"   0.00
                     ",2"   #,##0.00
                     "C0"   $#,##0_);($#,##0)
                     "C0-"  $#,##0_);[Red]($#,##0)
                     "C2"   $#,##0.00_);($#,##0.00)
                     "C2-"  $#,##0.00_);[Red]($#,##0.00)
                     "P0"   0%
                     "P2"   0.00%
                     "S2"   0.00E+00
                     "G"    # ?/? or # ??/??
                     "D4"   m/d/yy or m/d/yy h:mm or mm/dd/yy
                     "D1"   d-mmm-yy or dd-mmm-yy
                     "D2"   d-mmm or dd-mmm
                     "D3"   mmm-yy
                     "D5"   mm/dd
                     "D6"   h:mm:ss AM/PM
                     "D7"   h:mm AM/PM
                     "D8"   h:mm:ss
                     "D9"   h:mm
    "parentheses" Returns 1 if the cell is formatted with parentheses;
                  Otherwise, it returns 0.
    "prefix"      Label prefix for the cell.
                  - Returns a single quote (') if the cell is left-aligned.
                  - Returns a double quote (") if the cell is right-aligned.
                  - Returns a caret (^) if the cell is center-aligned.
                  - Returns a back slash (\) if the cell is fill-aligned.
                  - Returns an empty text value for all others.
    "protect"     Returns 1 if the cell is locked. Returns 0 if the cell is not locked.
    "row"         Row number of the cell.
    "type"        Returns "b" if the cell is empty.
                  Returns "l" if the cell contains a text constant.
                  Returns "v" for all others.
    "width"       Column width of the cell, rounded to the nearest integer.

  !!!! NOT ALL OF THEM ARE SUPPORTED HERE !!!

  "range" is optional in Excel. It is the cell (or range) that you wish to retrieve
  information for. If the range parameter is omitted, the CELL function will
  assume that you are retrieving information for the last cell that was changed.

  "range" is NOT OPTIONAL here because we don't know the last cell changed !!!
}
var
  stype: String;
  r1, c1: Cardinal;
  cell: PCell;
begin
  if Length(Args)=1 then
  begin
    // This case is not supported by us, but it is by Excel.
    // Therefore the error is not quite correct...
    Result := ErrorResult(errIllegalRef);
    exit;
  end;

  stype := lowercase(ArgToString(Args[0]));

  case Args[1].ResultType of
    rtCell:
      begin
        cell := ArgToCell(Args[1]);
        r1 := Args[1].ResRow;
        c1 := Args[1].ResCol;
      end;
    rtCellRange:
      begin
        r1 := Args[1].ResCellRange.Row1;
        c1 := Args[1].ResCellRange.Col1;
        cell := Args[1].Worksheet.FindCell(r1, c1);
      end;
    else
      Result := ErrorResult(errWrongType);
      exit;
  end;

  if stype = 'address' then
    Result := StringResult(GetCellString(r1, c1, []))
  else
  if stype = 'col' then
    Result := IntegerResult(c1+1)
  else
  if stype = 'color' then
  begin
    if (cell <> nil) and (cell^.NumberFormat = nfCurrencyRed) then
      Result := IntegerResult(1)
    else
      Result := IntegerResult(0);
  end else
  if stype = 'contents' then
  begin
    if cell = nil then
      Result := IntegerResult(0)
    else
      case cell^.ContentType of
        cctNumber     : if frac(cell^.NumberValue) = 0 then
                          Result := IntegerResult(trunc(cell^.NumberValue))
                        else
                          Result := FloatResult(cell^.NumberValue);
        cctDateTime   : Result := DateTimeResult(cell^.DateTimeValue);
        cctUTF8String : Result := StringResult(cell^.UTF8StringValue);
        cctBool       : Result := BooleanResult(cell^.BoolValue);
        cctError      : Result := ErrorResult(cell^.ErrorValue);
      end;
  end else
  if stype = 'filename' then
    Result := Stringresult(
      ExtractFilePath(Args[1].Worksheet.Workbook.FileName) + '[' +
      ExtractFileName(Args[1].Worksheet.Workbook.FileName) + ']' +
      Args[1].Worksheet.Name
    )
  else
  if stype = 'format' then begin
    Result := StringResult('G');
    if cell <> nil then
      case cell^.NumberFormat of
        nfGeneral:
          Result := StringResult('G');
        nfFixed:
          if cell^.NumberFormatStr= '0' then Result := StringResult('0') else
          if cell^.NumberFormatStr = '0.00' then  Result := StringResult('F0');
        nfFixedTh:
          if cell^.NumberFormatStr = '#,##0' then Result := StringResult(',0') else
          if cell^.NumberFormatStr = '#,##0.00' then Result := StringResult(',2');
        nfPercentage:
          if cell^.NumberFormatStr = '0%' then Result := StringResult('P0') else
          if cell^.NumberFormatStr = '0.00%' then Result := StringResult('P2');
        nfExp:
          if cell^.NumberFormatStr = '0.00E+00' then Result := StringResult('S2');
        nfShortDate, nfLongDate, nfShortDateTime:
          Result := StringResult('D4');
        nfLongTimeAM:
          Result := StringResult('D6');
        nfShortTimeAM:
          Result := StringResult('D7');
        nfLongTime:
          Result := StringResult('D8');
        nfShortTime:
          Result := StringResult('D9');
      end;
  end else
  if stype = 'prefix' then
  begin
    Result := StringResult('');
    if (cell^.ContentType = cctUTF8String) then
      case cell^.HorAlignment of
        haLeft  : Result := StringResult('''');
        haCenter: Result := StringResult('^');
        haRight : Result := StringResult('"');
      end;
  end else
  if stype = 'row' then
    Result := IntegerResult(r1+1)
  else
  if stype = 'type' then begin
    if (cell = nil) or (cell^.ContentType = cctEmpty) then
      Result := StringResult('b')
    else if cell^.ContentType = cctUTF8String then begin
      if (cell^.UTF8StringValue = '')
        then Result := StringResult('b')
        else Result := StringResult('l');
    end else
      Result := StringResult('v');
  end else
  if stype = 'width' then
    Result := FloatResult(Args[1].Worksheet.GetColWidth(c1))
  else
    Result := ErrorResult(errWrongType);
end;

procedure fpsISBLANK(var Result: TsExpressionResult; const Args: TsExprParameterArray);
//  ISBLANK( value )
// Checks for blank or null values.
// "value" is the value that you want to test.
// If "value" is blank, this function will return TRUE.
// If "value" is not blank, the function will return FALSE.
var
  cell: PCell;
begin
  case Args[0].ResultType of
    rtEmpty : Result := BooleanResult(true);
    rtString: Result := BooleanResult(Result.ResString = '');
    rtCell  : begin
                cell := ArgToCell(Args[0]);
                if (cell = nil) or (cell^.ContentType = cctEmpty) then
                  Result := BooleanResult(true)
                else
                  Result := BooleanResult(false);
              end;
  end;
end;

procedure fpsISERR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISERR( value )
// If "value" is an error value (except #N/A), this function will return TRUE.
// Otherwise, it will return FALSE.
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctError) and (cell^.ErrorValue <> errArgError)
      then Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtError) and (Args[0].ResError <> errArgError) then
    Result := BooleanResult(true);
end;

procedure fpsISERROR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISERROR( value )
// If "value" is an error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?
// or #NULL), this function will return TRUE. Otherwise, it will return FALSE.
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctError) and (cell^.ErrorValue <= errArgError)
      then Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtError) then
    Result := BooleanResult(true);
end;

procedure fpsISLOGICAL(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISLOGICAL( value )
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctBool) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtBoolean) then
    Result := BooleanResult(true);
end;

procedure fpsISNA(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISNA( value )
// If "value" is a #N/A error value , this function will return TRUE.
// Otherwise, it will return FALSE.
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctError) and (cell^.ErrorValue = errArgError)
      then Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtError) and (Args[0].ResError = errArgError) then
    Result := BooleanResult(true);
end;

procedure fpsISNONTEXT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISNONTEXT( value )
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell = nil) or ((cell <> nil) and (cell^.ContentType <> cctUTF8String)) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType <> rtString) then
    Result := BooleanResult(true);
end;

procedure fpsISNUMBER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISNUMBER( value )
// Tests "value" for a number (or date/time - checked with Excel).
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType in [cctNumber, cctDateTime]) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType in [rtFloat, rtInteger, rtDateTime]) then
    Result := BooleanResult(true);
end;

procedure fpsISREF(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISREF( value )
begin
  Result := BooleanResult(Args[0].ResultType in [rtCell, rtCellRange]);
end;

procedure fpsISTEXT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ISTEXT( value )
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctUTF8String) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtString) then
    Result := BooleanResult(true);
end;


{------------------------------------------------------------------------------}
{   Registration                                                               }
{------------------------------------------------------------------------------}

{@@ Registers the standard built-in functions. Called automatically. }
procedure RegisterStdBuiltins(AManager : TsBuiltInExpressionManager);
var
  cat: TsBuiltInExprCategory;
begin
  with AManager do
  begin
    // Math functions
    cat := bcMath;
    AddFunction(cat, 'ABS',       'F', 'F',    INT_EXCEL_SHEET_FUNC_ABS,        @fpsABS);
    AddFunction(cat, 'ACOS',      'F', 'F',    INT_EXCEL_SHEET_FUNC_ACOS,       @fpsACOS);
    AddFunction(cat, 'ACOSH',     'F', 'F',    INT_EXCEL_SHEET_FUNC_ACOSH,      @fpsACOSH);
    AddFunction(cat, 'ASIN',      'F', 'F',    INT_EXCEL_SHEET_FUNC_ASIN,       @fpsASIN);
    AddFunction(cat, 'ASINH',     'F', 'F',    INT_EXCEL_SHEET_FUNC_ASINH,      @fpsASINH);
    AddFunction(cat, 'ATAN',      'F', 'F',    INT_EXCEL_SHEET_FUNC_ATAN,       @fpsATAN);
    AddFunction(cat, 'ATANH',     'F', 'F',    INT_EXCEL_SHEET_FUNC_ATANH,      @fpsATANH);
    AddFunction(cat, 'CEILING',   'F', 'FF',   INT_EXCEL_SHEET_FUNC_CEILING,    @fpsCEILING);
    AddFunction(cat, 'COS',       'F', 'F',    INT_EXCEL_SHEET_FUNC_COS,        @fpsCOS);
    AddFunction(cat, 'COSH',      'F', 'F',    INT_EXCEL_SHEET_FUNC_COSH,       @fpsCOSH);
    AddFunction(cat, 'DEGREES',   'F', 'F',    INT_EXCEL_SHEET_FUNC_DEGREES,    @fpsDEGREES);
    AddFunction(cat, 'EVEN',      'I', 'F',    INT_EXCEL_SHEET_FUNC_EVEN,       @fpsEVEN);
    AddFunction(cat, 'EXP',       'F', 'F',    INT_EXCEL_SHEET_FUNC_EXP,        @fpsEXP);
    AddFunction(cat, 'FACT',      'F', 'I',    INT_EXCEL_SHEET_FUNC_FACT,       @fpsFACT);
    AddFunction(cat, 'FLOOR',     'F', 'FF',   INT_EXCEL_SHEET_FUNC_FLOOR,      @fpsFLOOR);
    AddFunction(cat, 'INT',       'I', 'F',    INT_EXCEL_SHEET_FUNC_INT,        @fpsINT);
    AddFunction(cat, 'LN',        'F', 'F',    INT_EXCEL_SHEET_FUNC_LN,         @fpsLN);
    AddFunction(cat, 'LOG',       'F', 'Ff',   INT_EXCEL_SHEET_FUNC_LOG,        @fpsLOG);
    AddFunction(cat, 'LOG10',     'F', 'F',    INT_EXCEL_SHEET_FUNC_LOG10,      @fpsLOG10);
    AddFunction(cat, 'MOD',       'I', 'II',   INT_EXCEL_SHEET_FUNC_MOD,        @fpsMOD);
    AddFunction(cat, 'ODD',       'I', 'F',    INT_EXCEL_SHEET_FUNC_ODD,        @fpsODD);
    AddFunction(cat, 'PI',        'F', '',     INT_EXCEL_SHEET_FUNC_PI,         @fpsPI);
    AddFunction(cat, 'POWER',     'F', 'FF',   INT_EXCEL_SHEET_FUNC_POWER,      @fpsPOWER);
    AddFunction(cat, 'RADIANS',   'F', 'F',    INT_EXCEL_SHEET_FUNC_RADIANS,    @fpsRADIANS);
    AddFunction(cat, 'RAND',      'F', '',     INT_EXCEL_SHEET_FUNC_RAND,       @fpsRAND);
    AddFunction(cat, 'ROUND',     'F', 'FF',   INT_EXCEL_SHEET_FUNC_ROUND,      @fpsROUND);
    AddFunction(cat, 'SIGN',      'F', 'F',    INT_EXCEL_SHEET_FUNC_SIGN,       @fpsSIGN);
    AddFunction(cat, 'SIN',       'F', 'F',    INT_EXCEL_SHEET_FUNC_SIN,        @fpsSIN);
    AddFunction(cat, 'SINH',      'F', 'F',    INT_EXCEL_SHEET_FUNC_SINH,       @fpsSINH);
    AddFunction(cat, 'SQRT',      'F', 'F',    INT_EXCEL_SHEET_FUNC_SQRT,       @fpsSQRT);
    AddFunction(cat, 'TAN',       'F', 'F',    INT_EXCEL_SHEET_FUNC_TAN,        @fpsTAN);
    AddFunction(cat, 'TANH',      'F', 'F',    INT_EXCEL_SHEET_FUNC_TANH,       @fpsTANH);

    // Date/time
    cat := bcDateTime;
    AddFunction(cat, 'DATE',      'D', 'III',  INT_EXCEL_SHEET_FUNC_DATE,       @fpsDATE);
    AddFunction(cat, 'DATEDIF',   'F', 'DDS',  INT_EXCEL_SHEET_FUNC_DATEDIF,    @fpsDATEDIF);
    AddFunction(cat, 'DATEVALUE', 'D', 'S',    INT_EXCEL_SHEET_FUNC_DATEVALUE,  @fpsDATEVALUE);
    AddFunction(cat, 'DAY',       'I', '?',    INT_EXCEL_SHEET_FUNC_DAY,        @fpsDAY);
    AddFunction(cat, 'HOUR',      'I', '?',    INT_EXCEL_SHEET_FUNC_HOUR,       @fpsHOUR);
    AddFunction(cat, 'MINUTE',    'I', '?',    INT_EXCEL_SHEET_FUNC_MINUTE,     @fpsMINUTE);
    AddFunction(cat, 'MONTH',     'I', '?',    INT_EXCEL_SHEET_FUNC_MONTH,      @fpsMONTH);
    AddFunction(cat, 'NOW',       'D', '',     INT_EXCEL_SHEET_FUNC_NOW,        @fpsNOW);
    AddFunction(cat, 'SECOND',    'I', '?',    INT_EXCEL_SHEET_FUNC_SECOND,     @fpsSECOND);
    AddFunction(cat, 'TIME' ,     'D', 'III',  INT_EXCEL_SHEET_FUNC_TIME,       @fpsTIME);
    AddFunction(cat, 'TIMEVALUE', 'D', 'S',    INT_EXCEL_SHEET_FUNC_TIMEVALUE,  @fpsTIMEVALUE);
    AddFunction(cat, 'TODAY',     'D', '',     INT_EXCEL_SHEET_FUNC_TODAY,      @fpsTODAY);
    AddFunction(cat, 'WEEKDAY',   'I', '?i',   INT_EXCEL_SHEET_FUNC_WEEKDAY,    @fpsWEEKDAY);
    AddFunction(cat, 'YEAR',      'I', '?',    INT_EXCEL_SHEET_FUNC_YEAR,       @fpsYEAR);

    // Strings
    cat := bcStrings;
    AddFunction(cat, 'CHAR',      'S', 'I',    INT_EXCEL_SHEET_FUNC_CHAR,       @fpsCHAR);
    AddFunction(cat, 'CODE',      'I', 'S',    INT_EXCEL_SHEET_FUNC_CODE,       @fpsCODE);
    AddFunction(cat, 'CONCATENATE','S','S+',   INT_EXCEL_SHEET_FUNC_CONCATENATE,@fpsCONCATENATE);
    AddFunction(cat, 'EXACT',     'B', 'SS',   INT_EXCEL_SHEET_FUNC_EXACT,      @fpsEXACT);
    AddFunction(cat, 'LEFT',      'S', 'Si',   INT_EXCEL_SHEET_FUNC_LEFT,       @fpsLEFT);
    AddFunction(cat, 'LEN',       'I', 'S',    INT_EXCEL_SHEET_FUNC_LEN,        @fpsLEN);
    AddFunction(cat, 'LOWER',     'S', 'S',    INT_EXCEL_SHEET_FUNC_LOWER,      @fpsLOWER);
    AddFunction(cat, 'MID',       'S', 'SII',  INT_EXCEL_SHEET_FUNC_MID,        @fpsMID);
    AddFunction(cat, 'REPLACE',   'S', 'SIIS', INT_EXCEL_SHEET_FUNC_REPLACE,    @fpsREPLACE);
    AddFunction(cat, 'REPT',      'S', 'SI',   INT_EXCEL_SHEET_FUNC_REPT,       @fpsREPT);
    AddFunction(cat, 'RIGHT',     'S', 'Si',   INT_EXCEL_SHEET_FUNC_RIGHT,      @fpsRIGHT);
    AddFunction(cat, 'SUBSTITUTE','S', 'SSSi', INT_EXCEL_SHEET_FUNC_SUBSTITUTE, @fpsSUBSTITUTE);
    AddFunction(cat, 'TRIM',      'S', 'S',    INT_EXCEL_SHEET_FUNC_TRIM,       @fpsTRIM);
    AddFunction(cat, 'UPPER',     'S', 'S',    INT_EXCEL_SHEET_FUNC_UPPER,      @fpsUPPER);
    AddFunction(cat, 'VALUE',     'F', 'S',    INT_EXCEL_SHEET_FUNC_VALUE,      @fpsVALUE);

    // Logical
    cat := bcLogical;
    AddFunction(cat, 'AND',       'B', 'B+',   INT_EXCEL_SHEET_FUNC_AND,        @fpsAND);
    AddFunction(cat, 'FALSE',     'B', '',     INT_EXCEL_SHEET_FUNC_FALSE,      @fpsFALSE);
    AddFunction(cat, 'IF',        'B', 'B?+',  INT_EXCEL_SHEET_FUNC_IF,         @fpsIF);
    AddFunction(cat, 'NOT',       'B', 'B',    INT_EXCEL_SHEET_FUNC_NOT,        @fpsNOT);
    AddFunction(cat, 'OR',        'B', 'B+',   INT_EXCEL_SHEET_FUNC_OR,         @fpsOR);
    AddFunction(cat, 'TRUE',      'B', '',     INT_EXCEL_SHEET_FUNC_TRUE ,      @fpsTRUE);

    // Statistical
    cat := bcStatistics;
    AddFunction(cat, 'AVEDEV',    'F', 'F+',   INT_EXCEL_SHEET_FUNC_AVEDEV,     @fpsAVEDEV);
    AddFunction(cat, 'AVERAGE',   'F', 'F+',   INT_EXCEL_SHEET_FUNC_AVERAGE,    @fpsAVERAGE);
    AddFunction(cat, 'COUNT',     'I', '?+',   INT_EXCEL_SHEET_FUNC_COUNT,      @fpsCOUNT);
    AddFunction(cat, 'COUNTA',    'I', '?+',   INT_EXCEL_SHEET_FUNC_COUNTA,     @fpsCOUNTA);
    AddFunction(cat, 'COUNTBLANK','I', 'R',    INT_EXCEL_SHEET_FUNC_COUNTBLANK, @fpsCOUNTBLANK);
    AddFunction(cat, 'MAX',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_MAX,        @fpsMAX);
    AddFunction(cat, 'MIN',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_MIN,        @fpsMIN);
    AddFunction(cat, 'PRODUCT',   'F', 'F+',   INT_EXCEL_SHEET_FUNC_PRODUCT,    @fpsPRODUCT);
    AddFunction(cat, 'STDEV',     'F', 'F+',   INT_EXCEL_SHEET_FUNC_STDEV,      @fpsSTDEV);
    AddFunction(cat, 'STDEVP',    'F', 'F+',   INT_EXCEL_SHEET_FUNC_STDEVP,     @fpsSTDEVP);
    AddFunction(cat, 'SUM',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_SUM,        @fpsSUM);
    AddFunction(cat, 'SUMSQ',     'F', 'F+',   INT_EXCEL_SHEET_FUNC_SUMSQ,      @fpsSUMSQ);
    AddFunction(cat, 'VAR',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_VAR,        @fpsVAR);
    AddFunction(cat, 'VARP',      'F', 'F+',   INT_EXCEL_SHEET_FUNC_VARP,       @fpsVARP);
    // to do: CountIF, SUMIF

    // Info functions
    cat := bcInfo;
    //AddFunction(cat, 'CELL',      '?', 'Sr',   INT_EXCEL_SHEET_FUNC_CELL,       @fpsCELL);
    AddFunction(cat, 'ISBLANK',   'B', '?',    INT_EXCEL_SHEET_FUNC_ISBLANK,    @fpsISBLANK);
    AddFunction(cat, 'ISERR',     'B', '?',    INT_EXCEL_SHEET_FUNC_ISERR,      @fpsISERR);
    AddFunction(cat, 'ISERROR',   'B', '?',    INT_EXCEL_SHEET_FUNC_ISERROR,    @fpsISERROR);
    AddFunction(cat, 'ISLOGICAL', 'B', '?',    INT_EXCEL_SHEET_FUNC_ISLOGICAL,  @fpsISLOGICAL);
    AddFunction(cat, 'ISNA',      'B', '?',    INT_EXCEL_SHEET_FUNC_ISNA,       @fpsISNA);
    AddFunction(cat, 'ISNONTEXT', 'B', '?',    INT_EXCEL_SHEET_FUNC_ISNONTEXT,  @fpsISNONTEXT);
    AddFunction(cat, 'ISNUMBER',  'B', '?',    INT_EXCEL_SHEET_FUNC_ISNUMBER,   @fpsISNUMBER);
    AddFunction(cat, 'ISREF',     'B', '?',    INT_EXCEL_SHEET_FUNC_ISREF,      @fpsISREF);
    AddFunction(cat, 'ISTEXT',    'B', '?',    INT_EXCEL_SHEET_FUNC_ISTEXT,     @fpsISTEXT);

    (*
    // Lookup / reference functions
    cat := bcLookup;
    AddFunction(cat, 'COLUMN',    'I', 'R',    INT_EXCEL_SHEET_FUNC_COLUMN,     @fpsCOLUMN);
                  *)

  end;
end;


  (*
function fpsCOUNTIF(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// COUNTIF( range, criteria )
// - "range" is to the cell range to be analyzed
// - "citeria" can be a cell, a value or a string starting with a symbol like ">" etc.
//   (in the former two cases a value is counted if equal to the criteria value)
var
  n: Integer;
  r, c: Cardinal;
  arg: TsArgument;
  cellarg: TsArgument;
  criteria: TsArgument;
  compare: TsCompareOperation;
  res: Integer;
  cell: PCell;
begin
  Unused(NumArgs);
  criteria := Args.Pop;
  arg := Args.Pop;
  compare := coEqual;
  case criteria.ArgumentType of
    atCellRange:
      criteria := CreateCellArg(criteria.Worksheet.FindCell(criteria.FirstRow, criteria.FirstCol));
    atString:
      criteria.Stringvalue := AnalyzeCompareStr(criteria.StringValue, compare);
  end;
  n := 0;
  for r := arg.FirstRow to arg.LastRow do
    for c := arg.FirstCol to arg.LastCol do begin
      cell := arg.Worksheet.FindCell(r, c);
      if cell <> nil then begin
        cellarg := CreateCellArg(cell);
        res := CompareArgs(cellarg, criteria, false);
        if res <> MaxInt then begin
          if (res < 0) and (compare in [coLess, coLessEqual, coNotEqual])
            then inc(n)
          else
          if (res = 0) and (compare in [coEqual, coLessEqual, coGreaterEqual])
            then inc(n)
          else
          if (res > 0) and (compare in [coGreater, coGreaterEqual, coNotEqual])
            then inc(n);
        end else
          if (compare = coNotEqual) then inc(n);
      end else
       if compare = coNotEqual then inc(n);
    end;
  Result := CreateNumberArg(n);
end;
*)

  (*
function fpsSUMIF(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// SUMIF( range, criteria [, sum_range] )
// - "range" is to the cell range to be analyzed
// - "citeria" can be a cell, a value or a string starting with a symbol like ">" etc.
//   (in the former two cases a value is counted if equal to the criteria value)
// - "sum_range" identifies the cells to sum. If omitted, the function uses
//   "range" as the "sum_range"
var
  cellval, sum: Double;
  r, c, rs, cs: Cardinal;
  range: TsArgument;
  sum_range: TsArgument;
  cellarg: TsArgument;
  criteria: TsArgument;
  compare: TsCompareOperation;
  res: Integer;
  cell: PCell;
  accept: Boolean;
begin
  if NumArgs = 3 then begin
    sum_range := Args.Pop;
    criteria := Args.Pop;
    range := Args.Pop;
  end else begin
    criteria := Args.Pop;
    range := Args.Pop;
    sum_range := range;
  end;

  if (range.LastCol - range.FirstCol <> sum_range.LastCol - sum_range.FirstCol) or
     (range.LastRow - range.FirstRow <> sum_range.LastRow - sum_range.FirstRow)
  then begin
    Result := CreateErrorArg(errArgError);
    exit;
  end;

  compare := coEqual;
  case criteria.ArgumentType of
    atCellRange:
      criteria := CreateCellArg(criteria.Worksheet.FindCell(criteria.FirstRow, criteria.FirstCol));
    atString:
      criteria.Stringvalue := AnalyzeCompareStr(criteria.StringValue, compare);
  end;

  sum := 0.0;
  for r := range.FirstRow to range.LastRow do begin
    rs := r - range.FirstRow + sum_range.FirstRow;
    for c := range.FirstCol to range.LastCol do begin
      cs := c - range.FirstCol + sum_range.FirstCol;
      cell := range.Worksheet.FindCell(r, c);
      accept := (compare = coNotEqual);
      if cell <> nil then begin
        cellarg := CreateCellArg(cell);
        res := CompareArgs(cellarg, criteria, false);
        if res <> MaxInt then
          accept := ( (res < 0) and (compare in [coLess, coLessEqual, coNotEqual]) )
                 or ( (res = 0) and (compare in [coEqual, coLessEqual, coGreaterEqual]) )
                 or ( (res > 0) and (compare in [coGreater, coGreaterEqual, coNotEqual]) )
      end;
      if accept then begin
        cell := sum_range.Worksheet.FindCell(rs, cs);
        if sum_range.Worksheet.ReadNumericValue(cell, cellval) then
          sum := sum + cellval;
      end;
    end;
  end;
  Result := CreateNumberArg(sum);
end;
*)

{ Lookup / reference functions }
  (*
function fpsCOLUMN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ COLUMN( [reference] )
  Returns the column number of a cell reference (starting at 1).
  "reference" is a reference to a cell or range of cells.
  If omitted, it is assumed that the reference is the cell address in which the
  COLUMN function has been entered in. }
var
  arg: TsArgument;
begin
  if NumArgs = 0 then
    Result := CreateErrorArg(errArgError);
    // We don't know here which cell contains the formula.

  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumberArg(arg.Cell^.Col + 1);
    atCellRange: Result := CreateNumberArg(arg.FirstCol + 1);
    else         Result := CreateErrorArg(errWrongType);
  end;
end;

function fpsCOLUMNS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ COLUMNS( [reference] )
  returns the number of column in a cell reference. }
var
  arg: TsArgument;
begin
  Unused(NumArgs);
  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumberArg(1);
    atCellRange: Result := CreateNumberArg(arg.LastCol - arg.FirstCol + 1);
    else         Result := CreateErrorArg(errWrongType);
  end;
end;

function fpsROW(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ ROW( [reference] )
  Returns the row number of a cell reference (starting at 1!)
  "reference" is a reference to a cell or range of cells.
  If omitted, it is assumed that the reference is the cell address in which the
  ROW function has been entered in. }
var
  arg: TsArgument;
begin
  if NumArgs = 0 then
    Result := CreateErrorArg(errArgError);
    // We don't know here which cell contains the formula.

  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumberArg(arg.Cell^.Row + 1);
    atCellRange: Result := CreateNumberArg(arg.FirstRow + 1);
    else         Result := CreateErrorArg(errWrongType);
  end;
end;


function fpsROWS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ ROWS( [reference] )
  returns the number of rows in a cell reference. }
var
  arg: TsArgument;
begin
  Unused(NumArgs);
  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumberArg(1);
    atCellRange: Result := CreateNumberArg(arg.LastRow - arg.FirstRow + 1);
    else         Result := CreateErrorArg(errWrongType);
  end;
end;

*)

(*
function fpsINFO(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{  INFO( type )
   returns information about the operating environment.
   type can be one of the following values:
     + "directory"    Path of the current directory.
     + "numfile"      Number of active worksheets.
     - "origin"       The cell that is in the top, left-most cell visible in the current Excel spreadsheet.
     - "osversion"    Operating system version.
     - "recalc"       Returns the recalculation mode - either Automatic or Manual.
     - "release"      Version of Excel that you are running.
     - "system"       Name of the operating environment.
   ONLY THOSE MARKED BY "+" ARE SUPPORTED! }
var
  arg: TsArgument;
  workbook: TsWorkbook;
  s: String;
begin
  Unused(NumArgs);
  arg := Args.Pop;
  if arg.ArgumentType <> atString then
    Result := CreateErrorArg(errWrongType)
  else begin
    s := Lowercase(arg.StringValue);
    workbook := arg.Worksheet.Workbook;
    if s = 'directory' then
      Result := CreateStringArg(ExtractFilePath(workbook.FileName))
    else
    if s = 'numfile' then
      Result := CreateNumberArg(workbook.GetWorksheetCount)
    else
      Result := CreateErrorArg(errFormulaNotSupported);
  end;
end;

*)

initialization

end.
