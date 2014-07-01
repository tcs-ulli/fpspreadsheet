unit fpsfunc;

{$mode objfpc}

interface

uses
  Classes, SysUtils, fpspreadsheet;

type
  TsArgumentType = (atNumber, atString, atBool, atError, atEmpty);

  TsArgument = record
    IsMissing: Boolean;
    case ArgumentType: TsArgumentType of
      atNumber  : (NumberValue: Double);
      atString  : (StringValue: String);
      atBool    : (BoolValue: Boolean);
      atError   : (ErrorValue: TsErrorValue);
  end;
  PsArgument = ^TsArgument;

  TsArgumentStack = class(TFPList)
  public
    destructor Destroy; override;
    function Pop: TsArgument;
    procedure Push(AValue: TsArgument);
    procedure PushBool(AValue: Boolean);
    procedure PushMissing;
    procedure PushNumber(AValue: Double);
    procedure PushString(AValue: String);
    procedure Clear;
    procedure Delete(AIndex: Integer);
  end;

function CreateBool(AValue: Boolean): TsArgument;
function CreateNumber(AValue: Double): TsArgument;
function CreateString(AValue: String): TsArgument;
function CreateError(AError: TsErrorValue): TsArgument;
function CreateEmpty: TsArgument;

{
These are the functions called when calculating an RPN formula.
}
type
  TsFormulaFunc = function(Args: TsArgumentStack; NumArgs: Integer): TsArgument;

function fpsAdd         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSub         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMul         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsDiv         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsPercent     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsPower       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsUMinus      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsUPlus       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsConcat      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsEqual       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsGreater     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsGreaterEqual(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsLess        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsLessEqual   (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsNotEqual    (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ Math }
function fpsABS         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsACOS        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsACOSH       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsASIN        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsASINH       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsATAN        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsATANH       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsCOS         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsCOSH        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsDEGREES     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsEXP         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsINT         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsLN          (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsLOG         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsLOG10       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsPI          (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsRADIANS     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsRAND        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsROUND       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSIGN        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSIN         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSINH        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSQRT        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTAN         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTANH        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ Date/time functions }
function fpsDATE        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsDATEDIF     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsDATEVALUE   (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsDAY         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsHOUR        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMINUTE      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMONTH       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsNOW         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSECOND      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTIME        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTIMEVALUE   (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTODAY       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsWEEKDAY     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsYEAR        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ Statistical functions }
function fpsAVEDEV      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsAVERAGE     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsCOUNT       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMAX         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMIN         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsPRODUCT     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSTDEV       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSTDEVP      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSUM         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSUMSQ       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsVAR         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsVARP        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ Logical functions }
function fpsAND         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsFALSE       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsIF          (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsNOT         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsOR          (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTRUE        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ String functions }
function fpsCHAR        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsCODE        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsLEFT        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsLOWER       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMID         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsREPLACE     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsRIGHT       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSUBSTITUTE  (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTRIM        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsUPPER       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ info functions }
function fpsISERR       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISERROR     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISLOGICAL   (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISNA        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISNONTEXT   (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISNUMBER    (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISTEXT      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsVALUE       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;

implementation

uses
  Math, lazutf8, StrUtils, DateUtils, fpsUtils;

type
  TBoolArray  = array of boolean;
  TFloatArray = array of double;
  TStrArray   = array of string;

{ TsArgumentStack }

destructor TsArgumentStack.Destroy;
begin
  Clear;
  inherited Destroy;
end;

procedure TsArgumentStack.Clear;
var
  i: Integer;
begin
  for i := Count-1 downto 0 do
    Delete(i);
  inherited Clear;
end;

procedure TsArgumentStack.Delete(AIndex: Integer);
var
  P: PsArgument;
begin
  P := PsArgument(Items[AIndex]);
  P^.StringValue := '';
  FreeMem(P, SizeOf(P));
  inherited Delete(AIndex);
end;

function TsArgumentStack.Pop: TsArgument;
var
  P: PsArgument;
begin
  P := PsArgument(Items[Count-1]);
  Result := P^;
  Result.StringValue := P^.StringValue;  // necessary?
  Delete(Count-1);
end;

procedure TsArgumentStack.Push(AValue: TsArgument);
var
  P: PsArgument;
begin
  GetMem(P, SizeOf(TsArgument));
  P^ := AValue;
  P^.StringValue := AValue.StringValue;
  Add(P);
end;

procedure TsArgumentStack.PushBool(AValue: Boolean);
var
  arg: TsArgument;
begin
  arg.ArgumentType := atBool;
  arg.BoolValue := AValue;
  arg.IsMissing := false;
  Push(arg);
end;

procedure TsArgumentStack.PushMissing;
var
  arg: TsArgument;
begin
  arg.IsMissing := true;
  Push(arg);
end;

procedure TsArgumentStack.PushNumber(AValue: Double);
var
  arg: TsArgument;
begin
  arg.ArgumentType := atNumber;
  arg.NumberValue := AValue;
  arg.IsMissing := false;
  Push(arg);
end;

procedure TsArgumentStack.PushString(AValue: String);
var
  arg: TsArgument;
begin
  arg.ArgumentType := atString;
  arg.StringValue := AValue;
  arg.IsMissing := false;
  Push(arg);
end;


{ Preparing arguments }

function GetBoolFromArgument(Arg: TsArgument; var AValue: Boolean): TsErrorValue;
begin
  case Arg.ArgumentType of
    atBool : begin
               AValue := Arg.BoolValue;
               Result := errOK;
             end;
    atError: Result := Arg.ErrorValue;
    else     Result := errWrongType;
  end;
end;

function GetNumberFromArgument(Arg: TsArgument; var ANumber: Double): TsErrorValue;
begin
  Result := errOK;
  case Arg.ArgumentType of
    atNumber : ANumber := Arg.NumberValue;
    atString : if not TryStrToFloat(arg.StringValue, ANumber) then Result := errWrongType;
    atBool   : if Arg.BoolValue then ANumber := 1.0 else ANumber := 0.0;
    atError  : Result := Arg.ErrorValue;
  end;
end;

function GetStringFromArgument(Arg: TsArgument; var AString: String): TsErrorValue;
begin
  case Arg.ArgumentType of
    atString : begin
                 AString := Arg.StringValue;
                 Result := errOK;
               end;
    atError  : Result := Arg.ErrorValue;
    else       Result := errWrongType;
  end;
end;

function CreateBool(AValue: Boolean): TsArgument;
begin
  Result.ArgumentType := atBool;
  Result.Boolvalue := AValue;
  Result.IsMissing := false;
end;

function CreateNumber(AValue: Double): TsArgument;
begin
  Result.ArgumentType := atNumber;
  Result.NumberValue := AValue;
  Result.IsMissing := false;
end;

function CreateString(AValue: String): TsArgument;
begin
  Result.ArgumentType := atString;
  Result.StringValue := AValue;
  Result.IsMissing := false;
end;

function CreateError(AError: TsErrorValue): TsArgument;
begin
  Result.ArgumentType := atError;
  Result.ErrorValue := AError;
  Result.IsMissing := false;
end;

function CreateEmpty: TsArgument;
begin
  Result.ArgumentType := atEmpty;
  Result.IsMissing := false;
end;

{@@
  Pops boolean values from the argument stack. Is called when calculating rpn
  formulas.
  @param  Args      Argument stack to be used.
  @param  NumArgs   Count of arguments to be popped from the stack
  @param  AValues   (output) Array containing the retrieved boolean values.
                    The array length is given by NumArgs. The data in the array
                    are in the same order in which they were pushed onto the stack.
                    Missing arguments are not included in the array, the case
                    of missing arguments must be handled separately if the are
                    important.
  @param  AErrArg   (output) Argument containing an error code, e.g. errWrongType
                    if non-boolean data were met on the stack.
  @return TRUE if everything was ok, FALSE, if AErrArg reports an error. }
function PopBoolValues(Args: TsArgumentStack; NumArgs: Integer;
  out AValues: TBoolArray; out AErrArg: TsArgument): Boolean;
var
  arg: TsArgument;
  err: TsErrorValue;
  counter, j: Integer;
  b: Boolean;
begin
  SetLength(AValues, NumArgs);
  j := 0;
  for counter := 1 to NumArgs do begin
    arg := Args.Pop;
    if not arg.IsMissing then begin
      err := GetBoolFromArgument(arg, b);
      if err = errOK then begin
        AValues[j] := b;
        inc(j);
      end else begin
        Result := false;
        AErrArg := CreateError(err);
        SetLength(AValues, 0);
        exit;
      end;
    end;
  end;
  Result := true;
  AErrArg := CreateError(errOK);
  SetLength(AValues, j);
  // Flip array - we want to have the arguments in the array in the same order
  // they were pushed.
  for j:=0 to Length(AValues) div 2 - 1 do begin
    b := AValues[j];
    AValues[j] := AValues[High(AValues)-j];
    AValues[High(AValues)-j] := b;
  end;
end;

function PopDateValue(Args: TsArgumentStack; out ADate: TDate;
  out AErrArg: TsArgument): Boolean;
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  case arg.ArgumentType of
    atError, atBool, atEmpty:
      begin
        Result := false;
        AErrArg := CreateError(errWrongType);
      end;
    atNumber:
      begin
        Result := true;
        ADate := arg.NumberValue;
      end;
    atString:
      begin
        Result := TryStrToDate(arg.StringValue, ADate);
        if not Result then AErrArg := CreateError(errWrongType);
      end
  end;
end;

{@@
  Pops floating point values from the argument stack. Is called when
  calculating rpn formulas.
  @param  Args      Argument stack to be used.
  @param  NumArgs   Count of arguments to be popped from the stack
  @param  AValues   (output) Array containing the retrieved float values.
                    The array length is given by NumArgs. The data in the array
                    are in the same order in which they were pushed onto the stack.
                    Missing arguments are not included in the array, the case
                    of missing arguments must be handled separately if the are
                    important.
  @param  AErrArg   (output) Argument containing an error code, e.g. errWrongType
                    if non-float data were met on the stack.
  @return TRUE if everything was ok, FALSE, if AErrArg reports an error. }
function PopFloatValues(Args: TsArgumentStack; NumArgs: Integer;
  out AValues: TFloatArray; out AErrArg: TsArgument): Boolean;
var
  arg: TsArgument;
  err: TsErrorValue;
  counter, j: Integer;
  val: double;
begin
  SetLength(AValues, NumArgs);
  j := 0;
  for counter := 1 to NumArgs do begin
    arg := Args.Pop;
    if not arg.IsMissing then begin
      err := GetNumberFromArgument(arg, val);
      if err = errOK then begin
        AValues[j] := val;
        inc(j);
      end else begin
        Result := false;
        SetLength(AValues, 0);
        AErrArg := CreateError(errWrongType);
        exit;
      end;
    end;
  end;
  Result := true;
  SetLength(AValues, j);
  // Flip array - we want to have the arguments in the array in the same order
  // they were pushed.
  for j:=0 to Length(AValues) div 2 - 1 do begin
    val := AValues[j];
    AValues[j] := AValues[High(AValues)-j];
    AValues[High(AValues)-j] := val;
  end;
  AErrArg := CreateError(errOK);
end;

{@@
  Pops string arguments from the argument stack. Is called when calculating
  rpn formulas.
  @param  Args      Argument stack to be used.
  @param  NumArgs   Count of arguments to be popped from the stack
  @param  AValues   (output) Array containing the retrieved strings. The array
                    length is given by NumArgs. The data in the array are in the
                    same order in which they were pushed onto the stack.
                    Missing arguments are not included in the array, the case
                    of missing arguments must be handled separately if the are
                    important.
  @param  AErrArg   (output) Argument containing an error code , e.g. errWrongType
                    if non-string data were met on the stack.
  @return TRUE if everything was ok, FALSE, if AErrArg reports an error. }
function PopStringValues(Args: TsArgumentStack; NumArgs: Integer;
  out AValues: TStrArray; out AErrArg: TsArgument): Boolean;
var
  arg: TsArgument;
  err: TsErrorValue;
  counter, j: Integer;
  s: String;
begin
  SetLength(AValues, NumArgs);
  j := 0;
  for counter := 1 to NumArgs do begin
    arg := Args.Pop;
    if not arg.IsMissing then begin
      err := GetStringFromArgument(arg, s);
      if err = errOK then begin
        AValues[j] := s;
        inc(j);
      end else begin
        Result := false;
        AErrArg := CreateError(errWrongType);
        SetLength(AValues, 0);
        exit;
      end;
    end;
  end;
  Result := true;
  AErrArg := CreateError(errOK);
  SetLength(AValues, j);
  // Flip array - we want to have the arguments in the array in the same order
  // they were pushed.
  for j:=0 to Length(AValues) div 2 - 1 do begin
    s := AValues[j];
    AValues[j] := AValues[High(AValues)-j];
    AValues[High(AValues)-j] := s;
  end;
end;


function PopTimeValue(Args: TsArgumentStack; out ATime: TTime;
  out AErrArg: TsArgument): Boolean;
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  case arg.ArgumentType of
    atError, atBool, atEmpty:
      begin
        Result := false;
        AErrArg := CreateError(errWrongType);
      end;
    atNumber:
      begin
        Result := true;
        ATime := frac(arg.NumberValue);
      end;
    atString:
      begin
        Result := TryStrToTime(arg.StringValue, ATime);
        if not Result then AErrArg := CreateError(errWrongType);
      end
  end;
end;


{ Operations }

function fpsAdd(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 2, data, Result) then
    Result := CreateNumber(data[0] + data[1]);
end;

function fpsSub(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 2, data, Result) then
    Result := CreateNumber(data[0] - data[1]);
end;

function fpsMul(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 2, data, Result) then
    Result := CreateNumber(data[0] * data[1]);
end;

function fpsDiv(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 2, data, Result) then begin
    if data[1] = 0 then
      Result := CreateError(errDivideByZero)
    else
      Result := CreateNumber(data[0] / data[1]);
  end;
end;

function fpsPercent(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(data[0] * 0.01);
end;

function fpsPower(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 2, data, Result) then
    try
      Result := CreateNumber(power(data[0], data[1]));
    except on E: EInvalidArgument do
      Result := CreateError(errOverflow);
      // this could happen, e.g., for "power( (neg value), (non-integer) )"
    end;
end;

function fpsUMinus(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(-data[0]);
end;

function fpsUPlus(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(data[0]);
end;

function fpsConcat(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TStrArray;
begin
  if PopStringValues(Args, 2, data, Result) then
    Result := CreateString(data[0] + data[1]);
end;

function fpsEqual(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  arg1, arg2: TsArgument;
begin
  arg2 := Args.Pop;
  arg1 := Args.Pop;
  if arg1.ArgumentType = arg2.ArgumentType then
    case arg1.ArgumentType of
      atNumber  : Result := CreateBool(arg1.NumberValue = arg2.NumberValue);
      atString  : Result := CreateBool(arg1.StringValue = arg2.StringValue);
      atBool    : Result := CreateBool(arg1.Boolvalue = arg2.BoolValue);
    end
  else
    Result := CreateBool(false);
end;

function fpsGreater(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  arg1, arg2: TsArgument;
begin
  arg2 := Args.Pop;
  arg1 := Args.Pop;
  if arg1.ArgumentType = arg2.ArgumentType then
    case arg1.ArgumentType of
      atNumber  : Result := CreateBool(arg1.NumberValue > arg2.NumberValue);
      atString  : Result := CreateBool(arg1.StringValue > arg2.StringValue);
      atBool    : Result := CreateBool(arg1.Boolvalue > arg2.BoolValue);
    end
  else
    Result := CreateBool(false);
end;

function fpsGreaterEqual(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  arg1, arg2: TsArgument;
begin
  arg2 := Args.Pop;
  arg1 := Args.Pop;
  if arg1.ArgumentType = arg2.ArgumentType then
    case arg1.ArgumentType of
      atNumber  : Result := CreateBool(arg1.NumberValue >= arg2.NumberValue);
      atString  : Result := CreateBool(arg1.StringValue >= arg2.StringValue);
      atBool    : Result := CreateBool(arg1.Boolvalue >= arg2.BoolValue);
    end
  else
    Result := CreateBool(false);
end;

function fpsLess(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  arg1, arg2: TsArgument;
begin
  arg2 := Args.Pop;
  arg1 := Args.Pop;
  if arg1.ArgumentType = arg2.ArgumentType then
    case arg1.ArgumentType of
      atNumber  : Result := CreateBool(arg1.NumberValue < arg2.NumberValue);
      atString  : Result := CreateBool(arg1.StringValue < arg2.StringValue);
      atBool    : Result := CreateBool(arg1.Boolvalue < arg2.BoolValue);
    end
  else
    Result := CreateBool(false);
end;

function fpsLessEqual(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  arg1, arg2: TsArgument;
begin
  arg2 := Args.Pop;
  arg1 := Args.Pop;
  if arg1.ArgumentType = arg2.ArgumentType then
    case arg1.ArgumentType of
      atNumber  : Result := CreateBool(arg1.NumberValue <= arg2.NumberValue);
      atString  : Result := CreateBool(arg1.StringValue <= arg2.StringValue);
      atBool    : Result := CreateBool(arg1.Boolvalue <= arg2.BoolValue);
    end
  else
    Result := CreateBool(false);
end;

function fpsNotEqual(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  arg1, arg2: TsArgument;
begin
  arg2 := Args.Pop;
  arg1 := Args.Pop;
  if arg1.ArgumentType = arg2.ArgumentType then
    case arg1.ArgumentType of
      atNumber  : Result := CreateBool(arg1.NumberValue <> arg2.NumberValue);
      atString  : Result := CreateBool(arg1.StringValue <> arg2.StringValue);
      atBool    : Result := CreateBool(arg1.Boolvalue <> arg2.BoolValue);
    end
  else
    Result := CreateBool(false);
end;


{ Math functions }

function fpsABS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(abs(data[0]));
end;

function fpsACOS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if InRange(data[0], -1, +1) then
      Result := CreateNumber(arccos(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsACOSH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if data[0] >= 1 then
      Result := CreateNumber(arccosh(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsASIN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if InRange(data[0], -1, +1) then
      Result := CreateNumber(arcsin(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsASINH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(arcsinh(data[0]));
end;

function fpsATAN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(arctan(data[0]));
end;

function fpsATANH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if (data[0] > -1) and (data[0] < +1) then
      Result := CreateNumber(arctanh(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsCOS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(cos(data[0]));
end;

function fpsCOSH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(cosh(data[0]));
end;

function fpsDEGREES(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(RadToDeg(data[0]));
end;

function fpsEXP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(exp(data[0]));
end;

function fpsINT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(floor(data[0]));
end;

function fpsLN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if (data[0] > 0) then
      Result := CreateNumber(ln(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsLOG(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// LOG( number, [base] )
var
  arg_base, arg_number: TsArgument;
  data: TFloatArray;
  base: Double;
begin
  base := 10;
  if NumArgs = 2 then begin
    arg_base := Args.Pop;
    if not arg_base.IsMissing then begin
      if arg_base.ArgumentType <> atNumber then begin
        Result := CreateError(errWrongType);
        exit;
      end;
      base := arg_base.NumberValue;
    end;
  end;
  arg_number := Args.Pop;

  if base < 0 then begin
    Result := CreateError(errOverflow);
    exit;
  end;

  if arg_number.ArgumentType <> atNumber then begin
    Result := CreateError(errWrongType);
    exit;
  end;

  if arg_number.NumberValue > 0 then
    Result := CreateNumber(logn(base, arg_number.NumberValue))
  else
    Result := CreateError(errOverflow);
end;

function fpsLOG10(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if (data[0] > 0) then
      Result := CreateNumber(log10(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsPI(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
begin
  Result := CreateNumber(pi);
end;

function fpsRADIANS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(degtorad(data[0]))
end;

function fpsRAND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
begin
  Result := CreateNumber(random);
end;

function fpsROUND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 2, data, Result) then
    Result := CreateNumber(RoundTo(data[0], round(data[1])))
end;

function fpsSIGN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(sign(data[0]))
end;

function fpsSIN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(sin(data[0]))
end;

function fpsSINH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(sinh(data[0]))
end;

function fpsSQRT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if data[0] >= 0.0 then
      Result := CreateNumber(sqrt(data[0]))
    else
      Result := CreateError(errOverflow);
  end;
end;

function fpsTAN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if frac(data[0] / (pi*0.5)) = 0 then
      Result := CreateError(errOverflow)
    else
      Result := CreateNumber(tan(data[0]))
  end;
end;

function fpsTANH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then
    Result := CreateNumber(tanh(data[0]))
end;


{ Date/time functions }

function fpsDATE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// DATE( year, month, day )
var
  data: TFloatArray;
  d: TDate;
begin
  if PopFloatValues(Args, 3, data, Result) then begin
    d := EncodeDate(round(data[0]), round(data[1]), round(data[2]));
    Result := CreateNumber(d);
  end;
end;

function fpsDATEDIF(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ DATEDIF( start_date, end_date, interval )
    start_date <= end_date !
    interval = Y  - The number of complete years.
             = M  - The number of complete months.
             = D  - The number of days.
             = MD - The difference between the days (months and years are ignored).
             = YM - The difference between the months (days and years are ignored).
             = YD - The difference between the days (years and dates are ignored). }
var
  data: TStrArray;
  start_date, end_date: TDate;
begin
  if not PopStringValues(Args, 1, data, Result) then exit;
  if not PopDateValue(Args, end_date, Result) then exit;
  if not PopDateValue(Args, start_date, Result) then exit;

  if end_date > start_date then
    Result := CreateError(errOverflow)
  else if data[0] = 'Y' then
    Result := CreateNumber(YearsBetween(end_date, start_date))
  else if data[0] = 'M' then
    Result := CreateNumber(MonthsBetween(end_date, start_date))
  else if data[0] = 'D' then
    Result := CreateNumber(DaysBetween(end_date, start_date))
  else
    Result := CreateError(errFormulaNotSupported);
end;

function fpsDATEVALUE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// DATEVALUE( date )   -- date can be a string or a date/time
var
  d: TDate;
begin
  if PopDateValue(Args, d, Result) then
    Result := CreateNumber(d);
end;

function fpsDAY(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  d: TDate;
begin
  if PopDateValue(Args, d, Result) then
    Result := CreateNumber(DayOf(d));
end;

function fpsHOUR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  t: TTime;
begin
  if PopTimeValue(Args, t, Result) then
    Result := CreateNumber(HourOf(t));
end;

function fpsMINUTE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  t: TTime;
begin
  if PopTimeValue(Args, t, Result) then
    Result := CreateNumber(MinuteOf(t));
end;

function fpsMONTH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  d: TDate;
begin
  if PopDateValue(Args, d, Result) then
    Result := CreateNumber(MonthOf(d));
end;

function fpsNOW(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// NOW()
begin
  Result := CreateNumber(now);
end;

function fpsSECOND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  t: TTime;
begin
  if PopTimeValue(Args, t, Result) then
    Result := CreateNumber(SecondOf(t));
end;

function fpsTIME(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// TIME( hour, minute, second )
var
  data: TFloatArray;
  t: TTime;
begin
  if PopFloatValues(Args, 3, data, Result) then begin
    t := EncodeTime(round(data[0]), round(data[1]), round(data[2]), 0);
    Result := CreateNumber(t);
  end;
end;

function fpsTIMEVALUE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// TIMEVALUE( time_value )
var
  t: TTime;
begin
  if PopTimeValue(Args, t, Result) then
    Result := CreateNumber(t);
end;

function fpsToday(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// TODAY()
begin
  Result := CreateNumber(Date());
end;

function fpsWEEKDAY(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ WEEKDAY( serial_number, [return_value]
   return_value = 1 - Returns a number from 1 (Sunday) to 7 (Saturday).
                      This is the default if parameter is omitted.
                = 2 - Returns a number from 1 (Monday) to 7 (Sunday).
                = 3 - Returns a number from 0 (Monday) to 6 (Sunday). }
var
  d: TDate;
  data: TFloatArray;
  n: Integer;
begin
  n := 1;
  if NumArgs = 2 then
    if PopFloatValues(Args, 1, data, Result) then n := round(data[0])
      else exit;
  if PopDateValue(Args, d, Result) then
    Result := CreateNumber(DayOfWeek(d));
end;

function fpsYEAR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  d: TDate;
begin
  if PopDateValue(Args, d, Result) then
    Result := CreateNumber(YearOf(d));
end;


{ Statistical functions }

function fpsAVEDEV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// Average value of absolute deviations of data from their mean.
// AVEDEV( argument1, [argument2, ... argument_n] )
var
  data: TFloatArray;
  m: Double;
  i: Integer;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then begin
    m := Mean(data);
    for i:=0 to High(data) do
      data[i] := abs(data[i] - m);
    m := Mean(data);
    Result := CreateNumber(m)
  end;
end;

function fpsAVERAGE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// AVERAGE( argument1, [argument2, ... argument_n] )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(Mean(data))
end;

function fpsCOUNT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ counts the number of cells that contain numbers as well as the number of
  arguments that contain numbers.
  COUNT( argument1, [argument2, ... argument_n] )
}
var
  n, i: Integer;
  arg: TsArgument;
begin
  n := 0;
  for i:=1 to NumArgs do begin
    arg := Args.Pop;
    if (not arg.IsMissing) and (arg.ArgumentType = atNumber) then inc(n);
  end;
  Result := CreateNumber(n);
end;

function fpsMAX(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// MAX( number1, number2, ... number_n )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(MaxValue(data))
end;

function fpsMIN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// MIN( number1, number2, ... number_n )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(MinValue(data))
end;

function fpsPRODUCT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// PRODUCT( number1, number2, ... number_n )
var
  data: TFloatArray;
  i: Integer;
  p: Double;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then begin
    p := 1.0;
    for i:=0 to High(data) do
      p := p * data[i];
    Result := CreateNumber(p);
  end;
end;

function fpsSTDEV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// STDEV( number1, [number2, ... number_n] )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(StdDev(data))
end;

function fpsSTDEVP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// STDEVP( number1, [number2, ... number_n] )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(PopnStdDev(data))
end;

function fpsSUM(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// SUM( value1, [value2, ... value_n] )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(Sum(data))
end;

function fpsSUMSQ(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// SUMSQ( value1, [value2, ... value_n] )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(SumOfSquares(data))
end;

function fpsVAR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// VAR( number1, number2, ... number_n )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(Variance(data))
end;

function fpsVARP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// VARP( number1, number2, ... number_n )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(PopnVariance(data))
end;


{ Logical functions }

function fpsAND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// AND( condition1, [condition2], ... )
var
  data: TBoolArray;
  i: Integer;
  b: Boolean;
begin
  if PopBoolValues(Args, NumArgs, data, Result) then begin
    // If at least one case is false the entire AND condition is false
    b := true;
    for i:=0 to High(data) do
      if not data[i] then begin
        b := false;
        break;
      end;
    Result := CreateBool(b);
  end;
end;

function fpsFALSE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// FALSE( )
begin
  Result := CreateBool(false);
end;

function fpsIF(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// IF( condition, [value_if_true], [value_if_false] )
var
  condition: TsArgument;
  case1, case2: TsArgument;
  err: TsErrorValue;
begin
  if NumArgs = 3 then
    case2 := Args.Pop;
  case1 := Args.Pop;
  condition := Args.Pop;
  if condition.ArgumentType <> atBool then
    Result := CreateError(errWrongType)
  else
    case NumArgs of
      2: if condition.BoolValue then Result := case1 else Result := Condition;
      3: if condition.BoolValue then Result := case1 else Result := case2;
    end;
end;

function fpsNOT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// NOT( logical_value )
var
  data: TBoolArray;
begin
  if PopBoolValues(Args, NumArgs, data, Result) then
    Result := CreateBool(not data[0]);
end;

function fpsOR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// OR( condition1, [condition2], ... )
var
  data: TBoolArray;
  i: Integer;
  b: Boolean;
begin
  if PopBoolValues(Args, NumArgs, data, Result) then begin
    // If at least one case is true, the entire OR condition is true
    b := false;
    for i:=0 to High(data) do
      if data[i] then begin
        b := true;
        break;
      end;
    Result := CreateBool(b);
  end;
end;

function fpsTRUE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// TRUE ( )
begin
  Result := CreateBool(true);
end;


{ String functions }

function fpsCHAR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// CHAR( ascii_value )
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, 1, data, Result) then begin
    if (data[0] >= 0) and (data[0] <= 255) then
      Result := CreateString(AnsiToUTF8(Char(Round(data[0]))));
  end;
end;

function fpsCODE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// CODE( text )
var
  data: TStrArray;
begin
  if PopStringValues(Args, 1, data, Result) then begin
    if Length(data) > 0 then
      Result := CreateNumber(Ord(UTF8ToAnsi(data[0])[1]))
    else
      Result := CreateEmpty;
  end;
end;

function fpsLEFT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// LEFT( text, [number_of_characters] )
var
  arg1, arg2: TsArgument;
  count: Integer;
  s: String;
begin
  count := 1;
  if NumArgs = 2 then begin
    arg2 := Args.Pop;
    if not arg2.IsMissing then begin
      if arg2.ArgumentType <> atNumber then begin
        Result := createError(errWrongType);
        exit;
      end;
      count := Round(arg2.NumberValue);
    end;
  end;
  arg1 := Args.Pop;
  if arg1.ArgumentType <> atString then begin
    Result := CreateError(errWrongType);
    exit;
  end;
  s := arg1.StringValue;
  Result := CreateString(UTF8LeftStr(s, count));
end;

function fpsLOWER(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// LOWER( text )
var
  data: TStrArray;
begin
  if PopStringValues(Args, NumArgs, data, Result) then
    Result := CreateString(UTF8LowerCase(data[0]));
end;

function fpsMID(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// MID( text, start_position, number_of_characters )
var
  fdata: TFloatArray;
  sdata: TStrArray;
begin
  if PopFloatValues(Args, 2, fdata, Result) then
    if PopStringValues(Args, 1, sdata, Result) then
      Result := CreateString(UTF8Copy(sdata[0], Round(fData[0]), Round(fdata[1])));
end;

function fpsREPLACE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// REPLACE( old_text, start, number_of_chars, new_text )
var
  arg_new, arg_old, arg_start, arg_count: TsArgument;
  s, s1, s2, snew: String;
  p, count: Integer;
begin
  arg_new := Args.Pop;
  if arg_new.ArgumentType <> atString then begin
    Result := CreateError(errWrongType);
    exit;
  end;
  arg_count := Args.Pop;
  if arg_count.ArgumentType <> atNumber then begin
    Result := CreateError(errWrongType);
    exit;
  end;
  arg_start := Args.Pop;
  if arg_start.ArgumentType <> atNumber then begin
    Result := CreateError(errWrongType);
    exit;
  end;
  arg_old := Args.Pop;
  if arg_old.ArgumentType <> atString then begin
    Result := CreateError(errWrongType);
    exit;
  end;

  s := arg_old.StringValue;
  snew := arg_new.StringValue;
  p := round(arg_start.NumberValue);
  count := round(arg_count.NumberValue);

  s1 := UTF8Copy(s, 1, p-1);
  s2 := UTF8Copy(s, p+count, UTF8Length(s));
  Result := CreateString(s1 + snew + s2);
end;

function fpsRIGHT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// RIGHT( text, [number_of_characters] )
var
  arg1, arg2: TsArgument;
  count: Integer;
  s: String;
begin
  count := 1;
  if NumArgs = 2 then begin
    arg2 := Args.Pop;
    if not arg2.IsMissing then begin
      if arg2.ArgumentType <> atNumber then begin
        Result := createError(errWrongType);
        exit;
      end;
      count := round(arg2.NumberValue);
    end;
  end;
  arg1 := Args.Pop;
  if arg1.ArgumentType <> atString then begin
    Result := CreateError(errWrongType);
    exit;
  end;
  s := arg1.StringValue;
  Result := CreateString(UTF8RightStr(s, count));
end;

function fpsSUBSTITUTE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// SUBSTITUTE( text, old_text, new_text, [nth_appearance] )
var
  n: Integer;
  arg: TsArgument;
  data: TStrArray;
  s, s_old, s_new: String;
begin
  Result := CreateError(errWrongType);
  n := -1;
  if (NumArgs = 4) then begin
    arg := Args.Pop;
    if not arg.IsMissing and (arg.ArgumentType <> atNumber) then
      exit;
    n := round(arg.NumberValue);
  end;

  if PopStringValues(Args, 3, data, Result) then begin
    s := data[0];
    s_old := data[1];
    s_new := data[2];
    if n = -1 then
      Result := CreateString(UTF8StringReplace(s, s_old, s_new, [rfReplaceAll]))
    else
      Result := CreateError(errFormulaNotSupported);
  end;
end;

function fpsTRIM(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// TRIM( text )
var
  data: TStrArray;
begin
  if PopStringValues(Args, NumArgs, data, Result) then
    Result := CreateString(UTF8Trim(data[0]));
end;

function fpsUPPER(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// UPPER( text )
var
  data: TStrArray;
begin
  if PopStringValues(Args, NumArgs, data, Result) then
    Result := CreateString(UTF8UpperCase(data[0]));
end;


{ Info functions }

function fpsISERR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISERR( value )
// If value is an error value (except #N/A), this function will return TRUE.
// Otherwise, it will return FALSE.
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool((arg.ArgumentType = atError) and (arg.ErrorValue <> errArgError));
end;

function fpsISERROR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISERROR( value )
// If value is an error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?
// or #NULL), this function will return TRUE. Otherwise, it will return FALSE.
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool((arg.ArgumentType = atError));
end;

function fpsISLOGICAL(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISLOGICAL( value )
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool(arg.ArgumentType = atBool);
end;

function fpsISNA(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISNA( value )
//  If value is a #N/A error value , this function will return TRUE.
// Otherwise, it will return FALSE.
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool((arg.ArgumentType = atError) and (arg.ErrorValue = errArgError));
end;

function fpsISNONTEXT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISNONTEXT( value )
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool(arg.ArgumentType <> atString);
end;

function fpsISNUMBER(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISNUMBER( value )
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool(arg.ArgumentType = atNumber);
end;

function fpsISTEXT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISTEXT( value )
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool(arg.ArgumentType = atString);
end;

function fpsVALUE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// VALUE( text )
// text is the text value to convert to a number. If text is not a number, the
// VALUE function will return #VALUE!.
var
  data: TStrArray;
  x: Double;
begin
  if PopStringValues(Args, 1, data, Result) then
    if TryStrToFloat(data[0], x) then
      Result := CreateNumber(x)
    else
      Result := CreateError(errWrongType);
end;


end.
