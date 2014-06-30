unit fpsmath;

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
function fpsLOWER       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsTRIM        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsUPPER       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;

implementation

uses
  Math;

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
  end;
end;

function GetStringFromArgument(Arg: TsArgument; var AString: String): TsErrorValue;
begin
  case Arg.ArgumentType of
    atString : begin
                 AString := Arg.StringValue;
                 Result := errOK;
               end;
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

  if base < 0 then begin
    Result := CreateError(errOverflow);
    exit;
  end;

  arg_number := Args.Pop;
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


{ Statistical functions }

function fpsAVEDEV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// Average value of absolute deviations of data from their mean.
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
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(Mean(data))
end;

function fpsCOUNT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ Count non-missing arguments }
var
  n, i: Integer;
begin
  n := 0;
  for i:=1 to NumArgs do
    if not Args.Pop.IsMissing then inc(n);
  Result := CreateNumber(n);
end;

function fpsMAX(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(MaxValue(data))
end;

function fpsMIN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(MinValue(data))
end;

function fpsPRODUCT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
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
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(StdDev(data))
end;

function fpsSTDEVP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(PopnStdDev(data))
end;

function fpsSUM(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(Sum(data))
end;

function fpsSUMSQ(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(SumOfSquares(data))
end;

function fpsVAR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(Variance(data))
end;

function fpsVARP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TFloatArray;
begin
  if PopFloatValues(Args, NumArgs, data, Result) then
    Result := CreateNumber(PopnVariance(data))
end;


{ Logical functions }

function fpsAND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
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
begin
  Result := CreateBool(false);
end;

function fpsIF(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
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
var
  data: TBoolArray;
begin
  if PopBoolValues(Args, NumArgs, data, Result) then
    Result := CreateBool(not data[0]);
end;

function fpsOR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
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
begin
  Result := CreateBool(true);
end;


{ String functions }

function fpsLOWER(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TStrArray;
begin
  if PopStringValues(Args, NumArgs, data, Result) then
    Result := CreateString(lowercase(data[0]));
end;

function fpsTRIM(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TStrArray;
begin
  if PopStringValues(Args, NumArgs, data, Result) then
    Result := CreateString(trim(data[0]));
end;

function fpsUPPER(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TStrArray;
begin
  if PopStringValues(Args, NumArgs, data, Result) then
    Result := CreateString(uppercase(data[0]));
end;

end.
