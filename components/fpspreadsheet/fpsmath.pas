unit fpsmath;

{$mode objfpc}

interface

uses
  Classes, SysUtils, fpspreadsheet;

type
  TsArgumentType = (atNumber, atString, atBool, atError);

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

procedure FixMissingBool  (var Arg: TsArgument; ABool: Boolean);
procedure FixMissingNumber(var Arg: TsArgument; ANumber: Double);
procedure FixMissingString(var Arg: TsArgument; AString: String);

function CreateBool(AValue: Boolean): TsArgument;
function CreateNumber(AValue: Double): TsArgument;
function CreateString(AValue: String): TsArgument;
function CreateError(AError: TsErrorValue): TsArgument;

{
These are the functions called when calculating an RPN formula.
}
type
  TsFormulaFunc = function(Args: TsArgumentStack): TsArgument;

function fpsAdd      (Args: TsArgumentStack): TsArgument;
function fpsSub      (Args: TsArgumentStack): TsArgument;
function fpsMul      (Args: TsArgumentStack): TsArgument;
function fpsDiv      (Args: TsArgumentStack): TsArgument;
function fpsPercent  (Args: TsArgumentStack): TsArgument;
function fpsPower    (Args: TsArgumentStack): TsArgument;
function fpsUMinus   (Args: TsArgumentStack): TsArgument;
function fpsUPlus    (Args: TsArgumentStack): TsArgument;
function fpsConcat   (Args: TsArgumentStack): TsArgument;
function fpsEqual    (Args: TsArgumentStack): TsArgument;
function fpsGreater  (Args: TsArgumentStack): TsArgument;
function fpsGreaterEqual(Args: TsArgumentStack): TsArgument;
function fpsLess     (Args: TsArgumentStack): TsArgument;
function fpsLessEqual(Args: TsArgumentStack): TsArgument;
function fpsNotEqual (Args: TsArgumentStack): TsArgument;

function fpsAnd      (Args: TsArgumentStack): TsArgument;

implementation

uses
  Math;


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


{ Missing arguments }

{@@
  Replaces a missing boolean argument by the passed boolean value
  @param  Arg    Argument to be considered
  @param  ABool  Replacement for the missing value
}
procedure FixMissingBool(var Arg: TsArgument; ABool: Boolean);
begin
  if Arg.IsMissing then Arg.BoolValue := ABool;
end;

{@@
  Replaces a missing number argument by the passed number value
  @param  Arg      Argument to be considered
  @param  ANumber  Replacement for the missing value
}
procedure FixMissingNumber(var Arg: TsArgument; ANumber: Double);
begin
  if Arg.IsMissing then Arg.NumberValue := ANumber;
end;

{@@
  Replaces a missing string argument by the passed string value
  @param  Arg      Argument to be considered
  @param  AString  Replacement for the missing value
}
procedure FixMissingString(var Arg: TsArgument; AString: String);
begin
  if Arg.IsMissing then Arg.StringValue := AString;
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
end;

function CreateNumber(AValue: Double): TsArgument;
begin
  Result.ArgumentType := atNumber;
  Result.NumberValue := AValue;
end;

function CreateString(AValue: String): TsArgument;
begin
  Result.ArgumentType := atString;
  Result.StringValue := AValue;
end;

function CreateError(AError: TsErrorValue): TsArgument;
begin
  Result.ArgumentType := atError;
  Result.ErrorValue := AError;
end;


function Pop_1Bool(Args: TsArgumentStack; out a: Boolean): TsErrorValue;
begin
  Result := GetBoolFromArgument(Args.Pop, a);
end;

function Pop_1Float(Args: TsArgumentStack; out a: Double): TsErrorValue;
begin
  Result := GetNumberFromArgument(Args.Pop, a);
end;

function Pop_1String(Args: TsArgumentStack; out a: String): TsErrorvalue;
begin
  Result := GetStringFromArgument(Args.Pop, a);
end;

function Pop_2Bools(Args: TsArgumentStack; out a, b: Boolean): TsErrorValue;
var
  erra, errb: TsErrorValue;
begin
  // Pop the data in reverse order they were pushed! Otherwise they will be
  // applied to the function in the wrong order.
  errb := GetBoolFromArgument(Args.Pop, b);
  erra := GetBoolFromArgument(Args.Pop, a);
  if erra <> errOK then
    Result := erra
  else if errb <> errOK then
    Result := errb
  else
    Result := errOK;
end;

function Pop_2Floats(Args: TsArgumentStack; out a, b: Double): TsErrorValue;
var
  erra, errb: TsErrorValue;
begin
  // Pop the data in reverse order they were pushed! Otherwise they will be
  // applied to the function in the wrong order.
  errb := GetNumberFromArgument(Args.Pop, b);
  erra := GetNumberFromArgument(Args.Pop, a);
  if erra <> errOK then
    Result := erra
  else if errb <> errOK then
    Result := errb
  else
    Result := errOK;
end;

function Pop_2Strings(Args: TsArgumentStack; out a, b: String): TsErrorValue;
var
  erra, errb: TsErrorValue;
begin
  // Pop the data in reverse order they were pushed! Otherwise they will be
  // applied to the function in the wrong order.
  errb := GetStringFromArgument(Args.Pop, b);
  erra := GetStringFromArgument(Args.Pop, a);
  if erra <> errOK then
    Result := erra
  else if errb <> errOK then
    Result := errb
  else
    Result := errOK;
end;


{ Operations }

function fpsAdd(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  err: TsErrorValue;
begin
  err := Pop_2Floats(Args, a, b);
  if err = errOK then
    Result := CreateNumber(a + b)
  else
    Result := CreateError(err);
end;

function fpsSub(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  err: TsErrorValue;
begin
  err := Pop_2Floats(Args, a, b);
  if err = errOK then
    Result := CreateNumber(a - b)
  else
    Result := CreateError(err);
end;

function fpsMul(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  err: TsErrorValue;
begin
  err := Pop_2Floats(Args, a, b);
  if err = errOK then
    Result := CreateNumber(a * b)
  else
    Result := CreateError(err);
end;

function fpsDiv(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  err: TsErrorValue;
begin
  err := Pop_2Floats(Args, a, b);
  if err <> errOK then
    Result := CreateError(err)
  else if b = 0 then
    Result := CreateError(errDivideByZero)
  else
    Result := CreateNumber(a / b);
end;

function fpsPercent(Args: TsArgumentStack): TsArgument;
var
  a: Double;
  err: TsErrorValue;
begin
  err := Pop_1Float(Args, a);
  if err = errOK then
    Result := CreateNumber(a * 0.01)
  else
    Result := CreateError(err);
end;

function fpsPower(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  err: TsErrorValue;
begin
  err := Pop_2Floats(Args, a, b);
  if err = errOK then begin
    try
      Result := CreateNumber(power(a, b));
    except on E: EInvalidArgument do
      Result := CreateError(errOverflow);
      // this could happen, e.g., for "power( (neg value), (non-integer) )"
    end;
  end else
    Result := CreateError(err);
end;

function fpsUMinus(Args: TsArgumentStack): TsArgument;
var
  a: Double;
  err: TsErrorValue;
begin
  err := Pop_1Float(Args, a);
  if err = errOK then
    Result := CreateNumber(-a)
  else
    Result := CreateError(err);
end;

function fpsUPlus(Args: TsArgumentStack): TsArgument;
var
  a: Double;
  err: TsErrorValue;
begin
  err := Pop_1Float(Args, a);
  if err = errOK then
    Result := CreateNumber(a)
  else
    Result := CreateError(err);
end;

function fpsConcat(Args: TsArgumentStack): TsArgument;
var
  a, b: String;
  err: TsErrorValue;
begin
  err := Pop_2Strings(Args, a, b);
  if err = errOK then
    Result := CreateString(a + b)
  else
    Result := CreateError(err);
end;

function fpsEqual(Args: TsArgumentStack): TsArgument;
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

function fpsGreater(Args: TsArgumentStack): TsArgument;
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

function fpsGreaterEqual(Args: TsArgumentStack): TsArgument;
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

function fpsLess(Args: TsArgumentStack): TsArgument;
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

function fpsLessEqual(Args: TsArgumentStack): TsArgument;
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


function fpsNotEqual(Args: TsArgumentStack): TsArgument;
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




// !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//   Variable parameter count !!!!!!!!!!!!
function fpsAnd(Args: TsArgumentStack): TsArgument;
var
  a, b: Boolean;
  err: TsErrorValue;
begin
  err := Pop_2Bools(Args, a, b);
  if err = errOK then
    Result := CreateBool(a and b)
  else
    Result := CreateError(err);
end;

end.
