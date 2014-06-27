unit fpsmath;

{$mode objfpc}

interface

uses
  Classes, SysUtils, fpspreadsheet;

type
  TsArgumentType = (atNumber, atString, atBool, atError);
  TsArgumentError = (aeOK, aeWrongType, aeDivideByZero, aeFuncNotDefined);

  TsArgument = record
    IsMissing: Boolean;
    case ArgumentType: TsArgumentType of
      atNumber  : (NumberValue: Double);
      atString  : (StringValue: String);
      atBool    : (BoolValue: Boolean);
      atError   : (ErrorValue: TsArgumentError);
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

procedure CheckMissingBool  (var Arg: TsArgument; ABool: Boolean);
procedure CheckMissingNumber(var Arg: TsArgument; ANumber: Double);
procedure CheckMissingString(var Arg: TsArgument; AString: String);

type
  TsFormulaFunc = function(Args: TsArgumentStack): TsArgument;

function fpsAdd(Args: TsArgumentStack): TsArgument;
function fpsSub(Args: TsArgumentStack): TsArgument;
function fpsMul(Args: TsArgumentStack): TsArgument;
function fpsDiv(Args: TsArgumentStack): TsArgument;

implementation


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
procedure CheckMissingBool(var Arg: TsArgument; ABool: Boolean);
begin
  if Arg.IsMissing then Arg.BoolValue := ABool;
end;

{@@
  Replaces a missing number argument by the passed number value
  @param  Arg      Argument to be considered
  @param  ANumber  Replacement for the missing value
}
procedure CheckMissingNumber(var Arg: TsArgument; ANumber: Double);
begin
  if Arg.IsMissing then Arg.NumberValue := ANumber;
end;

{@@
  Replaces a missing string argument by the passed string value
  @param  Arg      Argument to be considered
  @param  AString  Replacement for the missing value
}
procedure CheckMissingString(var Arg: TsArgument; AString: String);
begin
  if Arg.IsMissing then Arg.StringValue := AString;
end;


{ Preparing arguments }

function GetNumberFromArgument(Arg: TsArgument; var ANumber: Double): TsArgumentError;
begin
  Result := aeOK;
  case Arg.ArgumentType of
    atNumber : ANumber := Arg.NumberValue;
    atString : if not TryStrToFloat(arg.StringValue, ANumber) then Result := aeWrongType;
    atBool   : if Arg.BoolValue then ANumber := 1.0 else ANumber := 0.0;
  end;
end;

function CreateNumber(AValue: Double): TsArgument;
begin
  Result.ArgumentType := atNumber;
  Result.NumberValue := AValue;
end;


function CreateError(AError: TsArgumentError): TsArgument;
begin
  Result.ArgumentType := atError;
  Result.ErrorValue := AError;
end;


{ Operations }

function fpsAdd(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  erra, errb: TsArgumentError;
begin
  errb := GetNumberFromArgument(Args.Pop, b);
  erra := GetNumberFromArgument(Args.Pop, a);
  if erra <> aeOK then
    Result := CreateError(erra)
  else if errb <> aeOK then
    Result := CreateError(errb)
  else
    Result := CreateNumber(a + b);
end;

function fpsSub(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  erra, errb: TsArgumentError;
begin
  // Pop the data in reverse order they were pushed!
  errb := GetNumberFromArgument(Args.Pop, b);
  erra := GetNumberFromArgument(Args.Pop, a);
  if erra <> aeOK then
    Result := CreateError(erra)
  else if errb <> aeOK then
    Result := CreateError(errb)
  else
    Result := CreateNumber(a - b);
end;

function fpsMul(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  erra, errb: TsArgumentError;
begin
  errb := GetNumberFromArgument(Args.Pop, b);
  erra := GetNumberFromArgument(Args.Pop, a);
  if erra <> aeOK then
    Result := CreateError(erra)
  else if errb <> aeOK then
    Result := CreateError(errb)
  else
    Result := CreateNumber(a * b);
end;

function fpsDiv(Args: TsArgumentStack): TsArgument;
var
  a, b: Double;
  erra, errb: TsArgumentError;
begin
  // Pop the data in reverse order they were pushed!
  errb := GetNumberFromArgument(Args.Pop, b);
  erra := GetNumberFromArgument(Args.Pop, a);
  if erra <> aeOK then
    Result := CreateError(erra)
  else if errb <> aeOK then
    Result := CreateError(errb)
  else if b = 0 then
    Result := CreateError(aeDivideByZero)
  else
    Result := CreateNumber(a / b);
end;


end.
