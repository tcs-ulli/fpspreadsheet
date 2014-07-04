unit fpsfunc;

{$mode objfpc}

interface

uses
  Classes, SysUtils, fpspreadsheet;

type
  TsArgumentType = (atCell, atCellRange, atNumber, atString,
    atBool, atError, atEmpty);

  TsArgBoolArray  = array of boolean;
  TsArgNumberArray = array of double;
  TsArgStringArray   = array of string;

  TsArgument = record
    IsMissing: Boolean;
    Worksheet: TsWorksheet;
    case ArgumentType: TsArgumentType of
      atCell      : (Cell: PCell);
      atCellRange : (FirstRow,FirstCol,LastRow,LastCol: Cardinal);
      atNumber    : (NumberValue: Double);
      atString    : (StringValue: String);
      atBool      : (BoolValue: Boolean);
      atError     : (ErrorValue: TsErrorValue);
  end;
  PsArgument = ^TsArgument;

  TsArgumentStack = class(TFPList)
  protected
    function PopMultiple(ACount: Integer): TsArgumentStack;
  public
    destructor Destroy; override;
    function Pop: TsArgument;
    function PopNumber(out AValue: Double; out AErrArg: TsArgument): Boolean;
    function PopNumberValues(ANumArgs: Integer; ARangeAllowed: Boolean;
      out AValues: TsArgNumberArray; out AErrArg: TsArgument;
      AErrorOnNoNumber: Boolean = true): Boolean;
    function PopString(out AValue: String; out AErrArg: TsArgument): Boolean;
    function PopStringValues(ANumArgs: Integer; ARangeAllowed:Boolean;
      out AValues: TsArgStringArray; out AErrArg: TsArgument): Boolean;
    procedure Push(AValue: TsArgument; AWorksheet: TsWorksheet);
    procedure PushBool(AValue: Boolean; AWorksheet: TsWorksheet);
    procedure PushCell(AValue: PCell; AWorksheet: TsWorksheet);
    procedure PushCellRange(AFirstRow, AFirstCol, ALastRow, ALastCol: Cardinal;
      AWorksheet: TsWorksheet);
    procedure PushMissing(AWorksheet: TsWorksheet);
    procedure PushNumber(AValue: Double; AWorksheet: TsWorksheet);
    procedure PushString(AValue: String; AWorksheet: TsWorksheet);
    procedure Clear;
    procedure Delete(AIndex: Integer);
  end;

function CreateBool(AValue: Boolean): TsArgument;
function CreateCell(AValue: PCell): TsArgument;
function CreateCellRange(AFirstRow, AFirstCol, ALastRow, ALastCol: Cardinal): TsArgument;
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
function fpsCOUNTBLANK  (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsCOUNTIF     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMAX         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsMIN         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsPRODUCT     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSTDEV       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSTDEVP      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSUM         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsSUMIF       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
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
{ lookup / reference }
function fpsCOLUMN      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsCOLUMNS     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsROW         (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsROWS        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ info functions }
function fpsCELLINFO    (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsINFO        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISBLANK     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISERR       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISERROR     (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISLOGICAL   (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISNA        (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISNONTEXT   (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISNUMBER    (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISREF       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsISTEXT      (Args: TsArgumentStack; NumArgs: Integer): TsArgument;
function fpsVALUE       (Args: TsArgumentStack; NumArgs: Integer): TsArgument;


implementation

uses
  Math, lazutf8, StrUtils, DateUtils, fpsUtils;


{ Helpers }

function CreateArgument: TsArgument;
begin
  FillChar(Result, SizeOf(Result), 0);
end;

function CreateBool(AValue: Boolean): TsArgument;
begin
  Result := CreateArgument;
  Result.ArgumentType := atBool;
  Result.Boolvalue := AValue;
end;

function CreateCell(AValue: PCell): TsArgument;
begin
  Result := CreateArgument;
  Result.ArgumentType := atCell;
  Result.Cell := AValue;
end;

function CreateCellRange(AFirstRow, AFirstCol, ALastRow, ALastCol: Cardinal): TsArgument;
begin
  Result := CreateArgument;
  Result.ArgumentType := atCellRange;
  Result.FirstRow := AFirstRow;
  Result.FirstCol := AFirstCol;
  Result.LastRow := ALastRow;
  Result.LastCol := ALastCol;
end;

function CreateNumber(AValue: Double): TsArgument;
begin
  Result := CreateArgument;
  Result.ArgumentType := atNumber;
  Result.NumberValue := AValue;
end;

function CreateString(AValue: String): TsArgument;
begin
  Result := CreateArgument;
  Result.ArgumentType := atString;
  Result.StringValue := AValue;
end;

function CreateError(AError: TsErrorValue): TsArgument;
begin
  Result := CreateArgument;
  Result.ArgumentType := atError;
  Result.ErrorValue := AError;
end;

function CreateEmpty: TsArgument;
begin
  Result := CreateArgument;
  Result.ArgumentType := atEmpty;
end;

{ Compares two arguments and returns -1 if "Arg2 > Arg1", +1 if "Arg1 < Arg2",
  0 if "Arg1 = Arg2", MaxInt if result meaningless
  If AExact is true only matching types are compared, otherwise types are converted before comparing. }
function CompareArgs(Arg1, Arg2: TsArgument; AExact: Boolean): integer;
var
  val1, val2: Double;
  b1, b2: Boolean;
  cell1, cell2: PCell;
  s: String;
begin
  Result := MaxInt;

  // Number - Number
  if (Arg1.ArgumentType = atNumber) and (Arg2.ArgumentType = atNumber) then begin
    Result := CompareValue(Arg1.NumberValue, Arg2.NumberValue);
    exit;
  end;

  // String - String
  if (Arg1.ArgumentType = atString) and (Arg2.ArgumentType = atString) then begin
    if TryStrToFloat(Arg1.StringValue, val1) and TryStrToFloat(Arg2.StringValue, val2) then
      Result := CompareValue(val1, val2)
    else
      Result := UTF8CompareText(Arg1.StringValue, Arg2.StringValue);
    exit;
  end;

  // Bool - Bool
  if (Arg1.ArgumentType = atBool) and (Arg2.ArgumentType = atBool) then begin
    Result := CompareValue(ord(Arg1.BoolValue), ord(Arg2.BoolValue));
    exit;
  end;

  // Cell - Cell
  if (Arg1.ArgumentType in [atCell, atCellRange]) and (Arg2.ArgumentType in [atCell, atCellRange])
  then begin
    if Arg1.ArgumentType = atCell
      then cell1 := Arg1.Cell
      else cell1 := Arg1.Worksheet.FindCell(Arg1.FirstRow, Arg1.FirstCol);
    if Arg2.ArgumentType = atCell
      then cell2 := Arg2.Cell
      else cell2 := Arg2.Worksheet.FindCell(Arg2.FirstRow, Arg2.FirstCol);
    if Arg1.Worksheet.ReadNumericValue(cell1, val1) and Arg2.Worksheet.ReadNumericValue(cell2, val2) then begin
      Result := CompareValue(val1, val2);
      exit;
    end;
    Result := UTF8CompareText(cell1^.UTF8StringValue, cell2^.UTF8StringValue);
    exit;
  end;

  // Mixed type comparison only if AExact = true
  if AExact then
    exit;

  // Number - string
  if (Arg1.ArgumentType = atNumber) and (Arg2.ArgumentType = atString) then begin
    if TryStrToFloat(Arg2.StringValue, val2) then
      Result := CompareValue(Arg1.NumberValue, val2);
    exit;
  end;
  if (Arg1.ArgumentType = atString) and (Arg2.ArgumentType = atNumber) then begin
    if TryStrToFloat(Arg1.StringValue, val1) then
      Result := CompareValue(val1, Arg2.NumberValue);
    exit;
  end;

  // Number - bool
  if (Arg1.ArgumentType = atNumber) and (Arg2.ArgumentType = atBool) then begin
    Result := CompareValue(Arg1.NumberValue, ord(Arg2.BoolValue));
    exit;
  end;
  if (Arg1.ArgumentType = atBool) and (Arg2.ArgumentType = atNumber) then begin
    Result := CompareValue(ord(Arg1.BoolValue), Arg2.NumberValue);
    exit;
  end;

  // Number - cell
  if (Arg1.ArgumentType = atNumber) and (Arg2.ArgumentType in [atCell, atCellRange]) then begin
    if Arg2.ArgumentType = atCell
      then cell2 := Arg2.Cell
      else cell2 := Arg2.Worksheet.FindCell(Arg2.FirstRow, Arg2.FirstCol);
    if (cell2 <> nil) and Arg2.Worksheet.ReadNumericValue(cell2, val2) then
      Result := CompareValue(Arg1.NumberValue, val2);
    exit;
  end;
  if (Arg2.ArgumentType = atNumber) and (Arg1.ArgumentType in [atCell, atCellRange]) then begin
    Result := CompareArgs(Arg2, Arg1, AExact);
    if Result <> MaxInt then Result := -Result;
    exit;
  end;

  // String - bool
  if (Arg1.ArgumentType = atString) and (Arg2.ArgumentType = atBool) then begin
    if not TryStrToFloat(Arg1.StringValue, val1) then
      exit;
    val2 := ord(Arg2.BoolValue);
    Result := CompareValue(val1, val2);
    exit;
  end;
  if (Arg2.ArgumentType = atString) and (Arg1.ArgumentType = atBool) then begin
    Result := CompareArgs(Arg2, Arg1, AExact);
    if Result <> MaxInt then Result := -Result;
  end;

  // String - cell
  if (Arg1.ArgumentType = atString) and (Arg2.ArgumentType in [atCell, atCellRange]) then begin
    if Arg2.ArgumentType = atCell
      then cell2 := Arg2.Cell
      else cell2 := Arg2.Worksheet.FindCell(Arg2.FirstRow, Arg2.FirstCol);
    if cell2 = nil then
      exit;
    if TryStrToFloat(Arg1.stringValue, val1) then begin
      if Arg2.Worksheet.ReadNumericValue(cell2, val2) then
        Result := CompareValue(val1, val2);
      exit;
    end;
    Result := UTF8CompareText(Arg1.StringValue, cell2^.UTF8StringValue);
    exit;
  end;
  if (Arg2.ArgumentType = atString) and (Arg1.ArgumentType in [atCell, atCellRange]) then begin
    Result := CompareArgs(Arg2, Arg1, AExact);
    if Result <> MaxInt then Result := -Result;
    exit;
  end;

  // Bool - cell
  if (Arg1.ArgumentType = atBool) and (Arg2.ArgumentType in [atCell, atCellRange]) then begin
    val1 := ord(Arg1.BoolValue);
    if Arg2.ArgumentType = atCell
      then cell2 := Arg2.Cell
      else cell2 := Arg2.Worksheet.FindCell(Arg2.FirstRow, Arg2.FirstCol);
    if (cell2 <> nil) and Arg2.Worksheet.ReadNumericValue(cell2, val2) then
      Result := CompareValue(val1, val2);
    exit;
  end;
  if (Arg2.ArgumentType = atBool) and (Arg1.ArgumentType in [atCell, atCellRange]) then begin
    Result := CompareArgs(Arg2, Arg1, AExact);
    if Result <> MaxInt then Result := -Result;
    exit;
  end;
end;


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
  arg: PsArgument;
begin
  if Count = 0 then
    Result := CreateError(errArgError)
  else begin
    arg := PsArgument(Items[Count-1]);
    Result := arg^;
    Result.StringValue := arg^.StringValue;  // necessary?
    Result.Cell := arg^.Cell;
    Delete(Count-1);
  end;
end;

{ Pops ACount arguments from the stack and pushes them onto an intermediate
  stack. After popping the arguments from that stack, the arguments are in
  the correct order! }
function TsArgumentStack.PopMultiple(ACount: Integer): TsArgumentStack;
var
  arg: TsArgument;
  counter: Integer;
begin
  Result := TsArgumentStack.Create;
  for counter := 1 to ACount do begin
    arg := Pop;
    Result.Push(arg, arg.Worksheet);
  end;
end;

{@@
  Pops an argument from the stack and assumes that it is a number. Returns the
  number value if it is. Otherwise report an error in AErrArg and return false.
  In case of a cell range, only the left/top cell is considered.
}
function TsArgumentStack.PopNumber(out AValue: Double;
  out AErrArg: TsArgument): Boolean;
var
  arg: TsArgument;
  cell: PCell;
begin
  Result := true;
  arg := Pop;
  if arg.isMissing then
    AValue := NaN
  else
    case arg.ArgumentType of
      atNumber:
        AValue := arg.NumberValue;
      atCell, atCellRange:
        begin
          if arg.ArgumentType = atCell then
            cell := arg.Cell
          else   // In case of cell range, consider only top/left cell.
            cell := arg.Worksheet.FindCell(arg.FirstRow, arg.FirstCol);
          if cell = nil then begin
            Result := false;
            AErrArg := CreateError(errWrongType);
          end else
            case cell^.ContentType of
              cctNumber  : AValue := cell^.NumberValue;
              cctDateTime: AValue := cell^.DateTimeValue;
              else         begin
                             Result := false;
                             if cell^.ContentType = cctError then
                               AErrArg := CreateError(cell^.ErrorValue)
                             else
                               AErrArg := CreateError(errWrongType);
                           end;
            end;
        end;
      else
        begin
          Result := false;
          if arg.ArgumentType = atError then
            AErrArg := CreateError(arg.ErrorValue)
          else
            AErrArg := CreateError(errWrongType);
        end;
    end;  // case
  if Result then
    AErrArg := CreateError(errOK);
end;

{@@
  Pops a given number of arguments from the stack and returns an array with
  their number values. In case of a cell range, a value of each contained cell
  is extracted. The numbers are in the same order as they were pushed onto the
  stack.
  If not all argument types correspond to number arguments the function returns
  false and reports the error in the ErrArg parameter. }
function TsArgumentStack.PopNumberValues(ANumArgs: Integer; ARangeAllowed:Boolean;
  out AValues: TsArgNumberArray; out AErrArg: TsArgument;
  AErrorOnNoNumber: Boolean = true): Boolean;

  procedure AddNumber(ANumber: Double);
  begin
    SetLength(AValues, Length(AValues) + 1);
    AValues[Length(AValues)-1] := ANumber;
  end;

  function AddCellNumber(ACell: PCell): Boolean;
  begin
    Result := true;
    case ACell^.ContentType of
      cctNumber:
        AddNumber(ACell^.NumberValue);
      cctDateTime:
        AddNumber(ACell^.DateTimeValue);
      cctBool:
        AddNumber(IfThen(ACell^.BoolValue, 1.0, 0.0));
      cctError:
        if AErrorOnNoNumber then begin
          result := false;
          AErrArg := CreateError(ACell^.ErrorValue);
        end;
    end;
  end;

var
  arg: TsArgument;
  r,c: Cardinal;
  cell: PCell;
  ok: Boolean;
  stack: TsArgumentStack;
begin
  Result := true;
  SetLength(AValues, 0);
  stack := PopMultiple(ANumArgs);
  try
    while stack.Count > 0 do begin
      arg := stack.Pop;
      if arg.IsMissing then
        AddNumber(NaN)
      else
        case arg.ArgumentType of
          atNumber:
            AddNumber(arg.NumberValue);
          atBool:
            AddNumber(IfThen(arg.BoolValue, 1.0, 0.0));
          atCell:
            if arg.Cell <> nil then begin
              ok := AddCellNumber(arg.Cell);
              if not ok then Result := false;
            end;
          atCellRange:
            if ARangeAllowed then begin
              if arg.Worksheet <> nil then
                for r := arg.FirstRow to arg.LastRow do
                  for c := arg.FirstCol to arg.LastCol do begin
                    cell := arg.Worksheet.FindCell(r, c);
                    if cell <> nil then begin
                      ok := AddCellNumber(cell);
                      if not ok then Result := false;
                    end;
                 end;
            end else begin
              result := false;
              AErrArg := CreateError(errWrongType);
            end;
          atString:
            if AErrorOnNoNumber then begin
              result := false;
              AErrArg := CreateError(errWrongType);
            end;
          atError:
            begin
              result := false;
              AErrArg := CreateError(arg.ErrorValue);
            end;
        end;  // case
    end;  // while
    if Result then
      AErrArg := CreateError(errOK)
    else
      SetLength(AValues, 0);
  finally
    stack.Free;
  end;
end;

{@@
  Pops an argument from the stack and assumes that it is a string. Returns the
  text if it is. Otherwise report an error in AErrArg and return false.
  In case of a cell range, only the left/top cell is considered.
}
function TsArgumentStack.PopString(out AValue: String;
  out AErrArg: TsArgument): Boolean;
var
  arg: TsArgument;
  cell: PCell;
begin
  Result := true;
  AValue := '';
  arg := Pop;
  if not arg.isMissing then
    case arg.ArgumentType of
      atString:
        AValue := arg.StringValue;
      atCell, atCellRange:
        begin
          if arg.ArgumentType = atCell then
            cell := arg.Cell
          else    // In case of cell range, consider only top/left cell.
            cell := arg.Worksheet.FindCell(arg.FirstRow, arg.FirstCol);
          if (cell <> nil) and (cell^.ContentType = cctUTF8String) then
            AValue := cell^.UTF8StringValue
          else begin
            Result := false;
            AErrArg := CreateError(errWrongType);
          end;
        end;
      else
        begin
          if arg.ArgumentType = atError then
            AErrArg := CreateError(arg.ErrorValue)
          else
            AErrArg := CreateError(errWrongType);
          Result := false;
        end;
    end;  // case
  if Result then
    AErrArg := CreateError(errOK);
end;

{@@
  Pops a given count of arguments from the stack and returns an array with
  their string values. In case of a cell range, a value of each contained cell
  is extracted. The strings are in the same order as they were pushed onto the
  stack.
  If not all argument types correspond to string arguments the function returns
  false and reports the error in the ErrArg parameter. }
function TsArgumentStack.PopStringValues(ANumArgs: Integer; ARangeAllowed:Boolean;
  out AValues: TsArgStringArray; out AErrArg: TsArgument): Boolean;

  procedure AddString(AString: String);
  begin
    SetLength(AValues, Length(AValues) + 1);
    AValues[Length(AValues)-1] := AString;
  end;

  function AddCellString(ACell: PCell): Boolean;
  begin
    Result := true;
    case ACell^.ContentType of
      cctUTF8String:
        AddString(ACell^.UTF8StringValue);
      cctError:
        begin
          result := false;
          AErrArg := CreateError(ACell^.ErrorValue);
        end;
      else
        Result := false;
        AErrArg := CreateError(errWrongType);
    end;
  end;

var
  arg: TsArgument;
  r,c: Cardinal;
  cell: PCell;
  ok: Boolean;
  stack: TsArgumentStack;
begin
  Result := true;
  SetLength(AValues, 0);
  stack := PopMultiple(ANumArgs);
  try
    while stack.Count > 0 do begin
      arg := stack.Pop;
      if arg.IsMissing then
        AddString('')
      else
        case arg.ArgumentType of
          atString:
            AddString(arg.StringValue);

          atCell, atCellRange:
            if (arg.ArgumentType = atCellRange) and ARangeAllowed then begin
              if (arg.Worksheet <> nil) then begin
                for r := arg.FirstRow to arg.LastRow do
                  for c := arg.FirstCol to arg.LastCol do begin
                    cell := arg.Worksheet.FindCell(r, c);
                    if cell <> nil then begin
                      ok := AddCellString(cell);
                      if not ok then Result := false;
                    end;
                 end;
              end else begin
                result := false;
                AErrArg := CreateError(errWrongType);
              end;
            end else begin
              cell := nil;
              if arg.ArgumentType = atCell then
                cell := arg.Cell
              else if arg.Worksheet <> nil then
                cell := arg.Worksheet.FindCell(arg.FirstRow, arg.FirstCol);
              if cell <> nil then begin
                ok := AddCellString(arg.Cell);
                if not ok then Result := false;
              end;
            end;

          else
            begin
              Result := false;
              if arg.ArgumentTYpe = atError then
                AErrArg := CreateError(arg.ErrorValue)
              else
                AErrArg := CreateError(errWrongType);
            end;
        end;  // case
    end;  // while

    if Result then
      AErrArg := CreateError(errOK)
    else
      SetLength(AValues, 0);
  finally
    stack.Free;
  end;
end;

procedure TsArgumentStack.Push(AValue: TsArgument; AWorksheet: TsWorksheet);
var
  arg: PsArgument;
begin
  GetMem(arg, SizeOf(TsArgument));
  arg^ := AValue;
  arg^.StringValue := AValue.StringValue;
  arg^.Cell := AValue.Cell;
  arg^.Worksheet := AWorksheet;
  Add(arg);
end;

procedure TsArgumentStack.PushBool(AValue: Boolean; AWorksheet: TsWorksheet);
begin
  Push(CreateBool(AValue), AWorksheet);
end;

procedure TsArgumentStack.PushCell(AValue: PCell; AWorksheet: TsWorksheet);
begin
  Push(CreateCell(AValue), AWorksheet);
end;

procedure TsArgumentStack.PushCellRange(AFirstRow, AFirstCol, ALastRow, ALastCol: Cardinal;
  AWorksheet: TsWorksheet);
begin
  Push(CreateCellRange(AFirstRow, AFirstCol, ALastRow, ALastCol), AWorksheet);
end;

procedure TsArgumentStack.PushMissing(AWorksheet: TsWorksheet);
var
  arg: TsArgument;
begin
  arg := CreateArgument;
  arg.IsMissing := true;
  Push(arg, AWorksheet);
end;

procedure TsArgumentStack.PushNumber(AValue: Double; AWorksheet: TsWorksheet);
begin
  Push(CreateNumber(AValue), AWorksheet);
end;

procedure TsArgumentStack.PushString(AValue: String; AWorksheet: TsWorksheet);
begin
  Push(CreateString(AValue), AWorksheet);
end;


{ Preparing arguments }

function GetBoolFromArgument(Arg: TsArgument; var AValue: Boolean): TsErrorValue;
begin
  Result := errOK;
  case Arg.ArgumentType of
    atBool : AValue := Arg.BoolValue;
    atCell : if (Arg.Cell <> nil) and (Arg.Cell^.ContentType = cctBool)
               then AValue := Arg.Cell^.BoolValue
               else Result := errWrongType;
    atError: Result := Arg.ErrorValue;
    else     Result := errWrongType;
  end;
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
  out AValues: TsArgBoolArray; out AErrArg: TsArgument): Boolean;
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
      end;
    atCell:
      if (arg.Cell <> nil) then begin
        Result := true;
        case arg.Cell^.ContentType of
          cctDateTime: ADate := arg.Cell^.DateTimeValue;
          cctNumber  : ADate := arg.Cell^.NumberValue;
          else         Result := false;
                       AErrArg := CreateError(errWrongType);
        end;
      end;
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
      end;
    atCell:
      if (arg.Cell <> nil) then begin
        Result := true;
        case arg.Cell^.ContentType of
          cctDateTime: ATime := frac(arg.Cell^.DateTimeValue);
          cctNumber  : ATime := frac(arg.Cell^.NumberValue);
          else         Result := false;
                       AErrArg := CreateError(errWrongType);
        end;
      end;
  end;
end;


{ Operations }

function fpsAdd(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(2, false, data, Result) then
    Result := CreateNumber(data[0] + data[1]);
end;

function fpsSub(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(2, false, data, Result) then
    Result := CreateNumber(data[0] - data[1]);
end;

function fpsMul(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(2, false, data, Result) then
    Result := CreateNumber(data[0] * data[1]);
end;

function fpsDiv(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(2, false, data, Result) then begin
    if data[1] = 0 then
      Result := CreateError(errDivideByZero)
    else
      Result := CreateNumber(data[0] / data[1]);
  end;
end;

function fpsPercent(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(data[0] * 0.01);
end;

function fpsPower(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(2, false, data, Result) then
    try
      Result := CreateNumber(power(data[0], data[1]));
    except on E: EInvalidArgument do
      Result := CreateError(errOverflow);
      // this could happen, e.g., for "power( (neg value), (non-integer) )"
    end;
end;

function fpsUMinus(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(-data[0]);
end;

function fpsUPlus(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(data[0]);
end;

function fpsConcat(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgStringArray;
begin
  if Args.PopStringValues(2, false, data, Result) then
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
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(abs(data[0]));
end;

function fpsACOS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if InRange(data[0], -1, +1) then
      Result := CreateNumber(arccos(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsACOSH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if data[0] >= 1 then
      Result := CreateNumber(arccosh(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsASIN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if InRange(data[0], -1, +1) then
      Result := CreateNumber(arcsin(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsASINH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(arcsinh(data[0]));
end;

function fpsATAN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(arctan(data[0]));
end;

function fpsATANH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if (data[0] > -1) and (data[0] < +1) then
      Result := CreateNumber(arctanh(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsCOS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(cos(data[0]));
end;

function fpsCOSH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(cosh(data[0]));
end;

function fpsDEGREES(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(RadToDeg(data[0]));
end;

function fpsEXP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(exp(data[0]));
end;

function fpsINT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(floor(data[0]));
end;

function fpsLN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if (data[0] > 0) then
      Result := CreateNumber(ln(data[0]))
    else
      Result := CreateError(errOverflow);  // #NUM!
  end;
end;

function fpsLOG(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// LOG( number [, base] )  -  base is 10 if omitted.
var
  data: TsArgNumberArray;
  base: Double;
begin
  base := 10;
  if Args.PopNumberValues(NumArgs, false, data, Result) then begin
    if NumArgs = 2 then begin
      if IsNaN(data[1]) then begin
        Result := CreateError(errOverflow);
        exit;
      end;
      base := data[1];
    end;

    if base < 0 then begin
      Result := CreateError(errOverflow);
      exit;
    end;

    if data[0] > 0 then
      Result := CreateNumber(logn(base, data[0]))
    else
      Result := CreateError(errOverflow);
  end;
end;

function fpsLOG10(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
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
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(degtorad(data[0]))
end;

function fpsRAND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
begin
  Result := CreateNumber(random);
end;

function fpsROUND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(2, false, data, Result) then
    Result := CreateNumber(RoundTo(data[0], round(data[1])))
end;

function fpsSIGN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(sign(data[0]))
end;

function fpsSIN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(sin(data[0]))
end;

function fpsSINH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(sinh(data[0]))
end;

function fpsSQRT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if data[0] >= 0.0 then
      Result := CreateNumber(sqrt(data[0]))
    else
      Result := CreateError(errOverflow);
  end;
end;

function fpsTAN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if frac(data[0] / (pi*0.5)) = 0 then
      Result := CreateError(errOverflow)
    else
      Result := CreateNumber(tan(data[0]))
  end;
end;

function fpsTANH(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then
    Result := CreateNumber(tanh(data[0]))
end;


{ Date/time functions }

function fpsDATE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// DATE( year, month, day )
var
  data: TsArgNumberArray;
  d: TDate;
begin
  if Args.PopNumberValues(3, false, data, Result) then begin
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
  interval: String;
  data: TsArgStringArray;
  start_date, end_date: TDate;
  res1, res2, res3: TsArgument;
begin
  Args.PopString(interval, res1);
  PopDateValue(Args, end_date, res2);;
  PopDateValue(Args, start_date, res3);
  if res1.ErrorValue <> errOK then begin
    Result := CreateError(res1.ErrorValue);
    exit;
  end;
  if res2.ErrorValue <> errOK then begin
    Result := CreateError(res2.ErrorValue);
    exit;
  end;
  if res3.ErrorValue <> errOK then begin
    Result := CreateError(res3.ErrorValue);
    exit;
  end;

  interval := Uppercase(interval);

  if end_date > start_date then
    Result := CreateError(errOverflow)
  else if interval = 'Y' then
    Result := CreateNumber(YearsBetween(end_date, start_date))
  else if interval = 'M' then
    Result := CreateNumber(MonthsBetween(end_date, start_date))
  else if interval = 'D' then
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
  data: TsArgNumberArray;
  t: TTime;
begin
  if Args.PopNumberValues(3, false, data, Result) then begin
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
  data: TsArgNumberArray;
  n: Integer;
begin
  n := 1;
  if NumArgs = 2 then begin
    if Args.PopNumberValues(1, false, data, Result) then
      n := round(data[0])
    else begin
      Args.Pop;
      exit;
    end;
  end;
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
  data: TsArgNumberArray;
  m: Double;
  i: Integer;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then begin
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
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(Mean(data))
end;

function fpsCOUNT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ counts the number of cells that contain numbers as well as the number of
  arguments that contain numbers.
  COUNT( argument1, [argument2, ... argument_n] )
}
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, result, false) then
    Result := CreateNumber(Length(data));
end;

function fpsCOUNTBLANK(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// COUNTBLANK( range )
// counts the number of empty cells in a range.
var
  arg: TsArgument;
  r, c, n: Cardinal;
  cell: PCell;
begin
  arg := Args.Pop;
  case arg.ArgumentType of
    atCell:
      if arg.Cell = nil then Result := CreateNumber(1) else Result := CreateNumber(0);
    atCellRange:
      begin
        n := 0;
        for r := arg.FirstRow to arg.LastRow do
          for c := arg.FirstCol to arg.LastCol do
            if arg.Worksheet.FindCell(r, c) = nil then inc(n);
        Result := CreateNumber(n);
      end;
    else
      Result := CreateError(errWrongType);
  end;
end;

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
  criteria := Args.Pop;
  arg := Args.Pop;
  compare := coEqual;
  case criteria.ArgumentType of
    atCellRange:
      criteria := CreateCell(criteria.Worksheet.FindCell(criteria.FirstRow, criteria.FirstCol));
    atString:
      criteria.Stringvalue := AnalyzeCompareStr(criteria.StringValue, compare);
  end;
  n := 0;
  for r := arg.FirstRow to arg.LastRow do
    for c := arg.FirstCol to arg.LastCol do begin
      cell := arg.Worksheet.FindCell(r, c);
      if cell <> nil then begin
        cellarg := CreateCell(cell);
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
  Result := CreateNumber(n);
end;

function fpsMAX(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// MAX( number1, number2, ... number_n )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(MaxValue(data))
end;

function fpsMIN(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// MIN( number1, number2, ... number_n )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(MinValue(data))
end;

function fpsPRODUCT(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// PRODUCT( number1, number2, ... number_n )
var
  data: TsArgNumberArray;
  i: Integer;
  p: Double;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then begin
    p := 1.0;
    for i:=0 to High(data) do
      p := p * data[i];
    Result := CreateNumber(p);
  end;
end;

function fpsSTDEV(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// STDEV( number1, [number2, ... number_n] )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(StdDev(data))
end;

function fpsSTDEVP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// STDEVP( number1, [number2, ... number_n] )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(PopnStdDev(data))
end;

function fpsSUM(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// SUM( value1, [value2, ... value_n] )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(Sum(data))
end;

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
    Result := CreateError(errArgError);
    exit;
  end;

  compare := coEqual;
  case criteria.ArgumentType of
    atCellRange:
      criteria := CreateCell(criteria.Worksheet.FindCell(criteria.FirstRow, criteria.FirstCol));
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
        cellarg := CreateCell(cell);
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
  Result := CreateNumber(sum);
end;

function fpsSUMSQ(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// SUMSQ( value1, [value2, ... value_n] )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(SumOfSquares(data))
end;

function fpsVAR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// VAR( number1, number2, ... number_n )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(Variance(data))
end;

function fpsVARP(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// VARP( number1, number2, ... number_n )
var
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(NumArgs, true, data, Result) then
    Result := CreateNumber(PopnVariance(data))
end;


{ Logical functions }

function fpsAND(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// AND( condition1, [condition2], ... )
var
  data: TsArgBoolArray;
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
  data: TsArgBoolArray;
begin
  if PopBoolValues(Args, NumArgs, data, Result) then
    Result := CreateBool(not data[0]);
end;

function fpsOR(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// OR( condition1, [condition2], ... )
var
  data: TsArgBoolArray;
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
  data: TsArgNumberArray;
begin
  if Args.PopNumberValues(1, false, data, Result) then begin
    if (data[0] >= 0) and (data[0] <= 255) then
      Result := CreateString(AnsiToUTF8(Char(Round(data[0]))));
  end;
end;

function fpsCODE(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// CODE( text )
var
  s: String;
  ch: Char;
begin
  if Args.PopString(s, Result) then begin
    if s <> '' then begin
      ch := UTF8ToAnsi(s)[1];
      Result := CreateNumber(Ord(ch));
    end else
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
  s: String;
begin
  if Args.PopString(s, Result) then
    Result := CreateString(UTF8LowerCase(s));
end;

function fpsMID(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// MID( text, start_position, number_of_characters )
var
  data: TsArgNumberArray;
  s: String;
  res1, res2: TsArgument;
begin
  Args.PopNumberValues(2, false, data, res1);
  Args.PopString(s, res2);
  if res1.ErrorValue <> errOK then begin
    Result := CreateError(res1.ErrorValue);
    exit;
  end;
  if res2.ErrorValue <> errOK then begin
    Result := CreateError(res2.ErrorValue);
    exit;
  end;
  Result := CreateString(UTF8Copy(s, Round(data[0]), Round(data[1])));
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
  number: Double;
  n: Integer;
  arg: TsArgument;
  data: TsArgStringArray;
  s, s_old, s_new: String;
begin
  Result := CreateError(errWrongType);
  n := -1;
  if (NumArgs = 4) then begin
    if Args.PopNumber(number, Result) then
      n := round(number)
    else begin
      Args.Pop;
      Args.Pop;
      Args.Pop;
      exit;
    end;
  end;

  if Args.PopStringValues(3, false, data, Result) then begin
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
  s: String;
begin
  if Args.PopString(s, Result) then
    Result := CreateString(UTF8Trim(s));
end;

function fpsUPPER(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// UPPER( text )
var
  s: String;
begin
  if Args.PopString(s, Result) then
    Result := CreateString(UTF8UpperCase(s));
end;


{ Lookup / refernence functions }

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
    Result := CreateError(errArgError);
    // We don't know here which cell contains the formula.

  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumber(arg.Cell^.Col + 1);
    atCellRange: Result := CreateNumber(arg.FirstCol + 1);
    else         Result := CreateError(errWrongType);
  end;
end;

function fpsCOLUMNS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ COLUMNS( [reference] )
  returns the number of column in a cell reference. }
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumber(1);
    atCellRange: Result := CreateNumber(arg.LastCol - arg.FirstCol + 1);
    else         Result := CreateError(errWrongType);
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
    Result := CreateError(errArgError);
    // We don't know here which cell contains the formula.

  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumber(arg.Cell^.Row + 1);
    atCellRange: Result := CreateNumber(arg.FirstRow + 1);
    else         Result := CreateError(errWrongType);
  end;
end;


function fpsROWS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ ROWS( [reference] )
  returns the number of rows in a cell reference. }
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumber(1);
    atCellRange: Result := CreateNumber(arg.LastRow - arg.FirstRow + 1);
    else         Result := CreateError(errWrongType);
  end;
end;


{ Info functions }

function fpsCELLINFO(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
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
    "format"      Number format of the cell according to next table:
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
  arg: TsArgument;
  cell: PCell;
  sname: String;
  data: TsArgStringArray;
  res: TsArgument;
begin
  if NumArgs < 2 then begin
    Result := CreateError(errArgError);
    exit;
  end;

  arg := Args.Pop;
  Args.PopString(sname, res);

  if (arg.ArgumentType = atCellRange) then
    cell := arg.Worksheet.FindCell(arg.FirstRow, arg.FirstCol)
  else
  if (arg.ArgumentType = atCell) then
    cell := arg.Cell
  else begin
    Result := CreateError(errArgError);
    exit;
  end;

  if (cell = nil) then begin
    Result := CreateError(errArgError);
    exit;
  end;

  if (res.ErrorValue <> errOK) then begin
    Result := CreateError(res.ErrorValue);
    exit;
  end;

  sname := Lowercase(sname);

  if sname = 'address' then
    Result := CreateString(GetCellString(cell^.Row, cell^.Col, []))
  else if sname = 'col' then
    Result := CreateNumber(cell^.Col + 1)
  else if sname = 'color' then begin
    if (cell^.NumberFormat = nfCurrencyRed)
      then Result := CreateNumber(1)
      else Result := CreateNumber(0);
  end else if sname = 'contents' then
    case cell^.ContentType of
      cctNumber     : Result := CreateNumber(cell^.NumberValue);
      cctDateTime   : Result := CreateNumber(cell^.DateTimeValue);
      cctUTF8String : Result := CreateString(cell^.UTF8StringValue);
      cctBool       : Result := CreateString(BoolToStr(cell^.BoolValue));
      cctError      : Result := CreateString('Error');
    end
  else if sname = 'format' then begin
    Result := CreateString('');
    case cell^.NumberFormat of
      nfGeneral:
        Result := CreateString('G');
      nfFixed:
        if cell^.NumberFormatStr= '0' then Result := CreateString('0') else
        if cell^.NumberFormatStr = '0.00' then  Result := CreateString('F0');
      nfFixedTh:
        if cell^.NumberFormatStr = '#,##0' then Result := CreateString(',0') else
        if cell^.NumberFormatStr = '#,##0.00' then Result := CreateString(',2');
      nfPercentage:
        if cell^.NumberFormatStr = '0%' then Result := CreateString('P0') else
        if cell^.NumberFormatStr = '0.00%' then Result := CreateString('P2');
      nfExp:
        if cell^.NumberFormatStr = '0.00E+00' then Result := CreateString('S2');
      nfShortDate, nfLongDate, nfShortDateTime:
        Result := CreateString('D4');
      nfLongTimeAM:
        Result := CreateString('D6');
      nfShortTimeAM:
        Result := CreateString('D7');
      nfLongTime:
        Result := CreateString('D8');
      nfShortTime:
        Result := CreateString('D9');
    end;
  end else
  if (sname = 'prefix') then begin
    Result := CreateString('');
    if (cell^.ContentType = cctUTF8String) then
      case cell^.HorAlignment of
        haLeft  : Result := CreateString('''');
        haCenter: Result := CreateString('^');
        haRight : Result := CreateString('"');
      end;
  end else
  if sname = 'row' then
    Result := CreateNumber(cell^.Row + 1)
  else if sname = 'type' then begin
    if (cell^.ContentType = cctEmpty) then
      Result := CreateString('b')
    else if cell^.ContentType = cctUTF8String then begin
      if (cell^.UTF8StringValue = '')
        then Result := CreateString('b')
        else Result := CreateString('l');
    end else
      Result := CreateString('v');
  end;
end;

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
  arg := Args.Pop;
  if arg.ArgumentType <> atString then
    Result := CreateError(errWrongType)
  else begin
    s := Lowercase(arg.StringValue);
    workbook := arg.Worksheet.Workbook;
    if s = 'directory' then
      Result := CreateString(ExtractFilePath(workbook.FileName))
    else
    if s = 'numfile' then
      Result := CreateNumber(workbook.GetWorksheetCount)
    else
      Result := CreateError(errFormulaNotSupported);
  end;
end;

function fpsISBLANK(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISBLANK( value )
// Checks for blank cell
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool((arg.ArgumentType = atCell) and
    ((arg.Cell = nil) or (arg.Cell^.ContentType = cctEmpty))
  );
end;

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

function fpsISREF(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
// ISREF( value )
var
  arg: TsArgument;
begin
  arg := Args.Pop;
  Result := CreateBool(arg.ArgumentType in [atCell, atCellRange]);
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
  s: String;
  x: Double;
begin
  if Args.PopString(s, Result) then
    if TryStrToFloat(s, x) then
      Result := CreateNumber(x)
    else
      Result := CreateError(errWrongType);
end;


end.
