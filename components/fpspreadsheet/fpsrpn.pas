{@@ ----------------------------------------------------------------------------
  The unit fpsRPN contains methods for simple creation of an RPNFormula array
  to be used in fpspreadsheet.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpsRPN;

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

interface

uses
  SysUtils, fpstypes;

type
  {@@ Pointer to a TPRNItem record
      @see    TRPNItem }
  PRPNItem = ^TRPNItem;

  {@@ Helper record for simplification of RPN formula creation
      @param  FE     Formula element record stored in the RPN item
      @param  Next   Pointer to the next RPN item of the formula
      @see    TsFormulaElement }
  TRPNItem = record
    FE: TsFormulaElement;
    Next: PRPNItem;
  end;

function CreateRPNFormula(AItem: PRPNItem; AReverse: Boolean = false): TsRPNFormula;
procedure DestroyRPNFormula(AItem: PRPNItem);

function RPNBool(AValue: Boolean;
  ANext: PRPNItem): PRPNItem;
function RPNCellValue(ACellAddress: String;
  ANext: PRPNItem): PRPNItem; overload;
function RPNCellValue(ARow, ACol: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem; overload;
function RPNCellRef(ACellAddress: String;
  ANext: PRPNItem): PRPNItem; overload;
function RPNCellRef(ARow, ACol: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem; overload;
function RPNCellRange(ACellRangeAddress: String;
  ANext: PRPNItem): PRPNItem; overload;
function RPNCellRange(ARow, ACol, ARow2, ACol2: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem; overload;
function RPNCellOffset(ARowOffset, AColOffset: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
function RPNErr(AErrCode: TsErrorValue; ANext: PRPNItem): PRPNItem;
function RPNInteger(AValue: Word; ANext: PRPNItem): PRPNItem;
function RPNMissingArg(ANext: PRPNItem): PRPNItem;
function RPNNumber(AValue: Double; ANext: PRPNItem): PRPNItem;
function RPNParenthesis(ANext: PRPNItem): PRPNItem;
function RPNString(AValue: String; ANext: PRPNItem): PRPNItem;
function RPNFunc(AToken: TFEKind; ANext: PRPNItem): PRPNItem; overload;
function RPNFunc(AFuncName: String; ANext: PRPNItem): PRPNItem; overload;
function RPNFunc(AFuncName: String; ANumParams: Byte; ANext: PRPNItem): PRPNItem; overload;


implementation

uses
  fpsStrings, fpsUtils;

{******************************************************************************}
{                   Simplified creation of RPN formulas                        }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Creates a pointer to a new RPN item. This represents an element in the array
  of token of an RPN formula.

  @return  Pointer to the RPN item
-------------------------------------------------------------------------------}
function NewRPNItem: PRPNItem;
begin
  New(Result);
  FillChar(Result^.FE, SizeOf(Result^.FE), 0);
  Result^.FE.StringValue := '';
end;

{@@ ----------------------------------------------------------------------------
  Destroys an RPN item

  @param  AItem  Pointer to the RPN item to be disposed.
-------------------------------------------------------------------------------}
procedure DisposeRPNItem(AItem: PRPNItem);
begin
  if AItem <> nil then
    Dispose(AItem);
end;

{@@ ----------------------------------------------------------------------------
  Creates a boolean value entry in the RPN array.

  @param   AValue   Boolean value to be stored in the RPN item
  @param   ANext    Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNBool(AValue: Boolean; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekBool;
  if AValue then Result^.FE.DoubleValue := 1.0 else Result^.FE.DoubleValue := 0.0;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a cell value, specifed by its
  address, e.g. 'A1'. Takes care of absolute and relative cell addresses.

  @param  ACellAddress   Adress of the cell given in Excel A1 notation
  @param  ANext          Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNCellValue(ACellAddress: String; ANext: PRPNItem): PRPNItem;
var
  r,c: Cardinal;
  flags: TsRelFlags;
begin
  if not ParseCellString(ACellAddress, r, c, flags) then
    raise Exception.CreateFmt(rsNoValidCellAddress, [ACellAddress]);
  Result := RPNCellValue(r,c, flags, ANext);
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a cell value, specifed by its
  row and column index and a flag containing information on relative addresses.

  @param  ARow     Row index of the cell
  @param  ACol     Column index of the cell
  @param  AFlags   Flags specifying absolute or relative cell addresses
  @param  ANext    Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNCellValue(ARow, ACol: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCell;
  Result^.FE.Row := ARow;
  Result^.FE.Col := ACol;
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a cell reference, specifed by its
  address, e.g. 'A1'. Takes care of absolute and relative cell addresses.
  "Cell reference" means that all properties of the cell can be handled.
  Note that most Excel formulas with cells require the cell value only
  (--> RPNCellValue)

  @param  ACellAddress   Adress of the cell given in Excel A1 notation
  @param  ANext          Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNCellRef(ACellAddress: String; ANext: PRPNItem): PRPNItem;
var
  r,c: Cardinal;
  flags: TsRelFlags;
begin
  if not ParseCellString(ACellAddress, r, c, flags) then
    raise Exception.CreateFmt(rsNoValidCellAddress, [ACellAddress]);
  Result := RPNCellRef(r,c, flags, ANext);
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a cell reference, specifed by its
  row and column index and flags containing information on relative addresses.
  "Cell reference" means that all properties of the cell can be handled.
  Note that most Excel formulas with cells require the cell value only
  (--> RPNCellValue)

  @param  ARow     Row index of the cell
  @param  ACol     Column index of the cell
  @param  AFlags   Flags specifying absolute or relative cell addresses
  @param  ANext    Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNCellRef(ARow, ACol: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCellRef;
  Result^.FE.Row := ARow;
  Result^.FE.Col := ACol;
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a range of cells, specified by an
  Excel-style address, e.g. A1:G5. As in Excel, use a $ sign to indicate
  absolute addresses.

  @param  ACellRangeAddress   Adress of the cell range given in Excel notation,
                              such as A1:G5
  @param  ANext               Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNCellRange(ACellRangeAddress: String; ANext: PRPNItem): PRPNItem;
var
  r1,c1, r2,c2: Cardinal;
  flags: TsRelFlags;
begin
  if not ParseCellRangeString(ACellRangeAddress, r1,c1, r2,c2, flags) then
    raise Exception.CreateFmt(rsNoValidCellRangeAddress, [ACellRangeAddress]);
  Result := RPNCellRange(r1,c1, r2,c2, flags, ANext);
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a range of cells, specified by the
  row/column indexes of the top/left and bottom/right corners of the block.
  The flags indicate relative indexes.

  @param  ARow     Row index of the top/left cell
  @param  ACol     Column index of the top/left cell
  @param  ARow2    Row index of the bottom/right cell
  @param  ACol2    Column index of the bottom/right cell
  @param  AFlags   Flags specifying absolute or relative cell addresses
  @param  ANext    Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNCellRange(ARow, ACol, ARow2, ACol2: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCellRange;
  Result^.FE.Row := ARow;
  Result^.FE.Col := ACol;
  Result^.FE.Row2 := ARow2;
  Result^.FE.Col2 := ACol2;
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a relative cell reference as used in
  shared formulas. The given parameters indicate the relativ offset between
  the current cell coordinates and a reference rell.

  @param  ARowOffset  Offset between current row and the row of a reference cell
  @param  AColOffset  Offset between current column and the column of a reference cell
  @param  AFlags      Flags specifying absolute or relative cell addresses
  @param  ANext       Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNCellOffset(ARowOffset, AColOffset: Integer; AFlags: TsRelFlags;
  ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekCellOffset;
  Result^.FE.Row := Cardinal(ARowOffset);
  Result^.FE.Col := Cardinal(AColOffset);
  Result^.FE.RelFlags := AFlags;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array with an error value.

  @param  AErrCode  Error code to be inserted (see TsErrorValue
  @param  ANext     Pointer to the next RPN item in the list
  @see TsErrorValue
-------------------------------------------------------------------------------}
function RPNErr(AErrCode: TsErrorValue; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekErr;
  Result^.FE.IntValue := ord(AErrCode);
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a 2-byte unsigned integer

  @param  AValue  Integer value to be inserted into the formula
  @param  ANext   Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNInteger(AValue: Word; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekInteger;
  Result^.FE.IntValue := AValue;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a missing argument in of function call.
  Use this in a formula to indicate a missing argument

  @param ANext  Pointer to the next RPN item in the list.
-------------------------------------------------------------------------------}
function RPNMissingArg(ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekMissingArg;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a floating point number.

  @param  AValue  Number value to be inserted into the formula
  @param  ANext   Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNNumber(AValue: Double; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekNum;
  Result^.FE.DoubleValue := AValue;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array which puts the current operator in parenthesis.
  For display purposes only, does not affect calculation.

  @param  ANext   Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNParenthesis(ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekParen;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for a string.

  @param  AValue  String to be inserted into the formula
  @param  ANext   Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNString(AValue: String; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekString;
  Result^.FE.StringValue := AValue;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for an operation specified by its TokenID
  (--> TFEKind). Note that array elements for all needed parameters must have
  been created before.

  @param  AToken  Formula element indicating the function to be executed,
                  see the TFEKind enumeration for possible values.
  @param  ANext   Pointer to the next RPN item in the list

  @see TFEKind
-------------------------------------------------------------------------------}
function RPNFunc(AToken: TFEKind; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := AToken;
  Result^.Fe.FuncName := '';
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for an Excel function or operation
  specified by its name. Note that array elements for all needed parameters
  must have been created before.

  @param  AFuncName  Name of the spreadsheet function (as used by Excel)
  @param  ANext      Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNFunc(AFuncName: String; ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(AFuncName, 255, ANext);
end;

{@@ ----------------------------------------------------------------------------
  Creates an entry in the RPN array for an Excel spreadsheet function
  specified by its name. Specify the number of parameters used.
  They must have been created before.

  @param  AFuncName  Name of the spreadsheet function (as used by Excel).
  @param  ANumParams Number of arguments used in the formula.
  @param  ANext      Pointer to the next RPN item in the list
-------------------------------------------------------------------------------}
function RPNFunc(AFuncName: String; ANumParams: Byte; ANext: PRPNItem): PRPNItem;
begin
  Result := NewRPNItem;
  Result^.FE.ElementKind := fekFunc;
  Result^.Fe.FuncName := AFuncName;
  Result^.FE.ParamsNum := ANumParams;
  Result^.Next := ANext;
end;

{@@ ----------------------------------------------------------------------------
  Creates an RPN formula by a single call using nested RPN items.

  For each formula element, use one of the RPNxxxx functions implemented here.
  They are designed to be nested into each other. Terminate the chain by using nil.

  @param  AItem     Pointer to the first RPN item representing the formula.
                    Each item contains a pointer to the next item in the list.
                    The list is terminated by nil.
  @param  AReverse  If true the first rpn item in the chained list becomes the
                    last item in the token array. This feature is needed for
                    reading an xls file.

  @example
    The RPN formula for the string expression "$A1+2" can be created as follows:
    <pre>
      var
        f: TsRPNFormula;
      begin
        f := CreateRPNFormula(
          RPNCellValue('$A1',
          RPNNumber(2,
          RPNFunc(fekAdd,
          nil))));
    </pre>
-------------------------------------------------------------------------------}
function CreateRPNFormula(AItem: PRPNItem; AReverse: Boolean = false): TsRPNFormula;
var
  item: PRPNItem;
  nextitem: PRPNItem;
  n: Integer;
begin
  // Determine count of RPN elements
  n := 0;
  item := AItem;
  while item <> nil do begin
    inc(n);
    item := item^.Next;
  end;

  // Set array length of TsRPNFormula result
  SetLength(Result, n);

  // Copy FormulaElements to result and free temporary RPNItems
  item := AItem;
  if AReverse then n := Length(Result)-1 else n := 0;
  while item <> nil do begin
    nextitem := item^.Next;
    Result[n] := item^.FE;
    Result[n].StringValue := item^.FE.StringValue;
    if AReverse then dec(n) else inc(n);
    DisposeRPNItem(item);
    item := nextitem;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Destroys the RPN formula starting with the given RPN item.

  @param  AItem  Pointer to the first RPN items representing the formula.
                 Each item contains a pointer to the next item in the list.
                 The list is terminated by nil.
-------------------------------------------------------------------------------}
procedure DestroyRPNFormula(AItem: PRPNItem);
var
  nextitem: PRPNItem;
begin
  while AItem <> nil do begin
    nextitem := AItem^.Next;
    DisposeRPNItem(AItem);
    AItem := nextitem;
  end;
end;

end.

