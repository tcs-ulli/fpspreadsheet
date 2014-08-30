{
    This file is part of the Free Component Library (FCL)
    Copyright (c) 2008 Michael Van Canneyt.

    Expression parser, supports variables, functions and
    float/integer/string/boolean/datetime operations.

    See the file COPYING.FPC, included in this distribution,
    for details about the copyright.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

--------------------------------------------------------------------------------

    Modified for integration into fpspreadsheet by Werner Pamler:
    - Original file name: fpexprpars.pp
    - Rename identifiers to avoid naming conflicts with the original
    - TsExpressionParser and TsBuiltinExpressionManager are not components
      any more
    - TsExpressionParser is created with the worksheet as a parameter.
    - add new TExprNode classes:
      - TsCellExprNode for references to cells
      - TsCellRangeExprNode for references to cell ranges
      - TsPercentExprNode and token "%" to handle Excel's percent operation
      - TsParenthesisExprNode to handle the parenthesis token in RPN formulas
      - TsConcatExprNode and token "&" to handle string concatenation
      - TsUPlusExprNode for unary plus symbol
    - remove and modifiy built-in function such that the parser is compatible
      with Excel syntax (and Open/LibreOffice - which is the same).
    - use double quotes for strings (instead of single quotes)
    - add boolean constants "TRUE" and "FALSE".
    - add property RPNFormula to interface the parser to RPN formulas of xls files.
    - accept funtions with zero parameters

 ******************************************************************************}

// To do:
// Remove exceptions, use error message strings instead
// Cell reference not working (--> formula CELL!)
// Missing arguments
// Keep spaces in formula

{$mode objfpc}
{$h+}
unit fpsExprParser;

interface

uses
  Classes, SysUtils, contnrs, fpspreadsheet;

type
  { Tokens }

(*  { Basic operands }
  fekCell, fekCellRef, fekCellRange, fekCellOffset, fekNum, fekInteger,
  fekString, fekBool, fekErr, fekMissingArg,
  { Basic operations }
  fekAdd, fekSub, fekMul, fekDiv, fekPercent, fekPower, fekUMinus, fekUPlus,
  fekConcat,  // string concatenation
  fekEqual, fekGreater, fekGreaterEqual, fekLess, fekLessEqual, fekNotEqual,
  fekParen,
*)
  TsTokenType = (
    ttCell, ttCellRange, ttNumber, ttString, ttIdentifier,
    ttPlus, ttMinus, ttMul, ttDiv, ttConcat, ttPercent, ttPower, ttLeft, ttRight,
    ttLessThan, ttLargerThan, ttEqual, ttNotEqual, ttLessThanEqual, ttLargerThanEqual,
    ttComma, ttTrue, ttFalse, ttEOF
  );

  TsExprFloat = Double;
  TsExprFloatArray = array of TsExprFloat;

const
  ttDelimiters = [
    ttPlus, ttMinus, ttMul, ttDiv, ttLeft, ttRight, ttLessThan, ttLargerThan,
    ttEqual, ttNotEqual, ttLessThanEqual, ttLargerThanEqual
  ];

  ttComparisons = [
    ttLargerThan, ttLessThan, ttLargerThanEqual, ttLessThanEqual, ttEqual, ttNotEqual
  ];

type
  TsExpressionParser = class;
  TsBuiltInExpressionManager = class;

  TsResultType = (rtEmpty, rtBoolean, rtInteger, rtFloat, rtDateTime, rtString,
    rtCell, rtCellRange, rtError, rtAny);
  TsResultTypes = set of TsResultType;

  TsExpressionResult = record
    Worksheet       : TsWorksheet;
    ResString       : String;
    case ResultType : TsResultType of
      rtEmpty       : ();
      rtError       : (ResError       : TsErrorValue);
      rtBoolean     : (ResBoolean     : Boolean);
      rtInteger     : (ResInteger     : Int64);
      rtFloat       : (ResFloat       : TsExprFloat);
      rtDateTime    : (ResDateTime    : TDatetime);
      rtCell        : (ResRow, ResCol : Cardinal);
      rtCellRange   : (ResCellRange   : TsCellRange);
      rtString      : ();
  end;
  PsExpressionResult = ^TsExpressionResult;
  TsExprParameterArray = array of TsExpressionResult;

  { TsExprNode }
  TsExprNode = class(TObject)
  protected
    procedure CheckNodeType(ANode: TsExprNode; Allowed: TsResultTypes);
    // A procedure with var saves an implicit try/finally in each node
    // A marked difference in execution speed.
    procedure GetNodeValue(var Result: TsExpressionResult); virtual; abstract;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; virtual; abstract;
    function AsString: string; virtual; abstract;
    procedure Check; virtual; abstract;
    function NodeType: TsResultType; virtual; abstract;
    function NodeValue: TsExpressionResult;
  end;

  TsExprArgumentArray = array of TsExprNode;

  { TsBinaryOperationExprNode }
  TsBinaryOperationExprNode = class(TsExprNode)
  private
    FLeft: TsExprNode;
    FRight: TsExprNode;
  protected
    procedure CheckSameNodeTypes; virtual;
  public
    constructor Create(ALeft, ARight: TsExprNode);
    destructor Destroy; override;
    procedure Check; override;
    property Left: TsExprNode read FLeft;
    property Right: TsExprNode read FRight;
  end;
  TsBinaryOperationExprNodeClass = class of TsBinaryOperationExprNode;

  { TsBooleanOperationExprNode }
  TsBooleanOperationExprNode = class(TsBinaryOperationExprNode)
  public
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsBooleanResultExprNode }
  TsBooleanResultExprNode = class(TsBinaryOperationExprNode)
  protected
    procedure CheckSameNodeTypes; override;
  public
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;
  TsBooleanResultExprNodeClass = class of TsBooleanResultExprNode;

  { TsEqualExprNode }
  TsEqualExprNode = class(TsBooleanResultExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsNotEqualExprNode }
  TsNotEqualExprNode = class(TsEqualExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsOrderingExprNode }
  TsOrderingExprNode = class(TsBooleanResultExprNode)
  protected
    procedure CheckSameNodeTypes; override;
  public
    procedure Check; override;
  end;

  { TsLessExprNode }
  TsLessExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterExprNode }
  TsGreaterExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsLessEqualExprNode }
  TsLessEqualExprNode = class(TsGreaterExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterEqualExprNode }
  TsGreaterEqualExprNode = class(TsLessExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsConcatExprNode }
  TsConcatExprNode = class(TsBinaryOperationExprNode)
  protected
    procedure CheckSameNodeTypes; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsMathOperationExprNode }
  TsMathOperationExprNode = class(TsBinaryOperationExprNode)
  protected
    procedure CheckSameNodeTypes; override;
  public
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsAddExprNode }
  TsAddExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsSubtractExprNode }
  TsSubtractExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsMultiplyExprNode }
  TsMultiplyExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsDivideExprNode }
  TsDivideExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    function NodeType: TsResultType; override;
  end;

  { TsPowerExprNode }
  TsPowerExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    function NodeType: TsResultType; override;
  end;

  { TsUnaryOperationExprNode }
  TsUnaryOperationExprNode = class(TsExprNode)
  private
    FOperand: TsExprNode;
  protected
    procedure Check; override;
  public
    constructor Create(AOperand: TsExprNode);
    destructor Destroy; override;
    property Operand: TsExprNode read FOperand;
  end;

  { TsConvertExprNode }
  TsConvertExprNode = class(TsUnaryOperationExprNode)
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
  end;

  { TsNotExprNode }
  TsNotExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    function NodeType: TsResultType; override;
  end;

  { TsConvertToIntExprNode }
  TsConvertToIntExprNode = class(TsConvertExprNode)
  public
    procedure Check; override;
  end;

  { TsIntToFloatExprNode }
  TsIntToFloatExprNode = class(TsConvertToIntExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function NodeType: TsResultType; override;
  end;

  { TsIntToDateTimeExprNode }
  TsIntToDateTimeExprNode = class(TsConvertToIntExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function NodeType: TsResultType; override;
  end;

  { TsFloatToDateTimeExprNode }
  TsFloatToDateTimeExprNode = class(TsConvertExprNode)
  protected
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function NodeType: TsResultType; override;
  end;

  { TsUPlusExprNode }
  TsUPlusExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsUMinusExprNode }
  TsUMinusExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsPercentExprNode }
  TsPercentExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsParenthesisExprNode }
  TsParenthesisExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    function NodeType: TsResultType; override;
  end;

  { TsConstExprNode }
  TsConstExprNode = class(TsExprNode)
  private
    FValue: TsExpressionResult;
  protected
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    constructor CreateString(AValue: String);
    constructor CreateInteger(AValue: Int64);
    constructor CreateDateTime(AValue: TDateTime);
    constructor CreateFloat(AValue: TsExprFloat);
    constructor CreateBoolean(AValue: Boolean);
    constructor CreateError(AValue: TsErrorValue);
    function AsString: string; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function NodeType : TsResultType; override;
    // For inspection
    property ConstValue: TsExpressionResult read FValue;
  end;

  TsExprIdentifierType = (itVariable, itFunctionCallBack, itFunctionHandler);

  TsExprFunctionCallBack = procedure (var Result: TsExpressionResult;
    const Args: TsExprParameterArray);

  TsExprFunctionEvent = procedure (var Result: TsExpressionResult;
    const Args: TsExprParameterArray) of object;

  { TsExprIdentifierDef }
  TsExprIdentifierDef = class(TCollectionItem)
  private
    FStringValue: String;
    FValue: TsExpressionResult;
    FArgumentTypes: String;
    FIDType: TsExprIdentifierType;
    FName: ShortString;
    FExcelCode: Integer;
    FVariableArgumentCount: Boolean;
    FOnGetValue: TsExprFunctionEvent;
    FOnGetValueCB: TsExprFunctionCallBack;
    function GetAsBoolean: Boolean;
    function GetAsDateTime: TDateTime;
    function GetAsFloat: TsExprFloat;
    function GetAsInteger: Int64;
    function GetAsString: String;
    function GetResultType: TsResultType;
    function GetValue: String;
    procedure SetArgumentTypes(const AValue: String);
    procedure SetAsBoolean(const AValue: Boolean);
    procedure SetAsDateTime(const AValue: TDateTime);
    procedure SetAsFloat(const AValue: TsExprFloat);
    procedure SetAsInteger(const AValue: Int64);
    procedure SetAsString(const AValue: String);
    procedure SetName(const AValue: ShortString);
    procedure SetResultType(const AValue: TsResultType);
    procedure SetValue(const AValue: String);
  protected
    procedure CheckResultType(const AType: TsResultType);
    procedure CheckVariable;
  public
    function ArgumentCount: Integer;
    procedure Assign(Source: TPersistent); override;
    property AsFloat: TsExprFloat Read GetAsFloat Write SetAsFloat;
    property AsInteger: Int64 Read GetAsInteger Write SetAsInteger;
    property AsString: String Read GetAsString Write SetAsString;
    property AsBoolean: Boolean Read GetAsBoolean Write SetAsBoolean;
    property AsDateTime: TDateTime Read GetAsDateTime Write SetAsDateTime;
    function HasFixedArgumentCount: Boolean;
    function IsOptionalArgument(AIndex: Integer): Boolean;
    property OnGetFunctionValueCallBack: TsExprFunctionCallBack read FOnGetValueCB write FOnGetValueCB;
  published
    property IdentifierType: TsExprIdentifierType read FIDType write FIDType;
    property Name: ShortString read FName write SetName;
    property Value: String read GetValue write SetValue;
    property ParameterTypes: String read FArgumentTypes write SetArgumentTypes;
    property ResultType: TsResultType read GetResultType write SetResultType;
    property ExcelCode: Integer read FExcelCode write FExcelCode;
    property VariableArgumentCount: Boolean read FVariableArgumentCount write FVariableArgumentCount;
    property OnGetFunctionValue: TsExprFunctionEvent read FOnGetValue write FOnGetValue;
  end;

  TsBuiltInExprCategory = (bcMath, bcStatistics, bcStrings, bcLogical, bcDateTime,
    bcLookup, bcInfo, bcUser);

  TsBuiltInExprCategories = set of TsBuiltInExprCategory;

  { TsBuiltInExprIdentifierDef }
  TsBuiltInExprIdentifierDef = class(TsExprIdentifierDef)
  private
    FCategory: TsBuiltInExprCategory;
  public
    procedure Assign(Source: TPersistent); override;
  published
    property Category: TsBuiltInExprCategory read FCategory write FCategory;
  end;

  { TsExprIdentifierDefs }
  TsExprIdentifierDefs = class(TCollection)
  private
    FParser: TsExpressionParser;
    function GetI(AIndex: Integer): TsExprIdentifierDef;
    procedure SetI(AIndex: Integer; const AValue: TsExprIdentifierDef);
  protected
    procedure Update(Item: TCollectionItem); override;
    property Parser: TsExpressionParser read FParser;
  public
    function FindIdentifier(const AName: ShortString): TsExprIdentifierDef;
    function IdentifierByExcelCode(const AExcelCode: Integer): TsExprIdentifierDef;
    function IdentifierByName(const AName: ShortString): TsExprIdentifierDef;
    function IndexOfIdentifier(const AName: ShortString): Integer; overload;
    function IndexOfIdentifier(const AExcelCode: Integer): Integer; overload;
    function AddVariable(const AName: ShortString; AResultType: TsResultType;
      AValue: String): TsExprIdentifierDef;
    function AddBooleanVariable(const AName: ShortString;
      AValue: Boolean): TsExprIdentifierDef;
    function AddIntegerVariable(const AName: ShortString;
      AValue: Integer): TsExprIdentifierDef;
    function AddFloatVariable(const AName: ShortString;
      AValue: TsExprFloat): TsExprIdentifierDef;
    function AddStringVariable(const AName: ShortString;
      AValue: String): TsExprIdentifierDef;
    function AddDateTimeVariable(const AName: ShortString;
      AValue: TDateTime): TsExprIdentifierDef;
    function AddFunction(const AName: ShortString; const AResultType: Char;
      const AParamTypes: String; const AExcelCode: Integer;
      ACallBack: TsExprFunctionCallBack): TsExprIdentifierDef;
    function AddFunction(const AName: ShortString; const AResultType: Char;
      const AParamTypes: String; const AExcelCode: Integer;
      ACallBack: TsExprFunctionEvent): TsExprIdentifierDef;
    property Identifiers[AIndex: Integer]: TsExprIdentifierDef read GetI write SetI; default;
  end;

  { TsIdentifierExprNode }
  TsIdentifierExprNode = class(TsExprNode)
  private
    FID: TsExprIdentifierDef;
    PResult: PsExpressionResult;
    FResultType: TsResultType;
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    constructor CreateIdentifier(AID: TsExprIdentifierDef);
    function NodeType: TsResultType; override;
    property Identifier: TsExprIdentifierDef read FID;
  end;

  { TsVariableExprNode }
  TsVariableExprNode = class(TsIdentifierExprNode)
    procedure Check; override;
    function AsString: string; override;
    Function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
  end;

  { TsFunctionExprNode }
  TsFunctionExprNode = class(TsIdentifierExprNode)
  private
    FArgumentNodes: TsExprArgumentArray;
    FargumentParams: TsExprParameterArray;
  protected
    procedure CalcParams;
  public
    constructor CreateFunction(AID: TsExprIdentifierDef;
      const Args: TsExprArgumentArray); virtual;
    destructor Destroy; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    property ArgumentNodes: TsExprArgumentArray read FArgumentNodes;
    property ArgumentParams: TsExprParameterArray read FArgumentParams;
  end;

  { TsFunctionCallBackExprNode }
  TsFunctionCallBackExprNode = class(TsFunctionExprNode)
  private
    FCallBack: TsExprFunctionCallBack;
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    constructor CreateFunction(AID: TsExprIdentifierDef;
      const Args: TsExprArgumentArray); override;
    property CallBack: TsExprFunctionCallBack read FCallBack;
  end;

  { TFPFunctionEventHandlerExprNode }
  TFPFunctionEventHandlerExprNode = class(TsFunctionExprNode)
  private
    FCallBack: TsExprFunctionEvent;
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    constructor CreateFunction(AID: TsExprIdentifierDef;
      const Args: TsExprArgumentArray); override;
    property CallBack: TsExprFunctionEvent read FCallBack;
  end;

  { TsCellExprNode }
  TsCellExprNode = class(TsExprNode)
  private
    FWorksheet: TsWorksheet;
    FRow, FCol: Cardinal;
    FFlags: TsRelFlags;
    FCell: PCell;
    FIsRef: Boolean;
  protected
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    constructor Create(AWorksheet: TsWorksheet; ACellString: String); overload;
    constructor Create(AWorksheet: TsWorksheet; ARow, ACol: Cardinal;
      AFlags: TsRelFlags); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
    function NodeType: TsResultType; override;
    property Worksheet: TsWorksheet read FWorksheet;
  end;

  { TsCellRangeExprNode }
  TsCellRangeExprNode = class(TsExprNode)
  private
    FWorksheet: TsWorksheet;
    FRow1, FRow2: Cardinal;
    FCol1, FCol2: Cardinal;
    FFlags: TsRelFlags;
  protected
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    constructor Create(AWorksheet: TsWorksheet; ACellRangeString: String); overload;
    constructor Create(AWorksheet: TsWorksheet; ARow1,ACol1, ARow2,ACol2: Cardinal;
      AFlags: TsRelFlags); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    function NodeType: TsResultType; override;
    property Worksheet: TsWorksheet read FWorksheet;
  end;

  { TsExpressionScanner }
  TsExpressionScanner = class(TObject)
    FSource : String;
    LSource,
    FPos: Integer;
    FChar: PChar;
    FToken: String;
    FTokenType: TsTokenType;
  private
    function GetCurrentChar: Char;
    procedure ScanError(Msg: String);
  protected
    procedure SetSource(const AValue: String); virtual;
    function DoIdentifier: TsTokenType;
    function DoNumber: TsTokenType;
    function DoDelimiter: TsTokenType;
    function DoString: TsTokenType;
    function NextPos: Char; // inline;
    procedure SkipWhiteSpace; // inline;
    function IsWordDelim(C: Char): Boolean; // inline;
    function IsDelim(C: Char): Boolean; // inline;
    function IsDigit(C: Char): Boolean; // inline;
    function IsAlpha(C: Char): Boolean; // inline;
  public
    constructor Create;
    function GetToken: TsTokenType;
    property Token: String read FToken;
    property TokenType: TsTokenType read FTokenType;
    property Source: String read FSource write SetSource;
    property Pos: Integer read FPos;
    property CurrentChar: Char read GetCurrentChar;
  end;

  EExprScanner = class(Exception);

  { TsExpressionParser }
  TsExpressionParser = class
  private
    FBuiltIns: TsBuiltInExprCategories;
    FExpression: String;
    FScanner: TsExpressionScanner;
    FExprNode: TsExprNode;
    FIdentifiers: TsExprIdentifierDefs;
    FHashList: TFPHashObjectlist;
    FDirty: Boolean;
    FWorksheet: TsWorksheet;
    procedure CheckEOF;
    procedure CheckNodes(var ALeft, ARight: TsExprNode);
    function ConvertNode(Todo: TsExprNode; ToType: TsResultType): TsExprNode;
    function GetAsBoolean: Boolean;
    function GetAsDateTime: TDateTime;
    function GetAsFloat: TsExprFloat;
    function GetAsInteger: Int64;
    function GetAsString: String;
    function GetRPNFormula: TsRPNFormula;
    function MatchNodes(Todo, Match: TsExprNode): TsExprNode;
    procedure SetBuiltIns(const AValue: TsBuiltInExprCategories);
    procedure SetIdentifiers(const AValue: TsExprIdentifierDefs);
    procedure SetRPNFormula(const AFormula: TsRPNFormula);

  protected
    class function BuiltinExpressionManager: TsBuiltInExpressionManager;
    procedure ParserError(Msg: String);
    procedure SetExpression(const AValue: String); virtual;
    procedure CheckResultType(const Res: TsExpressionResult;
      AType: TsResultType); inline;
    function CurrentToken: String;
    function GetToken: TsTokenType;
    function Level1: TsExprNode;
    function Level2: TsExprNode;
    function Level3: TsExprNode;
    function Level4: TsExprNode;
    function Level5: TsExprNode;
    function Level6: TsExprNode;
    function Primitive: TsExprNode;
    function TokenType: TsTokenType;
    procedure CreateHashList;
    property Scanner: TsExpressionScanner read FScanner;
    property ExprNode: TsExprNode read FExprNode;
    property Dirty: Boolean read FDirty;

  public
    constructor Create(AWorksheet: TsWorksheet); virtual;
    destructor Destroy; override;
    function IdentifierByName(AName: ShortString): TsExprIdentifierDef; virtual;
    procedure Clear;
    function BuildStringFormula: String;
    function Evaluate: TsExpressionResult;
    procedure EvaluateExpression(var Result: TsExpressionResult);
    function ResultType: TsResultType;
    property AsFloat: TsExprFloat read GetAsFloat;
    property AsInteger: Int64 read GetAsInteger;
    property AsString: String read GetAsString;
    property AsBoolean: Boolean read GetAsBoolean;
    property AsDateTime: TDateTime read GetAsDateTime;
    // The expression to parse
    property Expression: String read FExpression write SetExpression;
    property RPNFormula: TsRPNFormula read GetRPNFormula write SetRPNFormula;
    property Identifiers: TsExprIdentifierDefs read FIdentifiers write SetIdentifiers;
    property BuiltIns: TsBuiltInExprCategories read FBuiltIns write SetBuiltIns;
    property Worksheet: TsWorksheet read FWorksheet;
  end;

  TsSpreadsheetParser = class(TsExpressionParser)
  public
    constructor Create(AWorksheet: TsWorksheet); override;
  end;


  { TsBuiltInExpressionManager }
  TsBuiltInExpressionManager = class(TComponent)
  private
    FDefs: TsExprIdentifierDefs;
    function GetCount: Integer;
    function GetI(AIndex: Integer): TsBuiltInExprIdentifierDef;
  protected
    property Defs: TsExprIdentifierDefs read FDefs;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function IndexOfIdentifier(const AName: ShortString): Integer;
    function FindIdentifier(const AName: ShortString): TsBuiltInExprIdentifierDef;
    function IdentifierByExcelCode(const AExcelCode: Integer): TsBuiltInExprIdentifierDef;
    function IdentifierByName(const AName: ShortString): TsBuiltInExprIdentifierDef;
    function AddVariable(const ACategory: TsBuiltInExprCategory; const AName: ShortString;
      AResultType: TsResultType; AValue: String): TsBuiltInExprIdentifierDef;
    function AddBooleanVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: Boolean): TsBuiltInExprIdentifierDef;
    function AddIntegerVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: Integer): TsBuiltInExprIdentifierDef;
    function AddFloatVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: TsExprFloat): TsBuiltInExprIdentifierDef;
    function AddStringVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: String): TsBuiltInExprIdentifierDef;
    function AddDateTimeVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: TDateTime): TsBuiltInExprIdentifierDef;
    function AddFunction(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; const AResultType: Char; const AParamTypes: String;
      const AExcelCode: Integer; ACallBack: TsExprFunctionCallBack): TsBuiltInExprIdentifierDef;
    function AddFunction(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; const AResultType: Char; const AParamTypes: String;
      const AExcelCode: Integer; ACallBack: TsExprFunctionEvent): TsBuiltInExprIdentifierDef;
    property IdentifierCount: Integer read GetCount;
    property Identifiers[AIndex: Integer]: TsBuiltInExprIdentifierDef read GetI;
  end;

  EExprParser = class(Exception);

function TokenName(AToken: TsTokenType): String;
function ResultTypeName(AResult: TsResultType): String;
function CharToResultType(C: Char): TsResultType;
function BuiltinIdentifiers: TsBuiltInExpressionManager;
procedure RegisterStdBuiltins(AManager: TsBuiltInExpressionManager);
function ArgToBoolean(Arg: TsExpressionResult): Boolean;
function ArgToCell(Arg: TsExpressionResult): PCell;
function ArgToDateTime(Arg: TsExpressionResult): TDateTime;
function ArgToInt(Arg: TsExpressionResult): Integer;
function ArgToFloat(Arg: TsExpressionResult): TsExprFloat;
function ArgToString(Arg: TsExpressionResult): String;
procedure ArgsToFloatArray(const Args: TsExprParameterArray; out AData: TsExprFloatArray);
function BooleanResult(AValue: Boolean): TsExpressionResult;
function DateTimeResult(AValue: TDateTime): TsExpressionResult;
function EmptyResult: TsExpressionResult;
function ErrorResult(const AValue: TsErrorValue): TsExpressionResult;
function FloatResult(const AValue: TsExprFloat): TsExpressionResult;
function IntegerResult(const AValue: Integer): TsExpressionResult;
function StringResult(const AValue: String): TsExpressionResult;

procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallBack: TsExprFunctionCallBack);

const
  AllBuiltIns = [bcMath, bcStatistics, bcStrings, bcLogical, bcDateTime, bcLookup,
    bcInfo, bcUser];

var
  ExprFormatSettings: TFormatSettings;

implementation

uses
  typinfo, math, lazutf8, dateutils, xlsconst, fpsutils;

const
  cNull = #0;
  cDoubleQuote = '"';

  Digits        = ['0'..'9', '.'];
  WhiteSpace    = [' ', #13, #10, #9];
  Operators     = ['+', '-', '<', '>', '=', '/', '*', '&', '%', '^'];
  Delimiters    = Operators + [',', '(', ')'];
  Symbols       = Delimiters;
  WordDelimiters = WhiteSpace + Symbols;

resourcestring
  SBadQuotes = 'Unterminated string';
  SUnknownDelimiter = 'Unknown delimiter character: "%s"';
  SErrUnknownCharacter = 'Unknown character at pos %d: "%s"';
  SErrUnexpectedEndOfExpression = 'Unexpected end of expression';
  SErrUnknownComparison = 'Internal error: Unknown comparison';
  SErrUnknownBooleanOp = 'Internal error: Unknown boolean operation';
  SErrBracketExpected = 'Expected ) bracket at position %d, but got %s';
  SerrUnknownTokenAtPos = 'Unknown token at pos %d : %s';
  SErrLeftBracketExpected = 'Expected ( bracket at position %d, but got %s';
  SErrInvalidFloat = '%s is not a valid floating-point value';
  SErrUnknownIdentifier = 'Unknown identifier: %s';
  SErrInExpression = 'Cannot evaluate: error in expression';
  SErrInExpressionEmpty = 'Cannot evaluate: empty expression';
  SErrCommaExpected =  'Expected comma (,) at position %d, but got %s';
  SErrInvalidNumberChar = 'Unexpected character in number : %s';
  SErrInvalidNumber = 'Invalid numerical value : %s';
  SErrNoOperand = 'No operand for unary operation %s';
  SErrNoLeftOperand = 'No left operand for binary operation %s';
  SErrNoRightOperand = 'No left operand for binary operation %s';
  SErrNoNegation = 'Cannot negate expression of type %s: %s';
  SErrNoUPlus = 'Cannot perform unary plus operation on type %s: %s';
  SErrNoNOTOperation = 'Cannot perform NOT operation on expression of type %s: %s';
  SErrNoPercentOperation = 'Cannot perform percent operation on expression of type %s: %s';
  SErrTypesDoNotMatch = 'Type mismatch: %s<>%s for expressions "%s" and "%s".';
  SErrTypesIncompatible = 'Incompatible types: %s<>%s for expressions "%s" and "%s".';
  SErrNoNodeToCheck = 'Internal error: No node to check !';
  SInvalidNodeType = 'Node type (%s) not in allowed types (%s) for expression: %s';
  SErrUnterminatedExpression = 'Badly terminated expression. Found token at position %d : %s';
  SErrDuplicateIdentifier = 'An identifier with name "%s" already exists.';
  SErrInvalidResultCharacter = '"%s" is not a valid return type indicator';
  ErrInvalidArgumentCount = 'Invalid argument count for function %s';
  SErrInvalidArgumentType = 'Invalid type for argument %d: Expected %s, got %s';
  SErrInvalidResultType = 'Invalid result type: %s';
  SErrNotVariable = 'Identifier %s is not a variable';
  SErrInactive = 'Operation not allowed while an expression is active';
  SErrCircularReference = 'Circular reference found when calculating worksheet formulas';

{ ---------------------------------------------------------------------
  Auxiliary functions
  ---------------------------------------------------------------------}

procedure RaiseParserError(Msg: String);
begin
  raise EExprParser.Create(Msg);
end;

procedure RaiseParserError(Fmt: String; Args: Array of const);
begin
  raise EExprParser.CreateFmt(Fmt, Args);
end;

function TokenName(AToken: TsTokenType): String;
begin
  Result := GetEnumName(TypeInfo(TsTokenType), ord(AToken));
end;

function ResultTypeName(AResult: TsResultType): String;
begin
  Result := GetEnumName(TypeInfo(TsResultType), ord(AResult));
end;

function CharToResultType(C: Char): TsResultType;
begin
  case Upcase(C) of
    'S' : Result := rtString;
    'D' : Result := rtDateTime;
    'B' : Result := rtBoolean;
    'I' : Result := rtInteger;
    'F' : Result := rtFloat;
    'R' : Result := rtCellRange;
    'C' : Result := rtCell;
    '?' : Result := rtAny;
  else
    RaiseParserError(SErrInvalidResultCharacter, [C]);
  end;
end;

var
  BuiltIns: TsBuiltInExpressionManager = nil;

function BuiltinIdentifiers: TsBuiltInExpressionManager;
begin
  If (BuiltIns = nil) then
    BuiltIns := TsBuiltInExpressionManager.Create(nil);
  Result := BuiltIns;
end;

procedure FreeBuiltIns;
begin
  FreeAndNil(Builtins);
end;


{------------------------------------------------------------------------------}
{  TsExpressionScanner                                                        }
{------------------------------------------------------------------------------}

constructor TsExpressionScanner.Create;
begin
  Source := '';
end;

function TsExpressionScanner.DoDelimiter: TsTokenType;
var
  B : Boolean;
  C, D : Char;
begin
  C := FChar^;
  FToken := C;
  B := C in ['<', '>'];
  D := C;
  C := NextPos;

  if B and (C in ['=', '>']) then
  begin
    FToken := FToken + C;
    NextPos;
    If D = '>' then
      Result := ttLargerThanEqual
    else if C = '>' then
      Result := ttNotEqual
    else
      Result := ttLessThanEqual;
  end
  else
    case D of
      '+' : Result := ttPlus;
      '-' : Result := ttMinus;
      '*' : Result := ttMul;
      '/' : Result := ttDiv;
      '^' : Result := ttPower;
      '%' : Result := ttPercent;
      '&' : Result := ttConcat;
      '<' : Result := ttLessThan;
      '>' : Result := ttLargerThan;
      '=' : Result := ttEqual;
      '(' : Result := ttLeft;
      ')' : Result := ttRight;
      ',' : Result := ttComma;
    else
      ScanError(Format(SUnknownDelimiter, [D]));
    end;
end;

function TsExpressionScanner.DoIdentifier: TsTokenType;
var
  C: Char;
  S: String;
  row, row2: Cardinal;
  col, col2: Cardinal;
  flags: TsRelFlags;
begin
  C := CurrentChar;
  while (not IsWordDelim(C)) and (C <> cNull) do
  begin
    FToken := FToken + C;
    C := NextPos;
  end;
  S := LowerCase(Token);
  if ParseCellString(S, row, col, flags) and (C <> '(') then
    Result := ttCell
  else if ParseCellRangeString(S, row, col, row2, col2, flags) and (C <> '(') then
    Result := ttCellRange
  else if (S = 'true') and (C <> '(') then
    Result := ttTrue
  else if (S = 'false') and (C <> '(') then
    Result := ttFalse
  else
    Result := ttIdentifier;
end;

function TsExpressionScanner.DoNumber: TsTokenType;
var
  C: Char;
  X: TsExprFloat;
  prevC: Char;
begin
  C := CurrentChar;
  prevC := #0;
  while (not IsWordDelim(C) or (prevC = 'E')) and (C <> cNull) do
  begin
    if not ( IsDigit(C)
             or ((FToken <> '') and (Upcase(C) = 'E'))
             or ((FToken <> '') and (C in ['+', '-']) and (prevC = 'E'))
           )
    then
      ScanError(Format(SErrInvalidNumberChar, [C]));
    FToken := FToken+C;
    prevC := Upcase(C);
    C := NextPos;
  end;
  if not TryStrToFloat(FToken, X, ExprFormatSettings) then
    ScanError(Format(SErrInvalidNumber, [FToken]));
  Result := ttNumber;
end;

function TsExpressionScanner.DoString: TsTokenType;

  function TerminatingChar(C: Char): boolean;
  begin
    Result := (C = cNull)
          or ((C = cDoubleQuote) and
               not ((FPos < LSource) and (FSource[FPos+1] = cDoubleQuote)));
  end;

var
  C: Char;
begin
  FToken := '';
  C := NextPos;
  while not TerminatingChar(C) do
  begin
    FToken := FToken+C;
    if C = cDoubleQuote then
      NextPos;
    C := NextPos;
  end;
  if (C = cNull) then
    ScanError(SBadQuotes);
  Result := ttString;
  FTokenType := Result;
  NextPos;
end;

function TsExpressionScanner.GetCurrentChar: Char;
begin
  if FChar <> nil then
    Result := FChar^
  else
    Result := #0;
end;

function TsExpressionScanner.GetToken: TsTokenType;
var
  C: Char;
begin
  FToken := '';
  SkipWhiteSpace;
  C := FChar^;
  if c = cNull then
    Result := ttEOF
  else if IsDelim(C) then
    Result := DoDelimiter
  else if (C = cDoubleQuote) then
    Result := DoString
  else if IsDigit(C) then
    Result := DoNumber
  else if IsAlpha(C) or (C = '$') then
    Result := DoIdentifier
  else
    ScanError(Format(SErrUnknownCharacter, [FPos, C]));
  FTokenType := Result;
end;

function TsExpressionScanner.IsAlpha(C: Char): Boolean;
begin
  Result := C in ['A'..'Z', 'a'..'z'];
end;

function TsExpressionScanner.IsDelim(C: Char): Boolean;
begin
  Result := C in Delimiters;
end;

function TsExpressionScanner.IsDigit(C: Char): Boolean;
begin
  Result := C in Digits;
end;

function TsExpressionScanner.IsWordDelim(C: Char): Boolean;
begin
  Result := C in WordDelimiters;
end;

function TsExpressionScanner.NextPos: Char;
begin
  Inc(FPos);
  Inc(FChar);
  Result := FChar^;
end;

procedure TsExpressionScanner.ScanError(Msg: String);
begin
  raise EExprScanner.Create(Msg)
end;

procedure TsExpressionScanner.SetSource(const AValue: String);
begin
  FSource := AValue;
  LSource := Length(FSource);
  FTokenType := ttEOF;
  if LSource = 0 then
    FPos := 0
  else
    FPos := 1;
  FChar := PChar(FSource);
  FToken := '';
end;

procedure TsExpressionScanner.SkipWhiteSpace;
begin
  while (FChar^ in WhiteSpace) and (FPos <= LSource) do
    NextPos;
end;


{------------------------------------------------------------------------------}
{  TsExpressionParser                                                         }
{------------------------------------------------------------------------------}

constructor TsExpressionParser.Create(AWorksheet: TsWorksheet);
begin
  inherited Create;
  FWorksheet := AWorksheet;
  FIdentifiers := TsExprIdentifierDefs.Create(TsExprIdentifierDef);
  FIdentifiers.FParser := Self;
  FScanner := TsExpressionScanner.Create;
  FHashList := TFPHashObjectList.Create(False);
end;

destructor TsExpressionParser.Destroy;
begin
  FreeAndNil(FHashList);
  FreeAndNil(FExprNode);
  FreeAndNil(FIdentifiers);
  FreeAndNil(FScanner);
  inherited Destroy;
end;

function TsExpressionParser.BuildStringFormula: String;
begin
  if FExprNode = nil then
    Result := ''
  else
    Result := FExprNode.AsString;
end;

class function TsExpressionParser.BuiltinExpressionManager: TsBuiltInExpressionManager;
begin
  Result := BuiltinIdentifiers;
end;

procedure TsExpressionParser.CheckEOF;
begin
  if (TokenType = ttEOF) then
    ParserError(SErrUnexpectedEndOfExpression);
end;

{ If the result types differ, they are converted to a common type if possible. }
procedure TsExpressionParser.CheckNodes(var ALeft, ARight: TsExprNode);
begin
  ALeft := MatchNodes(ALeft, ARight);
  ARight := MatchNodes(ARight, ALeft);
end;

procedure TsExpressionParser.CheckResultType(const Res: TsExpressionResult;
  AType: TsResultType); inline;
begin
  if (Res.ResultType <> AType) then
    RaiseParserError(SErrInvalidResultType, [ResultTypeName(Res.ResultType)]);
end;

procedure TsExpressionParser.Clear;
begin
  FExpression := '';
  FHashList.Clear;
  FreeAndNil(FExprNode);
end;

function TsExpressionParser.ConvertNode(ToDo: TsExprNode;
  ToType: TsResultType): TsExprNode;
begin
  Result := ToDo;
  case ToDo.NodeType of
    rtInteger :
      case ToType of
        rtFloat    : Result := TsIntToFloatExprNode.Create(Result);
        rtDateTime : Result := TsIntToDateTimeExprNode.Create(Result);
      end;
    rtFloat :
      case ToType of
        rtDateTime : Result := TsFloatToDateTimeExprNode.Create(Result);
      end;
  end;
end;

procedure TsExpressionParser.CreateHashList;
var
  ID: TsExprIdentifierDef;
  BID: TsBuiltInExprIdentifierDef;
  i: Integer;
  M: TsBuiltInExpressionManager;
begin
  FHashList.Clear;
  // Builtins
  M := BuiltinExpressionManager;
  If (FBuiltins <> []) and Assigned(M) then
    for i:=0 to M.IdentifierCount-1 do
    begin
      BID := M.Identifiers[i];
      If BID.Category in FBuiltins then
        FHashList.Add(UpperCase(BID.Name), BID);
    end;
  // User
  for i:=0 to FIdentifiers.Count-1 do
  begin
    ID := FIdentifiers[i];
    FHashList.Add(UpperCase(ID.Name), ID);
  end;
  FDirty := False;
end;

function TsExpressionParser.CurrentToken: String;
begin
  Result := FScanner.Token;
end;

function TsExpressionParser.Evaluate: TsExpressionResult;
begin
  EvaluateExpression(Result);
end;

procedure TsExpressionParser.EvaluateExpression(var Result: TsExpressionResult);
begin
  if (FExpression = '') then
    ParserError(SErrInExpressionEmpty);
  if not Assigned(FExprNode) then
    ParserError(SErrInExpression);
  FExprNode.GetNodeValue(Result);
end;

function TsExpressionParser.GetAsBoolean: Boolean;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtBoolean);
  Result := Res.ResBoolean;
end;

function TsExpressionParser.GetAsDateTime: TDateTime;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtDateTime);
  Result := Res.ResDatetime;
end;

function TsExpressionParser.GetAsFloat: TsExprFloat;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtFloat);
  Result := Res.ResFloat;
end;

function TsExpressionParser.GetAsInteger: Int64;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtInteger);
  Result := Res.ResInteger;
end;

function TsExpressionParser.GetAsString: String;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtString);
  Result := Res.ResString;
end;

function TsExpressionParser.GetRPNFormula: TsRPNFormula;
begin
  Result := CreateRPNFormula(FExprNode.AsRPNItem(nil), true);
end;

function TsExpressionParser.GetToken: TsTokenType;
begin
  Result := FScanner.GetToken;
end;

function TsExpressionParser.IdentifierByName(AName: ShortString): TsExprIdentifierDef;
var
  s: String;
begin
  if FDirty then
    CreateHashList;
  s := FHashList.NameOfIndex(0);
  s := FHashList.NameOfIndex(25);
  s := FHashList.NameOfIndex(36);
  Result := TsExprIdentifierDef(FHashList.Find(UpperCase(AName)));
end;

function TsExpressionParser.Level1: TsExprNode;
var
  tt: TsTokenType;
  Right: TsExprNode;
begin
{$ifdef debugexpr}Writeln('Level 1 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
{
  if TokenType = ttNot then
  begin
    GetToken;
    CheckEOF;
    Right := Level2;
    Result := TsNotExprNode.Create(Right);
  end
  else
  }
  Result := Level2;
   {
  try
    if TokenType = ttPower then
    begin
      tt := Tokentype;
      GetToken;
      CheckEOF;
      Right := Level2;
      Result := TsPowerExprNode.Create(Result, Right);
    end;
  except
    Result.Free;
    raise;
  end;
  }
{
  try

    while (TokenType in [ttAnd, ttOr, ttXor]) do
    begin
      tt := TokenType;
      GetToken;
      CheckEOF;
      Right := Level2;
      case tt of
        ttOr  : Result := TsBinaryOrExprNode.Create(Result, Right);
        ttAnd : Result := TsBinaryAndExprNode.Create(Result, Right);
        ttXor : Result := TsBinaryXorExprNode.Create(Result, Right);
      else
        ParserError(SErrUnknownBooleanOp)
      end;
    end;
  except
    Result.Free;
    raise;
  end;
}
end;

function TsExpressionParser.Level2: TsExprNode;
var
  right: TsExprNode;
  tt: TsTokenType;
  C: TsBinaryOperationExprNodeClass;
begin
{$ifdef debugexpr} Writeln('Level 2 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  Result := Level3;
  try
    if (TokenType in ttComparisons) then
    begin
      tt := TokenType;
      GetToken;
      CheckEOF;
      Right := Level3;
      CheckNodes(Result, right);
      case tt of
        ttLessthan         : C := TsLessExprNode;
        ttLessthanEqual    : C := TsLessEqualExprNode;
        ttLargerThan       : C := TsGreaterExprNode;
        ttLargerThanEqual  : C := TsGreaterEqualExprNode;
        ttEqual            : C := TsEqualExprNode;
        ttNotEqual         : C := TsNotEqualExprNode;
      else
        ParserError(SErrUnknownComparison)
      end;
      Result := C.Create(Result, right);
    end;
  except
    Result.Free;
    raise;
  end;
end;

function TsExpressionParser.Level3: TsExprNode;
var
  tt: TsTokenType;
  right: TsExprNode;
begin
{$ifdef debugexpr} Writeln('Level 3 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  Result := Level4;
  try
    while TokenType in [ttPlus, ttMinus, ttConcat] do begin
      tt := TokenType;
      GetToken;
      CheckEOF;
      right := Level4;
      CheckNodes(Result, right);
      case tt of
        ttPlus  : Result := TsAddExprNode.Create(Result, right);
        ttMinus : Result := TsSubtractExprNode.Create(Result, right);
        ttConcat: Result := TsConcatExprNode.Create(Result, right);
      end;
    end;
  except
    Result.Free;
    raise;
  end;
end;

function TsExpressionParser.Level4: TsExprNode;
var
  tt: TsTokenType;
  right: TsExprNode;
begin
{$ifdef debugexpr} Writeln('Level 4 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  Result := Level5;
  try
    while (TokenType in [ttMul, ttDiv]) do
    begin
      tt := TokenType;
      GetToken;
      right := Level5;
      CheckNodes(Result, right);
      case tt of
        ttMul : Result := TsMultiplyExprNode.Create(Result, right);
        ttDiv : Result := TsDivideExprNode.Create(Result, right);
      end;
    end;
  except
    Result.Free;
    Raise;
  end;
end;

function TsExpressionParser.Level5: TsExprNode;
var
  isPlus, isMinus: Boolean;
  tt: TsTokenType;
begin
{$ifdef debugexpr} Writeln('Level 5 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  isPlus := false;
  isMinus := false;
  if (TokenType in [ttPlus, ttMinus]) then
  begin
    isPlus := (TokenType = ttPlus);
    isMinus := (TokenType = ttMinus);
    GetToken;
  end;
  Result := Level6;
  if isPlus then
    Result := TsUPlusExprNode.Create(Result);
  if isMinus then
    Result := TsUMinusExprNode.Create(Result);
end;

function TsExpressionParser.Level6: TsExprNode;
var
  tt: TsTokenType;
  Right: TsExprNode;
begin
{$ifdef debugexpr} Writeln('Level 6 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  if (TokenType = ttLeft) then
  begin
    GetToken;
    Result := TsParenthesisExprNode.Create(Level1);
    try
      if (TokenType <> ttRight) then
        ParserError(Format(SErrBracketExpected, [SCanner.Pos, CurrentToken]));
      GetToken;
    except
      Result.Free;
      raise;
    end;
  end
  else
    Result := Primitive;

  if TokenType = ttPower then
  begin
    try
      CheckEOF;
      tt := Tokentype;
      GetToken;
      Right := Primitive;
      CheckNodes(Result, right);
      Result := TsPowerExprNode.Create(Result, Right);
      //GetToken;
    except
      Result.Free;
      raise;
    end;
  end;
end;

{ Checks types of todo and match. If ToDO can be converted to it matches
  the type of match, then a node is inserted.
  For binary operations, this function is called for both operands. }
function TsExpressionParser.MatchNodes(ToDo, Match: TsExprNode): TsExprNode;
var
  TT, MT : TsResultType;
begin
  Result := ToDo;
  TT := ToDo.NodeType;
  MT := Match.NodeType;
  if TT <> MT then
  begin
    if TT = rtInteger then
    begin
      if (MT in [rtFloat, rtDateTime]) then
        Result := ConvertNode(ToDo, MT);
    end
    else if (TT = rtFloat) then
    begin
      if (MT = rtDateTime) then
        Result := ConvertNode(ToDo, rtDateTime);
    end;
  end;
end;

procedure TsExpressionParser.ParserError(Msg: String);
begin
  raise EExprParser.Create(Msg);
end;

function TsExpressionParser.Primitive: TsExprNode;
var
  I: Int64;
  X: TsExprFloat;
  lCount: Integer;
  ID: TsExprIdentifierDef;
  Args: TsExprArgumentArray;
  AI: Integer;
  cell: PCell;
  optional: Boolean;
  token: String;
begin
{$ifdef debugexpr} Writeln('Primitive : ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  SetLength(Args, 0);
  if (TokenType = ttNumber) then
  begin
    if TryStrToInt64(CurrentToken, I) then
      Result := TsConstExprNode.CreateInteger(I)
    else
    begin
      if TryStrToFloat(CurrentToken, X, ExprFormatSettings) then
        Result := TsConstExprNode.CreateFloat(X)
      else
        ParserError(Format(SErrInvalidFloat, [CurrentToken]));
    end;
  end
  else if (TokenType = ttTrue) then
    Result := TsConstExprNode.CreateBoolean(true)
  else if (TokenType = ttFalse) then
    Result := TsConstExprNode.CreateBoolean(false)
  else if (TokenType = ttString) then
    Result := TsConstExprNode.CreateString(CurrentToken)
  else if (TokenType = ttCell) then
    Result := TsCellExprNode.Create(FWorksheet, CurrentToken)
  else if (TokenType = ttCellRange) then
    Result := TsCellRangeExprNode.Create(FWorksheet, CurrentToken)
  else if not (TokenType in [ttIdentifier{, ttIf}]) then
    ParserError(Format(SerrUnknownTokenAtPos, [Scanner.Pos, CurrentToken]))
  else
  begin
    token := Uppercase(CurrentToken);
    ID := self.IdentifierByName(token);
    if (ID = nil) then
      ParserError(Format(SErrUnknownIdentifier, [token]));
    if (ID.IdentifierType in [itFunctionCallBack, itFunctionHandler]) then
    begin
      lCount := ID.ArgumentCount;
      if lCount = 0 then  // we have to handle the () here, it will be skipped below.
      begin
        GetToken;
        if (TokenType <> ttLeft) then
          ParserError(Format(SErrLeftBracketExpected, [Scanner.Pos, CurrentToken]));
        GetToken;
        if (TokenType <> ttRight) then
          ParserError(Format(SErrBracketExpected, [Scanner.Pos, CurrentToken]));
        SetLength(Args, 0);
      end;
    end
    else
      lCount := 0;

    // Parse arguments.
    // Negative is for variable number of arguments, where Abs(value) is the minimum number of arguments
    if (lCount <> 0) then
    begin
      GetToken;
      if (TokenType <> ttLeft) then
        ParserError(Format(SErrLeftBracketExpected, [Scanner.Pos, CurrentToken]));
      SetLength(Args, abs(lCount));
      AI := 0;
      try
        repeat
          GetToken;
          // Check if we must enlarge the argument array
          if (lCount < 0) and (AI = Length(Args)) then
          begin
            SetLength(Args, AI+1);
            Args[AI] := nil;
          end;
          Args[AI] := Level1;
          inc(AI);
          optional := ID.IsOptionalArgument(AI+1);
          if not optional then
          begin
            if (TokenType <> ttComma) then
              if (AI < abs(lCount)) then
                ParserError(Format(SErrCommaExpected, [Scanner.Pos, CurrentToken]))
          end;
        until (AI = lCount) or (((lCount < 0) or optional) and (TokenType = ttRight));
        if TokenType <> ttRight then
          ParserError(Format(SErrBracketExpected, [Scanner.Pos, CurrentToken]));
        if AI < abs(lCount) then
          SetLength(Args, AI);
      except
        on E: Exception do
        begin
          dec(AI);
          while (AI >= 0) do
          begin
            FreeAndNil(Args[Ai]);
            dec(AI);
          end;
          raise;
        end;
      end;
    end;
    case ID.IdentifierType of
      itVariable         : Result := TsVariableExprNode.CreateIdentifier(ID);
      itFunctionCallBack : Result := TsFunctionCallBackExprNode.CreateFunction(ID, Args);
      itFunctionHandler  : Result := TFPFunctionEventHandlerExprNode.CreateFunction(ID, Args);
    end;
  end;
  GetToken;
  if TokenType = ttPercent then begin
    Result := TsPercentExprNode.Create(Result);
    GetToken;
  end;
end;

function TsExpressionParser.ResultType: TsResultType;
begin
  if not Assigned(FExprNode) then
    ParserError(SErrInExpression);
  Result := FExprNode.NodeType;;
end;

procedure TsExpressionParser.SetBuiltIns(const AValue: TsBuiltInExprCategories);
begin
  if FBuiltIns = AValue then
    exit;
  FBuiltIns := AValue;
  FDirty := true;
end;

procedure TsExpressionParser.SetExpression(const AValue: String);
begin
  if FExpression = AValue then
    exit;
  FExpression := AValue;
  if (AValue <> '') and (AValue[1] = '=') then
    FScanner.Source := Copy(AValue, 2, Length(AValue))
  else
    FScanner.Source := AValue;
  FreeAndNil(FExprNode);
  if (FExpression <> '') then
  begin
    GetToken;
    FExprNode := Level1;
    if (TokenType <> ttEOF) then
      ParserError(Format(SErrUnterminatedExpression, [Scanner.Pos, CurrentToken]));
    FExprNode.Check;
  end;
end;

procedure TsExpressionParser.SetIdentifiers(const AValue: TsExprIdentifierDefs);
begin
  FIdentifiers.Assign(AValue)
end;

procedure TsExpressionParser.SetRPNFormula(const AFormula: TsRPNFormula);

  procedure CreateNodeFromRPN(var ANode: TsExprNode; var AIndex: Integer);
  var
    node: TsExprNode;
    left: TsExprNode;
    right: TsExprNode;
    operand: TsExprNode;
    fek: TFEKind;
    r,c, r2,c2: Cardinal;
    flags: TsRelFlags;
    ID: TsExprIdentifierDef;
    i, n: Integer;
    args: TsExprArgumentArray;
  begin
    if AIndex < 0 then
      exit;

    fek := AFormula[AIndex].ElementKind;

    case fek of
      fekCell, fekCellRef:
        begin
          r := AFormula[AIndex].Row;
          c := AFormula[AIndex].Col;
          flags := AFormula[AIndex].RelFlags;
          ANode := TsCellExprNode.Create(FWorksheet, r, c, flags);
          dec(AIndex);
        end;
      fekCellRange:
        begin
          r := AFormula[AIndex].Row;
          c := AFormula[AIndex].Col;
          r2 := AFormula[AIndex].Row2;
          c2 := AFormula[AIndex].Col2;
          flags := AFormula[AIndex].RelFlags;
          ANode := TsCellRangeExprNode.Create(FWorksheet, r, c, r2, c2, flags);
          dec(AIndex);
        end;
      fekNum:
        begin
          ANode := TsConstExprNode.CreateFloat(AFormula[AIndex].DoubleValue);
          dec(AIndex);
        end;
      fekInteger:
        begin
          ANode := TsConstExprNode.CreateInteger(AFormula[AIndex].IntValue);
          dec(AIndex);
        end;
      fekString:
        begin
          ANode := TsConstExprNode.CreateString(AFormula[AIndex].StringValue);
          dec(AIndex);
        end;
      fekBool:
        begin
          ANode := TsConstExprNode.CreateBoolean(AFormula[AIndex].DoubleValue <> 0.0);
          dec(AIndex);
        end;
      fekErr:
        begin
          ANode := TsConstExprNode.CreateError(TsErrorValue(AFormula[AIndex].IntValue));
          dec(AIndex);
        end;

      // unary operations
      fekPercent, fekUMinus, fekUPlus, fekParen:
        begin
          dec(AIndex);
          CreateNodeFromRPN(operand, AIndex);
          case fek of
            fekPercent : ANode := TsPercentExprNode.Create(operand);
            fekUMinus  : ANode := TsUMinusExprNode.Create(operand);
            fekUPlus   : ANode := TsUPlusExprNode.Create(operand);
            fekParen   : ANode := TsParenthesisExprNode.Create(operand);
          end;
        end;

      // binary operations
      fekAdd, fekSub, fekMul, fekDiv,
      fekPower, fekConcat,
      fekEqual, fekNotEqual,
      fekGreater, fekGreaterEqual,
      fekLess, fekLessEqual:
        begin
          dec(AIndex);
          CreateNodeFromRPN(right, AIndex);
          CreateNodeFromRPN(left, AIndex);
          CheckNodes(left, right);
          case fek of
            fekAdd         : ANode := TsAddExprNode.Create(left, right);
            fekSub         : ANode := TsSubtractExprNode.Create(left, right);
            fekMul         : ANode := TsMultiplyExprNode.Create(left, right);
            fekDiv         : ANode := TsDivideExprNode.Create(left, right);
            fekPower       : ANode := TsPowerExprNode.Create(left, right);
            fekConcat      : ANode := tsConcatExprNode.Create(left, right);
            fekEqual       : ANode := TsEqualExprNode.Create(left, right);
            fekNotEqual    : ANode := TsNotEqualExprNode.Create(left, right);
            fekGreater     : ANode := TsGreaterExprNode.Create(left, right);
            fekGreaterEqual: ANode := TsGreaterEqualExprNode.Create(left, right);
            fekLess        : ANode := TsLessExprNode.Create(left, right);
            fekLessEqual   : ANode := tsLessEqualExprNode.Create(left, right);
          end;
        end;

      // functions
      fekFunc:
        begin
          ID := self.IdentifierByName(AFormula[AIndex].FuncName);
          if ID = nil then
          begin
            ParserError(Format(SErrUnknownIdentifier,[AFormula[AIndex].FuncName]));
            dec(AIndex);
          end else
          begin
            if ID.HasFixedArgumentCount then
              n := ID.ArgumentCount
            else
              n := AFormula[AIndex].ParamsNum;
            dec(AIndex);
            SetLength(args, n);
            for i:=n-1 downto 0 do
              CreateNodeFromRPN(args[i], AIndex);
            case ID.IdentifierType of
              itVariable         : ANode := TsVariableExprNode.CreateIdentifier(ID);
              itFunctionCallBack : ANode := TsFunctionCallBackExprNode.CreateFunction(ID, args);
              itFunctionHandler  : ANode := TFPFunctionEventHandlerExprNode.CreateFunction(ID, args);
            end;
          end;
        end;

    end;  //case
  end; //begin

var
  index: Integer;
  node: TsExprNode;
begin
  FExpression := '';
  FreeAndNil(FExprNode);
  index := Length(AFormula)-1;
  CreateNodeFromRPN(FExprNode, index);
  if Assigned(FExprNode) then FExprNode.Check;
end;

function TsExpressionParser.TokenType: TsTokenType;
begin
  Result := FScanner.TokenType;
end;


{------------------------------------------------------------------------------}
{  TsSpreadsheetParser                                                         }
{------------------------------------------------------------------------------}

constructor TsSpreadsheetParser.Create(AWorksheet: TsWorksheet);
begin
  inherited Create(AWorksheet);
  BuiltIns := AllBuiltIns;
end;


{------------------------------------------------------------------------------}
{  TsExprIdentifierDefs                                                        }
{------------------------------------------------------------------------------}

function TsExprIdentifierDefs.AddBooleanVariable(const AName: ShortString;
  AValue: Boolean): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtBoolean;
  Result.FValue.ResBoolean := AValue;
end;

function TsExprIdentifierDefs.AddDateTimeVariable(const AName: ShortString;
  AValue: TDateTime): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtDateTime;
  Result.FValue.ResDateTime := AValue;
end;

function TsExprIdentifierDefs.AddFloatVariable(const AName: ShortString;
  AValue: TsExprFloat): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtFloat;
  Result.FValue.ResFloat := AValue;
end;

function TsExprIdentifierDefs.AddFunction(const AName: ShortString;
  const AResultType: Char; const AParamTypes: String; const AExcelCode: Integer;
  ACallBack: TsExprFunctionCallBack): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.Name := AName;
  Result.IdentifierType := itFunctionCallBack;
  Result.ResultType := CharToResultType(AResultType);
  Result.ExcelCode := AExcelCode;
  Result.FOnGetValueCB := ACallBack;
  if (Length(AParamTypes) > 0) and (AParamTypes[Length(AParamTypes)]='+') then
  begin
    Result.ParameterTypes := Copy(AParamTypes, 1, Length(AParamTypes)-1);
    Result.VariableArgumentCount := true;
  end else
    Result.ParameterTypes := AParamTypes;
end;

function TsExprIdentifierDefs.AddFunction(const AName: ShortString;
  const AResultType: Char; const AParamTypes: String; const AExcelCode: Integer;
  ACallBack: TsExprFunctionEvent): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.Name := AName;
  Result.IdentifierType := itFunctionHandler;
  Result.ResultType := CharToResultType(AResultType);
  Result.ExcelCode := AExcelCode;
  Result.FOnGetValue := ACallBack;
  if (Length(AParamTypes) > 0) and (AParamTypes[Length(AParamTypes)]='+') then
  begin
    Result.ParameterTypes := Copy(AParamTypes, 1, Length(AParamTypes)-1);
    Result.VariableArgumentCount := true;
  end else
    Result.ParameterTypes := AParamTypes;
end;

function TsExprIdentifierDefs.AddIntegerVariable(const AName: ShortString;
  AValue: Integer): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtInteger;
  Result.FValue.ResInteger := AValue;
end;

function TsExprIdentifierDefs.AddStringVariable(const AName: ShortString;
  AValue: String): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtString;
  Result.FValue.ResString := AValue;
end;

function TsExprIdentifierDefs.AddVariable(const AName: ShortString;
  AResultType: TsResultType; AValue: String): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := AResultType;
  Result.Value := AValue;
end;

function TsExprIdentifierDefs.FindIdentifier(const AName: ShortString
  ): TsExprIdentifierDef;
var
  I: Integer;
begin
  I := IndexOfIdentifier(AName);
  if (I = -1) then
    Result := nil
  else
    Result := GetI(I);
end;

function TsExprIdentifierDefs.GetI(AIndex : Integer): TsExprIdentifierDef;
begin
  Result := TsExprIdentifierDef(Items[AIndex]);
end;

function TsExprIdentifierDefs.IdentifierByExcelCode(const AExcelCode: Integer
  ): TsExprIdentifierDef;
var
  I: Integer;
begin
  I := IndexOfIdentifier(AExcelCode);
  if I = -1 then
    Result := nil
  else
    Result := GetI(I);
end;

function TsExprIdentifierDefs.IdentifierByName(const AName: ShortString
  ): TsExprIdentifierDef;
begin
  Result := FindIdentifier(AName);
  if (Result = nil) then
    RaiseParserError(SErrUnknownIdentifier, [AName]);
end;

function TsExprIdentifierDefs.IndexOfIdentifier(const AName: ShortString): Integer;
begin
  Result := Count-1;
  while (Result >= 0) and (CompareText(GetI(Result).Name, AName) <> 0) do
    dec(Result);
end;

function TsExprIdentifierDefs.IndexOfIdentifier(const AExcelCode: Integer): Integer;
var
  ID: TsExprIdentifierDef;
begin
  Result := Count-1;
  while (Result >= 0) do begin
    ID := GetI(Result);
    if ID.ExcelCode = AExcelCode then exit;
    dec(Result);
  end;
  {
  while (Result >= 0) and (GetI(Result).ExcelCode = AExcelCode) do
    dec(Result);
    }
end;

procedure TsExprIdentifierDefs.SetI(AIndex: Integer;
  const AValue: TsExprIdentifierDef);
begin
  Items[AIndex] := AValue;
end;

procedure TsExprIdentifierDefs.Update(Item: TCollectionItem);
begin
  if Assigned(FParser) then
    FParser.FDirty := true;
end;


{------------------------------------------------------------------------------}
{  TsExprIdentifierDef                                                        }
{------------------------------------------------------------------------------}

function TsExprIdentifierDef.ArgumentCount: Integer;
begin
  if FVariableArgumentCount then
    Result := -Length(FArgumentTypes)
  else
    Result := Length(FArgumentTypes);
end;

procedure TsExprIdentifierDef.Assign(Source: TPersistent);
var
  EID: TsExprIdentifierDef;
begin
  if (Source is TsExprIdentifierDef) then
  begin
    EID := Source as TsExprIdentifierDef;
    FStringValue := EID.FStringValue;
    FValue := EID.FValue;
    FArgumentTypes := EID.FArgumentTypes;
    FVariableArgumentCount := EID.FVariableArgumentCount;
    FExcelCode := EID.ExcelCode;
    FIDType := EID.FIDType;
    FName := EID.FName;
    FOnGetValue := EID.FOnGetValue;
    FOnGetValueCB := EID.FOnGetValueCB;
  end
  else
    inherited Assign(Source);
end;

procedure TsExprIdentifierDef.CheckResultType(const AType: TsResultType);
begin
  if (FValue.ResultType <> AType) then
    RaiseParserError(SErrInvalidResultType, [ResultTypeName(AType)])
end;

procedure TsExprIdentifierDef.CheckVariable;
begin
  if Identifiertype <> itVariable then
    RaiseParserError(SErrNotVariable, [Name]);
end;

function TsExprIdentifierDef.GetAsBoolean: Boolean;
begin
  CheckResultType(rtBoolean);
  CheckVariable;
  Result := FValue.ResBoolean;
end;

function TsExprIdentifierDef.GetAsDateTime: TDateTime;
begin
  CheckResultType(rtDateTime);
  CheckVariable;
  Result := FValue.ResDateTime;
end;

function TsExprIdentifierDef.GetAsFloat: TsExprFloat;
begin
  CheckResultType(rtFloat);
  CheckVariable;
  Result := FValue.ResFloat;
end;

function TsExprIdentifierDef.GetAsInteger: Int64;
begin
  CheckResultType(rtInteger);
  CheckVariable;
  Result := FValue.ResInteger;
end;

function TsExprIdentifierDef.GetAsString: String;
begin
  CheckResultType(rtString);
  CheckVariable;
  Result := FValue.ResString;
end;

function TsExprIdentifierDef.GetResultType: TsResultType;
begin
  Result := FValue.ResultType;
end;

function TsExprIdentifierDef.GetValue: String;
begin
  case FValue.ResultType of
    rtBoolean  : if FValue.ResBoolean then
                   Result := 'True'
                 else
                   Result := 'False';
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtFloat    : Result := FloatToStr(FValue.ResFloat, ExprFormatSettings);
    rtDateTime : Result := FormatDateTime('cccc', FValue.ResDateTime);
    rtString   : Result := FValue.ResString;
  end;
end;

{ Returns true if the epxression has a fixed number of arguments. }
function TsExprIdentifierDef.HasFixedArgumentCount: Boolean;
var
  i: Integer;
begin
  if FVariableArgumentCount then
    Result := false
  else
  begin
    for i:= 1 to Length(FArgumentTypes) do
      if IsOptionalArgument(i) then
      begin
        Result := false;
        exit;
      end;
    Result := true;
  end;
end;

{ Checks whether an argument is optional. Index number starts at 1.
  Optional arguments are lower-case characters in the argument list. }
function TsExprIdentifierDef.IsOptionalArgument(AIndex: Integer): Boolean;
begin
  Result := (AIndex <= Length(FArgumentTypes))
    and (UpCase(FArgumentTypes[AIndex]) <> FArgumentTypes[AIndex]);
end;

procedure TsExprIdentifierDef.SetArgumentTypes(const AValue: String);
var
  i: integer;
begin
  if FArgumentTypes = AValue then
    exit;
  for i:=1 to Length(AValue) do
    CharToResultType(AValue[i]);
  FArgumentTypes := AValue;
end;

procedure TsExprIdentifierDef.SetAsBoolean(const AValue: Boolean);
begin
  CheckVariable;
  CheckResultType(rtBoolean);
  FValue.ResBoolean := AValue;
end;

procedure TsExprIdentifierDef.SetAsDateTime(const AValue: TDateTime);
begin
  CheckVariable;
  CheckResultType(rtDateTime);
  FValue.ResDateTime := AValue;
end;

procedure TsExprIdentifierDef.SetAsFloat(const AValue: TsExprFloat);
begin
  CheckVariable;
  CheckResultType(rtFloat);
  FValue.ResFloat := AValue;
end;

procedure TsExprIdentifierDef.SetAsInteger(const AValue: Int64);
begin
  CheckVariable;
  CheckResultType(rtInteger);
  FValue.ResInteger := AValue;
end;

procedure TsExprIdentifierDef.SetAsString(const AValue: String);
begin
  CheckVariable;
  CheckResultType(rtString);
  FValue.resString := AValue;
end;

procedure TsExprIdentifierDef.SetName(const AValue: ShortString);
begin
  if FName = AValue then
    exit;
  if (AValue <> '') then
    if Assigned(Collection) and (TsExprIdentifierDefs(Collection).IndexOfIdentifier(AValue) <> -1) then
      RaiseParserError(SErrDuplicateIdentifier,[AValue]);
  FName := AValue;
end;

procedure TsExprIdentifierDef.SetResultType(const AValue: TsResultType);
begin
  if AValue <> FValue.ResultType then
  begin
    FValue.ResultType := AValue;
    SetValue(FStringValue);
  end;
end;

procedure TsExprIdentifierDef.SetValue(const AValue: String);
begin
  FStringValue := AValue;
  if (AValue <> '') then
    case FValue.ResultType of
      rtBoolean  : FValue.ResBoolean := FStringValue='True';
      rtInteger  : FValue.ResInteger := StrToInt(AValue);
      rtFloat    : FValue.ResFloat := StrToFloat(AValue);
      rtDateTime : FValue.ResDateTime := StrToDateTime(AValue);
      rtString   : FValue.ResString := AValue;
    end
  else
    case FValue.ResultType of
      rtBoolean  : FValue.ResBoolean := false;
      rtInteger  : FValue.ResInteger := 0;
      rtFloat    : FValue.ResFloat := 0.0;
      rtDateTime : FValue.ResDateTime := 0;
      rtString   : FValue.ResString := '';
    end
end;


{------------------------------------------------------------------------------}
{  TsBuiltInExpressionManager                                                         }
{------------------------------------------------------------------------------}

constructor TsBuiltInExpressionManager.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDefs := TsExprIdentifierDefs.Create(TsBuiltInExprIdentifierDef)
end;

destructor TsBuiltInExpressionManager.Destroy;
begin
  FreeAndNil(FDefs);
  inherited Destroy;
end;

function TsBuiltInExpressionManager.AddVariable(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; AResultType: TsResultType; AValue: String
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.Addvariable(AName, AResultType, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddBooleanVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: Boolean
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddBooleanvariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddDateTimeVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: TDateTime
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddDateTimeVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFloatVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString;
  AValue: TsExprFloat): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFloatVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFunction(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; const AResultType: Char; const AParamTypes: String;
  const AExcelCode: Integer; ACallBack: TsExprFunctionCallBack): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFunction(AName, AResultType,
    AParamTypes, AExcelCode, ACallBack));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFunction(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; const AResultType: Char; const AParamTypes: String;
  const AExcelCode: Integer; ACallBack: TsExprFunctionEvent): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFunction(AName, AResultType,
    AParamTypes, AExcelCode, ACallBack));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddIntegerVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: Integer
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddIntegerVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddStringVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: String
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddStringVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.FindIdentifier(const AName: ShortString
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.FindIdentifier(AName));
end;

function TsBuiltInExpressionManager.GetCount: Integer;
begin
  Result := FDefs.Count;
end;

function TsBuiltInExpressionManager.GetI(AIndex: Integer): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs[Aindex])
end;

function TsBuiltInExpressionManager.IdentifierByExcelCode(const AExcelCode: Integer
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.IdentifierByExcelCode(AExcelCode));
end;

function TsBuiltInExpressionManager.IdentifierByName(const AName: ShortString
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.IdentifierByName(AName));
end;

function TsBuiltInExpressionManager.IndexOfIdentifier(const AName: ShortString): Integer;
begin
  Result := FDefs.IndexOfIdentifier(AName);
end;


{------------------------------------------------------------------------------}
{  Various Nodes                                                               }
{------------------------------------------------------------------------------}

{ TsExprNode }

procedure TsExprNode.CheckNodeType(ANode: TsExprNode; Allowed: TsResultTypes);
var
  S: String;
  A: TsResultType;
begin
  if (ANode = nil) then
    RaiseParserError(SErrNoNodeToCheck);
  if not (ANode.NodeType in Allowed) then
  begin
    S := '';
    for A := Low(TsResultType) to High(TsResultType) do
      if A in Allowed then
      begin
        if S <> '' then
          S := S + ',';
        S := S + ResultTypeName(A);
      end;
    RaiseParserError(SInvalidNodeType, [ResultTypeName(ANode.NodeType), S, ANode.AsString]);
  end;
end;

function TsExprNode.NodeValue: TsExpressionResult;
begin
  GetNodeValue(Result);
end;


{ TsUnaryOperationExprNode }

constructor TsUnaryOperationExprNode.Create(AOperand: TsExprNode);
begin
  FOperand := AOperand;
end;

destructor TsUnaryOperationExprNode.Destroy;
begin
  FreeAndNil(FOperand);
  inherited Destroy;
end;

procedure TsUnaryOperationExprNode.Check;
begin
  if not Assigned(Operand) then
    RaiseParserError(SErrNoOperand, [Self.ClassName]);
end;


{ TsBinaryOperationExprNode }

constructor TsBinaryOperationExprNode.Create(ALeft, ARight: TsExprNode);
begin
  FLeft := ALeft;
  FRight := ARight;
end;

destructor TsBinaryOperationExprNode.Destroy;
begin
  FreeAndNil(FLeft);
  FreeAndNil(FRight);
  inherited Destroy;
end;

procedure TsBinaryOperationExprNode.Check;
begin
  if not Assigned(Left) then
    RaiseParserError(SErrNoLeftOperand,[classname]);
  if not Assigned(Right) then
    RaiseParserError(SErrNoRightOperand,[classname]);
end;

procedure TsBinaryOperationExprNode.CheckSameNodeTypes;
var
  LT, RT: TsResultType;
begin
  LT := Left.NodeType;
  RT := Right.NodeType;
  if (RT <> LT) then
    RaiseParserError(SErrTypesDoNotMatch, [ResultTypeName(LT), ResultTypeName(RT), Left.AsString, Right.AsString])
end;


{ TsBooleanOperationExprNode }

procedure TsBooleanOperationExprNode.Check;
begin
  inherited Check;
  CheckNodeType(Left,  [rtBoolean, rtCell, rtError, rtEmpty]);
  CheckNodeType(Right, [rtBoolean, rtCell, rtError, rtEmpty]);
  CheckSameNodeTypes;
end;

function TsBooleanOperationExprNode.NodeType: TsResultType;
begin
  Result := Left.NodeType;
end;


{ TsConstExprNode }

constructor TsConstExprNode.CreateString(AValue: String);
begin
  FValue.ResultType := rtString;
  FValue.ResString := AValue;
end;

constructor TsConstExprNode.CreateInteger(AValue: Int64);
begin
  FValue.ResultType := rtInteger;
  FValue.ResInteger := AValue;
end;

constructor TsConstExprNode.CreateDateTime(AValue: TDateTime);
begin
  FValue.ResultType := rtDateTime;
  FValue.ResDateTime := AValue;
end;

constructor TsConstExprNode.CreateFloat(AValue: TsExprFloat);
begin
  Inherited Create;
  FValue.ResultType := rtFloat;
  FValue.ResFloat := AValue;
end;

constructor TsConstExprNode.CreateBoolean(AValue: Boolean);
begin
  FValue.ResultType := rtBoolean;
  FValue.ResBoolean := AValue;
end;

constructor TsConstExprNode.CreateError(AValue: TsErrorValue);
begin
  FValue.ResultType := rtError;
  FValue.ResError := AValue;
end;

procedure TsConstExprNode.Check;
begin
  // Nothing to check;
end;

function TsConstExprNode.NodeType: TsResultType;
begin
  Result := FValue.ResultType;
end;

procedure TsConstExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Result := FValue;
end;

function TsConstExprNode.AsString: string;
begin
  case NodeType of
    rtString   : Result := cDoubleQuote + FValue.ResString + cDoubleQuote;
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtDateTime : Result := '''' + FormatDateTime('cccc', FValue.ResDateTime) + '''';    // Probably wrong !!!
    rtBoolean  : if FValue.ResBoolean then Result := 'TRUE' else Result := 'FALSE';
    rtFloat    : Result := FloatToStr(FValue.ResFloat, ExprFormatSettings);
  end;
end;

function TsConstExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  case NodeType of
    rtString   : Result := RPNString(FValue.ResString, ANext);
    rtInteger  : Result := RPNInteger(FValue.ResInteger, ANext);
    rtDateTime : Result := RPNNumber(FValue.ResDateTime, ANext);
    rtBoolean  : Result := RPNBool(FValue.ResBoolean, ANext);
    rtFloat    : Result := RPNNumber(FValue.ResFloat, ANext);
  end;
end;


{ TsUPlusExprNode }

function TsUPlusExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekUPlus,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsUPlusExprNode.AsString: String;
begin
  Result := '+' + TrimLeft(Operand.AsString);
end;

procedure TsUPlusExprNode.Check;
const
  AllowedTokens = [rtInteger, rtFloat, rtCell, rtEmpty, rtError];
begin
  inherited;
  if not (Operand.NodeType in AllowedTokens) then
    RaiseParserError(SErrNoUPlus, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsUPlusExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  res: TsExpressionresult;
  cell: PCell;
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtInteger, rtFloat, rtError:
      exit;
    rtCell:
      begin
        cell := ArgToCell(Result);
        if cell = nil then
          Result := FloatResult(0.0)
        else
        if cell^.ContentType = cctNumber then
        begin
          if frac(cell^.NumberValue) = 0.0 then
            Result := IntegerResult(trunc(cell^.NumberValue))
          else
            Result := FloatResult(cell^.NumberValue);
        end;
      end;
    rtEmpty:
      Result := FloatResult(0.0);
    else
      Result := ErrorResult(errWrongType);
  end;
end;

function TsUPlusExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;


{ TsUMinusExprNode }

function TsUMinusExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekUMinus,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsUMinusExprNode.AsString: String;
begin
  Result := '-' + TrimLeft(Operand.AsString);
end;

procedure TsUMinusExprNode.Check;
const
  AllowedTokens = [rtInteger, rtFloat, rtCell, rtEmpty, rtError];
begin
  inherited;
  if not (Operand.NodeType in AllowedTokens) then
    RaiseParserError(SErrNoNegation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsUMinusExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  cell: PCell;
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtError:
      exit;
    rtFloat:
      Result := FloatResult(-Result.ResFloat);
    rtInteger:
      Result := IntegerResult(-Result.ResInteger);
    rtCell:
      begin
        cell := ArgToCell(Result);
        if (cell <> nil) and (cell^.ContentType = cctNumber) then
        begin
          if frac(cell^.NumberValue) = 0.0 then
            Result := IntegerResult(-trunc(cell^.NumberValue))
          else
            Result := FloatResult(cell^.NumberValue);
        end else
          Result := FloatResult(0.0);
      end;
    rtEmpty:
      Result := FloatResult(0.0);
    else
      Result := ErrorResult(errWrongType);
  end;
end;

function TsUMinusExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;


{ TsPercentExprNode }

function TsPercentExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekPercent,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsPercentExprNode.AsString: String;
begin
  Result := Operand.AsString + '%';
end;

procedure TsPercentExprNode.Check;
const
  AllowedTokens = [rtInteger, rtFloat, rtCell, rtEmpty, rtError];
begin
  inherited;
  if not (Operand.NodeType in AllowedTokens) then
    RaiseParserError(SErrNoPercentOperation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsPercentExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtError:
      exit;
    rtFloat, rtInteger, rtCell:
      Result := FloatResult(ArgToFloat(Result)*0.01);
    else
      Result := ErrorResult(errWrongType);
  end;
end;

function TsPercentExprNode.NodeType: TsResultType;
begin
  Result := rtFloat;
end;


{ TsParenthesisExprNode }

function TsParenthesisExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekParen,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsParenthesisExprNode.AsString: String;
begin
  Result := '(' + Operand.AsString + ')';
end;

function TsParenthesisExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;

procedure TsParenthesisExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Result := Operand.NodeValue;
end;


{ TsNotExprNode }

function TsNotExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc('NOT',
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsNotExprNode.AsString: String;
begin
  Result := 'not ' + Operand.AsString;
end;

procedure TsNotExprNode.Check;
const
  AllowedTokens = [rtBoolean, rtEmpty, rtError];
begin
  if not (Operand.NodeType in AllowedTokens) then
    RaiseParserError(SErrNoNotOperation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsNotExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtBoolean : Result.ResBoolean := not Result.ResBoolean;
    rtEmpty   : Result := BooleanResult(true);
  end
end;

function TsNotExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;


{ TsBooleanResultExprNode }

procedure TsBooleanResultExprNode.Check;
begin
  inherited Check;
  CheckSameNodeTypes;
end;

procedure TsBooleanResultExprNode.CheckSameNodeTypes;
begin
  // Same node types are checked in GetNodevalue
end;

function TsBooleanResultExprNode.NodeType: TsResultType;
begin
  Result := rtBoolean;
end;


{ TsEqualExprNode }

function TsEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '=' + Right.AsString;
end;

procedure TsEqualExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);

  if (Result.ResultType in [rtInteger, rtFloat, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtInteger, rtFloat, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToFloat(Result) = ArgToFloat(RRes))
  else
  if (Result.ResultType in [rtString, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtString, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToString(Result) = ArgToString(RRes))
  else
  if (Result.ResultType in [rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtDateTime, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToDateTime(Result) = ArgToDateTime(RRes))
  else
  if (Result.ResultType in [rtBoolean, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtBoolean, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToBoolean(Result) = ArgToBoolean(RRes))
  else
  if (Result.ResultType = rtError)
    then Result := ErrorResult(Result.ResError)
  else
  if (RRes.ResultType = rtError)
    then Result := ErrorResult(RRes.ResError)
  else
    Result := BooleanResult(false);
end;


{ TsNotEqualExprNode }

function TsNotEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekNotEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsNotEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '<>' + Right.AsString;
end;

procedure TsNotEqualExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  inherited GetNodeValue(Result);
  Result.ResBoolean := not Result.ResBoolean;
end;


{ TsOrderingExprNode }

procedure TsOrderingExprNode.Check;
const
  AllowedTypes = [rtBoolean, rtInteger, rtFloat, rtDateTime, rtString, rtEmpty, rtError, rtCell];
begin
  CheckNodeType(Left, AllowedTypes);
  CheckNodeType(Right, AllowedTypes);
  inherited Check;
end;

procedure TsOrderingExprNode.CheckSameNodeTypes;
var
  LT, RT: TsResultType;
begin
  {
  LT := Left.NodeType;
  RT := Right.NodeType;
  case LT of
    rtFloat, rtInteger:
      if (RT in [rtFloat, rtInteger]) or
         ((Rt = rtCell) and (Right.Res
  if (RT <> LT) then
    RaiseParserError(SErrTypesDoNotMatch, [ResultTypeName(LT), ResultTypeName(RT), Left.AsString, Right.AsString])
    }
end;



{ TsLessExprNode }

function TsLessExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekLess,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsLessExprNode.AsString: string;
begin
  Result := Left.AsString + '<' + Right.AsString;
end;

procedure TsLessExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  if (Result.ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToFloat(Result) < ArgToFloat(RRes))
  else
  if (Result.ResultType in [rtString, rtInteger, rtFloat, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtString, rtInteger, rtFloat, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToString(Result) < ArgToString(RRes))
  else
  if (Result.ResultType in [rtBoolean, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtBoolean, rtCell, rtEmpty])
  then
    Result := BooleanResult(ord(ArgToBoolean(Result)) < ord(ArgToBoolean(RRes)))
  else
  if (Result.ResultType = rtError)
    then Result := ErrorResult(Result.ResError)
  else
  if (RRes.ResultType = rtError)
    then Result := ErrorResult(RRes.ResError)
  else
    Result := ErrorResult(errWrongType);
end;


{ TsGreaterExprNode }

function TsGreaterExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekGreater,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsGreaterExprNode.AsString: string;
begin
  Result := Left.AsString + '>' + Right.AsString;
end;

procedure TsGreaterExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  if (Result.ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToFloat(Result) > ArgToFloat(RRes))
  else
  if (Result.ResultType in [rtString, rtInteger, rtFloat, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtString, rtInteger, rtFloat, rtCell, rtEmpty])
  then
    Result := BooleanResult(ArgToString(Result) > ArgToString(RRes))
  else
  if (Result.ResultType in [rtBoolean, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtBoolean, rtCell, rtEmpty])
  then
    Result := BooleanResult(ord(ArgToBoolean(Result)) > ord(ArgToBoolean(RRes)))
  else
  if (Result.ResultType = rtError)
    then Result := ErrorResult(Result.ResError)
  else
  if (RRes.ResultType = rtError)
    then Result := ErrorResult(RRes.ResError)
  else
    Result := ErrorResult(errWrongType);
end;


{ TsGreaterEqualExprNode }

function TsGreaterEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekGreaterEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsGreaterEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '>=' + Right.AsString;
end;

procedure TsGreaterEqualExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  inherited GetNodeValue(Result);
  Result.ResBoolean := not Result.ResBoolean;
end;


{ TsLessEqualExprNode }

function TsLessEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekLessEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsLessEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '<=' + Right.AsString;
end;

procedure TsLessEqualExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  inherited GetNodeValue(Result);
  Result.ResBoolean := not Result.ResBoolean;
end;


{ TsConcatExprNode }

function TsConcatExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekConcat,
    Right.AsRPNItem(
    Left.AsRPNItem(
    nil)));
end;

function TsConcatExprNode.AsString: string;
begin
  Result := Left.AsString + '&' + Right.AsString;
end;

procedure TsConcatExprNode.Check;
begin
  inherited Check;
  CheckNodeType(Left, [rtString, rtCell, rtEmpty, rtError]);
  CheckNodeType(Right, [rtString, rtCell, rtEmpty, rtError]);
end;

procedure TsConcatExprNode.CheckSameNodeTypes;
begin
  // Same node types are checked in GetNodevalue
end;

procedure TsConcatExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes : TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  if (Result.ResultType = rtError)
    then exit;
  Right.GetNodeValue(RRes);
  if (Result.ResultType in [rtString, rtCell]) and (RRes.ResultType in [rtString, rtCell])
    then Result := StringResult(ArgToString(Result) + ArgToString(RRes))
  else
  if (RRes.ResultType = rtError)
    then Result := ErrorResult(RRes.ResError)
  else
    Result := ErrorResult(errWrongType);
end;

function TsConcatExprNode.NodeType: TsResultType;
begin
  Result := rtString;
end;


{ TsMathOperationExprNode }

procedure TsMathOperationExprNode.Check;
const
  AllowedTypes  = [rtInteger, rtFloat, rtDateTime, rtCell, rtEmpty, rtError];
begin
  inherited Check;
  CheckNodeType(Left, AllowedTypes);
  CheckNodeType(Right, AllowedTypes);
  CheckSameNodeTypes;
end;

procedure TsMathOperationExprNode.CheckSameNodeTypes;
begin
  // Same node types are checked in GetNodevalue
end;

function TsMathOperationExprNode.NodeType: TsResultType;
begin
  Result := Left.NodeType;
end;


{ TsAddExprNode }

function TsAddExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekAdd,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsAddExprNode.AsString: string;
begin
  Result := Left.AsString + '+' + Right.AsString;
end;

procedure TsAddExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  if Result.ResultType = rtError then
    exit;

  Right.GetNodeValue(RRes);
  if RRes.ResultType = rtError then
  begin
    Result := ErrorResult(RRes.ResError);
    exit;
  end;

  if (Result.ResultType in [rtInteger, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtInteger, rtCell, rtEmpty])
  then
    Result := IntegerResult(ArgToInt(Result) + ArgToInt(RRes))
  else
  if (Result.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty])
  then
    Result := FloatResult(ArgToFloat(Result) + ArgToFloat(RRes));
end;


{ TsSubtractExprNode }

function TsSubtractExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekSub,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsSubtractExprNode.AsString: string;
begin
  Result := Left.AsString + '-' + Right.asString;
end;

procedure TsSubtractExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  if Result.ResultType = rtError then
    exit;

  Right.GetNodeValue(RRes);
  if RRes.ResultType = rtError then
  begin
    Result := ErrorResult(RRes.ResError);
    exit;
  end;

  if (Result.ResultType in [rtInteger, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtInteger, rtCell, rtEmpty])
  then
    Result := IntegerResult(ArgToInt(Result) - ArgToInt(RRes))
  else
  if (Result.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty])
  then
    Result := FloatResult(ArgToFloat(Result) - ArgToFloat(RRes));
end;


{ TsMultiplyExprNode }

function TsMultiplyExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekMul,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsMultiplyExprNode.AsString: string;
begin
  Result := Left.AsString + '*' + Right.AsString;
end;

procedure TsMultiplyExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  if Result.ResultType = rtError then
    exit;

  Right.GetNodeValue(RRes);
  if RRes.ResultType = rtError then
  begin
    Result := ErrorResult(RRes.ResError);
    exit;
  end;

  if (Result.ResultType in [rtInteger, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtInteger, rtCell, rtEmpty])
  then
    Result := IntegerResult(ArgToInt(Result) * ArgToInt(RRes))
  else
  if (Result.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty])
  then
    Result := FloatResult(ArgToFloat(Result) * ArgToFloat(RRes));
end;


{ TsDivideExprNode }

function TsDivideExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekDiv,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsDivideExprNode.AsString: string;
begin
  Result := Left.AsString + '/' + Right.asString;
end;

procedure TsDivideExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
  y: TsExprFloat;
begin
  Left.GetNodeValue(Result);
  if Result.ResultType = rtError then
    exit;

  Right.GetNodeValue(RRes);
  if RRes.ResultType = rtError then
  begin
    Result := ErrorResult(RRes.ResError);
    exit;
  end;

  if (Result.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty])
  then begin
    y := ArgToFloat(RRes);
    if y = 0.0 then
      Result := ErrorResult(errDivideByZero)
    else
      Result := FloatResult(ArgToFloat(Result) / y);
  end;
end;

function TsDivideExprNode.NodeType: TsResultType;
begin
  Result := rtFLoat;
end;


{ TsPowerExprNode }

function TsPowerExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekPower,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsPowerExprNode.AsString: string;
begin
  Result := Left.AsString + '^' + Right.AsString;
end;

procedure TsPowerExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
  ex: TsExprFloat;
begin
  Left.GetNodeValue(Result);
  if Result.ResultType = rtError then
    exit;

  Right.GetNodeValue(RRes);
  if RRes.ResultType = rtError then
  begin
    Result := ErrorResult(RRes.ResError);
    exit;
  end;

  if (Result.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty]) and
     (RRes.ResultType in [rtFloat, rtInteger, rtDateTime, rtCell, rtEmpty])
  then
    try
      Result := FloatResult(Power(ArgToFloat(Result), ArgToFloat(RRes)));
    except
      on E: EInvalidArgument do Result := ErrorResult(errOverflow);
    end;
end;

function TsPowerExprNode.NodeType: TsResultType;
begin
  Result := rtFloat;
end;


{ TsConvertExprNode }

function TsConvertExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := Operand.AsRPNItem(ANext);
end;

function TsConvertExprNode.AsString: String;
begin
  Result := Operand.AsString;
end;


{ TsIntToFloatExprNode }

procedure TsConvertToIntExprNode.Check;
begin
  inherited Check;
  CheckNodeType(Operand, [rtInteger, rtCell])
end;

procedure TsIntToFloatExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  if Result.ResultType in [rtInteger, rtCell] then
    Result := FloatResult(ArgToInt(Result));
end;

function TsIntToFloatExprNode.NodeType: TsResultType;
begin
  Result := rtFloat;
end;


{ TsIntToDateTimeExprNode }

function TsIntToDateTimeExprNode.NodeType: TsResultType;
begin
  Result := rtDatetime;
end;

procedure TsIntToDateTimeExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetnodeValue(Result);
  if Result.ResultType in [rtInteger, rtCell] then
    Result := DateTimeResult(ArgToInt(Result));
end;


{ TsFloatToDateTimeExprNode }

procedure TsFloatToDateTimeExprNode.Check;
begin
  inherited Check;
  CheckNodeType(Operand, [rtFloat, rtCell]);
end;

function TsFloatToDateTimeExprNode.NodeType: TsResultType;
begin
  Result := rtDateTime;
end;

procedure TsFloatToDateTimeExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  if Result.ResultType in [rtFloat, rtCell] then
    Result := DateTimeResult(ArgToFloat(Result));
end;


{ TsIdentifierExprNode }

constructor TsIdentifierExprNode.CreateIdentifier(AID: TsExprIdentifierDef);
begin
  inherited Create;
  FID := AID;
  PResult := @FID.FValue;
  FResultType := FID.ResultType;
end;

function TsIdentifierExprNode.NodeType: TsResultType;
begin
  Result := FResultType;
end;

procedure TsIdentifierExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Result := PResult^;
  Result.ResultType := FResultType;
end;


{ TsVariableExprNode }

procedure TsVariableExprNode.Check;
begin
  // Do nothing;
end;

function TsVariableExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  RaiseParserError('Cannot handle variables for RPN, so far.');
end;

function TsVariableExprNode.AsString: string;
begin
  Result := FID.Name;
end;


{ TsFunctionExprNode }

constructor TsFunctionExprNode.CreateFunction(AID: TsExprIdentifierDef;
  const Args: TsExprArgumentArray);
begin
  inherited CreateIdentifier(AID);
  FArgumentNodes := Args;
  SetLength(FArgumentParams, Length(Args));
end;

destructor TsFunctionExprNode.Destroy;
var
  i: Integer;
begin
  for i:=0 to Length(FArgumentNodes)-1 do
    FreeAndNil(FArgumentNodes[i]);
  inherited Destroy;
end;

function TsFunctionExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
var
  i, n: Integer;
begin
  if FID.HasFixedArgumentCount then
    n := FID.ArgumentCount
  else
    n := Length(FArgumentNodes);
  Result := ANext;
//  for i:=Length(FArgumentNodes)-1 downto 0 do
  for i:=0 to High(FArgumentNodes) do
    Result := FArgumentNodes[i].AsRPNItem(Result);
  Result := RPNFunc(FID.Name, n, Result);
end;

function TsFunctionExprNode.AsString: String;
var
  S : String;
  i : Integer;
begin
  S := '';
  for i := 0 to Length(FArgumentNodes)-1 do
  begin
    if (S <> '') then
      S := S + ',';
    S := S + FArgumentNodes[i].AsString;
  end;
  S := '(' + S + ')';
  Result := FID.Name + S;
end;

procedure TsFunctionExprNode.CalcParams;
var
  i : Integer;
begin
  for i := 0 to Length(FArgumentParams)-1 do
  {
    case FArgumentParams[i].ResultType of
      rtEmpty: FID.FValue.ResultType := rtEmpty;
      rtError: if FID.FValue.ResultType <> rtError then
               begin
                 FID.FValue.ResultType := rtError;
                 FID.FValue.ResError := FArgumentParams[i].ResError;
               end;
      else     FArgumentNodes[i].GetNodeValue(FArgumentParams[i]);
    end;
    }
    FArgumentNodes[i].GetNodeValue(FArgumentParams[i]);
end;

procedure TsFunctionExprNode.Check;
var
  i: Integer;
  rta,                  // parameter types passed to the function
  rtp: TsResultType;    // Parameter types expected from the parameter symbol
  lastrtp: TsResultType;
begin
  if Length(FArgumentNodes) <> FID.ArgumentCount then
  begin
    for i:=Length(FArgumentNodes)+1 to FID.ArgumentCount do
      if not FID.IsOptionalArgument(i) then
        RaiseParserError(ErrInvalidArgumentCount, [FID.Name]);
  end;

  for i := 0 to Length(FArgumentNodes)-1 do
  begin
    rta := FArgumentNodes[i].NodeType;

    // A "cell" can return any type --> no type conversion required here.
    if rta = rtCell then
      Continue;

    if i+1 <= Length(FID.ParameterTypes) then
    begin
      rtp := CharToResultType(FID.ParameterTypes[i+1]);
      lastrtp := rtp;
    end else
      rtp := lastrtp;
    if rtp = rtAny then
      Continue;
    if (rtp <> rta) and not (rta in [rtCellRange, rtError, rtEmpty]) then
    begin
      // Automatically convert integers to floats in functions that return a float
      if (rta = rtInteger) and (rtp = rtFloat) then
      begin
        FArgumentNodes[i] := TsIntToFloatExprNode(FArgumentNodes[i]);
        exit;
      end;
      // Floats are truncated automatically to integers - that's what Excel does.
      if (rta = rtFloat) and (rtp = rtInteger) then
        exit;
      RaiseParserError(SErrInvalidArgumentType, [i+1, ResultTypeName(rtp), ResultTypeName(rta)])
    end;
  end;
end;


{ TsFunctionCallBackExprNode }

constructor TsFunctionCallBackExprNode.CreateFunction(AID: TsExprIdentifierDef;
  const Args: TsExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValueCallBack;
end;

procedure TsFunctionCallBackExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Result.ResultType := NodeType;     // was at end!
  if Length(FArgumentParams) > 0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
end;


{ TFPFunctionEventHandlerExprNode }

constructor TFPFunctionEventHandlerExprNode.CreateFunction(AID: TsExprIdentifierDef;
  const Args: TsExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValue;
end;

procedure TFPFunctionEventHandlerExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Result.ResultType := NodeType;    // was at end
  if Length(FArgumentParams) > 0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
end;


{ TsCellExprNode }

constructor TsCellExprNode.Create(AWorksheet: TsWorksheet; ACellString: String);
var
  r, c: Cardinal;
  flags: TsRelFlags;
begin
  ParseCellString(ACellString, r, c, flags);
  Create(AWorksheet, r, c, flags);
end;

constructor TsCellExprNode.Create(AWorksheet: TsWorksheet; ARow,ACol: Cardinal;
  AFlags: TsRelFlags);
begin
  FWorksheet := AWorksheet;
  FRow := ARow;
  FCol := ACol;
  FFlags := AFlags;
  FCell := AWorksheet.FindCell(FRow, FCol);
end;

function TsCellExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  if FIsRef then
    Result := RPNCellRef(FRow, FCol, FFlags, ANext)
  else
    Result := RPNCellValue(FRow, FCol, FFlags, ANext);
end;

function TsCellExprNode.AsString: string;
begin
  Result := GetCellString(FRow, FCol, FFlags);
end;

procedure TsCellExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  if (FCell <> nil) and HasFormula(FCell) then
    case FCell^.CalcState of
      csNotCalculated:
        Worksheet.CalcFormula(FCell);
      csCalculating:
        raise Exception.Create(SErrCircularReference);
    end;

  Result.ResultType := rtCell;
  Result.ResRow := FRow;
  Result.ResCol := FCol;
  Result.Worksheet := FWorksheet;
end;

procedure TsCellExprNode.Check;
begin
  // Nothing to check;
end;

function TsCellExprNode.NodeType: TsResultType;
begin
  Result := rtCell;
  {
  if FIsRef then
    Result := rtCell
  else
  begin
    Result := rtEmpty;
    if FCell <> nil then
      case FCell^.ContentType of
        cctNumber:
          if frac(FCell^.NumberValue) = 0 then
            Result := rtInteger
          else
            Result := rtFloat;
        cctDateTime:
          Result := rtDateTime;
        cctUTF8String:
          Result := rtString;
        cctBool:
          Result := rtBoolean;
        cctError:
          Result := rtError;
      end;
  end;
  }
end;


{ TsCellRangeExprNode }

constructor TsCellRangeExprNode.Create(AWorksheet: TsWorksheet; ACellRangeString: String);
var
  r1, c1, r2, c2: Cardinal;
  flags: TsRelFlags;
begin
  if pos(':', ACellRangeString) = 0 then
  begin
    ParseCellString(ACellRangeString, r1, c1, flags);
    if rfRelRow in flags then Include(flags, rfRelRow2);
    if rfRelCol in flags then Include(flags, rfRelCol2);
    Create(AWorksheet, r1, c1, r1, c1, flags);
  end else
  begin
    ParseCellRangeString(ACellRangeString, r1, c1, r2, c2, flags);
    Create(AWorksheet, r1, c1, r2, c2, flags);
  end;
end;

constructor TsCellRangeExprNode.Create(AWorksheet: TsWorksheet;
  ARow1,ACol1,ARow2,ACol2: Cardinal; AFlags: TsRelFlags);
begin
  FWorksheet := AWorksheet;
  FRow1 := ARow1;
  FCol1 := ACol1;
  FRow2 := ARow2;
  FCol2 := ACol2;
  FFlags := AFlags;
end;

function TsCellRangeExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  {
  if (FRow1 = FRow2) and (FCol1 = FCol2) then
    Result := RPNCellRef(FRow1, FCol1, FFlags, ANext)
  else
  }
    Result := RPNCellRange(FRow1, FCol1, FRow2, FCol2, FFlags, ANext);
end;

function TsCellRangeExprNode.AsString: string;
begin
  if (FRow1 = FRow2) and (FCol1 = FCol2) then
    Result := GetCellString(FRow1, FCol1, FFlags)
  else
    Result := GetCellRangeString(FRow1, FCol1, FRow2, FCol2, FFlags);
end;

procedure TsCellRangeExprNode.Check;
begin
  // Nothing to check;
end;

procedure TsCellRangeExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  r,c: Cardinal;
  cell: PCell;
begin
  for r := FRow1 to FRow2 do
    for c := FCol1 to FCol2 do
    begin
      cell := FWorksheet.FindCell(r, c);
      if HasFormula(cell) then
        case cell^.CalcState of
          csNotCalculated: FWorksheet.CalcFormula(cell);
          csCalculating  : raise Exception.Create(SErrCircularReference);
        end;
    end;

  Result.ResultType := rtCellRange;
  Result.ResCellRange.Row1 := FRow1;
  Result.ResCellRange.Col1 := FCol1;
  Result.ResCellRange.Row2 := FRow2;
  Result.ResCellRange.Col2 := FCol2;
  Result.Worksheet := FWorksheet;
end;

function TsCellRangeExprNode.NodeType: TsResultType;
begin
  Result := rtCellRange;
end;


{------------------------------------------------------------------------------}
{   Conversion of arguments to simple data types                               }
{------------------------------------------------------------------------------}

function ArgToBoolean(Arg: TsExpressionResult): Boolean;
var
  cell: PCell;
begin
  Result := false;
  if Arg.ResultType = rtBoolean then
    Result := Arg.ResBoolean
  else
  if (Arg.ResultType = rtCell) then begin
    cell := ArgToCell(Arg);
    if (cell <> nil) and (cell^.ContentType = cctBool) then
      Result := cell^.BoolValue;
  end;
end;

function ArgToCell(Arg: TsExpressionResult): PCell;
begin
  if Arg.ResultType = rtCell then
    Result := Arg.Worksheet.FindCell(Arg.ResRow, Arg.ResCol)
  else
    Result := nil;
end;

function ArgToInt(Arg: TsExpressionResult): Integer;
var
  cell: PCell;
begin
  Result := 0;
  if Arg.ResultType = rtInteger then
    result := Arg.ResInteger
  else
  if Arg.ResultType = rtFloat then
    result := trunc(Arg.ResFloat)
  else
  if Arg.ResultType = rtDateTime then
    result := trunc(Arg.ResDateTime)
  else
  if (Arg.ResultType = rtCell) then
  begin
    cell := ArgToCell(Arg);
    if Assigned(cell) and (cell^.ContentType = cctNumber) then
      result := trunc(cell^.NumberValue);
  end;
end;

function ArgToFloat(Arg: TsExpressionResult): TsExprFloat;
// Utility function for the built-in math functions. Accepts also integers
// in place of the floating point arguments. To be called in builtins or
// user-defined callbacks having float results.
var
  cell: PCell;
begin
  Result := 0.0;
  if Arg.ResultType = rtInteger then
    result := Arg.ResInteger
  else
  if Arg.ResultType = rtDateTime then
    result := Arg.ResDateTime
  else
  if Arg.ResultType = rtFloat then
    result := Arg.ResFloat
  else
  if (Arg.ResultType = rtCell) then
  begin
    cell := ArgToCell(Arg);
    if Assigned(cell) then
      case cell^.ContentType of
        cctNumber   : Result := cell^.NumberValue;
        cctDateTime : Result := cell^.DateTimeValue;
      end;
  end;
end;

function ArgToDateTime(Arg: TsExpressionResult): TDateTime;
var
  cell: PCell;
begin
  Result := 0.0;
  if Arg.ResultType = rtDateTime then
    result := Arg.ResDateTime
  else
  if Arg.ResultType = rtInteger then
    Result := Arg.ResInteger
  else
  if Arg.ResultType = rtFloat then
    Result := Arg.ResFloat
  else
  if (Arg.ResultType = rtCell) then
  begin
    cell := ArgToCell(Arg);
    if Assigned(cell) and (cell^.ContentType = cctDateTime) then
      Result := cell^.DateTimeValue;
  end;
end;

function ArgToString(Arg: TsExpressionResult): String;
var
  cell: PCell;
begin
  Result := '';
  case Arg.ResultType of
    rtString  : result := Arg.ResString;
    rtInteger : Result := IntToStr(Arg.ResInteger);
    rtFloat   : Result := FloatToStr(Arg.ResFloat);
    rtCell    : begin
                  cell := ArgToCell(Arg);
                  if Assigned(cell) and (cell^.ContentType = cctUTF8String) then
                    Result := cell^.UTF8Stringvalue;
                end;
  end;
end;


{------------------------------------------------------------------------------}
{   Conversion simple data types to ExpressionResults                          }
{------------------------------------------------------------------------------}

function BooleanResult(AValue: Boolean): TsExpressionResult;
begin
  Result.ResultType := rtBoolean;
  Result.ResBoolean := AValue;
end;

function DateTimeResult(AValue: TDateTime): TsExpressionResult;
begin
  Result.ResultType := rtDateTime;
  Result.ResDateTime := AValue;
end;

function EmptyResult: TsExpressionResult;
begin
  Result.ResultType := rtEmpty;
end;

function ErrorResult(const AValue: TsErrorValue): TsExpressionResult;
begin
  Result.ResultType := rtError;
  Result.ResError := AValue;
end;

function FloatResult(const AValue: TsExprFloat): TsExpressionResult;
begin
  Result.ResultType := rtFloat;
  Result.ResFloat := AValue;
end;

function IntegerResult(const AValue: Integer): TsExpressionResult;
begin
  Result.ResultType := rtInteger;
  Result.ResInteger := AValue;
end;

function StringResult(const AValue: string): TsExpressionResult;
begin
  Result.ResultType := rtString;
  Result.ResString := AValue;
end;


{------------------------------------------------------------------------------}
{  Standard Builtins support                                                   }
{------------------------------------------------------------------------------}

// Builtin math functions

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

procedure fpsEXP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Result := FloatResult(exp(ArgToFloat(Args[0])));
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


// Builtin date/time functions

procedure fpsDATE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DATE( year, month, day )
begin
  Result := DateTimeResult(
    EncodeDate(ArgToInt(Args[0]), ArgToInt(Args[1]), ArgToInt(Args[2]))
  );
end;

procedure fpsDATEDIF(var Result: TsExpressionResult; const Args: TsExprParameterArray);
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


// Builtin string functions

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

procedure fpsLEFT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LEFT( text, [number_of_characters] )
// extracts a substring from a string, starting from the left-most character
var
  s: String;
  count: Integer;
begin
  s := Args[0].ResString;
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
  Result := IntegerResult(UTF8Length(Args[0].ResString));
end;

procedure fpsLOWER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LOWER( text )
// converts all letters in the specified string to lowercase. If there are
// characters in the string that are not letters, they are not affected.
begin
  Result := StringResult(UTF8Lowercase(Args[0].ResString));
end;

procedure fpsMID(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MID( text, start_position, number_of_characters )
// extracts a substring from a string (starting at any position).
begin
  Result := StringResult(UTF8Copy(Args[0].ResString, ArgToInt(Args[1]), ArgToInt(Args[2])));
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

procedure fpsRIGHT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// RIGHT( text, [number_of_characters] )
// extracts a substring from a string, starting from the last character
var
  s: String;
  count: Integer;
begin
  s := Args[0].ResString;
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
  s := Args[0].ResString;
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
  Result := StringResult(UTF8Trim(Args[0].ResString));
end;

procedure fpsUPPER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// UPPER( text )
// converts all letters in the specified string to uppercase. If there are
// characters in the string that are not letters, they are not affected.
begin
  Result := StringResult(UTF8Uppercase(Args[0].ResString));
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


{ Builtin logical functions }

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


{ Builtin statistical functions }

procedure ArgsToFloatArray(const Args: TsExprParameterArray; out AData: TsExprFloatArray);
const
  BLOCKSIZE = 128;
var
  i, n: Integer;
  r, c: Cardinal;
  cell: PCell;
  arg: TsExpressionResult;
begin
  SetLength(AData, BLOCKSIZE);
  n := 0;
  for i:=0 to High(Args) do
  begin
    arg := Args[i];
    if arg.ResultType = rtCellRange then
      for r := arg.ResCellRange.Row1 to arg.ResCellRange.Row2 do
        for c := arg.ResCellRange.Col1 to arg.ResCellRange.Col2 do
        begin
          cell := arg.Worksheet.FindCell(r, c);
          if (cell <> nil) and (cell^.ContentType in [cctNumber, cctDateTime]) then
          begin
            case cell^.ContentType of
              cctNumber   : AData[n] := cell^.NumberValue;
              cctDateTime : AData[n] := cell^.DateTimeValue
            end;
            inc(n);
            if n = Length(AData) then SetLength(AData, length(AData) + BLOCKSIZE);
          end;
        end
    else
    if (arg.ResultType in [rtInteger, rtFloat, rtDateTime, rtCell]) then
    begin
      AData[n] := ArgToFloat(arg);
      inc(n);
      if n = Length(AData) then SetLength(AData, Length(AData) + BLOCKSIZE);
    end;
  end;
  SetLength(AData, n);
end;


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
  arg: TsExpressionResult;
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


{ Builtin info functions }

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
  r1,r2, c1,c2: Cardinal;
  cell: PCell;
  res: TsExpressionResult;
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
        r2 := r1;
        c2 := c1;
      end;
    rtCellRange:
      begin
        r1 := Args[1].ResCellRange.Row1;
        r2 := Args[1].ResCellRange.Row2;
        c1 := Args[1].ResCellRange.Col1;
        c2 := Args[1].ResCellRange.Col2;
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
{@@
  Registers a non-built-in function:

  @param AName        Name of the function as used for calling it in the spreadsheet
  @param AResultType  A character classifying the data type of the function result:
                        'I' integer
                        'F' floating point number
                        'D' date/time value
                        'S' string
                        'B' boolean value (TRUE/FALSE)
                        'R' cell range, can also be used for functions requiring
                            a cell "reference", like "CELL(..)"
  @param AParamTypes A string with result type symbols for each parameter of the
                     function. Symbols as used for "ResultType" with these
                     additions:
                       - Use a lower-case character if a parameter is optional.
                         (must be at the end of the string)
                       - Add "+" if the last parameter type is valid for a variable
                         parameter count (Excel does pose a limit of 30, though).
                       - Use "?" if the data type should not be checked.
  @param AExcelCode  ID of the function needed in the xls biff file. Please see
                     the "OpenOffice Documentation of Microsoft Excel File Format"
                     section 3.11.
  @param ACallBack   Address of the procedure called when the formula is
                     calculated.
}
{------------------------------------------------------------------------------}
procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallback: TsExprFunctionCallBack);
begin
  with BuiltinIdentifiers do
    AddFunction(bcUser, AName, AResultType, AParamTypes, AExcelCode, ACallBack);
end;

{@@
  Registers the built-in functions. Called automatically.
}
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
    AddFunction(cat, 'COS',       'F', 'F',    INT_EXCEL_SHEET_FUNC_COS,        @fpsCOS);
    AddFunction(cat, 'COSH',      'F', 'F',    INT_EXCEL_SHEET_FUNC_COSH,       @fpsCOSH);
    AddFunction(cat, 'DEGREES',   'F', 'F',    INT_EXCEL_SHEET_FUNC_DEGREES,    @fpsDEGREES);
    AddFunction(cat, 'EXP',       'F', 'F',    INT_EXCEL_SHEET_FUNC_EXP,        @fpsEXP);
    AddFunction(cat, 'INT',       'I', 'F',    INT_EXCEL_SHEET_FUNC_INT,        @fpsINT);
    AddFunction(cat, 'LN',        'F', 'F',    INT_EXCEL_SHEET_FUNC_LN,         @fpsLN);
    AddFunction(cat, 'LOG',       'F', 'Ff',   INT_EXCEL_SHEET_FUNC_LOG,        @fpsLOG);
    AddFunction(cat, 'LOG10',     'F', 'F',    INT_EXCEL_SHEET_FUNC_LOG10,      @fpsLOG10);
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
    AddFunction(cat, 'LEFT',      'S', 'Si',   INT_EXCEL_SHEET_FUNC_LEFT,       @fpsLEFT);
    AddFunction(cat, 'LEN',       'I', 'S',    INT_EXCEL_SHEET_FUNC_LEN,        @fpsLEN);
    AddFunction(cat, 'LOWER',     'S', 'S',    INT_EXCEL_SHEET_FUNC_LOWER,      @fpsLOWER);
    AddFunction(cat, 'MID',       'S', 'SII',  INT_EXCEL_SHEET_FUNC_MID,        @fpsMID);
    AddFunction(cat, 'REPLACE',   'S', 'SIIS', INT_EXCEL_SHEET_FUNC_REPLACE,    @fpsREPLACE);
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
    AddFunction(cat, 'AVEDEV',    'F', '?+',   INT_EXCEL_SHEET_FUNC_AVEDEV,     @fpsAVEDEV);
    AddFunction(cat, 'AVERAGE',   'F', '?+',   INT_EXCEL_SHEET_FUNC_AVERAGE,    @fpsAVERAGE);
    AddFunction(cat, 'COUNT',     'I', '?+',   INT_EXCEL_SHEET_FUNC_COUNT,      @fpsCOUNT);
    AddFunction(cat, 'COUNTA',    'I', '?+',   INT_EXCEL_SHEET_FUNC_COUNTA,     @fpsCOUNTA);
    AddFunction(cat, 'COUNTBLANK','I', 'R',    INT_EXCEL_SHEET_FUNC_COUNTBLANK, @fpsCOUNTBLANK);
    AddFunction(cat, 'MAX',       'F', '?+',   INT_EXCEL_SHEET_FUNC_MAX,        @fpsMAX);
    AddFunction(cat, 'MIN',       'F', '?+',   INT_EXCEL_SHEET_FUNC_MIN,        @fpsMIN);
    AddFunction(cat, 'PRODUCT',   'F', '?+',   INT_EXCEL_SHEET_FUNC_PRODUCT,    @fpsPRODUCT);
    AddFunction(cat, 'STDEV',     'F', '?+',   INT_EXCEL_SHEET_FUNC_STDEV,      @fpsSTDEV);
    AddFunction(cat, 'STDEVP',    'F', '?+',   INT_EXCEL_SHEET_FUNC_STDEVP,     @fpsSTDEVP);
    AddFunction(cat, 'SUM',       'F', '?+',   INT_EXCEL_SHEET_FUNC_SUM,        @fpsSUM);
    AddFunction(cat, 'SUMSQ',     'F', '?+',   INT_EXCEL_SHEET_FUNC_SUMSQ,      @fpsSUMSQ);
    AddFunction(cat, 'VAR',       'F', '?+',   INT_EXCEL_SHEET_FUNC_VAR,        @fpsVAR);
    AddFunction(cat, 'VARP',      'F', '?+',   INT_EXCEL_SHEET_FUNC_VARP,       @fpsVARP);
    // to do: CountIF, SUMIF

    // Info functions
    cat := bcInfo;
    AddFunction(cat, 'CELL',      '?', 'Sr',   INT_EXCEL_SHEET_FUNC_CELL,       @fpsCELL);
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

{ TsBuiltInExprIdentifierDef }

procedure TsBuiltInExprIdentifierDef.Assign(Source: TPersistent);
begin
  inherited Assign(Source);
  if Source is TsBuiltInExprIdentifierDef then
    FCategory := (Source as TsBuiltInExprIdentifierDef).Category;
end;

initialization
  ExprFormatSettings := DefaultFormatSettings;
  ExprFormatSettings.DecimalSeparator := '.';
  ExprFormatSettings.ListSeparator := ',';

  RegisterStdBuiltins(BuiltinIdentifiers);

finalization
  FreeBuiltins;
end.
