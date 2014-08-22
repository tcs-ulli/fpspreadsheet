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
    - remove and modifiy built-in function such that the parser is compatible
      with Excel syntax (and OpenOffice - which is the same).

 ******************************************************************************}
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
  TTokenType = (
    ttPlus, ttMinus, ttMul, ttDiv, ttConcat, ttPercent, ttPower, ttLeft, ttRight,
    ttLessThan, ttLargerThan, ttEqual, ttNotEqual, ttLessThanEqual, ttLargerThanEqual,
    ttNumber, ttString, ttIdentifier, ttCell, ttCellRange,
    ttComma, ttAnd, ttOr, ttXor, ttTrue, ttFalse, ttNot, ttIf,
    ttEOF
  );

  TExprFloat    = Double;

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

  { TsExpressionScanner }
  TsExpressionScanner = class(TObject)
    FSource : String;
    LSource,
    FPos: Integer;
    FChar: PChar;
    FToken: String;
    FTokenType: TTokenType;
  private
    function GetCurrentChar: Char;
    procedure ScanError(Msg: String);
  protected
    procedure SetSource(const AValue: String); virtual;
    function DoIdentifier: TTokenType;
    function DoNumber: TTokenType;
    function DoDelimiter: TTokenType;
    function DoString: TTokenType;
    function NextPos: Char; // inline;
    procedure SkipWhiteSpace; // inline;
    function IsWordDelim(C: Char): Boolean; // inline;
    function IsDelim(C: Char): Boolean; // inline;
    function IsDigit(C: Char): Boolean; // inline;
    function IsAlpha(C: Char): Boolean; // inline;
  public
    constructor Create;
    function GetToken: TTokenType;
    property Token: String read FToken;
    property TokenType: TTokenType read FTokenType;
    property Source: String read FSource write SetSource;
    property Pos: Integer read FPos;
    property CurrentChar: Char read GetCurrentChar;
  end;

  EExprScanner = class(Exception);

  TsResultType = (rtBoolean, rtInteger, rtFloat, rtDateTime, rtString);
  TsResultTypes = set of TsResultType;

  TsExpressionResult = record
    ResString    : String;
    case ResultType : TsResultType of
      rtBoolean  : (ResBoolean  : Boolean);
      rtInteger  : (ResInteger  : Int64);
      rtFloat    : (ResFloat    : TExprFloat);
      rtDateTime : (ResDateTime : TDatetime);
      rtString   : ();
  end;
  PsExpressionResult = ^TsExpressionResult;
  TExprParameterArray = array of TsExpressionResult;

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

  TExprArgumentArray = array of TsExprNode;

  { TsBinaryOperationExprNode }
  TsBinaryOperationExprNode = class(TsExprNode)
  private
    FLeft: TsExprNode;
    FRight: TsExprNode;
  protected
    procedure CheckSameNodeTypes;
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

  { TsBinaryAndExprNode }
  TsBinaryAndExprNode = class(TsBooleanOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsBinaryOrExprNode }
  TsBinaryOrExprNode = class(TsBooleanOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsBinaryXorExprNode }
  TsBinaryXorExprNode = class(TsBooleanOperationExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsBooleanResultExprNode }
  TsBooleanResultExprNode = class(TsBinaryOperationExprNode)
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
    procedure Check; override;
  end;

  { TsLessThanExprNode }
  TsLessThanExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterThanExprNode }
  TsGreaterThanExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsLessEqualExprNode }
  TsLessEqualExprNode = class(TsGreaterThanExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterEqualExprNode }
  TsGreaterEqualExprNode = class(TsLessThanExprNode)
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TIfExprNode }
  TIfExprNode = class(TsBinaryOperationExprNode)
  private
    FCondition: TsExprNode;
  protected
    procedure GetNodeValue(var Result: TsExpressionResult); override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  public
    constructor Create(ACondition, ALeft, ARight: TsExprNode);
    destructor Destroy; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
    property Condition: TsExprNode read FCondition;
  end;

  { TsConcatExprNode }
  TsConcatExprNode = class(TsBinaryOperationExprNode)
  protected
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
    function NodeType: TsResultType; override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsMathOperationExprNode }
  TsMathOperationExprNode = class(TsBinaryOperationExprNode)
  protected
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
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsMultiplyExprNode }
  TsMultiplyExprNode = class(TsMathOperationExprNode)
  protected
    procedure check; override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsDivideExprNode }
  TsDivideExprNode = class(TsMathOperationExprNode)
  protected
    Procedure Check; override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
    function NodeType: TsResultType; override;
  end;

  { TsUnaryOperationExprNode }
  TsUnaryOperationExprNode = class(TsExprNode)
  private
    FOperand: TsExprNode;
  public
    constructor Create(AOperand: TsExprNode);
    destructor Destroy; override;
    procedure Check; override;
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
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsConvertToIntExprNode }
  TsConvertToIntExprNode = class(TsConvertExprNode)
  protected
    procedure Check; override;
  end;

  { TsIntToFloatExprNode }
  TsIntToFloatExprNode = class(TsConvertToIntExprNode)
  public
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsIntToDateTimeExprNode }
  TsIntToDateTimeExprNode = class(TsConvertToIntExprNode)
  public
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsFloatToDateTimeExprNode }
  TsFloatToDateTimeExprNode = class(TsConvertExprNode)
  protected
    procedure Check; override;
  public
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsNegateExprNode }
  TsNegateExprNode = class(TsUnaryOperationExprNode)
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsPercentExprNode }
  TsPercentExprNode = class(TsUnaryOperationExprNode)
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsParenthesisExprNode }
  TsParenthesisExprNode = class(TsUnaryOperationExprNode)
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

  { TsConstExprNode }
  TsConstExprNode = Class(TsExprNode)
  private
    FValue: TsExpressionResult;
  public
    constructor CreateString(AValue: String);
    constructor CreateInteger(AValue: Int64);
    constructor CreateDateTime(AValue: TDateTime);
    constructor CreateFloat(AValue: TExprFloat);
    constructor CreateBoolean(AValue: Boolean);
    function AsString: string; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    procedure Check; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
    function NodeType : TsResultType; override;
    // For inspection
    property ConstValue: TsExpressionResult read FValue;
  end;

  TsExprIdentifierType = (itVariable, itFunctionCallBack, itFunctionHandler);

  TsExprFunctionCallBack = procedure (var Result: TsExpressionResult;
      const Args: TExprParameterArray);

  TsExprFunctionEvent = procedure (var Result: TsExpressionResult;
      const Args: TExprParameterArray) of object;

  { TsExprIdentifierDef }
  TsExprIdentifierDef = class(TCollectionItem)
  private
    FStringValue: String;
    FValue: TsExpressionResult;
    FArgumentTypes: String;
    FIDType: TsExprIdentifierType;
    FName: ShortString;
    FOnGetValue: TsExprFunctionEvent;
    FOnGetValueCB: TsExprFunctionCallBack;
    function GetAsBoolean: Boolean;
    function GetAsDateTime: TDateTime;
    function GetAsFloat: TExprFloat;
    function GetAsInteger: Int64;
    function GetAsString: String;
    function GetResultType: TsResultType;
    function GetValue: String;
    procedure SetArgumentTypes(const AValue: String);
    procedure SetAsBoolean(const AValue: Boolean);
    procedure SetAsDateTime(const AValue: TDateTime);
    procedure SetAsFloat(const AValue: TExprFloat);
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
    property AsFloat: TExprFloat Read GetAsFloat Write SetAsFloat;
    property AsInteger: Int64 Read GetAsInteger Write SetAsInteger;
    property AsString: String Read GetAsString Write SetAsString;
    property AsBoolean: Boolean Read GetAsBoolean Write SetAsBoolean;
    property AsDateTime: TDateTime Read GetAsDateTime Write SetAsDateTime;
    property OnGetFunctionValueCallBack: TsExprFunctionCallBack read FOnGetValueCB write FOnGetValueCB;
  published
    property IdentifierType: TsExprIdentifierType read FIDType write FIDType;
    property Name: ShortString read FName write SetName;
    property Value: String read GetValue write SetValue;
    property ParameterTypes: String read FArgumentTypes write SetArgumentTypes;
    property ResultType: TsResultType read GetResultType write SetResultType;
    property OnGetFunctionValue: TsExprFunctionEvent read FOnGetValue write FOnGetValue;
  end;

  TsBuiltInExprCategory = (bcStrings, bcDateTime, bcMath, bcBoolean, bcConversion,
    bcData, bcVaria, bcUser);

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
    function IndexOfIdentifier(const AName: ShortString): Integer;
    function FindIdentifier(const AName: ShortString): TsExprIdentifierDef;
    function IdentifierByName(const AName: ShortString): TsExprIdentifierDef;
    function AddVariable(const AName: ShortString; AResultType: TsResultType;
      AValue: String): TsExprIdentifierDef;
    function AddBooleanVariable(const AName: ShortString;
      AValue: Boolean): TsExprIdentifierDef;
    function AddIntegerVariable(const AName: ShortString;
      AValue: Integer): TsExprIdentifierDef;
    function AddFloatVariable(const AName: ShortString;
      AValue: TExprFloat): TsExprIdentifierDef;
    function AddStringVariable(const AName: ShortString;
      AValue: String): TsExprIdentifierDef;
    function AddDateTimeVariable(const AName: ShortString;
      AValue: TDateTime): TsExprIdentifierDef;
    function AddFunction(const AName: ShortString; const AResultType: Char;
      const AParamTypes: String; ACallBack: TsExprFunctionCallBack): TsExprIdentifierDef;
    function AddFunction(const AName: ShortString; const AResultType: Char;
      const AParamTypes: String; ACallBack: TsExprFunctionEvent): TsExprIdentifierDef;
    property Identifiers[AIndex: Integer]: TsExprIdentifierDef read GetI write SetI; default;
  end;

  { TsIdentifierExprNode }
  TsIdentifierExprNode = class(TsExprNode)
  private
    FID: TsExprIdentifierDef;
    PResult: PsExpressionResult;
    FResultType: TsResultType;
  public
    constructor CreateIdentifier(AID: TsExprIdentifierDef);
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
    property Identifier: TsExprIdentifierDef read FID;
  end;

  { TFPExprVariable }
  TFPExprVariable = class(TsIdentifierExprNode)
    procedure Check; override;
    function AsString: string; override;
    Function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
  end;

  { TFPExprFunction }
  TFPExprFunction = class(TsIdentifierExprNode)
  private
    FArgumentNodes: TExprArgumentArray;
    FargumentParams: TExprParameterArray;
  protected
    procedure CalcParams;
    procedure Check; override;
  public
    constructor CreateFunction(AID: TsExprIdentifierDef;
      const Args: TExprArgumentArray); virtual;
    destructor Destroy; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    property ArgumentNodes: TExprArgumentArray read FArgumentNodes;
    property ArgumentParams: TExprParameterArray read FArgumentParams;
  end;

  { TFPFunctionCallBack }
  TFPFunctionCallBack = class(TFPExprFunction)
  private
    FCallBack: TsExprFunctionCallBack;
  public
    constructor CreateFunction(AID: TsExprIdentifierDef;
      const Args: TExprArgumentArray); override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
    property CallBack: TsExprFunctionCallBack read FCallBack;
  end;

  { TFPFunctionEventHandler }
  TFPFunctionEventHandler = class(TFPExprFunction)
  private
    FCallBack: TsExprFunctionEvent;
  public
    constructor CreateFunction(AID: TsExprIdentifierDef;
      const Args: TExprArgumentArray); override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
    property CallBack: TsExprFunctionEvent read FCallBack;
  end;

  { TsCellExprNode }
  TsCellExprNode = class(TsExprNode)
  private
    FWorksheet: TsWorksheet;
    FCell: PCell;
    FFlags: TsRelFlags;
  public
    constructor Create(AWorksheet: TsWorksheet; ACellString: String); overload;
    constructor Create(AWorksheet: TsWorksheet; ACell: PCell; AFlags: TsRelFlags); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
    procedure GetNodeValue(var Result: TsExpressionResult); override;
  end;

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
    function ConvertNode(Todo: TsExprNode; ToType: TsResultType): TsExprNode;
    function GetAsBoolean: Boolean;
    function GetAsDateTime: TDateTime;
    function GetAsFloat: TExprFloat;
    function GetAsInteger: Int64;
    function GetAsString: String;
    function MatchNodes(Todo, Match: TsExprNode): TsExprNode;
    procedure CheckNodes(var ALeft, ARight: TsExprNode);
    procedure SetBuiltIns(const AValue: TsBuiltInExprCategories);
    procedure SetIdentifiers(const AValue: TsExprIdentifierDefs);
  protected
    procedure ParserError(Msg: String);
    procedure SetExpression(const AValue: String); virtual;
    procedure CheckResultType(const Res: TsExpressionResult;
      AType: TsResultType); inline;
    class function BuiltinExpressionManager: TsBuiltInExpressionManager;
    function Level1: TsExprNode;
    function Level2: TsExprNode;
    function Level3: TsExprNode;
    function Level4: TsExprNode;
    function Level5: TsExprNode;
    function Level6: TsExprNode;
    function Primitive: TsExprNode;
    function GetToken: TTokenType;
    function TokenType: TTokenType;
    function CurrentToken: String;
    procedure CreateHashList;
    property Scanner: TsExpressionScanner read FScanner;
    property ExprNode: TsExprNode read FExprNode;
    property Dirty: Boolean read FDirty;
  public
    constructor Create(AWorksheet: TsWorksheet);
    destructor Destroy; override;
    function IdentifierByName(AName: ShortString): TsExprIdentifierDef; virtual;
    procedure Clear;
    function BuildRPNFormula: TsRPNFormula;
    function BuildFormula: String;
    procedure EvaluateExpression(var Result: TsExpressionResult);
    function Evaluate: TsExpressionResult;
    function ResultType: TsResultType;
    property AsFloat: TExprFloat read GetAsFloat;
    property AsInteger: Int64 read GetAsInteger;
    property AsString: String read GetAsString;
    property AsBoolean: Boolean read GetAsBoolean;
    property AsDateTime: TDateTime read GetAsDateTime;
    property Worksheet: TsWorksheet read FWorksheet;
    // The expression to parse
    property Expression: String read FExpression write SetExpression;
    property Identifiers: TsExprIdentifierDefs read FIdentifiers write SetIdentifiers;
    property BuiltIns: TsBuiltInExprCategories read FBuiltIns write SetBuiltIns;
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
    function IdentifierByName(const AName: ShortString): TsBuiltInExprIdentifierDef;
    function AddVariable(const ACategory: TsBuiltInExprCategory; const AName: ShortString;
      AResultType: TsResultType; AValue: String): TsBuiltInExprIdentifierDef;
    function AddBooleanVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: Boolean): TsBuiltInExprIdentifierDef;
    function AddIntegerVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: Integer): TsBuiltInExprIdentifierDef;
    function AddFloatVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: TExprFloat): TsBuiltInExprIdentifierDef;
    function AddStringVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: String): TsBuiltInExprIdentifierDef;
    function AddDateTimeVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: TDateTime): TsBuiltInExprIdentifierDef;
    function AddFunction(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; const AResultType: Char; const AParamTypes: String;
      ACallBack: TsExprFunctionCallBack): TsBuiltInExprIdentifierDef;
    function AddFunction(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; const AResultType: Char; const AParamTypes: String;
      ACallBack: TsExprFunctionEvent): TsBuiltInExprIdentifierDef;
    property IdentifierCount: Integer read GetCount;
    property Identifiers[AIndex: Integer]: TsBuiltInExprIdentifierDef read GetI;
  end;

  EExprParser = class(Exception);

function TokenName(AToken: TTokenType): String;
function ResultTypeName(AResult: TsResultType): String;
function CharToResultType(C: Char): TsResultType;
function BuiltinIdentifiers: TsBuiltInExpressionManager;
procedure RegisterStdBuiltins(AManager: TsBuiltInExpressionManager);
function ArgToFloat(Arg: TsExpressionResult): TExprFloat;

const
  AllBuiltIns = [bcStrings, bcDateTime, bcMath, bcBoolean, bcConversion,
                 bcData, bcVaria, bcUser];


implementation

uses
  typinfo, fpsutils;

const
  cNull = #0;
  cSingleQuote = '''';
  cDoubleQuote = '"';

  Digits        = ['0'..'9', '.'];
  WhiteSpace    = [' ', #13, #10, #9];
  Operators     = ['+', '-', '<', '>', '=', '/', '*', '&', '%'];
  Delimiters    = Operators + [',', '(', ')'];
  Symbols       = ['^'] + Delimiters;
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
  SErrNoNOTOperation = 'Cannot perform NOT operation on expression of type %s: %s';
  SErrNoPercentOperation = 'Cannot perform percent operation on expression of type %s: %s';
  SErrNoXOROperationRPN = 'Cannot create RPN item for "xor" expression';
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
  SErrIFNeedsBoolean = 'First argument to IF must be of type boolean: %s';
  SErrCaseLabelNotAConst = 'Case label %d "%s" is not a constant expression';
  SErrCaseLabelType = 'Case label %d "%s" needs type %s, but has type %s';
  SErrCaseValueType = 'Case value %d "%s" needs type %s, but has type %s';
  SErrNoCellOperand = 'Cell operand is NIL.';
  SErrCellError = 'Cell %s contains an error.';

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

function TokenName (AToken: TTokenType): String;
begin
  Result := GetEnumName(TypeInfo(TTokenType), ord(AToken));
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

function TsExpressionScanner.IsAlpha(C: Char): Boolean;
begin
  Result := C in ['A'..'Z', 'a'..'z'];
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

function TsExpressionScanner.NextPos: Char;
begin
  Inc(FPos);
  Inc(FChar);
  Result := FChar^;
end;

function TsExpressionScanner.IsWordDelim(C: Char): Boolean;
begin
  Result := C in WordDelimiters;
end;

function TsExpressionScanner.IsDelim(C: Char): Boolean;
begin
  Result := C in Delimiters;
end;

function TsExpressionScanner.IsDigit(C: Char): Boolean;
begin
  Result := C in Digits;
end;

procedure TsExpressionScanner.SkipWhiteSpace;
begin
  while (FChar^ in WhiteSpace) and (FPos <= LSource) do
    NextPos;
end;

function TsExpressionScanner.DoDelimiter: TTokenType;
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

procedure TsExpressionScanner.ScanError(Msg: String);
begin
  raise EExprScanner.Create(Msg)
end;

function TsExpressionScanner.DoString: TTokenType;

  function TerminatingChar(C: Char): boolean;
  begin
    Result := (C = cNull)
           or ((C = cSingleQuote) and
                not ((FPos < LSource) and (FSource[FPos+1] = cSingleQuote)))
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
    if C = cSingleQuote then
      NextPos;
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

function TsExpressionScanner.DoNumber: TTokenType;
var
  C: Char;
  X: TExprFloat;
  I: Integer;
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
  val(FToken, X, I);
  if (I <> 0) then
    ScanError(Format(SErrInvalidNumber, [FToken]));
  Result := ttNumber;
end;

function TsExpressionScanner.DoIdentifier: TTokenType;
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
  if (S = 'or') then
    Result := ttOr
  else if (S = 'xor') then
    Result := ttXOr
  else if (S = 'and') then
    Result := ttAnd
  else if (S = 'true') then
    Result := ttTrue
  else if (S = 'false') then
    Result := ttFalse
  else if (S = 'not') then
    Result := ttNot
  else if (S = 'if') then
    Result := ttIF
  else if ParseCellString(S, row, col, flags) then
    Result := ttCell
  else if ParseCellRangeString(S, row, col, row2, col2, flags) then
    Result := ttCellRange
  else
    Result := ttIdentifier;
end;

function TsExpressionScanner.GetToken: TTokenType;
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
  else if (C = cSingleQuote) or (C = cDoubleQuote) then
    Result := DoString
  else if IsDigit(C) then
    Result := DoNumber
  else if IsAlpha(C) then
    Result := DoIdentifier
  else
    ScanError(Format(SErrUnknownCharacter, [FPos, C]));
  FTokenType := Result;
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

function TsExpressionParser.BuildRPNFormula: TsRPNFormula;
begin
  Result := CreateRPNFormula(FExprNode.AsRPNItem(nil), true);
end;

function TsExpressionParser.BuildFormula: String;
begin
  Result := FExprNode.AsString;
end;

procedure TsExpressionParser.Clear;
begin
  FExpression := '';
  FHashList.Clear;
  FExprNode.Free;
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
        FHashList.Add(LowerCase(BID.Name), BID);
    end;
  // User
  for i:=0 to FIdentifiers.Count-1 do
  begin
    ID := FIdentifiers[i];
    FHashList.Add(LowerCase(ID.Name), ID);
  end;
  FDirty := False;
end;

function TsExpressionParser.CurrentToken: String;
begin
  Result := FScanner.Token;
end;

function TsExpressionParser.TokenType: TTokenType;
begin
  Result := FScanner.TokenType;
end;

function TsExpressionParser.IdentifierByName(AName: ShortString): TsExprIdentifierDef;
begin
  If FDirty then
    CreateHashList;
  Result := TsExprIdentifierDef(FHashList.Find(LowerCase(AName)));
end;

function TsExpressionParser.GetToken: TTokenType;
begin
  Result := FScanner.GetToken;
end;

procedure TsExpressionParser.CheckEOF;
begin
  if (TokenType = ttEOF) then
    ParserError(SErrUnexpectedEndOfExpression);
end;

procedure TsExpressionParser.SetIdentifiers(const AValue: TsExprIdentifierDefs);
begin
  FIdentifiers.Assign(AValue)
end;

procedure TsExpressionParser.EvaluateExpression(var Result: TsExpressionResult);
begin
  if (FExpression = '') then
    ParserError(SErrInExpressionEmpty);
  if not Assigned(FExprNode) then
    ParserError(SErrInExpression);
  FExprNode.GetNodeValue(Result);
end;

procedure TsExpressionParser.ParserError(Msg: String);
begin
  raise EExprParser.Create(Msg);
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

function TsExpressionParser.GetAsFloat: TExprFloat;
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

{
  Checks types of todo and match. If ToDO can be converted to it matches
  the type of match, then a node is inserted.
  For binary operations, this function is called for both operands.
}
function TsExpressionParser.MatchNodes(ToDo, Match: TsExprNode): TsExprNode;
Var
  TT, MT : TsResultType;
begin
  Result := ToDo;
  TT := ToDo.NodeType;
  MT := Match.NodeType;
  If TT <> MT then
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

{
  If the result types differ, they are converted to a common type if possible.
}
procedure TsExpressionParser.CheckNodes(var ALeft, ARight: TsExprNode);
begin
  ALeft := MatchNodes(ALeft, ARight);
  ARight := MatchNodes(ARight, ALeft);
end;

procedure TsExpressionParser.SetBuiltIns(const AValue: TsBuiltInExprCategories);
begin
  if FBuiltIns = AValue then
    exit;
  FBuiltIns := AValue;
  FDirty := true;
end;

function TsExpressionParser.Level1: TsExprNode;
var
  tt: TTokenType;
  Right: TsExprNode;
begin
{$ifdef debugexpr}Writeln('Level 1 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  if TokenType = ttNot then
  begin
    GetToken;
    CheckEOF;
    Right := Level2;
    Result := TsNotExprNode.Create(Right);
  end
  else
    Result := Level2;

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
end;

function TsExpressionParser.Level2: TsExprNode;
var
  right: TsExprNode;
  tt: TTokenType;
  C: TsBinaryOperationExprNodeClass;
begin
{$ifdef debugexpr}  Writeln('Level 2 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
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
        ttLessthan         : C := TsLessThanExprNode;
        ttLessthanEqual    : C := TsLessEqualExprNode;
        ttLargerThan       : C := TsGreaterThanExprNode;
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
  tt: TTokenType;
  right: TsExprNode;
begin
{$ifdef debugexpr}  Writeln('Level 3 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
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
  tt: TTokenType;
  right: TsExprNode;
begin
{$ifdef debugexpr}  Writeln('Level 4 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
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
  B: Boolean;
begin
{$ifdef debugexpr}  Writeln('Level 5 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  B := false;
  if (TokenType in [ttPlus, ttMinus]) then
  begin
    B := (TokenType = ttMinus);
    GetToken;
  end;
  Result := Level6;
  if B then
    Result := TsNegateExprNode.Create(Result);
end;

function TsExpressionParser.Level6: TsExprNode;
begin
{$ifdef debugexpr}  Writeln('Level 6 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
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
end;

function TsExpressionParser.Primitive: TsExprNode;
var
  I: Int64;
  C: Integer;
  X: TExprFloat;
  ACount: Integer;
  isIF: Boolean;
  ID: TsExprIdentifierDef;
  Args: TExprArgumentArray;
  AI: Integer;
  cell: PCell;
begin
{$ifdef debugexpr}  Writeln('Primitive : ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  SetLength(Args, 0);
  if (TokenType = ttNumber) then
  begin
    if TryStrToInt64(CurrentToken, I) then
      Result := TsConstExprNode.CreateInteger(I)
    else
    begin
      val(CurrentToken, X, C);
      if (I = 0) then
        Result := TsConstExprNode.CreateFloat(X)
      else
        ParserError(Format(SErrInvalidFloat, [CurrentToken]));
    end;
  end
  else if (TokenType = ttString) then
    Result := TsConstExprNode.CreateString(CurrentToken)
  else if (TokenType in [ttTrue, ttFalse]) then
    Result := TsConstExprNode.CreateBoolean(TokenType = ttTrue)
  else if (TokenType = ttCell) then
    Result := TsCellExprNode.Create(FWorksheet, CurrentToken)
  else if (TokenType = ttCellRange) then
    raise Exception.Create('Cell range missing')
  else if not (TokenType in [ttIdentifier, ttIf]) then
    ParserError(Format(SerrUnknownTokenAtPos, [Scanner.Pos, CurrentToken]))
  else
  begin
    isIF := (TokenType = ttIf);
    if not isIF then
    begin
      ID := self.IdentifierByName(CurrentToken);
      If (ID = nil) then
        ParserError(Format(SErrUnknownIdentifier,[CurrentToken]))
    end;
    // Determine number of arguments
    if isIF then
      ACount := 3
    else if (ID.IdentifierType in [itFunctionCallBack, itFunctionHandler]) then
      ACount := ID.ArgumentCount
    else
      ACount := 0;
    // Parse arguments.
    // Negative is for variable number of arguments, where Abs(value) is the minimum number of arguments
    if (ACount <> 0) then
    begin
      GetToken;
      if (TokenType <> ttLeft) then
         ParserError(Format(SErrLeftBracketExpected, [Scanner.Pos, CurrentToken]));
      SetLength(Args, abs(ACount));
      AI := 0;
      try
        repeat
          GetToken;
          // Check if we must enlarge the argument array
          if (ACount < 0) and (AI = Length(Args)) then
          begin
            SetLength(Args, AI+1);
            Args[AI] := nil;
          end;
          Args[AI] := Level1;
          inc(AI);
          if (TokenType <> ttComma) then
            if (AI < abs(ACount)) then
              ParserError(Format(SErrCommaExpected, [Scanner.Pos, CurrentToken]))
        until (AI = ACount) or ((ACount < 0) and (TokenType = ttRight));
        if TokenType <> ttRight then
          ParserError(Format(SErrBracketExpected, [Scanner.Pos, CurrentToken]));
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
    if isIF then
      Result := TIfExprNode.Create(Args[0], Args[1], Args[2])
    else
      case ID.IdentifierType of
        itVariable         : Result := TFPExprVariable.CreateIdentifier(ID);
        itFunctionCallBack : Result := TFPFunctionCallback.CreateFunction(ID, Args);
        itFunctionHandler  : Result := TFPFunctionEventHandler.CreateFunction(ID, Args);
      end;
  end;
  GetToken;
  if TokenType = ttPercent then begin
    Result := TsPercentExprNode.Create(Result);
    GetToken;
  end;
end;

procedure TsExpressionParser.SetExpression(const AValue: String);
begin
  if FExpression = AValue then
    exit;
  FExpression := AValue;
  FScanner.Source := AValue;
  if Assigned(FExprNode) then
    FreeAndNil(FExprNode);
  if (FExpression <> '') then
  begin
    GetToken;
    FExprNode := Level1;
    if (TokenType <> ttEOF) then
      ParserError(Format(SErrUnterminatedExpression, [Scanner.Pos, CurrentToken]));
    FExprNode.Check;
  end
  else
    FExprNode := nil;
end;

procedure TsExpressionParser.CheckResultType(const Res: TsExpressionResult;
  AType: TsResultType); inline;
begin
  if (Res.ResultType <> AType) then
    RaiseParserError(SErrInvalidResultType, [ResultTypeName(Res.ResultType)]);
end;

class function TsExpressionParser.BuiltinExpressionManager: TsBuiltInExpressionManager;
begin
  Result := BuiltinIdentifiers;
end;

function TsExpressionParser.Evaluate: TsExpressionResult;
begin
  EvaluateExpression(Result);
end;

function TsExpressionParser.ResultType: TsResultType;
begin
  if not Assigned(FExprNode) then
    ParserError(SErrInExpression);
  Result := FExprNode.NodeType;;
end;


{ ---------------------------------------------------------------------
  TsExprIdentifierDefs
  ---------------------------------------------------------------------}

function TsExprIdentifierDefs.GetI(AIndex : Integer): TsExprIdentifierDef;
begin
  Result := TsExprIdentifierDef(Items[AIndex]);
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

function TsExprIdentifierDefs.IndexOfIdentifier(const AName: ShortString): Integer;
begin
  Result := Count-1;
  while (Result >= 0) and (CompareText(GetI(Result).Name, AName) <> 0) do
    dec(Result);
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

function TsExprIdentifierDefs.IdentifierByName(const AName: ShortString
  ): TsExprIdentifierDef;
begin
  Result := FindIdentifier(AName);
  if (Result = nil) then
    RaiseParserError(SErrUnknownIdentifier, [AName]);
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

function TsExprIdentifierDefs.AddBooleanVariable(const AName: ShortString;
  AValue: Boolean): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtBoolean;
  Result.FValue.ResBoolean := AValue;
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

function TsExprIdentifierDefs.AddFloatVariable(const AName: ShortString;
  AValue: TExprFloat): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtFloat;
  Result.FValue.ResFloat := AValue;
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

function TsExprIdentifierDefs.AddDateTimeVariable(const AName: ShortString;
  AValue: TDateTime): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtDateTime;
  Result.FValue.ResDateTime := AValue;
end;

function TsExprIdentifierDefs.AddFunction(const AName: ShortString;
  const AResultType: Char; const AParamTypes: String;
  ACallBack: TsExprFunctionCallBack): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.Name := AName;
  Result.IdentifierType := itFunctionCallBack;
  Result.ParameterTypes := AParamTypes;
  Result.ResultType := CharToResultType(AResultType);
  Result.FOnGetValueCB := ACallBack;
end;

function TsExprIdentifierDefs.AddFunction(const AName: ShortString;
  const AResultType: Char; const AParamTypes: String;
  ACallBack: TsExprFunctionEvent): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.Name := AName;
  Result.IdentifierType := itFunctionHandler;
  Result.ParameterTypes := AParamTypes;
  Result.ResultType := CharToResultType(AResultType);
  Result.FOnGetValue := ACallBack;
end;


{------------------------------------------------------------------------------}
{  TsExprIdentifierDef                                                        }
{------------------------------------------------------------------------------}

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

procedure TsExprIdentifierDef.CheckResultType(const AType: TsResultType);
begin
  if FValue.ResultType <> AType then
    RaiseParserError(SErrInvalidResultType, [ResultTypeName(AType)])
end;

procedure TsExprIdentifierDef.CheckVariable;
begin
  if Identifiertype <> itVariable then
    RaiseParserError(SErrNotVariable, [Name]);
end;

function TsExprIdentifierDef.ArgumentCount: Integer;
begin
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
    FIDType := EID.FIDType;
    FName := EID.FName;
    FOnGetValue := EID.FOnGetValue;
    FOnGetValueCB := EID.FOnGetValueCB;
  end
  else
    inherited Assign(Source);
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

procedure TsExprIdentifierDef.SetAsFloat(const AValue: TExprFloat);
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

function TsExprIdentifierDef.GetValue: String;
begin
  case FValue.ResultType of
    rtBoolean  : if FValue.ResBoolean then
                   Result := 'True'
                 else
                   Result := 'False';
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtFloat    : Result := FloatToStr(FValue.ResFloat);
    rtDateTime : Result := FormatDateTime('cccc', FValue.ResDateTime);
    rtString   : Result := FValue.ResString;
  end;
end;

function TsExprIdentifierDef.GetResultType: TsResultType;
begin
  Result := FValue.ResultType;
end;

function TsExprIdentifierDef.GetAsFloat: TExprFloat;
begin
  CheckResultType(rtFloat);
  CheckVariable;
  Result := FValue.ResFloat;
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

function TsBuiltInExpressionManager.GetCount: Integer;
begin
  Result := FDefs.Count;
end;

function TsBuiltInExpressionManager.GetI(AIndex: Integer): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs[Aindex])
end;

function TsBuiltInExpressionManager.IndexOfIdentifier(const AName: ShortString): Integer;
begin
  Result := FDefs.IndexOfIdentifier(AName);
end;

function TsBuiltInExpressionManager.FindIdentifier(const AName: ShortString
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.FindIdentifier(AName));
end;

function TsBuiltInExpressionManager.IdentifierByName(const AName: ShortString
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.IdentifierByName(AName));
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

function TsBuiltInExpressionManager.AddIntegerVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: Integer
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddIntegerVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFloatVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString;
  AValue: TExprFloat): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFloatVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddStringVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: String
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddStringVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddDateTimeVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: TDateTime
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddDateTimeVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFunction(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; const AResultType: Char; const AParamTypes: String;
  ACallBack: TsExprFunctionCallBack): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFunction(AName, AResultType, AParamTypes, ACallBack));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFunction(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; const AResultType: Char; const AParamTypes: String;
  ACallBack: TsExprFunctionEvent): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFunction(AName, AResultType, AParamTypes, ACallBack));
  Result.Category := ACategory;
end;


{------------------------------------------------------------------------------}
{  Various Nodes                                                               }
{------------------------------------------------------------------------------}

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

constructor TsConstExprNode.CreateFloat(AValue: TExprFloat);
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
    rtString   : Result := '''' + FValue.ResString + '''';
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtDateTime : Result := '''' + FormatDateTime('cccc', FValue.ResDateTime) + '''';
    rtBoolean  : if FValue.ResBoolean then Result := 'True' else Result := 'False';
    rtFloat    : Str(FValue.ResFloat, Result);
  end;
end;

function TsConstExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  case NodeType of
    rtString   : Result := RPNString(FValue.ResString, ANext);
    rtInteger  : Result := RPNNumber(FValue.ResInteger, ANext);
    rtDateTime : Result := RPNNumber(FValue.ResDateTime, ANext);
    rtBoolean  : Result := RPNBool(FValue.ResBoolean, ANext);
    rtFloat    : Result := RPNNumber(FValue.ResFloat, ANext);
  end;
end;


{ TsNegateExprNode }

function TsNegateExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekUMinus,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsNegateExprNode.AsString: String;
begin
  Result := '-' + TrimLeft(Operand.AsString);
end;

procedure TsNegateExprNode.Check;
begin
  inherited;
  if not (Operand.NodeType in [rtInteger, rtFloat]) then
    RaiseParserError(SErrNoNegation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsNegateExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtInteger : Result.ResInteger := -Result.ResInteger;
    rtFloat   : Result.ResFloat := -Result.ResFloat;
  end;
end;

function TsNegateExprNode.NodeType: TsResultType;
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
begin
  inherited;
  if not (Operand.NodeType in [rtInteger, rtFloat]) then
    RaiseParserError(SErrNoPercentOperation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsPercentExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtInteger : Result.ResFloat := 0.01 * Result.ResInteger;
    rtFloat   : Result.ResFloat := 0.01 * Result.ResFloat;
  end;
  Result.ResultType := Nodetype;
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


{ TsBinaryAndExprNode }

function TsBinaryAndExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekAND,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsBinaryAndExprNode.AsString: string;
begin
  Result := Left.AsString + ' and ' + Right.AsString;
end;

procedure TsBooleanOperationExprNode.Check;
begin
  inherited Check;
  CheckNodeType(Left,  [rtInteger, rtBoolean]);
  CheckNodeType(Right, [rtInteger, rtBoolean]);
  CheckSameNodeTypes;
end;

function TsBooleanOperationExprNode.NodeType: TsResultType;
begin
  Result := Left.NodeType;
end;

procedure TsBinaryAndExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtBoolean : Result.resBoolean := Result.ResBoolean and RRes.ResBoolean;
    rtInteger : Result.resInteger := Result.ResInteger and RRes.ResInteger;
  end;
end;


{ TsExprNode }

procedure TsExprNode.CheckNodeType(Anode: TsExprNode; Allowed: TsResultTypes);
var
  S: String;
  A: TsResultType;
begin
  if (Anode = nil) then
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


{ TsBinaryOrExprNode }

function TsBinaryOrExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekOR,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsBinaryOrExprNode.AsString: string;
begin
  Result := Left.AsString + ' or ' + Right.AsString;
end;

procedure TsBinaryOrExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtBoolean : Result.resBoolean := Result.ResBoolean or RRes.ResBoolean;
    rtInteger : Result.resInteger := Result.ResInteger or RRes.ResInteger;
  end;
end;


{ TsBinaryXorExprNode }

function TsBinaryXorExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  RaiseParserError(SErrNoXOROperationRPN);
end;

function TsBinaryXorExprNode.AsString: string;
begin
  Result := Left.AsString + ' xor ' + Right.AsString;
end;

procedure TsBinaryXorExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtBoolean : Result.resBoolean := Result.ResBoolean xor RRes.ResBoolean;
    rtInteger : Result.resInteger := Result.ResInteger xor RRes.ResInteger;
  end;
end;


{ TsNotExprNode }

function TsNotExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekNOT,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsNotExprNode.AsString: String;
begin
  Result := 'not ' + Operand.AsString;
end;

procedure TsNotExprNode.Check;
begin
  if not (Operand.NodeType in [rtInteger, rtBoolean]) then
    RaiseParserError(SErrNoNotOperation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsNotExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtInteger : Result.ResInteger := not Result.ResInteger;
    rtBoolean : Result.ResBoolean := not Result.ResBoolean;
  end
end;

function TsNotExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;


{ TIfExprNode }

constructor TIfExprNode.Create(ACondition, ALeft, ARight: TsExprNode);
begin
  inherited Create(ALeft,ARight);
  FCondition := ACondition;
end;

destructor TIfExprNode.Destroy;
begin
  FreeAndNil(FCondition);
  inherited Destroy;
end;

function TIfExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  if Left = nil then
    Result := RPNFunc(fekIF,
      Right.AsRPNItem(
      ANext
    ))
  else
    Result := RPNFunc(fekIF,
      Right.AsRPNItem(
      Left.AsRPNItem(
      ANext
    )));
end;

function TIfExprNode.AsString: string;
begin
  if Right = nil then
    Result := Format('IF(%s, %s)', [Condition.AsString, Left.AsString])
  else
    Result := Format('IF(%s, %s, %s)',[Condition.AsString, Left.AsString, Right.AsString]);
end;

procedure TIfExprNode.Check;
begin
  inherited Check;
  if (Condition.NodeType <> rtBoolean) then
    RaiseParserError(SErrIFNeedsBoolean, [Condition.AsString]);
  CheckSameNodeTypes;
end;

procedure TIfExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  FCondition.GetNodeValue(Result);
  if Result.ResBoolean then
    Left.GetNodeValue(Result)
  else
    Right.GetNodeValue(Result)
end;

function TIfExprNode.NodeType: TsResultType;
begin
  Result := Left.NodeType;
end;


{ TsBooleanResultExprNode }

procedure TsBooleanResultExprNode.Check;
begin
  inherited Check;
  CheckSameNodeTypes;
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
  Result := Left.AsString + ' = ' + Right.AsString;
end;

procedure TsEqualExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtBoolean  : Result.resBoolean := Result.ResBoolean = RRes.ResBoolean;
    rtInteger  : Result.resBoolean := Result.ResInteger = RRes.ResInteger;
    rtFloat    : Result.resBoolean := Result.ResFloat = RRes.ResFLoat;
    rtDateTime : Result.resBoolean := Result.ResDateTime = RRes.ResDateTime;
    rtString   : Result.resBoolean := Result.ResString = RRes.ResString;
  end;
  Result.ResultType := rtBoolean;
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
  Result := Left.AsString + ' <> ' + Right.AsString;
end;

procedure TsNotEqualExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  inherited GetNodeValue(Result);
  Result.ResBoolean := not Result.ResBoolean;
end;


{ TsLessThanExprNode }

function TsLessThanExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekLess,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsLessThanExprNode.AsString: string;
begin
  Result := Left.AsString + ' < ' + Right.AsString;
end;

procedure TsLessThanExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtInteger  : Result.resBoolean := Result.ResInteger < RRes.ResInteger;
    rtFloat    : Result.resBoolean := Result.ResFloat < RRes.ResFLoat;
    rtDateTime : Result.resBoolean := Result.ResDateTime < RRes.ResDateTime;
    rtString   : Result.resBoolean := Result.ResString < RRes.ResString;
  end;
  Result.ResultType := rtBoolean;
end;


{ TsGreaterThanExprNode }

function TsGreaterThanExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekGreater,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsGreaterThanExprNode.AsString: string;
begin
  Result := Left.AsString + ' > ' + Right.AsString;
end;

procedure TsGreaterThanExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtInteger : case Right.NodeType of
                  rtInteger : Result.resBoolean := Result.ResInteger > RRes.ResInteger;
                  rtFloat   : Result.resBoolean := Result.ResInteger > RRes.ResFloat;
                end;
    rtFloat   : case Right.NodeType of
                  rtInteger : Result.resBoolean := Result.ResFloat > RRes.ResInteger;
                  rtFloat   : Result.resBoolean := Result.ResFloat > RRes.ResFLoat;
                end;
    rtDateTime : Result.resBoolean := Result.ResDateTime > RRes.ResDateTime;
    rtString   : Result.resBoolean := Result.ResString > RRes.ResString;
  end;
  Result.ResultType := rtBoolean;
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
  Result := Left.AsString + ' >= ' + Right.AsString;
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
  Result := Left.AsString + ' <= ' + Right.AsString;
end;

procedure TsLessEqualExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  inherited GetNodeValue(Result);
  Result.ResBoolean := not Result.ResBoolean;
end;


{ TsOrderingExprNode }

procedure TsOrderingExprNode.Check;
const
  AllowedTypes =[rtInteger, rtfloat, rtDateTime, rtString];
begin
  CheckNodeType(Left, AllowedTypes);
  CheckNodeType(Right, AllowedTypes);
  inherited Check;
end;


{ TsConcatExprNode }

procedure TsConcatExprNode.Check;
begin
  inherited Check;
  CheckNodeType(Left, [rtString]);
  CheckNodeType(Right, [rtString]);
end;

procedure TsConcatExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes : TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  Result.ResString := Result.ResString + RRes.ResString;
  Result.ResultType := rtString;
end;

function TsConcatExprNode.NodeType: TsResultType;
begin
  Result := rtString;
end;

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


{ TsMathOperationExprNode }

procedure TsMathOperationExprNode.Check;
const
  AllowedTypes  = [rtInteger, rtfloat, rtDateTime];
begin
  inherited Check;
  CheckNodeType(Left, AllowedTypes);
  CheckNodeType(Right, AllowedTypes);
  CheckSameNodeTypes;
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
  RRes : TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtInteger  : Result.ResInteger := Result.ResInteger + RRes.ResInteger;
//    rtString   : Result.ResString := Result.ResString + RRes.ResString;
    rtDateTime : Result.ResDateTime := Result.ResDateTime + RRes.ResDateTime;
    rtFloat    : Result.ResFloat := Result.ResFloat + RRes.ResFloat;
  end;
  Result.ResultType := NodeType;
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

procedure TsSubtractExprNode.Check;
const
  AllowedTypes =[rtInteger, rtfloat, rtDateTime];
begin
  CheckNodeType(Left, AllowedTypes);
  CheckNodeType(Right, AllowedTypes);
  inherited Check;
end;

procedure TsSubtractExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtInteger  : Result.ResInteger := Result.ResInteger - RRes.ResInteger;
    rtDateTime : Result.ResDateTime := Result.ResDateTime - RRes.ResDateTime;
    rtFloat    : Result.ResFLoat := Result.ResFLoat - RRes.ResFLoat;
  end;
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

procedure TsMultiplyExprNode.Check;
const
  AllowedTypes = [rtInteger, rtFloat];
begin
  CheckNodeType(Left, AllowedTypes);
  CheckNodeType(Right, AllowedTypes);
  inherited;
end;

procedure TsMultiplyExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtInteger : Result.ResInteger := Result.ResInteger * RRes.ResInteger;
    rtFloat   : Result.ResFloat := Result.ResFloat * RRes.ResFloat;
  end;
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

procedure TsDivideExprNode.Check;
const
  AllowedTypes =[rtInteger, rtFloat];
begin
  CheckNodeType(Left, AllowedTypes);
  CheckNodeType(Right, AllowedTypes);
  inherited Check;
end;

procedure TsDivideExprNode.GetNodeValue(var Result: TsExpressionResult);
var
  RRes: TsExpressionResult;
begin
  Left.GetNodeValue(Result);
  Right.GetNodeValue(RRes);
  case Result.ResultType of
    rtInteger : Result.ResFloat := Result.ResInteger / RRes.ResInteger;
    rtFloat   : Result.ResFloat := Result.ResFloat / RRes.ResFloat;
  end;
  Result.ResultType := rtFloat;
end;

function TsDivideExprNode.NodeType: TsResultType;
begin
  Result := rtFLoat;
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
  CheckNodeType(Operand, [rtInteger])
end;

procedure TsIntToFloatExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  Result.ResFloat := Result.ResInteger;
  Result.ResultType := rtFloat;
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
  Result.ResDateTime := Result.ResInteger;
  Result.ResultType := rtDateTime;
end;

{ TsFloatToDateTimeExprNode }

procedure TsFloatToDateTimeExprNode.Check;
begin
  inherited Check;
  CheckNodeType(Operand, [rtFloat]);
end;

function TsFloatToDateTimeExprNode.NodeType: TsResultType;
begin
  Result := rtDateTime;
end;

procedure TsFloatToDateTimeExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  Result.ResDateTime := Result.ResFloat;
  Result.ResultType := rtDateTime;
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


{ TFPExprVariable }

procedure TFPExprVariable.Check;
begin
  // Do nothing;
end;

function TFPExprVariable.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  RaiseParserError('Cannot handle variables for RPN, so far.');
end;

function TFPExprVariable.AsString: string;
begin
  Result := FID.Name;
end;


{ TFPExprFunction }

constructor TFPExprFunction.CreateFunction(AID: TsExprIdentifierDef;
  const Args: TExprArgumentArray);
begin
  inherited CreateIdentifier(AID);
  FArgumentNodes := Args;
  SetLength(FArgumentParams, Length(Args));
end;

destructor TFPExprFunction.Destroy;
var
  i: Integer;
begin
  for i:=0 to Length(FArgumentNodes)-1 do
    FreeAndNil(FArgumentNodes[i]);
  inherited Destroy;
end;

function TFPExprFunction.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := ANext;
//  RaiseParserError('Cannot handle functions for RPN, so far.');
end;

function TFPExprFunction.AsString: String;
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
  if (S <> '') then
    S := '(' + S + ')';
  Result := FID.Name + S;
end;

procedure TFPExprFunction.CalcParams;
var
  i : Integer;
begin
  for i := 0 to Length(FArgumentParams)-1 do
    FArgumentNodes[i].GetNodeValue(FArgumentParams[i]);
end;

procedure TFPExprFunction.Check;
var
  i: Integer;
  rtp, rta: TsResultType;
begin
  if Length(FArgumentNodes) <> FID.ArgumentCount then
    RaiseParserError(ErrInvalidArgumentCount, [FID.Name]);
  for i := 0 to Length(FArgumentNodes)-1 do
  begin
    rtp := CharToResultType(FID.ParameterTypes[i+1]);
    rta := FArgumentNodes[i].NodeType;
    if (rtp <> rta) then
    begin
      // Automatically convert integers to floats in functions that return
      // a float
      if (rta = rtInteger) and (rtp = rtFloat) then
      begin
        FArgumentNodes[i] := TsIntToFloatExprNode(FArgumentNodes[i]);
        exit;
      end;
      RaiseParserError(SErrInvalidArgumentType, [I+1, ResultTypeName(rtp), ResultTypeName(rta)])
    end;
  end;
end;


{ TFPFunctionCallBack }

constructor TFPFunctionCallBack.CreateFunction(AID: TsExprIdentifierDef;
  const Args: TExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValueCallBack;
end;

procedure TFPFunctionCallBack.GetNodeValue(var Result: TsExpressionResult);
begin
  if Length(FArgumentParams) > 0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
  Result.ResultType := NodeType;
end;


{ TFPFunctionEventHandler }

constructor TFPFunctionEventHandler.CreateFunction(AID: TsExprIdentifierDef;
  const Args: TExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValue;
end;

procedure TFPFunctionEventHandler.GetNodeValue(var Result: TsExpressionResult);
begin
  if Length(FArgumentParams)>0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
  Result.ResultType := NodeType;
end;


{ TsCellExprNode }

constructor TsCellExprNode.Create(AWorksheet: TsWorksheet; ACellString: String);
var
  r, c: Cardinal;
  flags: TsRelFlags;
begin
  ParseCellString(ACellString, r, c, flags);
  Create(AWorksheet, AWorksheet.FindCell(r, c), flags);
end;

constructor TsCellExprNode.Create(AWorksheet: TsWorksheet; ACell: PCell; AFlags: TsRelFlags);
begin
  FCell := ACell;
  FFlags := AFlags;
end;

function TsCellExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNCellValue(FCell^.Row, FCell^.Col, FFlags, ANext);
end;

function TsCellExprNode.AsString: string;
begin
  Result := GetCellString(FCell^.Row, FCell^.Col, FFlags);
end;

procedure TsCellExprNode.Check;
begin
  if not Assigned(FCell) then
    RaiseParserError(SErrNoCellOperand);
  if (FCell^.ContentType = cctError) and (FCell^.ErrorValue <> errOK) then
    raise EExprParser.CreateFmt(SErrCellError, [AsString]);
end;

procedure TsCellExprNode.GetNodeValue(var Result: TsExpressionResult);
begin
  case FCell^.ContentType of
    cctNumber:
      Result.ResFloat := FCell^.NumberValue;
    cctDateTime:
      Result.ResDateTime := FCell^.DateTimeValue;
    cctUTF8String:
      Result.ResString := FCell^.UTF8StringValue;
    cctBool:
      Result.ResBoolean := FCell^.BoolValue;
    cctEmpty:
      Result.ResString := '';
  end;
  Result.ResultType := NodeType;
end;

function TsCellExprNode.NodeType: TsResultType;
begin
  case FCell^.ContentType of
    cctNumber:
      Result := rtFloat;
    cctDateTime:
      Result := rtDateTime;
    cctUTF8String:
      Result := rtString;
    cctBool:
      Result := rtBoolean;
    cctEmpty:
      Result := rtString;
  end;
end;


{ ---------------------------------------------------------------------
  Standard Builtins support
  ---------------------------------------------------------------------}

{ Template for builtin.

Procedure MyCallback (Var Result : TsExpressionResult; Const Args : TExprParameterArray);
begin
end;
}

function ArgToFloat(Arg: TsExpressionResult): TExprFloat;
// Utility function for the built-in math functions. Accepts also integers
// in place of the floating point arguments. To be called in builtins or
// user-defined callbacks having float results.
begin
  if Arg.ResultType = rtInteger then
    result := Arg.resInteger
  else
    result := Arg.resFloat;
end;


// Math builtins

procedure BuiltInCos(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := cos(ArgToFloat(Args[0]));
end;

procedure BuiltInSin(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := sin(ArgToFloat(Args[0]));
end;

procedure BuiltInArcTan(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := arctan(ArgToFloat(Args[0]));
end;

procedure BuiltInAbs(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := abs(ArgToFloat(Args[0]));
end;

procedure BuiltInSqr(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := sqr(ArgToFloat(Args[0]));
end;

procedure BuiltInSqrt(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := sqrt(ArgToFloat(Args[0]));
end;

procedure BuiltInExp(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := exp(ArgToFloat(Args[0]));
end;

procedure BuiltInLn(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := ln(ArgToFloat(Args[0]));
end;

const
  ln10 = ln(10);

procedure BuiltInLog(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := ln(ArgToFloat(Args[0]))/ln10;
end;

procedure BuiltInRound(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  decs: Integer;
  f: TExprFloat;
begin           (*
  decs := round(ArgToFloat(Args[1]));
  f := 1.0;
  while decs > 0 do begin
    f := f * 10;
    dec(decs);
  end;        *)
  Result.ResInteger := round(ArgToFloat(Args[0]));
end;

procedure BuiltInTrunc(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := trunc(ArgToFloat(Args[0]));
end;

procedure BuiltInInt(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := int(ArgToFloat(Args[0]));
end;

procedure BuiltInFrac(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := frac(ArgToFloat(Args[0]));
end;


// String builtins

procedure BuiltInLength(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := Length(Args[0].ResString);
end;

procedure BuiltInCopy(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := copy(Args[0].ResString, Args[1].ResInteger, Args[2].ResInteger);
end;

procedure BuiltInDelete(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := Args[0].resString;
  Delete(Result.ResString, Args[1].ResInteger, Args[2].ResInteger);
end;

procedure BuiltInPos(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := pos(Args[0].ResString, Args[1].ResString);
end;

procedure BuiltInUppercase(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := Uppercase(Args[0].ResString);
end;

procedure BuiltInLowercase(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := Lowercase(Args[0].ResString);
end;

procedure BuiltInStringReplace(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  flags : TReplaceFlags;
begin
  flags := [];
  if Args[3].ResBoolean then
    Include(flags, rfReplaceAll);
  if Args[4].ResBoolean then
    Include(flags, rfIgnoreCase);
  Result.ResString := StringReplace(Args[0].ResString, Args[1].ResString, Args[2].ResString, flags);
end;

procedure BuiltInCompareText(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := CompareText(Args[0].ResString, Args[1].ResString);
end;


// Date/Time builtins

procedure BuiltInDate(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := Date;
end;

procedure BuiltInTime(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := Time;
end;

procedure BuiltInNow(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := Now;
end;

procedure BuiltInDayOfWeek(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger :=  DayOfWeek(Args[0].resDateTime);
end;

procedure BuiltInExtractYear(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  Y, M, D: Word;
begin
  DecodeDate(Args[0].ResDateTime, Y, M, D);
  Result.ResInteger := Y;
end;

procedure BuiltInExtractMonth(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  Y, M, D: Word;
begin
  DecodeDate(Args[0].ResDateTime, Y, M, D);
  Result.ResInteger := M;
end;

procedure BuiltInExtractDay(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  Y, M, D: Word;
begin
  DecodeDate(Args[0].ResDateTime, Y, M, D);
  Result.ResInteger := D;
end;

procedure BuiltInExtractHour(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  H, M, S, MS: Word;
begin
  DecodeTime(Args[0].ResDateTime, H, M, S, MS);
  Result.ResInteger := H;
end;

procedure BuiltInExtractMin(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  H, M, S, MS: word;
begin
  DecodeTime(Args[0].ResDateTime, H, M, S, MS);
  Result.ResInteger := M;
end;

procedure BuiltInExtractSec(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  H, M, S, MS: Word;
begin
  DecodeTime(Args[0].ResDateTime, H, M, S, MS);
  Result.ResInteger := S;
end;

procedure BuiltInExtractMSec(var Result: TsExpressionResult; const Args: TExprParameterArray);
var
  H, M, S, MS: Word;
begin
  DecodeTime(Args[0].ResDateTime, H, M, S, MS);
  Result.ResInteger := MS;
end;

procedure BuiltInEncodeDate(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := Encodedate(Args[0].ResInteger, Args[1].ResInteger, Args[2].ResInteger);
end;

procedure BuiltInEncodeTime(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := EncodeTime(Args[0].ResInteger, Args[1].ResInteger, Args[2].ResInteger, Args[3].ResInteger);
end;

procedure BuiltInEncodeDateTime(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := EncodeDate(Args[0].ResInteger, Args[1].ResInteger, Args[2].ResInteger)
                      + EncodeTime(Args[3].ResInteger, Args[4].ResInteger, Args[5].ResInteger, Args[6].ResInteger);
end;

procedure BuiltInShortDayName(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := ShortDayNames[Args[0].ResInteger];
end;

procedure BuiltInShortMonthName(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := ShortMonthNames[Args[0].ResInteger];
end;

Procedure BuiltInLongDayName(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := LongDayNames[Args[0].ResInteger];
end;

procedure BuiltInLongMonthName(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := LongMonthNames[Args[0].ResInteger];
end;

procedure BuiltInFormatDateTime(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := FormatDateTime(Args[0].ResString, Args[1].ResDateTime);
end;


// Conversion
procedure BuiltInIntToStr(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := IntToStr(Args[0].Resinteger);
end;

procedure BuiltInStrToInt(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := StrToInt(Args[0].ResString);
end;

procedure BuiltInStrToIntDef(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := StrToIntDef(Args[0].ResString, Args[1].ResInteger);
end;

procedure BuiltInFloatToStr(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := FloatToStr(Args[0].ResFloat);
end;

procedure BuiltInStrToFloat(var Result: TsExpressionResult; Const Args: TExprParameterArray);
begin
  Result.ResFloat := StrToFloat(Args[0].ResString);
end;

procedure BuiltInStrToFloatDef(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat := StrToFloatDef(Args[0].ResString, Args[1].ResFloat);
end;

procedure BuiltInDateToStr(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := DateToStr(Args[0].ResDateTime);
end;

procedure BuiltInTimeToStr(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := TimeToStr(Args[0].ResDateTime);
end;

procedure BuiltInStrToDate(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := StrToDate(Args[0].ResString);
end;

procedure BuiltInStrToDateDef(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := StrToDateDef(Args[0].ResString, Args[1].ResDateTime);
end;

procedure BuiltInStrToTime(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := StrToTime(Args[0].ResString);
end;

procedure BuiltInStrToTimeDef(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := StrToTimeDef(Args[0].ResString, Args[1].ResDateTime);
end;

procedure BuiltInStrToDateTime(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := StrToDateTime(Args[0].ResString);
end;

procedure BuiltInStrToDateTimeDef(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime := StrToDateTimeDef(Args[0].ResString, Args[1].ResDateTime);
end;

procedure BuiltInBoolToStr(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResString := BoolToStr(Args[0].ResBoolean);
end;

procedure BuiltInStrToBool(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResBoolean := StrToBool(Args[0].ResString);
end;

procedure BuiltInStrToBoolDef(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResBoolean := StrToBoolDef(Args[0].ResString, Args[1].ResBoolean);
end;


// Boolean

procedure BuiltInShl(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := Args[0].ResInteger shl Args[1].ResInteger
end;

procedure BuiltInShr(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResInteger := Args[0].ResInteger shr Args[1].ResInteger
end;

procedure BuiltinIFS(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  if Args[0].ResBoolean then
    Result.ResString := Args[1].ResString
  else
    Result.ResString := Args[2].ResString
end;

procedure BuiltinIFI(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  if Args[0].ResBoolean then
    Result.ResInteger := Args[1].ResInteger
  else
    Result.ResInteger := Args[2].ResInteger
end;

procedure BuiltinIFF(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  if Args[0].ResBoolean then
    Result.ResFloat := Args[1].ResFloat
  else
    Result.ResFloat := Args[2].ResFloat
end;

procedure BuiltinIFD(var Result: TsExpressionResult; const Args: TExprParameterArray);
begin
  if Args[0].ResBoolean then
    Result.ResDateTime := Args[1].ResDateTime
  else
    Result.ResDateTime := Args[2].ResDateTime
end;

procedure RegisterStdBuiltins(AManager : TsBuiltInExpressionManager);
begin
  with AManager do
  begin
    AddFloatVariable(bcMath, 'pi', Pi);
    // Math functions
    AddFunction(bcMath, 'cos',    'F', 'F', @BuiltinCos);
    AddFunction(bcMath, 'sin',    'F', 'F', @BuiltinSin);
    AddFunction(bcMath, 'arctan', 'F', 'F', @BuiltinArctan);
    AddFunction(bcMath, 'abs',    'F', 'F', @BuiltinAbs);
    AddFunction(bcMath, 'sqr',    'F', 'F', @BuiltinSqr);
    AddFunction(bcMath, 'sqrt',   'F', 'F', @BuiltinSqrt);
    AddFunction(bcMath, 'exp',    'F', 'F', @BuiltinExp);
    AddFunction(bcMath, 'ln',     'F', 'F', @BuiltinLn);
    AddFunction(bcMath, 'log',    'F', 'F', @BuiltinLog);
    AddFunction(bcMath, 'frac',   'F', 'F', @BuiltinFrac);
    AddFunction(bcMath, 'int',    'F', 'F', @BuiltinInt);
    AddFunction(bcMath, 'round',  'I', 'F', @BuiltinRound);
    AddFunction(bcMath, 'trunc',  'I', 'F', @BuiltinTrunc);
    // String
    AddFunction(bcStrings, 'length',       'I', 'S',    @BuiltinLength);
    AddFunction(bcStrings, 'copy',         'S', 'SII',  @BuiltinCopy);
    AddFunction(bcStrings, 'delete',       'S', 'SII',  @BuiltinDelete);
    AddFunction(bcStrings, 'pos',          'I', 'SS',   @BuiltinPos);
    AddFunction(bcStrings, 'lowercase',    'S', 'S',    @BuiltinLowercase);
    AddFunction(bcStrings, 'uppercase',    'S', 'S',    @BuiltinUppercase);
    AddFunction(bcStrings, 'stringreplace','S', 'SSSBB',@BuiltinStringReplace);
    AddFunction(bcStrings, 'comparetext',  'I', 'SS',   @BuiltinCompareText);
    // Date/Time
    AddFunction(bcDateTime, 'date',           'D', '',    @BuiltinDate);
    AddFunction(bcDateTime, 'time',           'D', '',    @BuiltinTime);
    AddFunction(bcDateTime, 'now',            'D', '',    @BuiltinNow);
    AddFunction(bcDateTime, 'dayofweek',      'I', 'D',   @BuiltinDayofweek);
    AddFunction(bcDateTime, 'extractyear',    'I', 'D',   @BuiltinExtractYear);
    AddFunction(bcDateTime, 'extractmonth',   'I', 'D',   @BuiltinExtractMonth);
    AddFunction(bcDateTime, 'extractday',     'I', 'D',   @BuiltinExtractDay);
    AddFunction(bcDateTime, 'extracthour',    'I', 'D',   @BuiltinExtractHour);
    AddFunction(bcDateTime, 'extractmin',     'I', 'D',   @BuiltinExtractMin);
    AddFunction(bcDateTime, 'extractsec',     'I', 'D',   @BuiltinExtractSec);
    AddFunction(bcDateTime, 'extractmsec',    'I', 'D',   @BuiltinExtractMSec);
    AddFunction(bcDateTime, 'encodedate',     'D', 'III', @BuiltinEncodedate);
    AddFunction(bcDateTime, 'encodetime',     'D', 'IIII',@BuiltinEncodeTime);
    AddFunction(bcDateTime, 'encodedatetime', 'D', 'IIIIIII',@BuiltinEncodeDateTime);
    AddFunction(bcDateTime, 'shortdayname',   'S', 'I',   @BuiltinShortDayName);
    AddFunction(bcDateTime, 'shortmonthname', 'S', 'I',   @BuiltinShortMonthName);
    AddFunction(bcDateTime, 'longdayname',    'S', 'I',   @BuiltinLongDayName);
    AddFunction(bcDateTime, 'longmonthname',  'S', 'I',   @BuiltinLongMonthName);
    AddFunction(bcDateTime, 'formatdatetime', 'S', 'SD',  @BuiltinFormatDateTime);
    // Boolean
    AddFunction(bcBoolean, 'shl', 'I', 'II',  @BuiltinShl);
    AddFunction(bcBoolean, 'shr', 'I', 'II',  @BuiltinShr);
    AddFunction(bcBoolean, 'IFS', 'S', 'BSS', @BuiltinIFS);
    AddFunction(bcBoolean, 'IFF', 'F', 'BFF', @BuiltinIFF);
    AddFunction(bcBoolean, 'IFD', 'D', 'BDD', @BuiltinIFD);
    AddFunction(bcBoolean, 'IFI', 'I', 'BII', @BuiltinIFI);
    // Conversion
    AddFunction(bcConversion, 'inttostr',         'S', 'I',  @BuiltInIntToStr);
    AddFunction(bcConversion, 'strtoint',         'I', 'S',  @BuiltInStrToInt);
    AddFunction(bcConversion, 'strtointdef',      'I', 'SI', @BuiltInStrToIntDef);
    AddFunction(bcConversion, 'floattostr',       'S', 'F',  @BuiltInFloatToStr);
    AddFunction(bcConversion, 'strtofloat',       'F', 'S',  @BuiltInStrToFloat);
    AddFunction(bcConversion, 'strtofloatdef',    'F', 'SF', @BuiltInStrToFloatDef);
    AddFunction(bcConversion, 'booltostr',        'S', 'B',  @BuiltInBoolToStr);
    AddFunction(bcConversion, 'strtobool',        'B', 'S',  @BuiltInStrToBool);
    AddFunction(bcConversion, 'strtobooldef',     'B', 'SB', @BuiltInStrToBoolDef);
    AddFunction(bcConversion, 'datetostr',        'S', 'D',  @BuiltInDateToStr);
    AddFunction(bcConversion, 'timetostr',        'S', 'D',  @BuiltInTimeToStr);
    AddFunction(bcConversion, 'strtodate',        'D', 'S',  @BuiltInStrToDate);
    AddFunction(bcConversion, 'strtodatedef',     'D', 'SD', @BuiltInStrToDateDef);
    AddFunction(bcConversion, 'strtotime',        'D', 'S',  @BuiltInStrToTime);
    AddFunction(bcConversion, 'strtotimedef',     'D', 'SD', @BuiltInStrToTimeDef);
    AddFunction(bcConversion, 'strtodatetime',    'D', 'S',  @BuiltInStrToDateTime);
    AddFunction(bcConversion, 'strtodatetimedef', 'D', 'SD', @BuiltInStrToDateTimeDef);
  end;
end;

{ TsBuiltInExprIdentifierDef }

procedure TsBuiltInExprIdentifierDef.Assign(Source: TPersistent);
begin
  inherited Assign(Source);
  if Source is TsBuiltInExprIdentifierDef then
    FCategory:=(Source as TsBuiltInExprIdentifierDef).Category;
end;

initialization
  RegisterStdBuiltins(BuiltinIdentifiers);

finalization
  FreeBuiltins;
end.
