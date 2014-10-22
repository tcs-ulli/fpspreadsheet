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
    - generalize scanner and parser to allow localized decimal and list separators
    - add to spreadsheet format to parser to take account of formula "dialect"
      (see OpenDocument using [] around cell addresses)

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
  Classes, SysUtils, contnrs, fpspreadsheet, fpsrpn;

type
  { Tokens }

  TsTokenType = (
    ttCell, ttCellRange, ttNumber, ttString, ttIdentifier,
    ttPlus, ttMinus, ttMul, ttDiv, ttConcat, ttPercent, ttPower, ttLeft, ttRight,
    ttLessThan, ttLargerThan, ttEqual, ttNotEqual, ttLessThanEqual, ttLargerThanEqual,
    ttListSep, ttTrue, ttFalse, ttError, ttEOF
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

  TsFormulaDialect = (fdExcel, fdOpenDocument);

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
  private
    FParser: TsExpressionParser;
  protected
    procedure CheckNodeType(ANode: TsExprNode; Allowed: TsResultTypes);
    // A procedure with var saves an implicit try/finally in each node
    // A marked difference in execution speed.
    procedure GetNodeValue(out Result: TsExpressionResult); virtual; abstract;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; virtual; abstract;
    function AsString: string; virtual; abstract;
    procedure Check; virtual; abstract;
    function NodeType: TsResultType; virtual; abstract;
    function NodeValue: TsExpressionResult;
    property Parser: TsExpressionParser read FParser;
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
    constructor Create(AParser: TsExpressionParser; ALeft, ARight: TsExprNode);
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsNotEqualExprNode }
  TsNotEqualExprNode = class(TsEqualExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsOrderingExprNode }
  TsOrderingExprNode = class(TsBooleanResultExprNode)
  public
    procedure Check; override;
  end;

  { TsLessExprNode }
  TsLessExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterExprNode }
  TsGreaterExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsLessEqualExprNode }
  TsLessEqualExprNode = class(TsGreaterExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterEqualExprNode }
  TsGreaterEqualExprNode = class(TsLessExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsConcatExprNode }
  TsConcatExprNode = class(TsBinaryOperationExprNode)
  protected
    procedure CheckSameNodeTypes; override;
    procedure GetNodeValue(out Result: TsExpressionResult); override;
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsSubtractExprNode }
  TsSubtractExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsMultiplyExprNode }
  TsMultiplyExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsDivideExprNode }
  TsDivideExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    function NodeType: TsResultType; override;
  end;

  { TsPowerExprNode }
  TsPowerExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    function NodeType: TsResultType; override;
  end;

  { TsUnaryOperationExprNode }
  TsUnaryOperationExprNode = class(TsExprNode)
  private
    FOperand: TsExprNode;
  public
    constructor Create(AParser: TsExpressionParser; AOperand: TsExprNode);
    procedure Check; override;
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function NodeType: TsResultType; override;
  end;

  { TsIntToDateTimeExprNode }
  TsIntToDateTimeExprNode = class(TsConvertToIntExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function NodeType: TsResultType; override;
  end;

  { TsFloatToDateTimeExprNode }
  TsFloatToDateTimeExprNode = class(TsConvertExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsUPlusExprNode }
  TsUPlusExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsUMinusExprNode }
  TsUMinusExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsPercentExprNode }
  TsPercentExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsParenthesisExprNode }
  TsParenthesisExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    procedure Check; override;
    constructor CreateString(AParser: TsExpressionParser; AValue: String);
    constructor CreateInteger(AParser: TsExpressionParser; AValue: Int64);
    constructor CreateDateTime(AParser: TsExpressionParser; AValue: TDateTime);
    constructor CreateFloat(AParser: TsExpressionParser; AValue: TsExprFloat);
    constructor CreateBoolean(AParser: TsExpressionParser; AValue: Boolean);
    constructor CreateError(AParser: TsExpressionParser; AValue: TsErrorValue); overload;
    constructor CreateError(AParser: TsExpressionParser; AValue: String); overload;
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
    function GetFormatSettings: TFormatSettings;
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor CreateIdentifier(AParser: TsExpressionParser; AID: TsExprIdentifierDef);
    function NodeType: TsResultType; override;
    property Identifier: TsExprIdentifierDef read FID;
  end;

  { TsVariableExprNode }
  TsVariableExprNode = class(TsIdentifierExprNode)
  public
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
    constructor CreateFunction(AParser: TsExpressionParser;
      AID: TsExprIdentifierDef; const Args: TsExprArgumentArray); virtual;
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor CreateFunction(AParser: TsExpressionParser;
      AID: TsExprIdentifierDef; const Args: TsExprArgumentArray); override;
    property CallBack: TsExprFunctionCallBack read FCallBack;
  end;

  { TFPFunctionEventHandlerExprNode }
  TFPFunctionEventHandlerExprNode = class(TsFunctionExprNode)
  private
    FCallBack: TsExprFunctionEvent;
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor CreateFunction(AParser: TsExpressionParser;
      AID: TsExprIdentifierDef; const Args: TsExprArgumentArray); override;
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
    function GetCol: Cardinal;
    function GetRow: Cardinal;
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ACellString: String); overload;
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ARow, ACol: Cardinal; AFlags: TsRelFlags); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
    procedure Check; override;
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
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ACellRangeString: String); overload;
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ARow1,ACol1, ARow2,ACol2: Cardinal; AFlags: TsRelFlags); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
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
    FParser: TsExpressionParser;
    function GetCurrentChar: Char;
    procedure ScanError(Msg: String);
  protected
    procedure SetSource(const AValue: String); virtual;
    function DoError: TsTokenType;
    function DoIdentifier: TsTokenType;
    function DoNumber: TsTokenType;
    function DoDelimiter: TsTokenType;
    function DoSquareBracket: TsTokenType;
    function DoString: TsTokenType;
    function NextPos: Char; // inline;
    procedure SkipWhiteSpace; // inline;
    function IsWordDelim(C: Char): Boolean; // inline;
    function IsDelim(C: Char): Boolean; // inline;
    function IsDigit(C: Char): Boolean; // inline;
    function IsAlpha(C: Char): Boolean; // inline;
  public
    constructor Create(AParser: TsExpressionParser);
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
    FDialect: TsFormulaDialect;
    FActiveCell: PCell;
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
    FFormatSettings: TFormatSettings;
    class function BuiltinExpressionManager: TsBuiltInExpressionManager;
    function BuildStringFormula(AFormatSettings: TFormatSettings): String;
    procedure ParserError(Msg: String);
    function GetExpression: String;
    function GetLocalizedExpression(const AFormatSettings: TFormatSettings): String; virtual;
    procedure SetExpression(const AValue: String);
    procedure SetLocalizedExpression(const AFormatSettings: TFormatSettings;
      const AValue: String); virtual;
    procedure CheckResultType(const Res: TsExpressionResult;
      AType: TsResultType); inline;
    function CurrentToken: String;
    function CurrentOrEOFToken: String;
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
    function Evaluate: TsExpressionResult;
    procedure EvaluateExpression(out Result: TsExpressionResult);
    function ResultType: TsResultType;
    function SharedFormulaMode: Boolean;

    property AsFloat: TsExprFloat read GetAsFloat;
    property AsInteger: Int64 read GetAsInteger;
    property AsString: String read GetAsString;
    property AsBoolean: Boolean read GetAsBoolean;
    property AsDateTime: TDateTime read GetAsDateTime;
    // The expression to parse
    property Expression: String read GetExpression write SetExpression;
    property LocalizedExpression[AFormatSettings: TFormatSettings]: String
        read GetLocalizedExpression write SetLocalizedExpression;
    property RPNFormula: TsRPNFormula read GetRPNFormula write SetRPNFormula;
    property Identifiers: TsExprIdentifierDefs read FIdentifiers write SetIdentifiers;
    property BuiltIns: TsBuiltInExprCategories read FBuiltIns write SetBuiltIns;
    property ActiveCell: PCell read FActiveCell write FActiveCell;
    property Worksheet: TsWorksheet read FWorksheet;
    property Dialect: TsFormulaDialect read FDialect write FDialect;
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
function ArgToBoolean(Arg: TsExpressionResult): Boolean;
function ArgToCell(Arg: TsExpressionResult): PCell;
function ArgToDateTime(Arg: TsExpressionResult): TDateTime;
function ArgToInt(Arg: TsExpressionResult): Integer;
function ArgToFloat(Arg: TsExpressionResult): TsExprFloat;
function ArgToString(Arg: TsExpressionResult): String;
procedure ArgsToFloatArray(const Args: TsExprParameterArray; out AData: TsExprFloatArray);
function BooleanResult(AValue: Boolean): TsExpressionResult;
function CellResult(AValue: String): TsExpressionResult; overload;
function CellResult(ACellRow, ACellCol: Cardinal): TsExpressionResult; overload;
function DateTimeResult(AValue: TDateTime): TsExpressionResult;
function EmptyResult: TsExpressionResult;
function ErrorResult(const AValue: TsErrorValue): TsExpressionResult;
function FloatResult(const AValue: TsExprFloat): TsExpressionResult;
function IntegerResult(const AValue: Integer): TsExpressionResult;
function StringResult(const AValue: String): TsExpressionResult;

procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallBack: TsExprFunctionCallBack); overload;
procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallBack: TsExprFunctionEvent); overload;

var
  ExprFormatSettings: TFormatSettings;

const
  AllBuiltIns = [bcMath, bcStatistics, bcStrings, bcLogical, bcDateTime, bcLookup,
    bcInfo, bcUser];


implementation

uses
  typinfo, math, lazutf8, dateutils, fpsutils; //, fpsfunc;

const
  cNull = #0;
  cDoubleQuote = '"';
  cError = '#';

  Digits         = ['0'..'9'];   // + decimalseparator
  WhiteSpace     = [' ', #13, #10, #9];
  Operators      = ['+', '-', '<', '>', '=', '/', '*', '&', '%', '^'];
  Delimiters     = Operators + ['(', ')'];  // + listseparator
  Symbols        = Delimiters;
  WordDelimiters = WhiteSpace + Symbols;

resourcestring
  SBadQuotes = 'Unterminated string';
  SUnknownDelimiter = 'Unknown delimiter character: "%s"';
  SErrUnknownCharacter = 'Unknown character at pos %d: "%s"';
  SErrUnexpectedEndOfExpression = 'Unexpected end of expression';
  SErrUnknownComparison = 'Internal error: Unknown comparison';
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
  SErrInvalidCell = 'No valid cell address specification : %s';
  SErrInvalidCellRange = 'No valid cell range specification : %s';
  SErrNoOperand = 'No operand for unary operation %s';
  SErrNoLeftOperand = 'No left operand for binary operation %s';
  SErrNoRightOperand = 'No left operand for binary operation %s';
  SErrNoNegation = 'Cannot negate expression of type %s: %s';
  SErrNoUPlus = 'Cannot perform unary plus operation on type %s: %s';
  SErrNoNOTOperation = 'Cannot perform NOT operation on expression of type %s: %s';
  SErrNoPercentOperation = 'Cannot perform percent operation on expression of type %s: %s';
  SErrTypesDoNotMatch = 'Type mismatch: %s<>%s for expressions "%s" and "%s".';
  SErrNoNodeToCheck = 'Internal error: No node to check !';
  SInvalidNodeType = 'Node type (%s) not in allowed types (%s) for expression: %s';
  SErrUnterminatedExpression = 'Badly terminated expression. Found token at position %d : %s';
  SErrDuplicateIdentifier = 'An identifier with name "%s" already exists.';
  SErrInvalidResultCharacter = '"%s" is not a valid return type indicator';
  ErrInvalidArgumentCount = 'Invalid argument count for function %s';
  SErrInvalidResultType = 'Invalid result type: %s';
  SErrNotVariable = 'Identifier %s is not a variable';
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

constructor TsExpressionScanner.Create(AParser: TsExpressionParser);
begin
  Source := '';
  FParser := AParser;
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
  if D = FParser.FFormatSettings.ListSeparator then
    Result := ttListSep
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
  //    ',' : Result := ttComma;
    else
      ScanError(Format(SUnknownDelimiter, [D]));
    end;
end;

function TsExpressionScanner.DoError: TsTokenType;
var
  C: Char;
begin
  C := CurrentChar;
  while (not IsWordDelim(C)) and (C <> cNull) do
  begin
    FToken := FToken + C;
    C := NextPos;
  end;
  Result := ttError;
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
  if not TryStrToFloat(FToken, X, FParser.FFormatSettings) then
    ScanError(Format(SErrInvalidNumber, [FToken]));
  Result := ttNumber;
end;

{ Scans until closing square bracket is reached. In OpenDocument, this is
  a cell or cell range identifier. }
function TsExpressionScanner.DoSquareBracket: TsTokenType;
var
  C: Char;
  p: Integer;
  r1,c1,r2,c2: Cardinal;
  flags: TsRelFlags;
begin
  FToken := '';
  C := NextPos;
  while (C <> ']') do
  begin
    if C = cNull then
      ScanError(SErrUnexpectedEndOfExpression);
    FToken := FToken + C;
    C := NextPos;
  end;
  C := NextPos;
  p := system.pos('.', FToken);  // Delete up tp "."  (--> to be considered later!)
  if p <> 0 then Delete(FToken, 1, p);
  if system.pos(':', FToken) > 0 then
  begin
    if ParseCellRangeString(FToken, r1, c1, r2, c2, flags) then
      Result := ttCellRange
    else
      ScanError(Format(SErrInvalidCellRange, [FToken]));
  end else
  if ParseCellString(FToken, r1, c1, flags) then
    Result := ttCell
  else
    ScanError(Format(SErrInvalidCell, [FToken]));
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
    FToken := FToken + C;
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
  if (FParser.Dialect = fdOpenDocument) and (C = '[') then
    Result := DoSquareBracket
  else if C = cNull then
    Result := ttEOF
  else if IsDelim(C) then
    Result := DoDelimiter
  else if (C = cDoubleQuote) then
    Result := DoString
  else if IsDigit(C) then
    Result := DoNumber
  else if (C = cError) then
    Result := DoError
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
  Result := (C in Delimiters) or (C = FParser.FFormatSettings.ListSeparator);
end;

function TsExpressionScanner.IsDigit(C: Char): Boolean;
begin
  Result := (C in Digits) or (C = FParser.FFormatSettings.DecimalSeparator);
end;

function TsExpressionScanner.IsWordDelim(C: Char): Boolean;
begin
  Result := (C in WordDelimiters) or (C = FParser.FFormatSettings.ListSeparator);
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
  FDialect := fdExcel;
  FWorksheet := AWorksheet;
  FIdentifiers := TsExprIdentifierDefs.Create(TsExprIdentifierDef);
  FIdentifiers.FParser := Self;
  FScanner := TsExpressionScanner.Create(self);
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

{ Constructs the string formula from the tree of expression nodes. Gets the
  decimal and list separator from the formatsettings provided. }
function TsExpressionParser.BuildStringFormula(AFormatSettings: TFormatSettings): String;
begin
  ExprFormatSettings := AFormatSettings;
  if FExprNode = nil then
    Result := ''
  else
  begin
    FFormatSettings := AFormatSettings;
    Result := FExprNode.AsString;
  end;
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
        rtFloat    : Result := TsIntToFloatExprNode.Create(self, Result);
        rtDateTime : Result := TsIntToDateTimeExprNode.Create(self, Result);
      end;
    rtFloat :
      case ToType of
        rtDateTime : Result := TsFloatToDateTimeExprNode.Create(self, Result);
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

function TsExpressionParser.CurrentOrEOFToken: String;
begin
  if (FScanner.TokenType = ttEOF) or (FScanner.Token = '') then
    Result := 'end of formula'
  else
    Result := FScanner.Token;
end;

function TsExpressionParser.Evaluate: TsExpressionResult;
begin
  EvaluateExpression(Result);
end;

procedure TsExpressionParser.EvaluateExpression(out Result: TsExpressionResult);
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
begin
  if FDirty then
    CreateHashList;
  Result := TsExprIdentifierDef(FHashList.Find(UpperCase(AName)));
end;

function TsExpressionParser.Level1: TsExprNode;
{
var
  tt: TsTokenType;
  Right: TsExprNode;
  }
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
      Result := C.Create(self, Result, right);
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
        ttPlus  : Result := TsAddExprNode.Create(self, Result, right);
        ttMinus : Result := TsSubtractExprNode.Create(self, Result, right);
        ttConcat: Result := TsConcatExprNode.Create(self, Result, right);
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
        ttMul : Result := TsMultiplyExprNode.Create(self, Result, right);
        ttDiv : Result := TsDivideExprNode.Create(self, Result, right);
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
    Result := TsUPlusExprNode.Create(self, Result);
  if isMinus then
    Result := TsUMinusExprNode.Create(self, Result);
  if TokenType = ttPercent then begin
    Result := TsPercentExprNode.Create(self, Result);
    GetToken;
  end;
end;

function TsExpressionParser.Level6: TsExprNode;
var
  Right: TsExprNode;
  currToken: String;
begin
{$ifdef debugexpr} Writeln('Level 6 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  if (TokenType = ttLeft) then
  begin
    GetToken;
    Result := TsParenthesisExprNode.Create(self, Level1);
    try
      if (TokenType <> ttRight) then begin
        currToken := CurrentToken;
        if TokenType = ttEOF then currToken := 'end of formula';
        ParserError(Format(SErrBracketExpected, [SCanner.Pos, currToken]));
      end;
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
      GetToken;
      Right := Primitive;
      CheckNodes(Result, right);
      Result := TsPowerExprNode.Create(self, Result, Right);
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
  optional: Boolean;
  token: String;
begin
{$ifdef debugexpr} Writeln('Primitive : ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  SetLength(Args, 0);
  if (TokenType = ttNumber) then
  begin
    if TryStrToInt64(CurrentToken, I) then
      Result := TsConstExprNode.CreateInteger(self, I)
    else
    begin
      if TryStrToFloat(CurrentToken, X, FFormatSettings) then
        Result := TsConstExprNode.CreateFloat(self, X)
      else
        ParserError(Format(SErrInvalidFloat, [CurrentToken]));
    end;
  end
  else if (TokenType = ttTrue) then
    Result := TsConstExprNode.CreateBoolean(self, true)
  else if (TokenType = ttFalse) then
    Result := TsConstExprNode.CreateBoolean(self, false)
  else if (TokenType = ttString) then
    Result := TsConstExprNode.CreateString(self, CurrentToken)
  else if (TokenType = ttCell) then
    Result := TsCellExprNode.Create(self, FWorksheet, CurrentToken)
  else if (TokenType = ttCellRange) then
    Result := TsCellRangeExprNode.Create(self, FWorksheet, CurrentToken)
  else if (TokenType = ttError) then
    Result := tsConstExprNode.CreateError(self, CurrentToken)
  else if not (TokenType in [ttIdentifier]) then
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
          ParserError(Format(SErrLeftBracketExpected, [Scanner.Pos, CurrentOrEOFToken]));
        GetToken;
        if (TokenType <> ttRight) then
          ParserError(Format(SErrBracketExpected, [Scanner.Pos, CurrentOrEOFToken]));
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
        ParserError(Format(SErrLeftBracketExpected, [Scanner.Pos, CurrentOrEofToken]));
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
            if (TokenType <> ttListSep) then
              if (AI < abs(lCount)) then
                ParserError(Format(SErrCommaExpected, [Scanner.Pos, CurrentOrEofToken]))
          end;
        until (AI = lCount) or (((lCount < 0) or optional) and (TokenType = ttRight));
        if TokenType <> ttRight then
          ParserError(Format(SErrBracketExpected, [Scanner.Pos, CurrentOrEofToken]));
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
      itVariable:
        Result := TsVariableExprNode.CreateIdentifier(self, ID);
      itFunctionCallBack:
        Result := TsFunctionCallBackExprNode.CreateFunction(self, ID, Args);
      itFunctionHandler:
        Result := TFPFunctionEventHandlerExprNode.CreateFunction(self, ID, Args);
    end;
  end;
  GetToken;
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

function TsExpressionParser.GetExpression: String;
var
  fs: TFormatsettings;
begin
  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';
  fs.ListSeparator := ',';
  Result := BuildStringFormula(fs);
end;

function TsExpressionParser.GetLocalizedExpression(const AFormatSettings: TFormatSettings): String;
begin
  ExprFormatSettings := AFormatSettings;
  Result := BuildStringFormula(AFormatSettings);
end;

procedure TsExpressionParser.SetExpression(const AValue: String);
var
  fs: TFormatSettings;
begin
  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';
  fs.ListSeparator := ',';
  SetLocalizedExpression(fs, AValue);
end;

procedure TsExpressionParser.SetLocalizedExpression(const AFormatSettings: TFormatSettings;
  const AValue: String);
begin
  if FExpression = AValue then
    exit;
  FFormatSettings := AFormatSettings;
  ExprFormatSettings := AFormatSettings;
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
    left: TsExprNode = nil;
    right: TsExprNode = nil;
    operand: TsExprNode = nil;
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
          ANode := TsCellExprNode.Create(self, FWorksheet, r, c, flags);
          dec(AIndex);
        end;
      fekCellRange:
        begin
          r := AFormula[AIndex].Row;
          c := AFormula[AIndex].Col;
          r2 := AFormula[AIndex].Row2;
          c2 := AFormula[AIndex].Col2;
          flags := AFormula[AIndex].RelFlags;
          ANode := TsCellRangeExprNode.Create(self, FWorksheet, r, c, r2, c2, flags);
          dec(AIndex);
        end;
      fekNum:
        begin
          ANode := TsConstExprNode.CreateFloat(self, AFormula[AIndex].DoubleValue);
          dec(AIndex);
        end;
      fekInteger:
        begin
          ANode := TsConstExprNode.CreateInteger(self, AFormula[AIndex].IntValue);
          dec(AIndex);
        end;
      fekString:
        begin
          ANode := TsConstExprNode.CreateString(self, AFormula[AIndex].StringValue);
          dec(AIndex);
        end;
      fekBool:
        begin
          ANode := TsConstExprNode.CreateBoolean(self, AFormula[AIndex].DoubleValue <> 0.0);
          dec(AIndex);
        end;
      fekErr:
        begin
          ANode := TsConstExprNode.CreateError(self, TsErrorValue(AFormula[AIndex].IntValue));
          dec(AIndex);
        end;

      // unary operations
      fekPercent, fekUMinus, fekUPlus, fekParen:
        begin
          dec(AIndex);
          CreateNodeFromRPN(operand, AIndex);
          case fek of
            fekPercent : ANode := TsPercentExprNode.Create(self, operand);
            fekUMinus  : ANode := TsUMinusExprNode.Create(self, operand);
            fekUPlus   : ANode := TsUPlusExprNode.Create(self, operand);
            fekParen   : ANode := TsParenthesisExprNode.Create(self, operand);
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
            fekAdd         : ANode := TsAddExprNode.Create(self, left, right);
            fekSub         : ANode := TsSubtractExprNode.Create(self, left, right);
            fekMul         : ANode := TsMultiplyExprNode.Create(self, left, right);
            fekDiv         : ANode := TsDivideExprNode.Create(self, left, right);
            fekPower       : ANode := TsPowerExprNode.Create(self, left, right);
            fekConcat      : ANode := tsConcatExprNode.Create(self, left, right);
            fekEqual       : ANode := TsEqualExprNode.Create(self, left, right);
            fekNotEqual    : ANode := TsNotEqualExprNode.Create(self, left, right);
            fekGreater     : ANode := TsGreaterExprNode.Create(self, left, right);
            fekGreaterEqual: ANode := TsGreaterEqualExprNode.Create(self, left, right);
            fekLess        : ANode := TsLessExprNode.Create(self, left, right);
            fekLessEqual   : ANode := tsLessEqualExprNode.Create(self, left, right);
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
              itVariable:
                ANode := TsVariableExprNode.CreateIdentifier(self, ID);
              itFunctionCallBack:
                ANode := TsFunctionCallBackExprNode.CreateFunction(self, ID, args);
              itFunctionHandler:
                ANode := TFPFunctionEventHandlerExprNode.CreateFunction(self, ID, args);
            end;
          end;
        end;

    end;  //case
  end; //begin

var
  index: Integer;
begin
  FExpression := '';
  FreeAndNil(FExprNode);
  index := Length(AFormula)-1;
  CreateNodeFromRPN(FExprNode, index);
  if Assigned(FExprNode) then FExprNode.Check;
end;

{ Signals that the parser is in SharedFormulaMode, i.e. there is an active cell
  to which all relative addresses have to be adapted. }
function TsExpressionParser.SharedFormulaMode: Boolean;
begin
  Result := (ActiveCell <> nil) and (ActiveCell^.SharedFormulaBase <> nil);
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
  Unused(Item);
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

function TsExprIdentifierDef.GetFormatSettings: TFormatSettings;
begin
  Result := TsExprIdentifierDefs(Collection).Parser.FFormatSettings;
end;

function TsExprIdentifierDef.GetResultType: TsResultType;
begin
  Result := FValue.ResultType;
end;

function TsExprIdentifierDef.GetValue: String;
begin
  case FValue.ResultType of
    rtBoolean  : if FValue.ResBoolean then
                   Result := 'TRUE'
                 else
                   Result := 'FALSE';
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtFloat    : Result := FloatToStr(FValue.ResFloat, GetFormatSettings);
    rtDateTime : Result := FormatDateTime('cccc', FValue.ResDateTime, GetFormatSettings);
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
  FValue.ResString := AValue;
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
      rtBoolean  : FValue.ResBoolean := (FStringValue='True');
      rtInteger  : FValue.ResInteger := StrToInt(AValue);
      rtFloat    : FValue.ResFloat := StrToFloat(AValue, GetFormatSettings);
      rtDateTime : FValue.ResDateTime := StrToDateTime(AValue, GetFormatSettings);
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

constructor TsUnaryOperationExprNode.Create(AParser: TsExpressionParser;
  AOperand: TsExprNode);
begin
  FParser := AParser;
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

constructor TsBinaryOperationExprNode.Create(AParser: TsExpressionParser;
  ALeft, ARight: TsExprNode);
begin
  FParser := AParser;
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

constructor TsConstExprNode.CreateString(AParser: TsExpressionParser;
  AValue: String);
begin
  FParser := AParser;
  FValue.ResultType := rtString;
  FValue.ResString := AValue;
end;

constructor TsConstExprNode.CreateInteger(AParser: TsExpressionParser;
  AValue: Int64);
begin
  FParser := AParser;
  FValue.ResultType := rtInteger;
  FValue.ResInteger := AValue;
end;

constructor TsConstExprNode.CreateDateTime(AParser: TsExpressionParser;
  AValue: TDateTime);
begin
  FParser := AParser;
  FValue.ResultType := rtDateTime;
  FValue.ResDateTime := AValue;
end;

constructor TsConstExprNode.CreateFloat(AParser: TsExpressionParser;
  AValue: TsExprFloat);
begin
  FParser := AParser;
  FValue.ResultType := rtFloat;
  FValue.ResFloat := AValue;
end;

constructor TsConstExprNode.CreateBoolean(AParser: TsExpressionParser;
  AValue: Boolean);
begin
  FParser := AParser;
  FValue.ResultType := rtBoolean;
  FValue.ResBoolean := AValue;
end;

constructor TsConstExprNode.CreateError(AParser: TsExpressionParser;
  AValue: TsErrorValue);
begin
  FParser := AParser;
  FValue.ResultType := rtError;
  FValue.ResError := AValue;
end;

constructor TsConstExprNode.CreateError(AParser: TsExpressionParser;
  AValue: String);
var
  err: TsErrorValue;
begin
  if AValue = '#NULL!' then
    err := errEmptyIntersection
  else if AValue = '#DIV/0!' then
    err := errDivideByZero
  else if AValue = '#VALUE!' then
    err := errWrongType
  else if AVAlue = '#REF!' then
    err := errIllegalRef
  else if AVAlue = '#NAME?' then
    err := errWrongName
  else if AValue = '#FORMULA?' then
    err := errFormulaNotSupported
  else
    AParser.ParserError('Unknown error type.');
  CreateError(AParser, err);
end;

procedure TsConstExprNode.Check;
begin
  // Nothing to check;
end;

function TsConstExprNode.NodeType: TsResultType;
begin
  Result := FValue.ResultType;
end;

procedure TsConstExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Result := FValue;
end;

function TsConstExprNode.AsString: string;
begin
  case NodeType of
    rtString   : Result := cDoubleQuote + FValue.ResString + cDoubleQuote;
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtDateTime : Result := '''' + FormatDateTime('cccc', FValue.ResDateTime, Parser.FFormatSettings) + '''';    // Probably wrong !!!
    rtBoolean  : if FValue.ResBoolean then Result := 'TRUE' else Result := 'FALSE';
    rtFloat    : Result := FloatToStr(FValue.ResFloat, Parser.FFormatSettings);
    rtError    : Result := GetErrorValueStr(FValue.ResError);
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
    rtError    : Result := RPNErr(FValue.ResError, ANext);
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

procedure TsUPlusExprNode.GetNodeValue(out Result: TsExpressionResult);
var
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

procedure TsUMinusExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsPercentExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsParenthesisExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsNotExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsEqualExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsNotEqualExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsLessExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsGreaterExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsGreaterEqualExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsLessEqualExprNode.GetNodeValue(out Result: TsExpressionResult);
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
    ANext)));
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

procedure TsConcatExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsAddExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsSubtractExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsMultiplyExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsDivideExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsPowerExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsIntToFloatExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsIntToDateTimeExprNode.GetNodeValue(out Result: TsExpressionResult);
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

procedure TsFloatToDateTimeExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  if Result.ResultType in [rtFloat, rtCell] then
    Result := DateTimeResult(ArgToFloat(Result));
end;


{ TsIdentifierExprNode }

constructor TsIdentifierExprNode.CreateIdentifier(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef);
begin
  FParser := AParser;
  FID := AID;
  PResult := @FID.FValue;
  FResultType := FID.ResultType;
end;

function TsIdentifierExprNode.NodeType: TsResultType;
begin
  Result := FResultType;
end;

procedure TsIdentifierExprNode.GetNodeValue(out Result: TsExpressionResult);
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
  Result := ANext;  // Just a dummy assignment to silence the compiler...
  RaiseParserError('Cannot handle variables for RPN, so far.');
end;

function TsVariableExprNode.AsString: string;
begin
  Result := FID.Name;
end;


{ TsFunctionExprNode }

constructor TsFunctionExprNode.CreateFunction(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef; const Args: TsExprArgumentArray);
begin
  inherited CreateIdentifier(AParser, AID);
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
      S := S + Parser.FFormatSettings.ListSeparator;
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
  rta,                  // Parameter types passed to the function
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

    if i+1 <= Length(FID.ParameterTypes) then
    begin
      rtp := CharToResultType(FID.ParameterTypes[i+1]);
      lastrtp := rtp;
    end else
      rtp := lastrtp;

    if rtp = rtAny then
      Continue;
    // A "cell" can return any type --> no type conversion required here.

    if rta = rtCell then
      Continue;
                    (*
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
    *)
  end;
end;


{ TsFunctionCallBackExprNode }

constructor TsFunctionCallBackExprNode.CreateFunction(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef; const Args: TsExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValueCallBack;
end;

procedure TsFunctionCallBackExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Result.ResultType := NodeType;     // was at end!
  if Length(FArgumentParams) > 0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
end;


{ TFPFunctionEventHandlerExprNode }

constructor TFPFunctionEventHandlerExprNode.CreateFunction(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef; const Args: TsExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValue;
end;

procedure TFPFunctionEventHandlerExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Result.ResultType := NodeType;    // was at end
  if Length(FArgumentParams) > 0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
end;


{ TsCellExprNode }

constructor TsCellExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ACellString: String);
var
  r, c: Cardinal;
  flags: TsRelFlags;
begin
  ParseCellString(ACellString, r, c, flags);
  Create(AParser, AWorksheet, r, c, flags);
end;

constructor TsCellExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ARow,ACol: Cardinal; AFlags: TsRelFlags);
begin
  FParser := AParser;
  FWorksheet := AWorksheet;
  FRow := ARow;
  FCol := ACol;
  FFlags := AFlags;
  FCell := AWorksheet.FindCell(FRow, FCol);
end;

function TsCellExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  if FIsRef then
    Result := RPNCellRef(GetRow, GetCol, FFlags, ANext)
  else
    Result := RPNCellValue(GetRow, GetCol, FFlags, ANext);
end;

function TsCellExprNode.AsString: string;
begin
  Result := GetCellString(GetRow, GetCol, FFlags);
  if FParser.Dialect = fdOpenDocument then
    Result := '[.' + Result + ']';
end;

procedure TsCellExprNode.Check;
begin
  // Nothing to check;
end;

{ Calculates the row address of the node's cell for various cases:
  (1) SharedFormula mode:
      The "ActiveCell" of the parser is the cell for which the formula is
      calculated. If the formula contains a relative address in the cell node
      the function calculates the row address of the cell represented by the
      node as seen from the active cell.
      If the formula contains an absolute address the function returns the row
      address of the SharedFormulaBase of the ActiveCell.
  (2) Normal mode:
      Returns the "true" row address of the cell assigned to the formula node. }
function TsCellExprNode.GetCol: Cardinal;
begin
  if FParser.SharedFormulaMode then
  begin
    // A shared formula is stored in the SharedFormulaBase cell of the ActiveCell
    // Since the cell data stored in the node are those used by the formula in
    // the SharedFormula, the current node is relative to the SharedFormulaBase
    if rfRelCol in FFlags then
      Result := FCol - FParser.ActiveCell^.SharedFormulaBase^.Col + FParser.ActiveCell^.Col
    else
      Result := FCol; //FParser.ActiveCell^.SharedFormulaBase^.Col;
  end
  else
    // Normal mode
    Result := FCol;
end;

procedure TsCellExprNode.GetNodeValue(out Result: TsExpressionResult);
var
  cell: PCell;
begin
  if Parser.SharedFormulaMode then
    cell := FWorksheet.FindCell(GetRow, GetCol)
  else
    cell := FCell;

  if (cell <> nil) and HasFormula(cell) then
    case cell^.CalcState of
      csNotCalculated:
        Worksheet.CalcFormula(cell);
      csCalculating:
        raise Exception.Create(SErrCircularReference);
    end;

  Result.ResultType := rtCell;
  Result.ResRow := GetRow;
  Result.ResCol := GetCol;
  Result.Worksheet := FWorksheet;
end;

{ See GetCol }
function TsCellExprNode.GetRow: Cardinal;
begin
  if Parser.SharedFormulaMode then
  begin
    if rfRelRow in FFlags then
      Result := FRow - FParser.ActiveCell^.SharedFormulaBase^.Row + FParser.ActiveCell^.Row
    else
      Result := FRow; //FParser.ActiveCell^.SharedFormulaBase^.Row;
  end
  else
    Result := FRow;
end;

function TsCellExprNode.NodeType: TsResultType;
begin
  Result := rtCell;
end;


{ TsCellRangeExprNode }

constructor TsCellRangeExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ACellRangeString: String);
var
  r1, c1, r2, c2: Cardinal;
  flags: TsRelFlags;
begin
  if pos(':', ACellRangeString) = 0 then
  begin
    ParseCellString(ACellRangeString, r1, c1, flags);
    if rfRelRow in flags then Include(flags, rfRelRow2);
    if rfRelCol in flags then Include(flags, rfRelCol2);
    Create(AParser, AWorksheet, r1, c1, r1, c1, flags);
  end else
  begin
    ParseCellRangeString(ACellRangeString, r1, c1, r2, c2, flags);
    Create(AParser, AWorksheet, r1, c1, r2, c2, flags);
  end;
end;

constructor TsCellRangeExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ARow1,ACol1,ARow2,ACol2: Cardinal; AFlags: TsRelFlags);
begin
  FParser := AParser;
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

procedure TsCellRangeExprNode.GetNodeValue(out Result: TsExpressionResult);
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
  case Arg.ResultType of
    rtInteger  : result := Arg.ResInteger;
    rtFloat    : result := trunc(Arg.ResFloat);
    rtDateTime : result := trunc(Arg.ResDateTime);
    rtBoolean  : if Arg.ResBoolean then Result := 1 else Result := 0;
    rtString   : if not TryStrToInt(Arg.ResString, Result) then Result := 0;
    rtCell     : begin
                   cell := ArgToCell(Arg);
                   if Assigned(cell) then
                     case cell^.ContentType of
                       cctNumber    : result := trunc(cell^.NumberValue);
                       cctDateTime  : result := trunc(cell^.DateTimeValue);
                       cctBool      : if cell^.BoolValue then result := 1;
                       cctUTF8String: if not TryStrToInt(cell^.UTF8StringValue, result)
                                        then Result := 0;
                     end;
                 end;
  end;
end;

function ArgToFloat(Arg: TsExpressionResult): TsExprFloat;
// Utility function for the built-in math functions. Accepts also integers and
// other data types in place of floating point arguments. To be called in
// builtins or user-defined callbacks having float results or arguments.
var
  cell: PCell;
  s: String;
  fs: TFormatSettings;
begin
  Result := 0.0;
  case Arg.ResultType of
    rtInteger  : result := Arg.ResInteger;
    rtDateTime : result := Arg.ResDateTime;
    rtFloat    : result := Arg.ResFloat;
    rtBoolean  : if Arg.ResBoolean then Result := 1.0 else Result := 0.0;
    rtString   : if not TryStrToFloat(Arg.ResString, Result) then Result := 0.0;
    rtCell     : begin
                   cell := ArgToCell(Arg);
                   if Assigned(cell) then
                     case cell^.ContentType of
                       cctNumber    : Result := cell^.NumberValue;
                       cctDateTime  : Result := cell^.DateTimeValue;
                       cctBool      : if cell^.BoolValue then result := 1.0;
                       cctUTF8String: begin
                                        fs := Arg.Worksheet.Workbook.FormatSettings;
                                        s := cell^.UTF8StringValue;
                                        if not TryStrToFloat(s, result, fs) then
                                          result := 0.0;
                                      end;
                     end;
                 end;
  end;
end;

function ArgToDateTime(Arg: TsExpressionResult): TDateTime;
var
  cell: PCell;
  fs: TFormatSettings;
begin
  Result := 0.0;
  case Arg.ResultType of
    rtDateTime  : result := Arg.ResDateTime;
    rtInteger   : Result := Arg.ResInteger;
    rtFloat     : Result := Arg.ResFloat;
    rtBoolean   : if Arg.ResBoolean then Result := 1.0;
    rtString    : begin
                    fs := Arg.Worksheet.Workbook.FormatSettings;
                    if not TryStrToDateTime(Arg.ResString, Result, fs) then
                      Result := 1.0;
                  end;
    rtCell      : begin
                    cell := ArgToCell(Arg);
                    if Assigned(cell) then
                      if (cell^.ContentType = cctDateTime) then
                        Result := cell^.DateTimeValue;
                  end;
  end;
end;

function ArgToString(Arg: TsExpressionResult): String;
// The Office applications are very fuzzy about data types...
var
  cell: PCell;
  fs: TFormatSettings;
  dt: TDateTime;
begin
  Result := '';
  case Arg.ResultType of
    rtString  : result := Arg.ResString;
    rtInteger : Result := IntToStr(Arg.ResInteger);
    rtFloat   : Result := FloatToStr(Arg.ResFloat);
    rtBoolean : if Arg.ResBoolean then Result := '1' else Result := '0';
    rtCell    : begin
                  cell := ArgToCell(Arg);
                  if Assigned(cell) then
                    case cell^.ContentType of
                      cctUTF8String : Result := cell^.UTF8Stringvalue;
                      cctNumber     : Result := Format('%g', [cell^.NumberValue]);
                      cctBool       : if cell^.BoolValue then Result := '1' else Result := '0';
                      cctDateTime   : begin
                                        fs := Arg.Worksheet.Workbook.FormatSettings;
                                        dt := cell^.DateTimeValue;
                                        if frac(dt) = 0.0 then
                                          Result := FormatDateTime(fs.LongTimeFormat, dt, fs)
                                        else
                                        if trunc(dt) = 0 then
                                          Result := FormatDateTime(fs.ShortDateFormat, dt, fs)
                                        else
                                          Result := FormatDateTime('cc', dt, fs);
                                      end;
                    end;
                end;
  end;
end;

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
    if (arg.ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtBoolean]) then
    begin
      AData[n] := ArgToFloat(arg);
      inc(n);
      if n = Length(AData) then SetLength(AData, Length(AData) + BLOCKSIZE);
    end;
  end;
  SetLength(AData, n);
end;


{------------------------------------------------------------------------------}
{   Conversion simple data types to ExpressionResults                          }
{------------------------------------------------------------------------------}

function BooleanResult(AValue: Boolean): TsExpressionResult;
begin
  Result.ResultType := rtBoolean;
  Result.ResBoolean := AValue;
end;

function CellResult(AValue: String): TsExpressionResult;
begin
  Result.ResultType := rtCell;
  ParseCellString(AValue, Result.ResRow, Result.ResCol);
end;

function CellResult(ACellRow, ACellCol: Cardinal): TsExpressionResult;
begin
  Result.ResultType := rtCell;
  Result.ResRow := ACellRow;
  Result.ResCol := ACellCol;
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

procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallback: TsExprFunctionEvent);
begin
  with BuiltinIdentifiers do
    AddFunction(bcUser, AName, AResultType, AParamTypes, AExcelCode, ACallBack);
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

//  RegisterStdBuiltins(BuiltinIdentifiers);

finalization
  FreeBuiltins;
end.
