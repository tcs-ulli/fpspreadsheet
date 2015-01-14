unit SynHighlighterWikitable;

{$IFDEF FPC}
  {$MODE OBJFPC}{$H+}
{$ENDIF}

interface

uses
  SysUtils, Classes,
  LCLIntf, LCLType,
  Controls, Graphics,
  SynEditTypes, SynEditHighlighter;

type
  TtkTokenKind = (tkAmpersand, tkComment, tkIdentifier, tkNull, tkNumber,
    tkSpace, tkString, tkSymbol, tkText, tkUnknown);
  {
  TtkTokenKind = (tkComment, tkIdentifier, tkKey, tkNull, tkNumber, tkSpace,
    tkString, tkSymbol, tkUnknown);
   }

  TRangeState = (rsUnknown, rsSymbol, rsParam, rsValue, rsComment, rsText, rsAmpersand);

  TProcTableProc = procedure of object;

  PIdentFuncTableFunc = ^TIdentFuncTableFunc;
  TIdentFuncTableFunc = function: TtkTokenKind of object;

  TSynWikiTableSyn = class(TSynCustomHighlighter)
  private
    FLine: PChar;
    FLineNumber: Integer;
    FTokenPos: Integer;
    FTokenID: TtkTokenKind;
    FRange: TRangeState;
    Run: LongInt;
    FAmpersandCode: Integer;
    FStringLen: Integer;
    FToIdent: PChar;
    FProcTable: array[#0..#255] of TProcTableProc;
    FIdentFuncTable: array[0..255] of TIdentFuncTableFunc;
    FCommentAttri: TSynHighlighterAttributes;
    FNumberAttri: TSynHighlighterAttributes;
    FSpaceAttri: TSynHighlighterAttributes;
    FSymbolAttri: TSynHighlighterAttributes;
    FIdentifierAttri: TSynHighlighterAttributes;
    FStringAttri: TSynHighlighterAttributes;
    (*
    FKeyAttri: TSynHighlighterAttributes;
    FNumberAttri: TSynHighlighterAttributes;
    *)
    procedure InitIdent;
    function IdentKind(MayBe: PChar): TtkTokenKind;
    function KeyHash(ToHash: PChar): Integer;
    function KeyComp(const aKey: string): Boolean;
    procedure MakeMethodTables;

    procedure AmpersandProc;
    procedure BarProc;
    procedure BeginProc;
    procedure CRProc;
    procedure ExclamProc;
    procedure LFProc;
    procedure IdentProc;
    procedure NullProc;
    procedure NumberProc;
    procedure OpenBraceProc;
    procedure SpaceProc;
    procedure StringProc;
    procedure TextProc;
    procedure UnknownProc;
  protected
  public
    constructor Create(AOwner: TComponent); override;
    function GetDefaultAttribute(Index: integer): TSynHighlighterAttributes; override;
    function GetEol: Boolean; override;
    function GetToken: string; override;
    procedure GetTokenEx(out TokenStart: PChar; out TokenLength: integer); override;
    function GetTokenID: TtkTokenKind;
    function GetTokenAttribute: TSynHighlighterAttributes; override;
    function GetTokenKind: integer; override;
    function GetTokenPos: Integer; override;
    procedure Next; override;
    procedure SetLine(const NewValue: String; LineNumber: Integer); override;
  published
    property CommentAttri: TSynHighlighterAttributes
      read FCommentAttri write FCommentAttri;
    property IdentifierAttri: TSynHighlighterAttributes
      read FIdentifierAttri write FIdentifierAttri;
    property NumberAttri: TSynHighlighterAttributes
      read FNumberAttri write FNumberAttri;
    property SpaceAttri: TSynHighlighterAttributes
      read FSpaceAttri write FSpaceAttri;
    property StringAttri: TSynHighlighterAttributes
      read FStringAttri write FStringAttri;
    property SymbolAttri: TSynHighlighterAttributes
      read FSymbolAttri write FSymbolAttri;

  end;


implementation

uses
  SynEditStrConst;

const
  // to do: remove next line when this identifier is in stable
  SYN_ATTR_NUMBER = 6; // not available in Laz 1.2.4

  MAX_ESCAPEAMPS = 159;

  EscapeAmps: array[0..MAX_ESCAPEAMPS - 1] of PChar = (
    ('&amp;'),               {   &   }
    ('&lt;'),                {   >   }
    ('&gt;'),                {   <   }
    ('&quot;'),              {   "   }
    ('&trade;'),             {      }
    ('&nbsp;'),              { space }
    ('&copy;'),              {   ©   }
    ('&reg;'),               {   ®   }
    ('&Agrave;'),            {   À   }
    ('&Aacute;'),            {   Á   }
    ('&Acirc;'),             {   Â   }
    ('&Atilde;'),            {   Ã   }
    ('&Auml;'),              {   Ä   }
    ('&Aring;'),             {   Å   }
    ('&AElig;'),             {   Æ   }
    ('&Ccedil;'),            {   Ç   }
    ('&Egrave;'),            {   È   }
    ('&Eacute;'),            {   É   }
    ('&Ecirc;'),             {   Ê   }
    ('&Euml;'),              {   Ë   }
    ('&Igrave;'),            {   Ì   }
    ('&Iacute;'),            {   Í   }
    ('&Icirc;'),             {   Î   }
    ('&Iuml;'),              {   Ï   }
    ('&ETH;'),               {   Ð   }
    ('&Ntilde;'),            {   Ñ   }
    ('&Ograve;'),            {   Ò   }
    ('&Oacute;'),            {   Ó   }
    ('&Ocirc;'),             {   Ô   }
    ('&Otilde;'),            {   Õ   }
    ('&Ouml;'),              {   Ö   }
    ('&Oslash;'),            {   Ø   }
    ('&Ugrave;'),            {   Ù   }
    ('&Uacute;'),            {   Ú   }
    ('&Ucirc;'),             {   Û   }
    ('&Uuml;'),              {   Ü   }
    ('&Yacute;'),            {   Ý   }
    ('&THORN;'),             {   Þ   }
    ('&szlig;'),             {   ß   }
    ('&agrave;'),            {   à   }
    ('&aacute;'),            {   á   }
    ('&acirc;'),             {   â   }
    ('&atilde;'),            {   ã   }
    ('&auml;'),              {   ä   }
    ('&aring;'),             {   å   }
    ('&aelig;'),             {   æ   }
    ('&ccedil;'),            {   ç   }
    ('&egrave;'),            {   è   }
    ('&eacute;'),            {   é   }
    ('&ecirc;'),             {   ê   }
    ('&euml;'),              {   ë   }
    ('&igrave;'),            {   ì   }
    ('&iacute;'),            {   í   }
    ('&icirc;'),             {   î   }
    ('&iuml;'),              {   ï   }
    ('&eth;'),               {   ð   }
    ('&ntilde;'),            {   ñ   }
    ('&ograve;'),            {   ò   }
    ('&oacute;'),            {   ó   }
    ('&ocirc;'),             {   ô   }
    ('&otilde;'),            {   õ   }
    ('&ouml;'),              {   ö   }
    ('&oslash;'),            {   ø   }
    ('&ugrave;'),            {   ù   }
    ('&uacute;'),            {   ú   }
    ('&ucirc;'),             {   û   }
    ('&uuml;'),              {   ü   }
    ('&yacute;'),            {   ý   }
    ('&thorn;'),             {   þ   }
    ('&yuml;'),              {   ÿ   }
    ('&iexcl;'),             {   ¡   }
    ('&cent;'),              {   ¢   }
    ('&pound;'),             {   £   }
    ('&curren;'),            {   ¤   }
    ('&yen;'),               {   ¥   }
    ('&brvbar;'),            {   ¦   }
    ('&sect;'),              {   §   }
    ('&uml;'),               {   ¨   }
    ('&ordf;'),              {   ª   }
    ('&laquo;'),             {   «   }
    ('&shy;'),               {   ¬   }
    ('&macr;'),              {   ¯   }
    ('&deg;'),               {   °   }
    ('&plusmn;'),            {   ±   }
    ('&sup2;'),              {   ²   }
    ('&sup3;'),              {   ³   }
    ('&acute;'),             {   ´   }
    ('&micro;'),             {   µ   }
    ('&middot;'),            {   ·   }
    ('&cedil;'),             {   ¸   }
    ('&sup1;'),              {   ¹   }
    ('&ordm;'),              {   º   }
    ('&raquo;'),             {   »   }
    ('&frac14;'),            {   ¼   }
    ('&frac12;'),            {   ½   }
    ('&frac34;'),            {   ¾   }
    ('&iquest;'),            {   ¿   }
    ('&times;'),             {   ×   }
    ('&divide'),             {   ÷   }
    ('&euro;'),              {      }
    ('&permil;'),
    ('&bdquo;'),
    ('&rdquo;'),
    ('&lsquo;'),
    ('&rsquo;'),
    ('&ndash;'),
    ('&mdash;'),
    ('&bull;'),
    //used by very old HTML editors
    ('&#9;'),                {  TAB  }
    ('&#127;'),              {      }
    ('&#128;'),              {      }
    ('&#129;'),              {      }
    ('&#130;'),              {      }
    ('&#131;'),              {      }
    ('&#132;'),              {      }
    ('&ldots;'),             {      }
    ('&#134;'),              {      }
    ('&#135;'),              {      }
    ('&#136;'),              {      }
    ('&#137;'),              {      }
    ('&#138;'),              {      }
    ('&#139;'),              {      }
    ('&#140;'),              {      }
    ('&#141;'),              {      }
    ('&#142;'),              {      }
    ('&#143;'),              {      }
    ('&#144;'),              {      }
    ('&#152;'),              {      }
    ('&#153;'),              {      }
    ('&#154;'),              {      }
    ('&#155;'),              {      }
    ('&#156;'),              {      }
    ('&#157;'),              {      }
    ('&#158;'),              {      }
    ('&#159;'),              {      }
    ('&#161;'),              {   ¡   }
    ('&#162;'),              {   ¢   }
    ('&#163;'),              {   £   }
    ('&#164;'),              {   ¤   }
    ('&#165;'),              {   ¥   }
    ('&#166;'),              {   ¦   }
    ('&#167;'),              {   §   }
    ('&#168;'),              {   ¨   }
    ('&#170;'),              {   ª   }
    ('&#175;'),              {   »   }
    ('&#176;'),              {   °   }
    ('&#177;'),              {   ±   }
    ('&#178;'),              {   ²   }
    ('&#180;'),              {   ´   }
    ('&#181;'),              {   µ   }
    ('&#183;'),              {   ·   }
    ('&#184;'),              {   ¸   }
    ('&#185;'),              {   ¹   }
    ('&#186;'),              {   º   }
    ('&#188;'),              {   ¼   }
    ('&#189;'),              {   ½   }
    ('&#190;'),              {   ¾   }
    ('&#191;'),              {   ¿   }
    ('&#215;'));             {   Ô   }

var
  Identifiers: array[#0..#255] of ByteBool;
  mHashTable: array[#0..#255] of Integer;

procedure MakeIdentTable;
var
  I, J: Char;
begin
  for I := #0 to #255 do begin
    case I of
      'a'..'z', 'A'..'Z', '-', '_', '0'..'9','@': Identifiers[I] := True;
    else
      Identifiers[I] := False;
    end;
    J := UpCase(I);
    if I in ['a'..'z', 'A'..'Z', '-', '_','@'] then
      mHashTable[I] := Ord(J) - 64
    else
      mHashTable[I] := 0;
  end;
end;

constructor TSynWikiTableSyn.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FCommentAttri := TSynHighlighterAttributes.Create(SYNS_AttrComment, SYNS_XML_AttrComment);
  FCommentAttri.Style := [fsItalic];
  FCommentAttri.Foreground := clTeal;
  AddAttribute(FCommentAttri);

  FNumberAttri := TSynHighlighterAttributes.Create(SYNS_AttrNumber, SYNS_XML_AttrNumber);
  FNumberAttri.Foreground := clBlue;
  AddAttribute(fNumberAttri);

  FSpaceAttri := TSynHighlighterAttributes.Create(SYNS_AttrSpace, SYNS_XML_AttrSpace);
  AddAttribute(FSpaceAttri);

  FSymbolAttri := TSynHighlighterAttributes.Create(SYNS_AttrSymbol, SYNS_XML_AttrSymbol);
  FSymbolAttri.Style := [fsBold];
  FSymbolAttri.Foreground := clPurple;
  AddAttribute(fSymbolAttri);

  FIdentifierAttri := TSynHighlighterAttributes.Create(SYNS_AttrIdentifier, SYNS_XML_AttrIdentifier);
  FIdentifierAttri.Foreground := clNavy;
  FIdentifierAttri.Style := [fsBold];
  AddAttribute(fIdentifierAttri);

  FStringAttri := TSynHighlighterAttributes.Create(SYNS_AttrString, SYNS_XML_AttrString);
  FStringAttri.Foreground := clOlive;
  AddAttribute(FStringAttri);
   (*
  fKeyAttri := TSynHighlighterAttributes.Create(SYNS_AttrKey, SYNS_XML_AttrKey);
  fKeyAttri.Style := [fsBold];
  AddAttribute(fKeyAttri);

  *)

  SetAttributesOnChange(@DefHighlightChange);
  InitIdent;
  MakeMethodTables;
//  fDefaultFilter := SYNS_FilterCSS;
  FRange := rsUnknown;
end;

procedure TSynWikiTableSyn.AmpersandProc;
begin
  case FAmpersandCode of
    Low(EscapeAmps)..High(EscapeAmps):
      begin
        FTokenID := tkAmpersand;
        inc(Run, StrLen(EscapeAmps[FAmpersandCode]));
      end;
  end;
  FAmpersandCode := -1;
  FRange := rsText;
end;

procedure TSynWikiTableSyn.BarProc;
begin
  FTokenID := tkSymbol;
  FRange := rsSymbol;
  inc(Run);
  if FLine[Run] in ['-', '}'] then inc(Run);
end;

procedure TSynWikiTableSyn.BeginProc;
begin
  inc(Run);
  if FLine[Run] = '|' then begin
    FTokenID := tkSymbol;
    FRange := rsSymbol;
    inc(Run);
  end;
end;

procedure TSynWikiTableSyn.CRProc;
begin
  FTokenID := tkSpace;
  inc(Run);
  if FLine[Run] = #10 then inc(Run);
end;

procedure TSynWikiTableSyn.ExclamProc;
begin
  FTokenID := tkSymbol;
  inc(Run);
  if FLine[Run] = '-' then inc(Run);
end;

function TSynWikiTableSyn.GetDefaultAttribute(Index: Integer): TSynHighlighterAttributes;
begin
  case Index of
    SYN_ATTR_COMMENT    : Result := FCommentAttri;
    SYN_ATTR_SYMBOL     : Result := FSymbolAttri;
    SYN_ATTR_NUMBER     : Result := FNumberAttri;
    SYN_ATTR_WHITESPACE : Result := FSpaceAttri;
    SYN_ATTR_IDENTIFIER : Result := FIdentifierAttri;
    SYN_ATTR_STRING     : Result := FStringAttri;
    (*
    SYN_ATTR_KEYWORD    : Result := FKeyAttri;
    *)
  else
    Result := nil;
  end;
end;

function TSynWikiTableSyn.GetEol: Boolean;
begin
  Result := (FTokenID = tkNull);
end;

function TSynWikiTableSyn.GetToken: string;
var
  Len: LongInt;
begin
  Result := '';
  Len := Run - FTokenPos;
  SetString(Result, (FLine + FTokenPos), Len);
end;

function TSynWikiTableSyn.GetTokenAttribute: TSynHighlighterAttributes;
begin
  case GetTokenID of
    tkComment    : Result := FCommentAttri;
    tkSymbol     : Result := FSymbolAttri;
    tkNumber     : Result := FNumberAttri;
    tkSpace      : Result := FSpaceAttri;
    tkIdentifier : Result := FIdentifierAttri;
    tkString     : Result := FStringAttri;
    {
    tkKey        : Result := FKeyAttri;
    tkNumber     : Result := FNumberAttri;
    tkUnknown    : Result := FIdentifierAttri;
    }
  else
    Result := nil;
  end;
end;

procedure TSynWikiTableSyn.GetTokenEx(out TokenStart: PChar;
  out TokenLength: integer);
begin
  TokenLength := Run - FTokenPos;
  TokenStart := FLine + FTokenPos;
end;

function TSynWikiTableSyn.GetTokenID: TtkTokenKind;
begin
  Result := FTokenId;
end;

function TSynWikiTableSyn.GetTokenKind: integer;
begin
  Result := Ord(FTokenId);
end;

function TSynWikiTableSyn.GetTokenPos: Integer;
begin
  Result := FTokenPos;
end;

function TSynWikiTableSyn.IdentKind(MayBe: PChar): TtkTokenKind;
var
  HashKey: Integer;
begin
  FToIdent := MayBe;
  HashKey := KeyHash(MayBe);
  if (HashKey >= 16) and (HashKey <= 275) then
    Result := FIdentFuncTable[HashKey]()
  else
    Result := tkIdentifier;
end;

procedure TSynWikiTableSyn.IdentProc;
begin
  FTokenID := IdentKind((FLine + Run));
  inc(Run, FStringLen);
  while Identifiers[FLine[Run]] do
    Inc(Run);
end;

procedure TSynWikiTableSyn.InitIdent;
var
  i: Integer;
begin                             (*
  for i := 0 to 255 do
    case i of
      1:   FIdentFuncTable[i] := @Func1;
      2:   FIdentFuncTable[i] := @Func2;
      8:   FIdentFuncTable[i] := @Func8;
      9:   FIdentFuncTable[i] := @Func9;
      10:  FIdentFuncTable[i] := @Func10;
      11:  FIdentFuncTable[i] := @Func11;
      12:  FIdentFuncTable[i] := @Func12;
      13:  FIdentFuncTable[i] := @Func13;
      14:  FIdentFuncTable[i] := @Func14;
      15:  FIdentFuncTable[i] := @Func15;
      16:  FIdentFuncTable[i] := @Func16;
      17:  FIdentFuncTable[i] := @Func17;
      18:  FIdentFuncTable[i] := @Func18;
      19:  FIdentFuncTable[i] := @Func19;
      20:  FIdentFuncTable[i] := @Func20;
      21:  FIdentFuncTable[i] := @Func21;
      23:  FIdentFuncTable[i] := @Func23;
      24:  FIdentFuncTable[i] := @Func24;
      25:  FIdentFuncTable[i] := @Func25;
  end;
  *)
end;

function TSynWikiTableSyn.KeyComp(const aKey: string): Boolean;
var
  i: Integer;
  Temp: PChar;
begin
  Temp := FToIdent;
  if Length(aKey) = FStringLen then begin
    Result := True;
    for i := 1 to FStringLen do begin
      if mHashTable[Temp^] <> mHashTable[aKey[i]] then begin
        Result := False;
        Break;
      end;
      inc(Temp);
    end;
  end else
    Result := False;
end;

function TSynWikiTableSyn.KeyHash(ToHash: PChar): Integer;
begin
  Result := 0;
  While (ToHash^ In ['a'..'z', 'A'..'Z', '!', '/']) do begin
    Inc(Result, mHashTable[ToHash^]);
    Inc(ToHash);
  end;
  While (ToHash^ In ['0'..'9']) do begin
    Inc(Result, (Ord(ToHash^) - Ord('0')) );
    Inc(ToHash);
  end;
  FStringLen := (ToHash - FToIdent);
end;

procedure TSynWikiTableSyn.LFProc;
begin
  FTokenID := tkSpace;
  inc(Run);
end;

procedure TSynWikiTableSyn.MakeMethodTables;
var
  ch: Char;
begin
  for ch := #0 to #255 do
    case ch of
      #0                          : FProcTable[ch] := @NullProc;
      #10                         : FProcTable[ch] := @LFProc;
      #13                         : FProcTable[ch] := @CRProc;
      #1..#9, #11, #12, #14..#32  : FProcTable[ch] := @SpaceProc;
      '"'                         : FProcTable[ch] := @StringProc;
      '0'..'9'                    : FProcTable[ch] := @NumberProc;
      'A'..'Z', 'a'..'z', '_','@' : FProcTable[ch] := @IdentProc;
      '&'                         : FProcTable[ch] := @AmpersandProc;
      '<'                         : FProcTable[ch] := @OpenBraceProc;
      '{'                         : FProcTable[ch] := @BeginProc;
      '|'                         : FProcTable[ch] := @BarProc;
      '!'                         : FProcTable[ch] := @ExclamProc;

   //   '{', '}'                    : FProcTable[ch] := @AsciiCharProc;
//      '-'                         : FProcTable[ch] := @DashProc;
//      '#', '$'                    : FProcTable[ch] := @IntegerProc;
//      ')', '('                    : FProcTable[ch] := @RoundOpenProc;
//      '/'                         : FProcTable[ch] := @SlashProc;
    else
      FProcTable[ch] := @UnknownProc;
    end;
end;
(*
var
  i: Char;
begin
  For i := #0 To #255 do begin
    case i of
    #0                         : FProcTable[i] := @NullProc;
    #10                        : FProcTable[i] := @LFProc;
    #13                        : FProcTable[i] := @CRProc;
    #1..#9, #11, #12, #14..#32 : FProcTable[i] := @SpaceProc;
    '"'                        : FProcTable[i] := @StringProc;
//    '<'                        : FProcTable[i] := @BraceOpenProc;
//    '>'                        : FProcTable[i] := @BraceCloseProc;
{    '&':
      begin
        fProcTable[i] := @AmpersandProc;
      end;
    '=':
      begin
        fProcTable[i] := @EqualProc;
      end;
      }
    else
      fProcTable[i] := @IdentProc;
    end;
  end;
end;
  *)
procedure TSynWikiTableSyn.Next;
begin
  FTokenPos := Run;
  case FRange of
    rsText    : TextProc;
    rsComment : OpenBraceProc;
    else        FProcTable[FLine[Run]];
  end;


  {
  if FRange = rsCStyle then
    CStyleCommentProc
  else
    FProcTable[FLine[Run]]();
    }
end;

procedure TSynWikiTableSyn.NullProc;
begin
  fTokenID := tkNull;
end;

procedure TSynWikiTableSyn.NumberProc;
begin
  inc(Run);
  FTokenID := tkNumber;
  while FLine[Run] in ['0'..'9', '.', 'e', 'E'] do begin
    if ((FLine[Run] = '.') and (FLine[Run + 1] = '.')) or
       ((FLine[Run] = 'e') and ((FLine[Run + 1] = 'x') or (FLine[Run + 1] = 'm'))) then
      Break;
    Inc(Run);
  end;
end;

procedure TSynWikitableSyn.OpenBraceProc;
begin
  if (FLine[Run+1] = '!') and (FLine[Run+2] = '-') and (FLine[Run+3] = '-') then
  begin
    FTokenID := tkComment;
    while not (FLine[Run] in [#0, #10, #13]) do begin
      if (FLine[Run] = '>') and (FLine[Run - 1] = '-') and (FLine[Run - 2] = '-')
      then begin
        FRange := rsText;
        inc(Run);
        {
        if TopHtmlCodeFoldBlockType = cfbtHtmlComment then
          EndHtmlNodeCodeFoldBlock;
          }
        break;
      end;
      inc(Run);
    end;
  end else begin
    FTokenID := tkSymbol;
    while not (FLine[Run] in [#0, #10, #13]) do begin
      if FLine[Run] = '>' then begin
        FRange := rsText;
        inc(Run);
        break;
      end;
      inc(Run);
    end;
  end;

(*
  if (FLine[Run] in [#0, #10, #13]) then begin
    FProcTable[FLine[Run]];
    Exit;
  end;

  while not (FLine[Run] in [#0, #10, #13]) do begin
    if (FLine[Run] = '>') and (FLine[Run - 1] = '-') and (FLine[Run - 2] = '-')
    then begin
      FRange := rsText;
      inc(Run);
      {
      if TopHtmlCodeFoldBlockType = cfbtHtmlComment then
        EndHtmlNodeCodeFoldBlock;
        }
      break;
    end;
    inc(Run);
  end;
  *)
end;

procedure TSynWikiTableSyn.SetLine(const NewValue: String; LineNumber: Integer);
begin
  inherited;
  FLine := PChar(NewValue);
  Run := 0;
  FLineNumber := LineNumber;
  Next;
end;

procedure TSynWikiTableSyn.SpaceProc;
begin
  inc(Run);
  FTokenID := tkSpace;
  while FLine[Run] <= #32 do begin
    if FLine[Run] in [#0, #9, #10, #13] then break;
    inc(Run);
  end;
end;

procedure TSynWikiTableSyn.StringProc;
begin
  (*
  if (FRange = rsValue) then begin
    FRange := rsParam;
    FTokenID := tkValue;
  end else begin
    fTokenID := tkString;
  end;
  *)
  FTokenID := tkString;
  inc(Run);  // first '"'
  while not (FLine[Run] in [#0, #10, #13, '"']) do inc(Run);
  if FLine[Run] = '"' then inc(Run);  // last '"'
end;

procedure TSynWikiTableSyn.TextProc;
const
  StopSet = [#0..#31, '<', '&', '{', '|'];
var
  i: Integer;
begin
  if FLine[Run] in (StopSet - ['&']) then begin
    FProcTable[fLine[Run]];
    exit;
  end;

  FTokenID := tkText;
  while True do begin
    while not (FLine[Run] in StopSet) do inc(Run);

    if (FLine[Run] = '&') then begin
      for i:=Low(EscapeAmps) To High(EscapeAmps) do begin
        if (StrLIComp((fLine + Run), PChar(EscapeAmps[i]), StrLen(EscapeAmps[i])) = 0) then begin
          fAmpersandCode := i;
          fRange := rsAmpersand;
          Exit;
        end;
      end;

      Inc(Run);
    end else begin
      Break;
    end;
  end;
end;

procedure TSynWikiTableSyn.UnknownProc;
begin
  inc(Run);
  while (FLine[Run] in [#128..#191]) or // continued utf8 subcode
    ((FLine[Run] <> #0) and (FProcTable[fLine[Run]] = @UnknownProc)) do inc(Run);
  FTokenID := tkUnknown;
end;


end.

