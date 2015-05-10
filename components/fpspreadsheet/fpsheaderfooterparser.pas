unit fpsHeaderFooterParser;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpsTypes;

type
  TsHeaderFooterToken = (hftText, hftNewLine,
    hftSheetName, hftPath, hftFileName, hftDate, hftTime, hftPage, hftPageCount);

  TsHeaderFooterFontStyle = (hfsBold, hfsItalic, hfsUnderline, hfsDblUnderline,
    hfsStrikeout, hfsShadow, hfsOutline, hfsSubscript, hfsSuperScript);

  TsHeaderFooterFontStyles = set of TsHeaderFooterFontStyle;

  TsHeaderFooterFont = class(TObject)
    FontName: String;
    Size: Double;
    Style: TsHeaderFooterFontStyles;
    Color: TsColorValue;
    constructor Create; overload;
    constructor Create(AFont: TsFont); overload;
    constructor Create(AFontName: String; ASize: Double;
      AStyle: TsHeaderFooterFontStyles; AColor: TsColorValue); overload;
    procedure Assign(AFont: TObject);
  end;

  TsHeaderFooterFontClass = class of TsHeaderFooterFont;

  TsHeaderFooterElement = record
    Token: TsHeaderFooterToken;
    TextValue: String;
    FontIndex: Integer;
  end;

  TsHeaderFooterSectionIndex = (hfsLeft, hfsCenter, hfsRight);

  TsHeaderFooterSection = array of TsHeaderFooterElement;

  TsHeaderFooterSections = array[TsHeaderFooterSectionIndex] of TsHeaderFooterSection;

  TsHeaderFooterParser = class(TObject)
  private
    FParseText: String;
    FToken: Char;
    FCurrent: PChar;
    FStart: PChar;
    FEnd: PChar;
    FCurrFont: TsHeaderFooterFont;
    function NextToken: Char;
    function PrevToken: Char;
    procedure ScanFont;
    procedure ScanFontColor;
    procedure ScanFontSize;
    procedure ScanNewLine;
    procedure ScanSymbol;
  protected
    FSections: TsHeaderFooterSections;
    FDefaultFont: TsHeaderFooterFont;
    FCurrSection: TsHeaderFooterSectionIndex;
    FStatus: Integer;
    FFontList: TList;
    FPointSeparatorSettings: TFormatSettings;
    FCurrFontIndex: Integer;
    FCurrText: String;
    FFontClass: TsHeaderFooterFontClass;
    procedure AddCurrTextElement;
    procedure AddElement(AToken: TsHeaderFooterToken);
    procedure AddFontStyle(AStyle: TsHeaderFooterFontStyle);
    function FindCurrFont: Integer;
    function GetCurrFontIndex: Integer; virtual;
    procedure Parse; virtual;
    procedure UseSection(AIndex: TsHeaderFooterSectionIndex); virtual;
  public
    constructor Create; overload;
    constructor Create(AText: String; AFontList: TList;
      ADefaultFont: TsHeaderFooterFont); overload;
    destructor Destroy; override;
    function BuildHeaderFooter: String;
  end;

const
  hfpsOK = 0;

implementation

uses
  Math,
  fpsUtils;

const
  FONTSTYLE_SYMBOLS: array[TsHeaderFooterFontStyle] of char =
    ('B', 'I', 'U', 'E', 'S', 'H', 'O', 'X', 'Y');

constructor TsHeaderFooterFont.Create;
begin
  inherited;
end;

constructor TsHeaderFooterFont.Create(AFontName: String; ASize: Double;
  AStyle: TsHeaderFooterFontStyles; AColor: TsColorValue);
begin
  FontName := AFontName;
  Size := ASize;
  Style := AStyle;
  Color := AColor;
end;

constructor TsHeaderFooterFont.Create(AFont: TsFont);
begin
  Create;
  Assign(AFont);
end;

procedure TsHeaderFooterFont.Assign(AFont: TObject);
begin
  if AFont is TsFont then
  begin
    FontName := TsFont(AFont).FontName;
    Size := TsFont(AFont).Size;
    Style := [];
    if fssBold in TsFont(AFont).Style then Include(Style, hfsBold);
    if fssItalic in TsFont(AFont).Style then Include(Style, hfsItalic);
    if fssUnderline in TsFont(AFont).Style then Include(Style, hfsUnderline);
    if fssStrikeout in TsFont(AFont).Style then Include(Style, hfsStrikeout);
    Color := 0; // black --- to be replaced by TsFont.Color once it is no longer paletted
  end else
  if AFont is TsHeaderFooterFont then
  begin
    FontName := TsHeaderFooterFont(AFont).FontName;
    Size := TsHeaderFooterFont(AFont).Size;
    Style := TsHeaderFooterFont(AFont).Style;
    Color := TsHeaderFooterFont(AFont).Color;
  end else
    raise Exception.Create('[TsHeaderFooterFont.Assign] Argument can only be a TsFont or a TsHeaderFooterFont');
end;


{ TsHeaderFooterParser }

constructor TsHeaderFooterParser.Create;
begin
  FFontClass := TsHeaderFooterFont;

  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  FCurrSection := hfsCenter;
  FCurrText := '';
end;

constructor TsHeaderFooterParser.Create(AText: String; AFontList: TList;
  ADefaultFont: TsHeaderFooterFont);
begin
  if AFontList = nil then
    raise Exception.Create('[TsHeaderFooterParser.Create] FontList must not be nil.');
  if ADefaultFont = nil then
    raise Exception.Create('[TsHeaderFooterParser.Create] DefaultFont must not be nil.');

  Create;

  FFontList := AFontList;
  FDefaultFont := ADefaultFont;
  FCurrFont := TsHeaderFooterFont.Create;
  FCurrFont.Assign(ADefaultFont);
  FParseText := AText;

  Parse;
end;

destructor TsHeaderFooterParser.Destroy;
begin
  FCurrFont.Free;
  inherited Destroy;
end;

procedure TsHeaderFooterParser.AddCurrTextElement;
begin
  AddElement(hftText);
end;

procedure TsHeaderFooterParser.AddElement(AToken: TsHeaderFooterToken);
var
  n: Integer;
begin
  n := Length(FSections[FCurrSection]);
  SetLength(FSections[FCurrSection], n+1);
  with FSections[FCurrSection][n] do
  begin
    Token := AToken;
    if Token = hftText then
    begin
      TextValue := FCurrText;
      FCurrText := '';
    end else
      TextValue := '';
    FontIndex := GetCurrFontIndex;
  end;
end;

procedure TsHeaderFooterParser.AddFontStyle(AStyle: TsHeaderFooterFontStyle);
begin
  if FCurrText <> '' then
    AddCurrTextElement;

  if AStyle in FCurrFont.Style then
    Exclude(FCurrFont.Style, AStyle)
  else
    Include(FCurrFont.Style, AStyle);
end;

function TsHeaderFooterParser.BuildHeaderFooter: String;
var
  sec: TsHeaderFooterSectionIndex;
  element: TsHeaderFooterElement;
  fnt, prevfnt: TsHeaderFooterFont;
  fs: TsHeaderFooterFontStyle;
  i: Integer;
begin
  Result := '';
  for sec := hfsLeft to hfsRight do
  begin
    prevfnt := FDefaultFont;
    if Length(FSections[sec]) > 0 then
      case sec of
        hfsLeft      : Result := Result + '&L';
        hfsCenter    : Result := Result + '&C';
        hfsRight     : Result := Result + '&R';
      end;
    for element in FSections[sec] do
    begin
      if (element.FontIndex > -1) and (element.FontIndex < FFontList.Count) then
      begin
        fnt := TsHeaderFooterFont(FFontList[element.FontIndex]);
        if fnt.FontName = '' then fnt.FontName := FDefaultFont.FontName;
        if not SameText(fnt.FontName, prevFnt.FontName) then
          Result := Result + '&"' + fnt.FontName + '"';
        if not SameValue(fnt.Size, prevfnt.Size, 1e-2) then
          Result := Result + '&' + Format('%d', [round(fnt.Size)]);  // Excel wants only integers!
        for fs in TsHeaderFooterFontStyle do
          if ((fs in fnt.Style) and not (fs in prevfnt.Style)) or
             (not (fs in fnt.Style) and (fs in prevfnt.Style))
          then
            Result := Result + '&' + FONTSTYLE_SYMBOLS[fs];
        if fnt.Color <> prevfnt.Color then
          Result := Result + '&K' + ColorToHTMLColorStr(fnt.Color, true);
        prevfnt := fnt;
      end;
      case element.Token of
        hftText      : for i:=1 to length(element.TextValue) do
                         if element.TextValue[i]='&'
                           then Result := Result + '&&'
                           else Result := Result + element.TextValue[i];
        hftSheetName : Result := Result + '&A';
        hftPath      : Result := Result + '&Z';
        hftFileName  : Result := Result + '&F';
        hftDate      : Result := Result + '&D';
        hftTime      : Result := Result + '&T';
        hftPage      : Result := Result + '&P';
        hftPageCount : Result := Result + '&N';
        hftNewLine   : Result := Result + LineEnding;
      end;
    end; // for element
  end;  // for sesc
end;

function TsHeaderFooterParser.FindCurrFont: Integer;
var
  fnt: TsHeaderFooterFont;
begin
  for Result := 0 to FFontList.Count-1 do
  begin
    fnt := TsHeaderFooterFont(FFontList[Result]);
    if SameText(fnt.FontName, FCurrFont.FontName) and
       SameValue(fnt.Size, FCurrFont.Size) and
       (fnt.Style = FCurrFont.Style) and
       (fnt.Color = FCurrFont.Color)
    then
      exit;
  end;
  Result := -1;
end;

function TsHeaderFooterParser.GetCurrFontIndex: Integer;
var
  fnt: TsHeaderFooterFont;
begin
  Result := FindCurrFont;
  if Result = -1 then
  begin
    fnt := FFontClass.Create;
    fnt.Assign(FCurrFont);
    Result := FFontList.Add(fnt);
  end;
end;

function TsHeaderFooterParser.NextToken: Char;
begin
  if FCurrent < FEnd then begin
    inc(FCurrent);
    Result := FCurrent^;
  end else
    Result := #0;
end;

function TsHeaderFooterParser.PrevToken: Char;
begin
  if FCurrent > nil then begin
    dec(FCurrent);
    Result := FCurrent^;
  end else
    Result := #0;
end;

procedure TsHeaderFooterParser.Parse;
begin
  if FParseText = '' then
    exit;

  FStart := @FParseText[1];
  FEnd := FStart + Length(FParseText);
  FCurrent := FStart;
  FToken := FCurrent^;
  FCurrSection := hfsCenter;

  while (FCurrent < FEnd) and (FStatus = hfpsOK) do begin
    case FToken of
      '&': ScanSymbol;
      #13, #10: ScanNewLine;
      else FCurrText := FCurrText + FToken;
    end;
    FToken := NextToken;
  end;
  if Length(FCurrText) > 0 then
    AddCurrTextElement;
end;

procedure TsHeaderFooterParser.ScanFont;
var
  s: String;
begin
  s := '';
  FToken := NextToken;
  while (FCurrent < FEnd) and (FStatus = hfpsOK) and not (FToken in ['"', ',']) do
  begin
    // Excel allows to add a font-style identifier to the font name, separated
    // by a comma. We do not support this feature because the font style
    // identifier is a localized string! --> Skip text after the comma
    if FToken = ',' then
    begin
      while (FCurrent < FEnd) and (FToken <> '"') do
        FToken := NextToken;
      break;
    end else
    begin
      s := s + FToken;
      FToken := NextToken;
    end;
  end;
  FCurrFont.FontName := s;
end;

procedure TsHeaderFooterParser.ScanFontColor;
var
  s: String;
begin
  s := '#';
  FToken := NextToken;
  while (FCurrent < FEnd) and (FStatus = hfpsOK) and (FToken in ['0'..'9', 'A'..'F']) do
  begin
    s := s + FToken;
    FToken := NextToken;
  end;
  FToken := PrevToken;
  FCurrFont.Color := HTMLColorStrToColor(s);
end;

procedure TsHeaderFooterParser.ScanFontSize;
var
  s: String;
begin
  s := '';
  while (FCurrent < FEnd) and (FStatus = hfpsOK) and (FToken in ['0'..'9','.']) do
  begin
    s := s + FToken;
    FToken := NextToken;
  end;
  FToken := PrevToken;
  FCurrFont.Size := StrToFloat(s, FPointSeparatorSettings);
end;

procedure TsHeaderFooterParser.ScanNewLine;
begin
  if FCurrText <> '' then
    AddCurrTextElement;

  case FToken of
    #13: begin
           AddElement(hftNewLine);
           FToken := NextToken;
           if FToken <> #10 then FToken := PrevToken;
         end;
    #10: AddElement(hftNewLine);
  end;
end;

procedure TsHeaderFooterParser.ScanSymbol;
begin
  FToken := NextToken;

  if FToken = '&' then
    FCurrText := FCurrText + '&'
  else
  begin
    if FCurrText <> '' then
      AddCurrTextElement;
    case FToken of
      'L': UseSection(hfsLeft);
      'C': UseSection(hfsCenter);
      'R': UseSection(hfsRight);
      'A': AddElement(hftSheetName);
      'F': AddElement(hftFileName);
      'Z': AddElement(hftPath);
      'D': AddElement(hftDate);
      'T': AddElement(hftTime);
      'P': AddElement(hftPage);
      'N': AddElement(hftPageCount);
      '"': ScanFont;
      '0'..'9', '.': ScanFontSize;
      'K': ScanFontColor;
      'B': AddFontStyle(hfsBold);
      'I': AddFontStyle(hfsItalic);
      'U': AddFontStyle(hfsUnderline);
      'E': AddFontStyle(hfsDblUnderline);
      'S': AddFontStyle(hfsStrikeout);
      'H': AddFontStyle(hfsShadow);
      'O': AddFontStyle(hfsOutline);
      'X': AddFontStyle(hfsSuperscript);
      'Y': AddFontStyle(hfsSubscript);
    end;
  end;
end;

procedure TsHeaderFooterParser.UseSection(AIndex: TsHeaderFooterSectionIndex);
begin
  if FCurrText <> '' then
    AddCurrTextElement;
  FCurrFont.FontName := FDefaultFont.FontName;
  FCurrFont.Size := FDefaultFont.Size;
  FCurrFont.Style := FDefaultFont.Style;
  FCurrFont.Color := FDefaultFont.Color;
  FCurrSection := AIndex;
end;

end.

