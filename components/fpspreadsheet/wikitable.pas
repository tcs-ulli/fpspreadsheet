(*
wikitable.pas

One unit which handles multiple wiki table formats

Format simplepipes:

|| || title1 || title2 || title3
| [link_to_something|http://google.com]| {color:red}FAILED{color}| {color:red}FAILED{color}| {color:green}PASS{color}

Format mediawiki:

{| border="1" cellpadding="2" class="wikitable sortable"
|-
|
! Title
|-
| [http://google.com link_to_something]
! style="background-color:green;color:white;" | PASS
|}

AUTHORS: Felipe Monteiro de Carvalho
*)
unit wikitable;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  fpimage, fgl,
  fpspreadsheet, fpsutils, lconvencoding;

type

  TWikiTableToken = class
  public
    BackgroundColor: TsColor;
    UseBackgroundColor: Boolean;
    Bold: Boolean;
    Value: string;
  end;

  TWikiTableTokenList = specialize TFPGList<TWikiTableToken>;

  { TWikiTableTokenizer }

  TWikiTableTokenizer = class
  public
    Tokens: TWikiTableTokenList;
    constructor Create; virtual;
    destructor Destroy; override;
    procedure Clear;
    function AddToken(AValue: string): TWikiTableToken;
    procedure TokenizeString_Pipes(AStr: string);
  end;

  { TsWikiTableNumFormatList }
  TsWikiTableNumFormatList = class(TsCustomNumFormatList)
  protected
//    procedure AddBuiltinFormats; override;
  public
//    function FormatStringForWriting(AIndex: Integer): String; override;
  end;


  { TsWikiTableReader }

  TsWikiTableReader = class(TsCustomSpreadReader)
  protected
    procedure CreateNumFormatList; override;
  public
    SubFormat: TsSpreadsheetFormat;
    { General reading methods }
    procedure ReadFromStrings(AStrings: TStrings; AData: TsWorkbook); override;
    procedure ReadFromStrings_Pipes(AStrings: TStrings; AData: TsWorkbook);
  end;

  { TsWikiTable_PipesReader }

  TsWikiTable_PipesReader = class(TsWikiTableReader)
  public
    constructor Create(AWorkbook: TsWorkbook); override;
  end;

  { TsWikiTableWriter }

  TsWikiTableWriter = class(TsCustomSpreadWriter)
  protected
    // Helpers
    procedure CreateNumFormatList; override;
  public
    SubFormat: TsSpreadsheetFormat;
    { General writing methods }
    procedure WriteToStrings(AStrings: TStrings); override;
    procedure WriteToStrings_WikiMedia(AStrings: TStrings);
  end;

  { TsWikiTable_WikiMediaWriter }

  TsWikiTable_WikiMediaWriter = class(TsWikiTableWriter)
  public
    constructor Create(AWorkbook: TsWorkbook); override;
  end;

implementation

{ TsWikiTableNumFormatList }


{ TWikiTableTokenizer }

constructor TWikiTableTokenizer.Create;
begin
  inherited Create;
  Tokens := TWikiTableTokenList.Create;
end;

destructor TWikiTableTokenizer.Destroy;
begin
  Clear;
  Tokens.Free;
  inherited Destroy;
end;

procedure TWikiTableTokenizer.Clear;
var
  i: Integer;
begin
  for i := 0 to Tokens.Count-1 do
    Tokens.Items[i].Free;
  Tokens.Clear;
end;

function TWikiTableTokenizer.AddToken(AValue: string): TWikiTableToken;
begin
  Result := TWikiTableToken.Create;
  Result.Value := AValue;
  Tokens.Add(Result);
end;

(*
Format simplepipes:

|| || title1 || title2 || title3
| [link_to_something|http://google.com]| {color:red}FAILED{color}| {color:red}FAILED{color}| {color:green}PASS{color}
*)
procedure TWikiTableTokenizer.TokenizeString_Pipes(AStr: string);
const
  Str_Pipe: Char = '|';
  Str_LinkStart: Char = '[';
  Str_LinkEnd: Char = ']';
  Str_FormatStart: Char = '{';
  Str_FormatEnd: Char = '}';
  Str_EmptySpaces: set of Char = [' '];
var
  i: Integer;
  lTmpStr: string = '';
  lFormatStr: string = '';
  lState: Integer;
  lLookAheadChar, lCurChar: Char;
  lIsTitle: Boolean = False;
  lCurBackgroundColor: TsColor;
  lUseBackgroundColor: Boolean = False;
  lCurToken: TWikiTableToken;

  procedure DoAddToken();
  begin
    lCurToken := AddToken(lTmpStr);
    lCurToken.Bold := lIsTitle;
    lCurToken.UseBackgroundColor := lUseBackgroundColor;
    if lUseBackgroundColor then
      lCurToken.BackgroundColor := lCurBackgroundColor;
  end;

begin
  Clear;

  lState := 0;

  i := 1;
  while i <= Length(AStr) do
  begin
    lCurChar := AStr[i];
    if i < Length(AStr) then lLookAheadChar := AStr[i+1];

    case lState of
    0: // Line-start or otherwise reading a pipe separator, expecting a | or ||
    begin
      if lCurChar = Str_Pipe then
      begin
        lState := 1;
        lIsTitle := False;
        if lLookAheadChar = Str_Pipe then
        begin
          Inc(i);
          lIsTitle := True;
        end;
        Inc(i);

        lUseBackgroundColor := False;
        lTmpStr := '';
      end
      else if lCurChar in Str_EmptySpaces then
      begin
        // Do nothing
        Inc(i);
      end
      else
      begin
        // Error!!!
        raise Exception.Create('[TWikiTableTokenizer.TokenizeString] Wrong char!');
      end;
    end;
    1: // Reading cell text
    begin
      if lCurChar = Str_Pipe then
      begin
        lState := 0;
        DoAddToken();
      end
      else if lCurChar = Str_LinkStart then
      begin
        lState := 2;
        Inc(i);
      end
      else if lCurChar = Str_FormatStart then
      begin
        lState := 4;
        Inc(i);
      end
      else
      begin
        lTmpStr := lTmpStr + lCurChar;
        Inc(i);
      end;
    end;
    2: // Link text reading
    begin
      if lCurChar = Str_Pipe then
      begin
        lState := 3;
        Inc(i);
      end
      else
      begin
        lTmpStr := lTmpStr + lCurChar;
        Inc(i);
      end;
    end;
    3: // Link target reading
    begin
      if lCurChar = Str_LinkEnd then
      begin
        lState := 1;
        Inc(i);
      end
      else
      begin
        Inc(i);
      end;
    end;
    4: // Color start reading
    begin
      if lCurChar = Str_FormatEnd then
      begin
        lState := 1;
        Inc(i);
        lFormatStr := LowerCase(Trim(lFormatStr));
        if lFormatStr = 'color:red' then lCurBackgroundColor := scRED
        else if lFormatStr = 'color:green' then lCurBackgroundColor := scGREEN
        else if lFormatStr = 'color:yellow' then lCurBackgroundColor := scYELLOW
        //
        else if lFormatStr = 'color:orange' then lCurBackgroundColor := scOrange
        else lCurBackgroundColor := scWHITE;
        lUseBackgroundColor := True;
        lFormatStr := '';
      end
      else
      begin
        lFormatStr := lFormatStr + lCurChar;
        Inc(i);
      end;
    end;
    end;
  end;

  // rest after the last || is also a token
  if lTmpStr <> '' then DoAddToken();

  // If there is a token still to be added, add it now
  if (lState = 0) and (lTmpStr <> '') then AddToken(lTmpStr);
end;

{ TsWikiTableReader }

procedure TsWikiTableReader.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsWikiTableNumFormatList.Create(Workbook);
end;

procedure TsWikiTableReader.ReadFromStrings(AStrings: TStrings;
  AData: TsWorkbook);
begin
  case SubFormat of
  sfWikiTable_Pipes: ReadFromStrings_Pipes(AStrings, AData);
  end;
end;

procedure TsWikiTableReader.ReadFromStrings_Pipes(AStrings: TStrings;
  AData: TsWorkbook);
var
  i, j: Integer;
  lCurLine: String;
  lLineSplitter: TWikiTableTokenizer;
  lCurToken: TWikiTableToken;
begin
  FWorksheet := AData.AddWorksheet('Table');
  lLineSplitter := TWikiTableTokenizer.Create;
  try
    for i := 0 to AStrings.Count-1 do
    begin
      lCurLine := AStrings[i];
      lLineSplitter.TokenizeString_Pipes(lCurLine);
      for j := 0 to lLineSplitter.Tokens.Count-1 do
      begin
        lCurToken := lLineSplitter.Tokens[j];
        FWorksheet.WriteUTF8Text(i, j, lCurToken.Value);
        if lCurToken.Bold then FWorksheet.WriteUsedFormatting(i, j, [uffBold]);
        if lCurToken.UseBackgroundColor then FWorksheet.WriteBackgroundColor(i, j, lCurToken.BackgroundColor);
      end;
    end;
  finally
    lLineSplitter.Free;
  end;
end;

{ TsWikiTable_PipesReader }

constructor TsWikiTable_PipesReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  SubFormat := sfWikiTable_Pipes;
end;

{ TsWikiTableWriter }

procedure TsWikiTableWriter.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsWikiTableNumFormatList.Create(Workbook);
end;

procedure TsWikiTableWriter.WriteToStrings(AStrings: TStrings);
begin
  case SubFormat of
    sfWikiTable_WikiMedia: WriteToStrings_WikiMedia(AStrings);
  end;
end;

(*
Format mediawiki:

{| border="1" cellpadding="2" class="wikitable sortable"
|-
|
! Title
|-
| [http://google.com link_to_something]
! style="background-color:green;color:white;" | PASS
|}
*)
procedure TsWikiTableWriter.WriteToStrings_WikiMedia(AStrings: TStrings);
var
  i, j: Integer;
  lCurStr: string = '';
  lCurUsedFormatting: TsUsedFormattingFields;
  lCurColor: TsColor;
  lColorStr: String;
begin
  AStrings.Add('{| border="1" cellpadding="2" class="wikitable sortable"');
  FWorksheet := Workbook.GetFirstWorksheet();
  for i := 0 to FWorksheet.GetLastRowIndex() do
  begin
    AStrings.Add('|-');
    for j := 0 to FWorksheet.GetLastColIndex() do
    begin
      lCurStr := FWorksheet.ReadAsUTF8Text(i, j);
      lCurUsedFormatting := FWorksheet.ReadUsedFormatting(i, j);

      if uffBackgroundColor in lCurUsedFormatting then
      begin
        lCurColor := FWorksheet.ReadBackgroundColor(i, j);
        case lCurColor of
        scBlack:    lColorStr := 'style="background-color:black;color:white;"';
        scWhite:    lColorStr := 'style="background-color:white;color:black;"';
        scRed:      lColorStr := 'style="background-color:red;color:white;"';
        scGREEN:    lColorStr := 'style="background-color:green;color:white;"';
        scBLUE:     lColorStr := 'style="background-color:blue;color:white;"';
        scYELLOW:   lColorStr := 'style="background-color:yellow;color:black;"';
        {scMAGENTA,  // FF00FFH
        scCYAN,     // 00FFFFH
        scDarkRed,  // 800000H
        scDarkGreen,// 008000H
        scDarkBlue, // 000080H
        scOLIVE,    // 808000H
        scPURPLE,   // 800080H
        scTEAL,     // 008080H
        scSilver,   // C0C0C0H
        scGrey,     // 808080H
        //
        scGrey10pct,// E6E6E6H
        scGrey20pct // CCCCCCH  }
        scOrange:   lColorStr := 'style="background-color:orange;color:white;"';
        end;
        lCurStr := lColorStr + ' |' + lCurStr;
      end;

      if uffBold in lCurUsedFormatting then lCurStr := '!' + lCurStr
      else lCurStr := '|' + lCurStr;

      AStrings.Add(lCurStr);
    end;
  end;
  AStrings.Add('|}');
end;

{ TsWikiTable_WikiMediaWriter }

constructor TsWikiTable_WikiMediaWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  SubFormat := sfWikiTable_WikiMedia;
end;

initialization

  RegisterSpreadFormat(TsWikiTable_PipesReader, nil, sfWikiTable_Pipes);
  RegisterSpreadFormat(nil, TsWikiTable_WikiMediaWriter, sfWikiTable_WikiMedia);

end.
