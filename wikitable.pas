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
  fpimage, fgl, lconvencoding,
  fpsTypes, fpspreadsheet, fpsutils;

type

  TWikiTokenState = (wtsLineStart, wtsCellText, wtsLinkText, wtsLinkTarget, wtsColor);

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
  private
    FWorkbook: TsWorkbook;
  public
    Tokens: TWikiTableTokenList;
    constructor Create(AWorkbook: TsWorkbook); virtual;
    destructor Destroy; override;
    procedure Clear;
    function AddToken(AValue: string): TWikiTableToken;
    procedure TokenizeString_Pipes(AStr: string);
  end;

  { TsWikiTableReader }

  TsWikiTableReader = class(TsCustomSpreadReader)
  protected
    procedure ReadFromStrings_Pipes(AStrings: TStrings);
  public
    SubFormat: TsSpreadsheetFormat;
    { General reading methods }
    procedure ReadFromStrings(AStrings: TStrings); override;
  end;

  { TsWikiTable_PipesReader }

  TsWikiTable_PipesReader = class(TsWikiTableReader)
  public
    constructor Create(AWorkbook: TsWorkbook); override;
  end;

  { TsWikiTableWriter }

  TsWikiTableWriter = class(TsCustomSpreadWriter)
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

uses
  fpsStrings;


{ TWikiTableTokenizer }

constructor TWikiTableTokenizer.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
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
  lColorStr: String = '';
  lState: TWikiTokenState;
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

  lState := wtsLineStart;

  i := 1;
  while i <= Length(AStr) do
  begin
    lCurChar := AStr[i];
    if i < Length(AStr) then lLookAheadChar := AStr[i+1];

    case lState of
      wtsLineStart: // Line-start or otherwise reading a pipe separator, expecting a | or ||
        begin
          if lCurChar = Str_Pipe then
          begin
            lState := wtsCellText;
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

      wtsCellText: // Reading cell text
        begin
          if lCurChar = Str_Pipe then
          begin
            lState := wtsLineStart;
            DoAddToken();
          end
          else if lCurChar = Str_LinkStart then
          begin
            lState := wtsLinkText;
            Inc(i);
          end
          else if lCurChar = Str_FormatStart then
          begin
            lState := wtsColor;
            Inc(i);
          end
          else
          begin
            lTmpStr := lTmpStr + lCurChar;
            Inc(i);
          end;
        end;

      wtsLinkText: // Link text reading
        begin
          if lCurChar = Str_Pipe then
          begin
            lState := wtsLinkTarget;
            Inc(i);
          end
          else
          begin
            lTmpStr := lTmpStr + lCurChar;
            Inc(i);
          end;
        end;

      wtsLinkTarget: // Link target reading
        begin
          if lCurChar = Str_LinkEnd then
          begin
            lState := wtsCellText;
            Inc(i);
          end
          else
          begin
            Inc(i);
          end;
        end;

      wtsColor: // Color start reading
        begin
          if lCurChar = Str_FormatEnd then
          begin
            lState := wtsCellText;
            Inc(i);
            lFormatStr := LowerCase(Trim(lFormatStr));
            if copy(lFormatstr, 1, 6) = 'color:' then
            begin
              lColorstr := Copy(lFormatstr, 7, Length(lFormatStr));
              lCurBackgroundColor := FWorkbook.AddColorToPalette(HTMLColorStrToColor(lColorStr));
              lUseBackgroundColor := True;
              lFormatStr := '';
            end;
          end
          else
          begin
            lFormatStr := lFormatStr + lCurChar;
            Inc(i);
          end;
        end;
    end; // case
  end;  // while

  // rest after the last || is also a token
  if lTmpStr <> '' then DoAddToken();

  // If there is a token still to be added, add it now
  if (lState = wtsLineStart) and (lTmpStr <> '') then AddToken(lTmpStr);
end;


{ TsWikiTableReader }

procedure TsWikiTableReader.ReadFromStrings(AStrings: TStrings);
begin
  case SubFormat of
    sfWikiTable_Pipes: ReadFromStrings_Pipes(AStrings);
  end;
end;

procedure TsWikiTableReader.ReadFromStrings_Pipes(AStrings: TStrings);
var
  i, j: Integer;
  lCurLine: String;
  lLineSplitter: TWikiTableTokenizer;
  lCurToken: TWikiTableToken;
begin
  FWorksheet := FWorkbook.AddWorksheet('Table', true);
  lLineSplitter := TWikiTableTokenizer.Create(FWorkbook);
  try
    for i := 0 to AStrings.Count-1 do
    begin
      lCurLine := AStrings[i];
      lLineSplitter.TokenizeString_Pipes(lCurLine);
      for j := 0 to lLineSplitter.Tokens.Count-1 do
      begin
        lCurToken := lLineSplitter.Tokens[j];
        FWorksheet.WriteUTF8Text(i, j, lCurToken.Value);
        if lCurToken.Bold then
          FWorksheet.WriteFontStyle(i, j, [fssBold]);
        if lCurToken.UseBackgroundColor then
          FWorksheet.WriteBackgroundColor(i, j, lCurToken.BackgroundColor);
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

  function DoBorder(ABorder: TsCellBorder; ACell: PCell): String;
  const
    // (cbNorth, cbWest, cbEast, cbSouth, cbDiagUp, cbDiagDown)
    BORDERNAMES: array[TsCellBorder] of string =
      ('top', 'left', 'right', 'bottom', '', '');
    // (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair)
    LINESTYLES: array[TsLineStyle] of string =
      ('1pt solid', 'medium solid', 'dashed', 'dotted', 'thick solid', 'double', 'dotted');
  var
    ls: TsLineStyle;
    clr: TsColor;
    fmt: PsCellFormat;
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    ls := fmt^.BorderStyles[ABorder].LineStyle;
    clr := fmt^.BorderStyles[ABorder].Color;
    Result := Format('border-%s:%s', [BORDERNAMES[ABorder], LINESTYLES[ls]]);
    if clr <> scBlack then
      Result := Result + ' ' + FWorkbook.GetPaletteColorAsHTMLStr(clr) + '; ';
  end;

const
  PIPE_CHAR: array[boolean] of String = ('|', '!');
var
  i, j: cardinal;
  lCurStr: ansistring = '';
  lCurUsedFormatting: TsUsedFormattingFields;
  lCurColor: TsColor;
  lStyleStr: String;
  lColSpanStr: String;
  lRowSpanStr: String;
  lColWidthStr: String;
  lRowHeightStr: String;
  lCell: PCell;
  lCol: PCol;
  lRow: PRow;
  lFont: TsFont;
  horalign: TsHorAlignment;
  vertalign: TsVertAlignment;
  r1, c1, r2, c2: Cardinal;
  isHeader: Boolean;
  borders: TsCellBorders;
begin
  FWorksheet := Workbook.GetFirstWorksheet();
  FWorksheet.UpdateCaches;

  AStrings.Add('<!-- generated by fpspreadsheet -->');

  // Show/hide grid lines
  if soShowGridLines in FWorksheet.Options then
    lCurStr := '{| class="wikitable"' // sortable"'
  else
    lCurStr := '{| border="0" cellpadding="2"';

  // Default font
  lStyleStr := '';
  lFont := FWorkbook.GetDefaultFont;
  if lFont.FontName <> DEFAULT_FONTNAME then
    lStyleStr := lStyleStr + Format('font-family:%s;', [lFont.FontName]);
  if fssBold in lFont.Style then
    lStyleStr := lStyleStr + 'font-weight:bold;';
  if fssItalic in lFont.Style then
    lStyleStr := lStyleStr + 'font-style:italic;';
  if fssUnderline in lFont.Style then
    lStyleStr := lStyleStr + 'text-decoration:underline;';
  if lFont.Size <> DEFAULT_FONTSIZE then
    lStyleStr := lStyleStr + Format('font-size:%.0fpt;', [lFont.Size]);
  if lStyleStr <> '' then
    lCurStr := lCurStr + ' style="' + lStyleStr + '"';

  AStrings.Add(lCurStr);

  for i := 0 to FWorksheet.GetLastRowIndex() do
  begin
    AStrings.Add('|-');

    for j := 0 to FWorksheet.GetLastColIndex do
    begin
      lCell := FWorksheet.FindCell(i, j);
      lCurStr := FWorksheet.ReadAsUTF8Text(lCell);
//      if lCurStr = '' then lCurStr := '&nbsp;';

      // Check for invalid characters
      if not ValidXMLText(lCurStr, false) then
        Workbook.AddErrorMsg(rsInvalidCharacterInCell, [
          GetCellString(i, j)
        ]);

      lStyleStr := '';
      lColSpanStr := '';
      lRowSpanStr := '';
      lColWidthStr := '';
      lRowHeightStr := '';
      lCurUsedFormatting := FWorksheet.ReadUsedFormatting(lCell);

      // Row header
      isHeader := (soHasFrozenPanes in FWorksheet.Options) and
         ((i < cardinal(FWorksheet.TopPaneHeight)) or (j < cardinal(FWorksheet.LeftPaneWidth)));

      // Column width (to be considered in first row)
      if i = 0 then
      begin
        lCol := FWorksheet.FindCol(j);
        if lCol <> nil then
          lColWidthStr := Format(' width="%.0fpt"', [lCol^.Width*FWorkbook.GetDefaultFontSize*0.5]);
      end;

      // Row height (to be considered in first column)
      if j = 0 then
      begin
        lRow := FWorksheet.FindRow(i);
        if lRow <> nil then
          lRowHeightStr := Format(' height="%.0fpt"', [lRow^.Height*FWorkbook.GetDefaultFontSize]);
      end;

      // Font
      lFont := FWorkbook.GetDefaultFont;
      if (uffFont in lCurUsedFormatting) then
      begin
        lFont := FWorksheet.ReadCellFont(lCell);
        if fssBold in lFont.Style then lCurStr := '<b>' + lCurStr + '</b>';
        if fssItalic in lFont.Style then lCurStr := '<i>' + lCurStr + '</i>';
        if fssUnderline in lFont.Style then lCurStr := '<u>' + lCurStr + '</u>';
        if fssStrikeout in lFont.Style then lCurStr := '<s>' + lCurStr + '</s>';
      end else
      if uffBold in lCurUsedFormatting then
        lCurStr := '<b>' + lCurStr + '</b>';

      // Background color
      if uffBackground in lCurUsedFormatting then
      begin
        lCurColor := FWorksheet.ReadBackgroundColor(lCell);
        lStyleStr := Format('background-color:%s;color:%s;', [
          FWorkbook.GetPaletteColorAsHTMLStr(lCurColor),
          FWorkbook.GetPaletteColorAsHTMLStr(lFont.Color)
        ]);
      end;

      // Horizontal alignment
      if uffHorAlign in lCurUsedFormatting then
      begin
        horAlign := FWorksheet.ReadHorAlignment(lCell);
        if horAlign = haDefault then
          case lCell^.ContentType of
            cctNumber,
            cctDateTime : horAlign := haRight;
            cctBool     : horAlign := haCenter;
            else          horAlign := haLeft;
          end;
        case horAlign of
          haLeft   : lStyleStr := lStyleStr + 'text-align:left;';
          haCenter : lStyleStr := lStyleStr + 'text-align:center;';
          haRight  : lStyleStr := lStyleStr + 'text-align:right';
        end;
      end;

      // vertical alignment
      if uffVertAlign in lCurUsedFormatting then
      begin
        vertAlign := FWorksheet.ReadVertAlignment(lCell);
        case vertAlign of
          vaTop    : lStyleStr := lStyleStr + 'vertical-align:top;';
          vaCenter : lStyleStr := lStyleStr + 'vertical-align:center;';
          vaBottom : lStyleStr := lStyleStr + 'vertical-align:bottom;';
        end;
      end;

      // borders
      if uffBorder in lCurUsedFormatting then
      begin
        borders := FWorksheet.ReadCellBorders(lCell);
        if (cbWest in borders) then
          lStyleStr := lStyleStr + DoBorder(cbWest, lCell);
        if (cbEast in borders) then
          lStyleStr := lStyleStr + DoBorder(cbEast, lCell);
        if (cbNorth in borders) then
          lStyleStr := lStyleStr + DoBorder(cbNorth, lCell);
        if (cbSouth in borders) then
          lStyleStr := lStyleStr + DoBorder(cbSouth, lCell);
      end;

      // Merged cells
      if FWorksheet.IsMerged(lCell) then
      begin
        FWorksheet.FindMergedRange(lCell, r1, c1, r2, c2);
        if (i = r1) and (j = c1) then
        begin
          if r1 < r2 then
            lRowSpanStr := Format(' rowspan="%d"', [r2-r1+1]);
          if c1 < c2 then
            lColSpanStr := Format(' colspan="%d"', [c2-c1+1]);
        end
        else
        if (i >= r1) and (i <= r2) and (j >= c1) and (j <= c2) then
          Continue;
      end;

      // Put everything together...
      if lStyleStr <> '' then
        lStyleStr := Format(' style="%s"', [lStyleStr]);

      if lRowSpanStr <> '' then
        lStyleStr := lRowSpanStr + lStyleStr;

      if lColSpanStr <> '' then
        lStyleStr := lColSpanStr + lStyleStr;

      if lColWidthStr <> '' then
        lStyleStr := lColWidthStr + lStyleStr;

      if lRowHeightStr <> '' then
        lStyleStr := lRowHeightStr + lStyleStr;

      if lCurStr <> '' then
        lCurStr := ' ' + lCurStr;

      if lStyleStr <> '' then
        lCurStr := lStyleStr + ' |' + lCurStr;

      lCurStr := PIPE_CHAR[isHeader] + lCurStr;

      // Add to list
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
