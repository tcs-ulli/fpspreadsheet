unit xlsbiffcommon;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  fpspreadsheet,
  fpsutils;

{ Excel Constants which don't change across versions }
const
  { Built In Color Pallete Indexes }
  BUILT_IN_COLOR_PALLETE_BLACK     = $08; // 000000H
  BUILT_IN_COLOR_PALLETE_WHITE     = $09; // FFFFFFH
  BUILT_IN_COLOR_PALLETE_RED       = $0A; // FF0000H
  BUILT_IN_COLOR_PALLETE_GREEN     = $0B; // 00FF00H
  BUILT_IN_COLOR_PALLETE_BLUE      = $0C; // 0000FFH
  BUILT_IN_COLOR_PALLETE_YELLOW    = $0D; // FFFF00H
  BUILT_IN_COLOR_PALLETE_MAGENTA   = $0E; // FF00FFH
  BUILT_IN_COLOR_PALLETE_CYAN      = $0F; // 00FFFFH
  BUILT_IN_COLOR_PALLETE_DARK_RED  = $10; // 800000H
  BUILT_IN_COLOR_PALLETE_DARK_GREEN= $11; // 008000H
  BUILT_IN_COLOR_PALLETE_DARK_BLUE = $12; // 000080H
  BUILT_IN_COLOR_PALLETE_OLIVE     = $13; // 808000H
  BUILT_IN_COLOR_PALLETE_PURPLE    = $14; // 800080H
  BUILT_IN_COLOR_PALLETE_TEAL      = $15; // 008080H
  BUILT_IN_COLOR_PALLETE_SILVER    = $16; // C0C0C0H
  BUILT_IN_COLOR_PALLETE_GREY      = $17; // 808080H

  EXTRA_COLOR_PALETTE_GREY10PCT    = $18; // E6E6E6H
  EXTRA_COLOR_PALETTE_GREY20PCT    = $19; // E6E6E6H

type

  { TsSpreadBIFFReader }

  TsSpreadBIFFReader = class(TsCustomSpreadReader)
  protected
  end;

  { TsSpreadBIFFWriter }

  TsSpreadBIFFWriter = class(TsCustomSpreadWriter)
  protected
    {
      An array with cells which are models for the used styles
      In this array the Row property holds the Index to the corresponding XF field
    }
    FFormattingStyles: array of TCell;
    NextXFIndex: Integer; // Indicates which should be the next XF Index when filling the styles list
    function FPSColorToEXCELPallete(AColor: TsColor): Word;
    function FindFormattingInList(AFormat: PCell): Integer;
    procedure AddDefaultFormats(); virtual;
    procedure ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
    procedure ListAllFormattingStyles(AData: TsWorkbook);
  end;

implementation

function TsSpreadBIFFWriter.FPSColorToEXCELPallete(AColor: TsColor): Word;
begin
  case AColor of
    scBlack: Result := BUILT_IN_COLOR_PALLETE_BLACK;
    scWhite: Result := BUILT_IN_COLOR_PALLETE_WHITE;
    scRed: Result := BUILT_IN_COLOR_PALLETE_RED;
    scGREEN: Result := BUILT_IN_COLOR_PALLETE_GREEN;
    scBLUE: Result := BUILT_IN_COLOR_PALLETE_BLUE;
    scYELLOW: Result := BUILT_IN_COLOR_PALLETE_YELLOW;
    scMAGENTA: Result := BUILT_IN_COLOR_PALLETE_MAGENTA;
    scCYAN: Result := BUILT_IN_COLOR_PALLETE_CYAN;
    scDarkRed: Result := BUILT_IN_COLOR_PALLETE_DARK_RED;
    scDarkGreen: Result := BUILT_IN_COLOR_PALLETE_DARK_GREEN;
    scDarkBlue: Result := BUILT_IN_COLOR_PALLETE_DARK_BLUE;
    scOLIVE: Result := BUILT_IN_COLOR_PALLETE_OLIVE;
    scPURPLE: Result := BUILT_IN_COLOR_PALLETE_PURPLE;
    scTEAL: Result := BUILT_IN_COLOR_PALLETE_TEAL;
    scSilver: Result := BUILT_IN_COLOR_PALLETE_SILVER;
    scGrey: Result := BUILT_IN_COLOR_PALLETE_GREY;
    //
    scGrey10pct: Result := EXTRA_COLOR_PALETTE_GREY10PCT;
    scGrey20pct: Result := EXTRA_COLOR_PALETTE_GREY20PCT;
  end;
end;

{
  Checks if the style of a cell is in the list FFormattingStyles and returns the index
  or -1 if it isn't
}
function TsSpreadBIFFWriter.FindFormattingInList(AFormat: PCell): Integer;
var
  i: Integer;
begin
  Result := -1;

  for i := 0 to Length(FFormattingStyles) - 1 do
  begin
    if (FFormattingStyles[i].UsedFormattingFields <> AFormat^.UsedFormattingFields) then Continue;

    if uffTextRotation in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].TextRotation <> AFormat^.TextRotation) then Continue;

    if uffBorder in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].Border <> AFormat^.Border) then Continue;

    if uffBackgroundColor in AFormat^.UsedFormattingFields then
      if (FFormattingStyles[i].BackgroundColor <> AFormat^.BackgroundColor) then Continue;

    // If we arrived here it means that the styles match
    Exit(i);
  end;
end;

{ Each descendent should define it's own default formats, if any.
  Always add the normal, unformatted style first to speed up. }
procedure TsSpreadBIFFWriter.AddDefaultFormats();
begin
  SetLength(FFormattingStyles, 0);
  NextXFIndex := 0;
end;

procedure TsSpreadBIFFWriter.ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
var
  Len: Integer;
begin
  if ACell^.UsedFormattingFields = [] then Exit;

  if FindFormattingInList(ACell) <> -1 then Exit;

  Len := Length(FFormattingStyles);
  SetLength(FFormattingStyles, Len+1);
  FFormattingStyles[Len] := ACell^;
  FFormattingStyles[Len].Row := NextXFIndex;
  Inc(NextXFIndex);
end;

procedure TsSpreadBIFFWriter.ListAllFormattingStyles(AData: TsWorkbook);
var
  i: Integer;
begin
  SetLength(FFormattingStyles, 0);

  AddDefaultFormats();

  for i := 0 to AData.GetWorksheetCount - 1 do
  begin
    IterateThroughCells(nil, AData.GetWorksheetByIndex(i).Cells, ListAllFormattingStylesCallback);
  end;
end;

end.

