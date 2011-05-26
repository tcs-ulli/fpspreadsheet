unit xlsbiffcommon;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  fpspreadsheet,
  fpsutils;

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
    function FindFormattingInList(AFormat: PCell): Integer;
    procedure AddDefaultFormats(); virtual;
    procedure ListAllFormattingStylesCallback(ACell: PCell; AStream: TStream);
    procedure ListAllFormattingStyles(AData: TsWorkbook);
  end;

implementation

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

