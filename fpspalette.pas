{ fpsPalette }

{@@ ----------------------------------------------------------------------------
  Palette support for fpspreadsheet file formats

  AUTHORS: Werner Pamler, Felipe Monteiro de Carvalho, Reinier Olislagers

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}

unit fpsPalette;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpstypes, fpspreadsheet;

type

  { TsPalette }
  TsPalette = class
  private
    FColors: array of TsColor;
    function GetColor(AIndex: Integer): TsColor;
    procedure SetColor(AIndex: Integer; AColor: TsColor);
  public
    constructor Create;
    procedure AddBuiltinColors; virtual;
    function AddColor(AColor: TsColor; ABigEndian: Boolean = false): Integer;
    procedure AddExcelColors;
    function AddUniqueColor(AColor: TsColor; ABigEndian: Boolean = false): Integer;
    procedure Clear;
    procedure CollectFromWorkbook(AWorkbook: TsWorkbook);
    function ColorUsedInWorkbook(APaletteIndex: Integer; AWorkbook: TsWorkbook): Boolean;
    function FindClosestColorIndex(AColor: TsColor; AMaxPaletteCount: Integer = -1): Integer;
    function FindColor(AColor: TsColor; AMaxPaletteCount: Integer = -1): Integer;
    function Count: Integer;
    procedure Trim(AMaxSize: Integer);
    procedure UseColors(const AColors: array of TsColor; ABigEndian: Boolean = false);
    property Colors[AIndex: Integer]: TsColor read GetColor write SetColor; default;
  end;

  procedure MakeLEPalette(var AColors: array of TsColor);


implementation

uses
  fpsutils;

{@@ ----------------------------------------------------------------------------
  If a palette is coded as big-endian (e.g. by copying the rgb values from
  the OpenOffice documentation) the palette values can be converted by means
  of this procedure to little-endian which is required by fpspreadsheet.

  @param AColors      Color array to be converted.
                      After conversion, its color values are replaced.
-------------------------------------------------------------------------------}
procedure MakeLEPalette(var AColors: array of TsColor);
var
  i: Integer;
begin
  for i := 0 to High(AColors) do
    AColors[i] := LongRGBToExcelPhysical(AColors[i])
end;


{@@ ----------------------------------------------------------------------------
  Constructor of the palette: initializes the color array
-------------------------------------------------------------------------------}
constructor TsPalette.Create;
begin
  inherited;
  SetLength(FColors, 0);
end;

{@@ ----------------------------------------------------------------------------
  Adds an rgb color value to the palette and returns the palette index
  of the new color.

  Existing colors are not checked.

  If ABigEndian is TRUE then the rgb values are assumed to be in big endian
  order (r = high byte).
  By default, rgb is in little-endian order (r = low byte)
-------------------------------------------------------------------------------}
function TsPalette.AddColor(AColor: TsColor; ABigEndian: Boolean = false): Integer;
begin
  if ABigEndian then
    AColor := LongRGBToExcelPhysical(AColor);

  SetLength(FColors, Length(FColors) + 1);
  FColors[High(FColors)] := AColor;

  Result := High(FColors);
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in colors
-------------------------------------------------------------------------------}
procedure TsPalette.AddBuiltinColors;
begin
  AddColor(scBlack);   // 0
  AddColor(scWhite);   // 1
  AddColor(scRed);     // 2
  AddColor(scGreen);   // 3
  AddColor(scBlue);    // 4
  AddColor(scYellow);  // 5
  AddColor(scMagenta); // 6
  AddColor(scCyan);    // 7
end;

{@@ ----------------------------------------------------------------------------
  Adds the standard palette of Excel 8

  NOTE: To get the full Excel8 palette call this after AddBuiltinColors
-------------------------------------------------------------------------------}
procedure TsPalette.AddExcelColors;
begin
  AddColor($000000, true);   // $08: EGA black
  AddColor($FFFFFF, true);   // $09: EGA white
  AddColor($FF0000, true);   // $0A: EGA red
  AddColor($00FF00, true);   // $0B: EGA green
  AddColor($0000FF, true);   // $0C: EGA blue
  AddColor($FFFF00, true);   // $0D: EGA yellow
  AddColor($FF00FF, true);   // $0E: EGA magenta
  AddColor($00FFFF, true);   // $0F: EGA cyan

  AddColor($800000, true);   // $10: EGA dark red
  AddColor($008000, true);   // $11: EGA dark green
  AddColor($000080, true);   // $12: EGA dark blue
  AddColor($808000, true);   // $13: EGA olive
  AddColor($800080, true);   // $14: EGA purple
  AddColor($008080, true);   // $15: EGA teal
  AddColor($C0C0C0, true);   // $16: EGA silver
  AddColor($808080, true);   // $17: EGA gray

  AddColor($9999FF, true);   // $18:
  AddColor($993366, true);   // $19:
  AddColor($FFFFCC, true);   // $1A:
  AddColor($CCFFFF, true);   // $1B:
  AddColor($660066, true);   // $1C:
  AddColor($FF8080, true);   // $1D:
  AddColor($0066CC, true);   // $1E:
  AddColor($CCCCFF, true);   // $1F:

  AddColor($000080, true);   // $20:
  AddColor($FF00FF, true);   // $21:
  AddColor($FFFF00, true);   // $22:
  AddColor($00FFFF, true);   // $23:
  AddColor($800080, true);   // $24:
  AddColor($800000, true);   // $25:
  AddColor($008080, true);   // $26:
  AddColor($0000FF, true);   // $27:
  AddColor($00CCFF, true);   // $28:
  AddColor($CCFFFF, true);   // $29:
  AddColor($CCFFCC, true);   // $2A:
  AddColor($FFFF99, true);   // $2B:
  AddColor($99CCFF, true);   // $2C:
  AddColor($FF99CC, true);   // $2D:
  AddColor($CC99FF, true);   // $2E:
  AddColor($FFCC99, true);   // $2F:

  AddColor($3366FF, true);   // $30:
  AddColor($33CCCC, true);   // $31:
  AddColor($99CC00, true);   // $32:
  AddColor($FFCC00, true);   // $33:
  AddColor($FF9900, true);   // $34:
  AddColor($FF6600, true);   // $35:
  AddColor($666699, true);   // $36:
  AddColor($969696, true);   // $37:
  AddColor($003366, true);   // $38:
  AddColor($339966, true);   // $39:
  AddColor($003300, true);   // $3A:
  AddColor($333300, true);   // $3B:
  AddColor($993300, true);   // $3C:
  AddColor($993366, true);   // $3D:
  AddColor($333399, true);   // $3E:
  AddColor($333333, true);   // $3F:
end;

{@@ ----------------------------------------------------------------------------
  Adds the specified color to the palette if it does not yet exist.

  Returns the palette index of the new or existing color
-------------------------------------------------------------------------------}
function TsPalette.AddUniqueColor(AColor: TsColor;
  ABigEndian: Boolean = false): Integer;
begin
  if ABigEndian then
    AColor := LongRGBToExcelPhysical(AColor);

  Result := FindColor(AColor);
  if Result = -1 then result := AddColor(AColor);
end;

{@@ ----------------------------------------------------------------------------
  Clears the palette
-------------------------------------------------------------------------------}
procedure TsPalette.Clear;
begin
  SetLength(FColors, 0);
end;

{@@ ----------------------------------------------------------------------------
  Collects the colors used in the specified workbook
-------------------------------------------------------------------------------}
procedure TsPalette.CollectFromWorkbook(AWorkbook: TsWorkbook);
var
  i: Integer;
  sheet: TsWorksheet;
  cell: PCell;
  fmt: TsCellFormat;
  fnt: TsFont;
  cb: TsCellBorder;
begin
  for i:=0 to AWorkbook.GetWorksheetCount-1 do
  begin
    sheet := AWorkbook.GetWorksheetByIndex(i);
    for cell in sheet.Cells do begin
      fmt := sheet.ReadCellFormat(cell);
      if (uffBackground in fmt.UsedFormattingFields) then
      begin
        AddUniqueColor(fmt.Background.BgColor);
        AddUniqueColor(fmt.Background.FgColor);
      end;
      if (uffFont in fmt.UsedFormattingFields) then
      begin
        fnt := AWorkbook.GetFont(fmt.FontIndex);
        AddUniqueColor(fnt.Color);
      end;
      if (uffBorder in fmt.UsedFormattingFields) then
      begin
        for cb in TsCellBorder do
          if (cb in fmt.Border) then
            AddUniqueColor(fmt.BorderStyles[cb].Color);
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a given color is used somewhere within the entire workbook

  @param  APaletteIndex   Palette index of the color
  @result True if the color is used by at least one cell, false if not.
-------------------------------------------------------------------------------}
function TsPalette.ColorUsedInWorkbook(APaletteIndex: Integer;
  AWorkbook: TsWorkbook): Boolean;
var
  sheet: TsWorksheet;
  cell: PCell;
  i: Integer;
  fnt: TsFont;
  b: TsCellBorder;
  fmt: PsCellFormat;
  color: TsColor;
begin
  color := GetColor(APaletteIndex);
  if (color = scNotDefined) or (AWorkbook = nil) then
    exit(false);

  Result := true;
  for i:=0 to AWorkbook.GetWorksheetCount-1 do
  begin
    sheet := AWorkbook.GetWorksheetByIndex(i);
    for cell in sheet.Cells do
    begin
      fmt := AWorkbook.GetPointerToCellFormat(cell^.FormatIndex);
      if (uffBackground in fmt^.UsedFormattingFields) then
      begin
        if fmt^.Background.BgColor = color then exit;
        if fmt^.Background.FgColor = color then exit;
      end;
      if (uffBorder in fmt^.UsedFormattingFields) then
        for b in TsCellBorders do
          if (b in fmt^.Border) and (fmt^.BorderStyles[b].Color = color) then
            exit;
      if (uffFont in fmt^.UsedFormattingFields) then
      begin
        fnt := AWorkbook.GetFont(fmt^.FontIndex);
        if fnt.Color = color then
          exit;
      end;
    end;
  end;
  Result := false;
end;

{@@ ----------------------------------------------------------------------------
  Finds the palette color index which points to a color that is "closest" to a
  given color. "Close" means here smallest length of the rgb-difference vector.

  @param   AColor            Rgb color value to be considered
  @param   AMaxPaletteCount  Number of palette entries considered. Example:
                             BIFF5/BIFF8 can write only 64 colors, i.e
                             AMaxPaletteCount = 64
  @return  Palette index of the color closest to AColor
-------------------------------------------------------------------------------}
function TsPalette.FindClosestColorIndex(AColor: TsColor;
  AMaxPaletteCount: Integer = -1): Integer;
type
  TRGBA = record r,g,b,a: Byte end;
var
  rgb: TRGBA;
  rgb0: TRGBA absolute AColor;
  dist: Double;
  minDist: Double;
  i: Integer;
  n: Integer;
begin
  Result := -1;
  minDist := 1E108;
  n := Length(FColors);
  if AMaxPaletteCount > n then n := AMaxPaletteCount;
  for i := 0 to n - 1 do
  begin
    rgb := TRGBA(GetColor(i));
    dist := sqr(rgb.r - rgb0.r) + sqr(rgb.g - rgb0.g) + sqr(rgb.b - rgb0.b);
    if dist < minDist then
    begin
      Result := i;
      minDist := dist;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Finds the palette color index which belongs to the specified color.
  Returns -1 if the color is not contained in the palette.

  @param   AColor            Rgb color value to be considered
  @param   AMaxPaletteCount  Number of palette entries considered. Example:
                             BIFF5/BIFF8 can write only 64 colors, i.e
                             AMaxPaletteCount = 64
  @return  Palette index of AColor
-------------------------------------------------------------------------------}
function TsPalette.FindColor(AColor: TsColor;
  AMaxPaletteCount: Integer = -1): Integer;
var
  n: Integer;
begin
  n := Length(FColors);
  if AMaxPaletteCount > n then n := AMaxPaletteCount;
  for Result := 0 to n - 1 do
    if GetColor(Result) = AColor then
      exit;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Reads the rgb color for the given index from the palette.
  Can be type-cast to TColor for usage in GUI applications.

  @param  AIndex  Index of the color considered
  @return A number containing the rgb components in little-endian notation.
-------------------------------------------------------------------------------}
function TsPalette.GetColor(AIndex: Integer): TsColor;
begin
  if (AIndex >= 0) and (AIndex < Length(FColors)) then
    Result := FColors[AIndex]
  else
    Result := scNotDefined;
end;

{@@ ----------------------------------------------------------------------------
  Returns the number of palette colors
-------------------------------------------------------------------------------}
function TsPalette.Count: Integer;
begin
  Result := Length(FColors);
end;

{@@ ----------------------------------------------------------------------------
  Replaces a color value of the palette by a new value.
  The color must be given in little-endian notation (ABGR, with A=0).

  @param  AIndex   Palette index of the color to be replaced
  @param  AColor   Number containing the rgb components of the new color
-------------------------------------------------------------------------------}
procedure TsPalette.SetColor(AIndex: Integer; AColor: TsColor);
begin
  if (AIndex >= 0) and (AIndex < Length(FColors)) then
    FColors[AIndex] := AColor;
end;

{@@ ----------------------------------------------------------------------------
  Trims the size of the palette
-------------------------------------------------------------------------------}
procedure TsPalette.Trim(AMaxSize: Integer);
begin
  if Length(FColors) > AMaxSize then
    SetLength(FColors, AMaxSize);
end;

{@@ ----------------------------------------------------------------------------
  Uses the color array to with "APalette" points in the palette.
  If ABigEndian is true it is assumed that the input colors are specified in
  big-endian notation, i.e. "blue" in the low-value byte.
-------------------------------------------------------------------------------}
procedure TsPalette.UseColors(const AColors: array of TsColor; ABigEndian: Boolean = false);
var
  i: Integer;
begin
  SetLength(FColors, High(AColors)+1);
  if ABigEndian then
    for i:=0 to High(AColors) do FColors[i] := LongRGBToExcelPhysical(AColors[i])
  else
    for i:=0 to High(AColors) do FColors[i] := AColors[i];
end;


end.
