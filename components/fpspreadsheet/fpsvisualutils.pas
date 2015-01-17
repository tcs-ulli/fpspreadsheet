unit fpsvisualutils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics,
  fpstypes, fpspreadsheet;

procedure Convert_sFont_to_Font(AWorkbook: TsWorkbook; sFont: TsFont; AFont: TFont);
procedure Convert_Font_to_sFont(AWorkbook: TsWorkbook; AFont: TFont; sFont: TsFont);
function FindNearestPaletteIndex(AWorkbook: TsWorkbook; AColor: TColor): TsColor;
function WrapText(ACanvas: TCanvas; const AText: string; AMaxWidth: integer): string;


implementation

uses
  Types, LCLType, LCLIntf, Math;

{@@ ----------------------------------------------------------------------------
  Converts a spreadsheet font to a font used for painting (TCanvas.Font).

  @param  AWorkbook  Workbook in which the font is used
  @param  sFont      Font as used by fpspreadsheet (input)
  @param  AFont      Font as used by TCanvas for painting (output)
-------------------------------------------------------------------------------}
procedure Convert_sFont_to_Font(AWorkbook: TsWorkbook; sFont: TsFont; AFont: TFont);
begin
  if Assigned(AFont) and Assigned(sFont) then begin
    AFont.Name := sFont.FontName;
    AFont.Size := round(sFont.Size);
    AFont.Style := [];
    if fssBold in sFont.Style then AFont.Style := AFont.Style + [fsBold];
    if fssItalic in sFont.Style then AFont.Style := AFont.Style + [fsItalic];
    if fssUnderline in sFont.Style then AFont.Style := AFont.Style + [fsUnderline];
    if fssStrikeout in sFont.Style then AFont.Style := AFont.Style + [fsStrikeout];
    AFont.Color := AWorkbook.GetPaletteColor(sFont.Color);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Converts a font used for painting (TCanvas.Font) to a spreadsheet font.

  @param  AFont  Font as used by TCanvas for painting (input)
  @param  sFont  Font as used by fpspreadsheet (output)
-------------------------------------------------------------------------------}
procedure Convert_Font_to_sFont(AWorkbook: TsWorkbook; AFont: TFont; sFont: TsFont);
begin
  if Assigned(AFont) and Assigned(sFont) then begin
    sFont.FontName := AFont.Name;
    sFont.Size := AFont.Size;
    sFont.Style := [];
    if fsBold in AFont.Style then Include(sFont.Style, fssBold);
    if fsItalic in AFont.Style then Include(sFont.Style, fssItalic);
    if fsUnderline in AFont.Style then Include(sFont.Style, fssUnderline);
    if fsStrikeout in AFont.Style then Include(sFont.Style, fssStrikeout);
    sFont.Color := FindNearestPaletteIndex(AWorkbook, AFont.Color);
  end;
end;

function FindNearestPaletteIndex(AWorkbook: TsWorkbook; AColor: TColor): TsColor;

  procedure ColorToHSL(RGB: TColor; out H, S, L : double);
  // Taken from https://code.google.com/p/thtmlviewer/source/browse/trunk/source/HSLUtils.pas?r=277
  // The procedure in GraphUtils crashes for some colors in Laz < 1.3
  var
    R, G, B, D, Cmax, Cmin: double;
  begin
    R := GetRValue(RGB) / 255;
    G := GetGValue(RGB) / 255;
    B := GetBValue(RGB) / 255;
    Cmax := Max(R, Max(G, B));
    Cmin := Min(R, Min(G, B));

    // calculate luminosity
    L := (Cmax + Cmin) / 2;

    if Cmax = Cmin then begin // it's grey
      H := 0; // it's actually undefined
      S := 0
    end else
    begin
      D := Cmax - Cmin;

      // calculate Saturation
      if L < 0.5 then
        S := D / (Cmax + Cmin)
      else
        S := D / (2 - Cmax - Cmin);

      // calculate Hue
      if R = Cmax then
        H := (G - B) / D
      else
      if G = Cmax then
        H := 2 + (B - R) /D
      else
        H := 4 + (R - G) / D;

      H := H / 6;
      if H < 0 then
        H := H + 1
    end
  end;

  function ColorDistance(color1, color2: TColor): Double;
  var
    H1,S1,L1, H2,S2,L2: Double;
  begin
    ColorToHSL(color1, H1, S1, L1);
    ColorToHSL(color2, H2, S2, L2);
    Result := sqr(H1-H2) + sqr(S1-S2) + sqr(L1-L2);
  end;

  {
  // To be activated when Lazarus 1.4 is available. (RgbToHLS bug in Laz < 1.3)

  function ColorDistance(color1, color2: TColor): Integer;
  type
    TRGBA = packed record R, G, B, A: Byte end;
  var
    H1,L1,S1, H2,L2,S2: Byte;
  begin
    ColorToHLS(color1, H1,L1,S1);
    ColorToHLS(color2, H2,L2,S2);
    result := sqr(Integer(H1)-H2) + sqr(Integer(L1)-L2) + sqr(Integer(S1)-S2);
  end;
  }

var
  i: Integer;
  dist, mindist: Double;
begin
  Result := 0;
  if AWorkbook <> nil then
  begin
    mindist := 1E308;
    for i:=0 to AWorkbook.GetPaletteSize-1 do
    begin
      dist := ColorDistance(AColor, TColor(AWorkbook.GetPaletteColor(i)));
      if dist < mindist then
      begin
        mindist := dist;
        Result := i;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Wraps text by inserting line ending characters so that the lines are not
  longer than AMaxWidth.

  @param   ACanvas       Canvas on which the text will be drawn
  @param   AText         Text to be drawn
  @param   AMaxWidth     Maximimum line width (in pixels)
  @return  Text with inserted line endings such that the lines are shorter than
           AMaxWidth.

  @note    Based on ocde posted by user "taazz" in the Lazarus forum
           http://forum.lazarus.freepascal.org/index.php/topic,21305.msg124743.html#msg124743
-------------------------------------------------------------------------------}
function WrapText(ACanvas: TCanvas; const AText: string; AMaxWidth: integer): string;
var
  DC: HDC;
  textExtent: TSize = (cx:0; cy:0);
  S, P, E: PChar;
  line: string;
  isFirstLine: boolean;
begin
  Result := '';
  DC := ACanvas.Handle;
  isFirstLine := True;
  P := PChar(AText);
  while P^ = ' ' do
    Inc(P);
  while P^ <> #0 do begin
    S := P;
    E := nil;
    while (P^ <> #0) and (P^ <> #13) and (P^ <> #10) do begin
      LCLIntf.GetTextExtentPoint(DC, S, P - S + 1, textExtent);
      if (textExtent.CX > AMaxWidth) and (E <> nil) then begin
        if (P^ <> ' ') and (P^ <> ^I) then begin
          while (E >= S) do
            case E^ of
              '.', ',', ';', '?', '!', '-', ':',
              ')', ']', '}', '>', '/', '\', ' ':
                break;
              else
                Dec(E);
            end;
          if E < S then
            E := P - 1;
        end;
        Break;
      end;
      E := P;
      Inc(P);
    end;
    if E <> nil then begin
      while (E >= S) and (E^ = ' ') do
        Dec(E);
    end;
    if E <> nil then
      SetString(Line, S, E - S + 1)
    else
      SetLength(Line, 0);
    if (P^ = #13) or (P^ = #10) then begin
      Inc(P);
      if (P^ <> (P - 1)^) and ((P^ = #13) or (P^ = #10)) then
        Inc(P);
      if P^ = #0 then
        line := line + LineEnding;
    end
    else if P^ <> ' ' then
      P := E + 1;
    while P^ = ' ' do
      Inc(P);
    if isFirstLine then begin
      Result := Line;
      isFirstLine := False;
    end else
      Result := Result + LineEnding + line;
  end;
end;

end.
