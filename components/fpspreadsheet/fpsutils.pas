unit fpsutils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils; 

function WordToLE(AValue: Word): Word;
function DWordToLE(AValue: Cardinal): Cardinal;
function IntegerToLE(AValue: Integer): Integer;

implementation

{
  Endianess helper functions

  Excel files are all written with Little Endian byte order,
  so it's necessary to swap the data to be able to build a
  correct file on big endian systems.
}

function WordToLE(AValue: Word): Word;
begin
  {$IFDEF BIG_ENDIAN}
    Result := ((AValue shl 8) and $FF00) or ((AValue shr 8) and $00FF);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

function DWordToLE(AValue: Cardinal): Cardinal;
begin
  {$IFDEF BIG_ENDIAN}
// ??  Result := ((AValue shl 8) and $FF00) or ((AValue shr 8) and $00FF);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

function IntegerToLE(AValue: Integer): Integer;
begin
  {$IFDEF BIG_ENDIAN}
// ??  Result := ((AValue shl 8) and $FF00) or ((AValue shr 8) and $00FF);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

end.

