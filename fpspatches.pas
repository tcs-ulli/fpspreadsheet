{@@ ----------------------------------------------------------------------------
  Provides functions and procedures if FPSpreadsheet is compiled in an older
  version of Lazarus / fpc.

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpspatches;

{$mode objfpc}{$H+}
{$I fps.inc}

interface

uses
  Classes, SysUtils;


{$IFDEF FPS_VARISBOOL}
{ Needed only if FPC version is < 2.6.4 }
  function VarIsBool(const V: Variant): Boolean;
{$ENDIF}


{$IFDEF FPS_LAZUTF8}
  // implemented in LazUTF8 in r43348 (Laz 1.2)
  function UTF8LeftStr(const AText: String; const ACount: Integer): String;
  function UTF8RightStr(const AText: String; const ACount: Integer): String;
  function UTF8StringReplace(const S, OldPattern, NewPattern: String;
    Flags: TReplaceFlags; ALanguage: string=''): String;
  function UTF8LowerCase(const AInStr: string; ALanguage: string=''): string;
  function UTF8UpperCase(const AInStr: string; ALanguage: string=''): string;
{$ENDIF}


{$IFDEF FPS_FORMATDATETIME}
type
  TFormatDateTimeOption = (fdoInterval);
  TFormatDateTimeOptions =  set of TFormatDateTimeOption;

  // Needed if fpc version is < 3.0.
  function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
    Options : TFormatDateTimeOptions = []): string;
  function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
    const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []): string;
{$ENDIF}


implementation

{$IFDEF FPS_VARISBOOL}
function VarIsBool(const V: Variant): Boolean;
begin
  Result := (TVarData(V).vType and varTypeMask) = varboolean;
end;
{$ENDIF}

{$IFDEF FPS_LAZUTF8}
function UTF8CharacterLength(p: PChar): integer;
begin
  if p<>nil then begin
    if ord(p^)<%11000000 then begin
      // regular single byte character (#0 is a character, this is pascal ;)
      Result:=1;
    end
    else begin
      // multi byte
      if ((ord(p^) and %11100000) = %11000000) then begin
        // could be 2 byte character
        if (ord(p[1]) and %11000000) = %10000000 then
          Result:=2
        else
          Result:=1;
      end
      else if ((ord(p^) and %11110000) = %11100000) then begin
        // could be 3 byte character
        if ((ord(p[1]) and %11000000) = %10000000)
        and ((ord(p[2]) and %11000000) = %10000000) then
          Result:=3
        else
          Result:=1;
      end
      else if ((ord(p^) and %11111000) = %11110000) then begin
        // could be 4 byte character
        if ((ord(p[1]) and %11000000) = %10000000)
        and ((ord(p[2]) and %11000000) = %10000000)
        and ((ord(p[3]) and %11000000) = %10000000) then
          Result:=4
        else
          Result:=1;
      end
      else
        Result:=1;
    end;
  end else
    Result:=0;
end;

function UTF8CharStart(UTF8Str: PChar; Len, CharIndex: PtrInt): PChar;
var
  CharLen: LongInt;
begin
  Result:=UTF8Str;
  if Result<>nil then begin
    while (CharIndex>0) and (Len>0) do begin
      CharLen:=UTF8CharacterLength(Result);
      dec(Len,CharLen);
      dec(CharIndex);
      inc(Result,CharLen);
    end;
    if (CharIndex<>0) or (Len<0) then
      Result:=nil;
  end;
end;

function UTF8Copy(const s: string; StartCharIndex, CharCount: PtrInt): string;
// returns substring
var
  StartBytePos: PChar;
  EndBytePos: PChar;
  MaxBytes: PtrInt;
begin
  StartBytePos:=UTF8CharStart(PChar(s),length(s),StartCharIndex-1);
  if StartBytePos=nil then
    Result:=''
  else begin
    MaxBytes:=PtrInt(PChar(s)+length(s)-StartBytePos);
    EndBytePos:=UTF8CharStart(StartBytePos,MaxBytes,CharCount);
    if EndBytePos=nil then
      Result:=copy(s,StartBytePos-PChar(s)+1,MaxBytes)
    else
      Result:=copy(s,StartBytePos-PChar(s)+1,EndBytePos-StartBytePos);
  end;
end;

function UTF8LeftStr(const AText: String; const ACount: Integer): String;
begin
  Result := Utf8Copy(AText,1,ACount);
end;


function UTF8Length(p: PChar; ByteCount: PtrInt): PtrInt;
var
  CharLen: LongInt;
begin
  Result:=0;
  while (ByteCount>0) do begin
    inc(Result);
    CharLen:=UTF8CharacterLength(p);
    inc(p,CharLen);
    dec(ByteCount,CharLen);
  end;
end;

function UTF8Length(const s: string): PtrInt;
begin
  Result:=UTF8Length(PChar(s),length(s));
end;

function Utf8RightStr(const AText: String; const ACount: Integer): String;
var
  j,l:integer;
begin
  l := Utf8Length(AText);
  j := ACount;
  if (j > l) then j := l;
  Result := Utf8Copy(AText,l-j+1,j);
end;

function UTF8StringReplace(const S, OldPattern, NewPattern: String;
  Flags: TReplaceFlags; ALanguage: string): String;
// same algorithm as StringReplace, but using UTF8LowerCase
// for case insensitive search
var
  Srch, OldP, RemS: string;
  P: Integer;
begin
  Srch := S;
  OldP := OldPattern;
  if rfIgnoreCase in Flags then
  begin
    Srch := UTF8LowerCase(Srch,ALanguage);
    OldP := UTF8LowerCase(OldP,ALanguage);
  end;
  RemS := S;
  Result := '';
  while Length(Srch) <> 0 do
  begin
    P := Pos(OldP, Srch);
    if P = 0 then
    begin
      Result := Result + RemS;
      Srch := '';
    end
    else
    begin
      Result := Result + Copy(RemS,1,P-1) + NewPattern;
      P := P + Length(OldP);
      RemS := Copy(RemS, P, Length(RemS)-P+1);
      if not (rfReplaceAll in Flags) then
      begin
        Result := Result + RemS;
        Srch := '';
      end
      else
        Srch := Copy(Srch, P, Length(Srch)-P+1);
    end;
  end;
end;

function UTF8LowerCase(const AInStr: string; ALanguage: string=''): string;
var
  CounterDiff: PtrInt;
  InStr, InStrEnd, OutStr: PChar;
  // Language identification
  IsTurkish: Boolean;
  c1, c2, c3, new_c1, new_c2, new_c3: Char;
  p: SizeInt;
begin
  Result:=AInStr;
  InStr := PChar(AInStr);
  InStrEnd := InStr + length(AInStr); // points behind last char

  // Do a fast initial parsing of the string to maybe avoid doing
  // UniqueString if the resulting string will be identical
  while (InStr < InStrEnd) do
  begin
    c1 := InStr^;
    case c1 of
    'A'..'Z': Break;
    #$C3..#$FF:
      case c1 of
      #$C3..#$C9, #$CE, #$CF, #$D0..#$D5, #$E1..#$E2,#$E5:
        begin
          c2 := InStr[1];
          case c1 of
          #$C3: if c2 in [#$80..#$9E] then Break;
          #$C4:
          begin
            case c2 of
            #$80..#$AF, #$B2..#$B6: if ord(c2) mod 2 = 0 then Break;
            #$B8..#$FF: if ord(c2) mod 2 = 1 then Break;
            #$B0: Break;
            end;
          end;
          #$C5:
          begin
            case c2 of
              #$8A..#$B7: if ord(c2) mod 2 = 0 then Break;
              #$00..#$88, #$B9..#$FF: if ord(c2) mod 2 = 1 then Break;
              #$B8: Break;
            end;
          end;
          // Process E5 to avoid stopping on chinese chars
          #$E5: if (c2 = #$BC) and (InStr[2] in [#$A1..#$BA]) then Break;
          // Others are too complex, better not to pre-inspect them
          else
            Break;
          end;
          // already lower, or otherwhise not affected
        end;
      end;
    end;
    inc(InStr);
  end;

  if InStr >= InStrEnd then Exit;

  // Language identification
  IsTurkish := (ALanguage = 'tr') or (ALanguage = 'az'); // Turkish and Azeri have a special handling

  UniqueString(Result);
  OutStr := PChar(Result) + (InStr - PChar(AInStr));
  CounterDiff := 0;

  while InStr < InStrEnd do
  begin
    c1 := InStr^;
    case c1 of
      // codepoints      UTF-8 range           Description                Case change
      // $0041..$005A    $41..$5A              Capital ASCII              X+$20
      'A'..'Z':
      begin
        { First ASCII chars }
        // Special turkish handling
        // capital undotted I to small undotted i
        if IsTurkish and (c1 = 'I') then
        begin
          p:=OutStr - PChar(Result);
          SetLength(Result,Length(Result)+1);// Increase the buffer
          OutStr := PChar(Result)+p;
          OutStr^ := #$C4;
          inc(OutStr);
          OutStr^ := #$B1;
          dec(CounterDiff);
        end
        else
        begin
          OutStr^ := chr(ord(c1)+32);
        end;
        inc(InStr);
        inc(OutStr);
      end;

      // Chars with 2-bytes which might be modified
      #$C3..#$D5:
      begin
        c2 := InStr[1];
        new_c1 := c1;
        new_c2 := c2;
        case c1 of
        // Latin Characters 0000–0FFF http://en.wikibooks.org/wiki/Unicode/Character_reference/0000-0FFF
        // codepoints      UTF-8 range           Description                Case change
        // $00C0..$00D6    C3 80..C3 96          Capital Latin with accents X+$20
        // $D7             C3 97                 Multiplication Sign        N/A
        // $00D8..$00DE    C3 98..C3 9E          Capital Latin with accents X+$20
        // $DF             C3 9F                 German beta ß              already lowercase
        #$C3:
        begin
          case c2 of
          #$80..#$96, #$98..#$9E: new_c2 := chr(ord(c2) + $20)
          end;
        end;
        // $0100..$012F    C4 80..C4 AF        Capital/Small Latin accents  if mod 2 = 0 then X+1
        // $0130..$0131    C4 B0..C4 B1        Turkish
        //  C4 B0 turkish uppercase dotted i -> 'i'
        //  C4 B1 turkish lowercase undotted ı
        // $0132..$0137    C4 B2..C4 B7        Capital/Small Latin accents  if mod 2 = 0 then X+1
        // $0138           C4 B8               ĸ                            N/A
        // $0139..$024F    C4 B9..C5 88        Capital/Small Latin accents  if mod 2 = 1 then X+1
        #$C4:
        begin
          case c2 of
            #$80..#$AF, #$B2..#$B7: if ord(c2) mod 2 = 0 then new_c2 := chr(ord(c2) + 1);
            #$B0: // Turkish
            begin
              OutStr^ := 'i';
              inc(InStr, 2);
              inc(OutStr);
              inc(CounterDiff, 1);
              Continue;
            end;
            #$B9..#$BE: if ord(c2) mod 2 = 1 then new_c2 := chr(ord(c2) + 1);
            #$BF: // This crosses the borders between the first byte of the UTF-8 char
            begin
              new_c1 := #$C5;
              new_c2 := #$80;
            end;
          end;
        end;
        // $C589 ŉ
        // $C58A..$C5B7: if OldChar mod 2 = 0 then NewChar := OldChar + 1;
        // $C5B8:        NewChar := $C3BF; // Ÿ
        // $C5B9..$C8B3: if OldChar mod 2 = 1 then NewChar := OldChar + 1;
        #$C5:
        begin
          case c2 of
            #$8A..#$B7: //0
            begin
              if ord(c2) mod 2 = 0 then
                new_c2 := chr(ord(c2) + 1);
            end;
            #$00..#$88, #$B9..#$BE: //1
            begin
              if ord(c2) mod 2 = 1 then
                new_c2 := chr(ord(c2) + 1);
            end;
            #$B8:  // Ÿ
            begin
              new_c1 := #$C3;
              new_c2 := #$BF;
            end;
          end;
        end;
        {A convoluted part: C6 80..C6 8F

        0180;LATIN SMALL LETTER B WITH STROKE;Ll;0;L;;;;;N;LATIN SMALL LETTER B BAR;;0243;;0243
        0181;LATIN CAPITAL LETTER B WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER B HOOK;;;0253; => C6 81=>C9 93
        0182;LATIN CAPITAL LETTER B WITH TOPBAR;Lu;0;L;;;;;N;LATIN CAPITAL LETTER B TOPBAR;;;0183;
        0183;LATIN SMALL LETTER B WITH TOPBAR;Ll;0;L;;;;;N;LATIN SMALL LETTER B TOPBAR;;0182;;0182
        0184;LATIN CAPITAL LETTER TONE SIX;Lu;0;L;;;;;N;;;;0185;
        0185;LATIN SMALL LETTER TONE SIX;Ll;0;L;;;;;N;;;0184;;0184
        0186;LATIN CAPITAL LETTER OPEN O;Lu;0;L;;;;;N;;;;0254; ==> C9 94
        0187;LATIN CAPITAL LETTER C WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER C HOOK;;;0188;
        0188;LATIN SMALL LETTER C WITH HOOK;Ll;0;L;;;;;N;LATIN SMALL LETTER C HOOK;;0187;;0187
        0189;LATIN CAPITAL LETTER AFRICAN D;Lu;0;L;;;;;N;;;;0256; => C9 96
        018A;LATIN CAPITAL LETTER D WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER D HOOK;;;0257; => C9 97
        018B;LATIN CAPITAL LETTER D WITH TOPBAR;Lu;0;L;;;;;N;LATIN CAPITAL LETTER D TOPBAR;;;018C;
        018C;LATIN SMALL LETTER D WITH TOPBAR;Ll;0;L;;;;;N;LATIN SMALL LETTER D TOPBAR;;018B;;018B
        018D;LATIN SMALL LETTER TURNED DELTA;Ll;0;L;;;;;N;;;;;
        018E;LATIN CAPITAL LETTER REVERSED E;Lu;0;L;;;;;N;LATIN CAPITAL LETTER TURNED E;;;01DD; => C7 9D
        018F;LATIN CAPITAL LETTER SCHWA;Lu;0;L;;;;;N;;;;0259; => C9 99
        }
        #$C6:
        begin
          case c2 of
            #$81:
            begin
              new_c1 := #$C9;
              new_c2 := #$93;
            end;
            #$82..#$85:
            begin
              if ord(c2) mod 2 = 0 then
                new_c2 := chr(ord(c2) + 1);
            end;
            #$87..#$88,#$8B..#$8C:
            begin
              if ord(c2) mod 2 = 1 then
                new_c2 := chr(ord(c2) + 1);
            end;
            #$86:
            begin
              new_c1 := #$C9;
              new_c2 := #$94;
            end;
            #$89:
            begin
              new_c1 := #$C9;
              new_c2 := #$96;
            end;
            #$8A:
            begin
              new_c1 := #$C9;
              new_c2 := #$97;
            end;
            #$8E:
            begin
              new_c1 := #$C7;
              new_c2 := #$9D;
            end;
            #$8F:
            begin
              new_c1 := #$C9;
              new_c2 := #$99;
            end;
          {
          And also C6 90..C6 9F

          0190;LATIN CAPITAL LETTER OPEN E;Lu;0;L;;;;;N;LATIN CAPITAL LETTER EPSILON;;;025B; => C9 9B
          0191;LATIN CAPITAL LETTER F WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER F HOOK;;;0192; => +1
          0192;LATIN SMALL LETTER F WITH HOOK;Ll;0;L;;;;;N;LATIN SMALL LETTER SCRIPT F;;0191;;0191 <=
          0193;LATIN CAPITAL LETTER G WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER G HOOK;;;0260; => C9 A0
          0194;LATIN CAPITAL LETTER GAMMA;Lu;0;L;;;;;N;;;;0263; => C9 A3
          0195;LATIN SMALL LETTER HV;Ll;0;L;;;;;N;LATIN SMALL LETTER H V;;01F6;;01F6 <=
          0196;LATIN CAPITAL LETTER IOTA;Lu;0;L;;;;;N;;;;0269; => C9 A9
          0197;LATIN CAPITAL LETTER I WITH STROKE;Lu;0;L;;;;;N;LATIN CAPITAL LETTER BARRED I;;;0268; => C9 A8
          0198;LATIN CAPITAL LETTER K WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER K HOOK;;;0199; => +1
          0199;LATIN SMALL LETTER K WITH HOOK;Ll;0;L;;;;;N;LATIN SMALL LETTER K HOOK;;0198;;0198 <=
          019A;LATIN SMALL LETTER L WITH BAR;Ll;0;L;;;;;N;LATIN SMALL LETTER BARRED L;;023D;;023D <=
          019B;LATIN SMALL LETTER LAMBDA WITH STROKE;Ll;0;L;;;;;N;LATIN SMALL LETTER BARRED LAMBDA;;;; <=
          019C;LATIN CAPITAL LETTER TURNED M;Lu;0;L;;;;;N;;;;026F; => C9 AF
          019D;LATIN CAPITAL LETTER N WITH LEFT HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER N HOOK;;;0272; => C9 B2
          019E;LATIN SMALL LETTER N WITH LONG RIGHT LEG;Ll;0;L;;;;;N;;;0220;;0220 <=
          019F;LATIN CAPITAL LETTER O WITH MIDDLE TILDE;Lu;0;L;;;;;N;LATIN CAPITAL LETTER BARRED O;;;0275; => C9 B5
          }
          #$90:
          begin
            new_c1 := #$C9;
            new_c2 := #$9B;
          end;
          #$91, #$98: new_c2 := chr(ord(c2)+1);
          #$93:
          begin
            new_c1 := #$C9;
            new_c2 := #$A0;
          end;
          #$94:
          begin
            new_c1 := #$C9;
            new_c2 := #$A3;
          end;
          #$96:
          begin
            new_c1 := #$C9;
            new_c2 := #$A9;
          end;
          #$97:
          begin
            new_c1 := #$C9;
            new_c2 := #$A8;
          end;
          #$9C:
          begin
            new_c1 := #$C9;
            new_c2 := #$AF;
          end;
          #$9D:
          begin
            new_c1 := #$C9;
            new_c2 := #$B2;
          end;
          #$9F:
          begin
            new_c1 := #$C9;
            new_c2 := #$B5;
          end;
          {
          And also C6 A0..C6 AF

          01A0;LATIN CAPITAL LETTER O WITH HORN;Lu;0;L;004F 031B;;;;N;LATIN CAPITAL LETTER O HORN;;;01A1; => +1
          01A1;LATIN SMALL LETTER O WITH HORN;Ll;0;L;006F 031B;;;;N;LATIN SMALL LETTER O HORN;;01A0;;01A0 <=
          01A2;LATIN CAPITAL LETTER OI;Lu;0;L;;;;;N;LATIN CAPITAL LETTER O I;;;01A3; => +1
          01A3;LATIN SMALL LETTER OI;Ll;0;L;;;;;N;LATIN SMALL LETTER O I;;01A2;;01A2 <=
          01A4;LATIN CAPITAL LETTER P WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER P HOOK;;;01A5; => +1
          01A5;LATIN SMALL LETTER P WITH HOOK;Ll;0;L;;;;;N;LATIN SMALL LETTER P HOOK;;01A4;;01A4 <=
          01A6;LATIN LETTER YR;Lu;0;L;;;;;N;LATIN LETTER Y R;;;0280; => CA 80
          01A7;LATIN CAPITAL LETTER TONE TWO;Lu;0;L;;;;;N;;;;01A8; => +1
          01A8;LATIN SMALL LETTER TONE TWO;Ll;0;L;;;;;N;;;01A7;;01A7 <=
          01A9;LATIN CAPITAL LETTER ESH;Lu;0;L;;;;;N;;;;0283; => CA 83
          01AA;LATIN LETTER REVERSED ESH LOOP;Ll;0;L;;;;;N;;;;;
          01AB;LATIN SMALL LETTER T WITH PALATAL HOOK;Ll;0;L;;;;;N;LATIN SMALL LETTER T PALATAL HOOK;;;; <=
          01AC;LATIN CAPITAL LETTER T WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER T HOOK;;;01AD; => +1
          01AD;LATIN SMALL LETTER T WITH HOOK;Ll;0;L;;;;;N;LATIN SMALL LETTER T HOOK;;01AC;;01AC <=
          01AE;LATIN CAPITAL LETTER T WITH RETROFLEX HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER T RETROFLEX HOOK;;;0288; => CA 88
          01AF;LATIN CAPITAL LETTER U WITH HORN;Lu;0;L;0055 031B;;;;N;LATIN CAPITAL LETTER U HORN;;;01B0; => +1
          }
          #$A0..#$A5,#$AC:
          begin
            if ord(c2) mod 2 = 0 then
              new_c2 := chr(ord(c2) + 1);
          end;
          #$A7,#$AF:
          begin
            if ord(c2) mod 2 = 1 then
              new_c2 := chr(ord(c2) + 1);
          end;
          #$A6:
          begin
            new_c1 := #$CA;
            new_c2 := #$80;
          end;
          #$A9:
          begin
            new_c1 := #$CA;
            new_c2 := #$83;
          end;
          #$AE:
          begin
            new_c1 := #$CA;
            new_c2 := #$88;
          end;
          {
          And also C6 B0..C6 BF

          01B0;LATIN SMALL LETTER U WITH HORN;Ll;0;L;0075 031B;;;;N;LATIN SMALL LETTER U HORN;;01AF;;01AF <= -1
          01B1;LATIN CAPITAL LETTER UPSILON;Lu;0;L;;;;;N;;;;028A; => CA 8A
          01B2;LATIN CAPITAL LETTER V WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER SCRIPT V;;;028B; => CA 8B
          01B3;LATIN CAPITAL LETTER Y WITH HOOK;Lu;0;L;;;;;N;LATIN CAPITAL LETTER Y HOOK;;;01B4; => +1
          01B4;LATIN SMALL LETTER Y WITH HOOK;Ll;0;L;;;;;N;LATIN SMALL LETTER Y HOOK;;01B3;;01B3 <=
          01B5;LATIN CAPITAL LETTER Z WITH STROKE;Lu;0;L;;;;;N;LATIN CAPITAL LETTER Z BAR;;;01B6; => +1
          01B6;LATIN SMALL LETTER Z WITH STROKE;Ll;0;L;;;;;N;LATIN SMALL LETTER Z BAR;;01B5;;01B5 <=
          01B7;LATIN CAPITAL LETTER EZH;Lu;0;L;;;;;N;LATIN CAPITAL LETTER YOGH;;;0292; => CA 92
          01B8;LATIN CAPITAL LETTER EZH REVERSED;Lu;0;L;;;;;N;LATIN CAPITAL LETTER REVERSED YOGH;;;01B9; => +1
          01B9;LATIN SMALL LETTER EZH REVERSED;Ll;0;L;;;;;N;LATIN SMALL LETTER REVERSED YOGH;;01B8;;01B8 <=
          01BA;LATIN SMALL LETTER EZH WITH TAIL;Ll;0;L;;;;;N;LATIN SMALL LETTER YOGH WITH TAIL;;;; <=
          01BB;LATIN LETTER TWO WITH STROKE;Lo;0;L;;;;;N;LATIN LETTER TWO BAR;;;; X
          01BC;LATIN CAPITAL LETTER TONE FIVE;Lu;0;L;;;;;N;;;;01BD; => +1
          01BD;LATIN SMALL LETTER TONE FIVE;Ll;0;L;;;;;N;;;01BC;;01BC <=
          01BE;LATIN LETTER INVERTED GLOTTAL STOP WITH STROKE;Ll;0;L;;;;;N;LATIN LETTER INVERTED GLOTTAL STOP BAR;;;; X
          01BF;LATIN LETTER WYNN;Ll;0;L;;;;;N;;;01F7;;01F7  <=
          }
          #$B8,#$BC:
          begin
            if ord(c2) mod 2 = 0 then
              new_c2 := chr(ord(c2) + 1);
          end;
          #$B3..#$B6:
          begin
            if ord(c2) mod 2 = 1 then
              new_c2 := chr(ord(c2) + 1);
          end;
          #$B1:
          begin
            new_c1 := #$CA;
            new_c2 := #$8A;
          end;
          #$B2:
          begin
            new_c1 := #$CA;
            new_c2 := #$8B;
          end;
          #$B7:
          begin
            new_c1 := #$CA;
            new_c2 := #$92;
          end;
          end;
        end;
        #$C7:
        begin
          case c2 of
          #$84..#$8C,#$B1..#$B3:
          begin
            if (ord(c2) and $F) mod 3 = 1 then new_c2 := chr(ord(c2) + 2)
            else if (ord(c2) and $F) mod 3 = 2 then new_c2 := chr(ord(c2) + 1);
          end;
          #$8D..#$9C:
          begin
            if ord(c2) mod 2 = 1 then
              new_c2 := chr(ord(c2) + 1);
          end;
          #$9E..#$AF,#$B4..#$B5,#$B8..#$BF:
          begin
            if ord(c2) mod 2 = 0 then
              new_c2 := chr(ord(c2) + 1);
          end;
          {
          01F6;LATIN CAPITAL LETTER HWAIR;Lu;0;L;;;;;N;;;;0195;
          01F7;LATIN CAPITAL LETTER WYNN;Lu;0;L;;;;;N;;;;01BF;
          }
          #$B6:
          begin
            new_c1 := #$C6;
            new_c2 := #$95;
          end;
          #$B7:
          begin
            new_c1 := #$C6;
            new_c2 := #$BF;
          end;
          end;
        end;
        {
        Codepoints 0200 to 023F
        }
        #$C8:
        begin
          // For this one we can simply start with a default and override for some specifics
          if (c2 in [#$80..#$B3]) and (ord(c2) mod 2 = 0) then new_c2 := chr(ord(c2) + 1);

          case c2 of
          #$A0:
          begin
            new_c1 := #$C6;
            new_c2 := #$9E;
          end;
          #$A1: new_c2 := c2;
          {
          023A;LATIN CAPITAL LETTER A WITH STROKE;Lu;0;L;;;;;N;;;;2C65; => E2 B1 A5
          023B;LATIN CAPITAL LETTER C WITH STROKE;Lu;0;L;;;;;N;;;;023C; => +1
          023C;LATIN SMALL LETTER C WITH STROKE;Ll;0;L;;;;;N;;;023B;;023B <=
          023D;LATIN CAPITAL LETTER L WITH BAR;Lu;0;L;;;;;N;;;;019A; => C6 9A
          023E;LATIN CAPITAL LETTER T WITH DIAGONAL STROKE;Lu;0;L;;;;;N;;;;2C66; => E2 B1 A6
          023F;LATIN SMALL LETTER S WITH SWASH TAIL;Ll;0;L;;;;;N;;;2C7E;;2C7E <=
          0240;LATIN SMALL LETTER Z WITH SWASH TAIL;Ll;0;L;;;;;N;;;2C7F;;2C7F <=
          }
          #$BA,#$BE:
          begin
            p:= OutStr - PChar(Result);
            SetLength(Result,Length(Result)+1);// Increase the buffer
            OutStr := PChar(Result)+p;
            OutStr^ := #$E2;
            inc(OutStr);
            OutStr^ := #$B1;
            inc(OutStr);
            if c2 = #$BA then OutStr^ := #$A5
            else OutStr^ := #$A6;
            dec(CounterDiff);
            inc(OutStr);
            inc(InStr, 2);
            Continue;
          end;
          #$BD:
          begin
            new_c1 := #$C6;
            new_c2 := #$9A;
          end;
          #$BB: new_c2 := chr(ord(c2) + 1);
          end;
        end;
        {
        Codepoints 0240 to 027F

        Here only 0240..024F needs lowercase
        }
        #$C9:
        begin
          case c2 of
          #$81..#$82:
          begin
            if ord(c2) mod 2 = 1 then
              new_c2 := chr(ord(c2) + 1);
          end;
          #$86..#$8F:
          begin
            if ord(c2) mod 2 = 0 then
              new_c2 := chr(ord(c2) + 1);
          end;
          #$83:
          begin
            new_c1 := #$C6;
            new_c2 := #$80;
          end;
          #$84:
          begin
            new_c1 := #$CA;
            new_c2 := #$89;
          end;
          #$85:
          begin
            new_c1 := #$CA;
            new_c2 := #$8C;
          end;
          end;
        end;
        // $CE91..$CE9F: NewChar := OldChar + $20; // Greek Characters
        // $CEA0..$CEA9: NewChar := OldChar + $E0; // Greek Characters
        #$CE:
        begin
          case c2 of
            // 0380 = CE 80
            #$86: new_c2 := #$AC;
            #$88: new_c2 := #$AD;
            #$89: new_c2 := #$AE;
            #$8A: new_c2 := #$AF;
            #$8C: new_c1 := #$CF; // By coincidence new_c2 remains the same
            #$8E:
            begin
              new_c1 := #$CF;
              new_c2 := #$8D;
            end;
            #$8F:
            begin
              new_c1 := #$CF;
              new_c2 := #$8E;
            end;
            // 0390 = CE 90
            #$91..#$9F:
            begin
              new_c2 := chr(ord(c2) + $20);
            end;
            // 03A0 = CE A0
            #$A0..#$AB:
            begin
              new_c1 := #$CF;
              new_c2 := chr(ord(c2) - $20);
            end;
          end;
        end;
        // 03C0 = CF 80
        // 03D0 = CF 90
        // 03E0 = CF A0
        // 03F0 = CF B0
        #$CF:
        begin
          case c2 of
            // 03CF;GREEK CAPITAL KAI SYMBOL;Lu;0;L;;;;;N;;;;03D7; CF 8F => CF 97
            #$8F: new_c2 := #$97;
            // 03D8;GREEK LETTER ARCHAIC KOPPA;Lu;0;L;;;;;N;;;;03D9;
            #$98: new_c2 := #$99;
            // 03DA;GREEK LETTER STIGMA;Lu;0;L;;;;;N;GREEK CAPITAL LETTER STIGMA;;;03DB;
            #$9A: new_c2 := #$9B;
            // 03DC;GREEK LETTER DIGAMMA;Lu;0;L;;;;;N;GREEK CAPITAL LETTER DIGAMMA;;;03DD;
            #$9C: new_c2 := #$9D;
            // 03DE;GREEK LETTER KOPPA;Lu;0;L;;;;;N;GREEK CAPITAL LETTER KOPPA;;;03DF;
            #$9E: new_c2 := #$9F;
            {
            03E0;GREEK LETTER SAMPI;Lu;0;L;;;;;N;GREEK CAPITAL LETTER SAMPI;;;03E1;
            03E1;GREEK SMALL LETTER SAMPI;Ll;0;L;;;;;N;;;03E0;;03E0
            03E2;COPTIC CAPITAL LETTER SHEI;Lu;0;L;;;;;N;GREEK CAPITAL LETTER SHEI;;;03E3;
            03E3;COPTIC SMALL LETTER SHEI;Ll;0;L;;;;;N;GREEK SMALL LETTER SHEI;;03E2;;03E2
            ...
            03EE;COPTIC CAPITAL LETTER DEI;Lu;0;L;;;;;N;GREEK CAPITAL LETTER DEI;;;03EF;
            03EF;COPTIC SMALL LETTER DEI;Ll;0;L;;;;;N;GREEK SMALL LETTER DEI;;03EE;;03EE
            }
            #$A0..#$AF: if ord(c2) mod 2 = 0 then
                          new_c2 := chr(ord(c2) + 1);
            // 03F4;GREEK CAPITAL THETA SYMBOL;Lu;0;L;<compat> 0398;;;;N;;;;03B8;
            #$B4:
            begin
              new_c1 := #$CE;
              new_c2 := #$B8;
            end;
            // 03F7;GREEK CAPITAL LETTER SHO;Lu;0;L;;;;;N;;;;03F8;
            #$B7: new_c2 := #$B8;
            // 03F9;GREEK CAPITAL LUNATE SIGMA SYMBOL;Lu;0;L;<compat> 03A3;;;;N;;;;03F2;
            #$B9: new_c2 := #$B2;
            // 03FA;GREEK CAPITAL LETTER SAN;Lu;0;L;;;;;N;;;;03FB;
            #$BA: new_c2 := #$BB;
            // 03FD;GREEK CAPITAL REVERSED LUNATE SIGMA SYMBOL;Lu;0;L;;;;;N;;;;037B;
            #$BD:
            begin
              new_c1 := #$CD;
              new_c2 := #$BB;
            end;
            // 03FE;GREEK CAPITAL DOTTED LUNATE SIGMA SYMBOL;Lu;0;L;;;;;N;;;;037C;
            #$BE:
            begin
              new_c1 := #$CD;
              new_c2 := #$BC;
            end;
            // 03FF;GREEK CAPITAL REVERSED DOTTED LUNATE SIGMA SYMBOL;Lu;0;L;;;;;N;;;;037D;
            #$BF:
            begin
              new_c1 := #$CD;
              new_c2 := #$BD;
            end;
          end;
        end;
        // $D080..$D08F: NewChar := OldChar + $110; // Cyrillic alphabet
        // $D090..$D09F: NewChar := OldChar + $20; // Cyrillic alphabet
        // $D0A0..$D0AF: NewChar := OldChar + $E0; // Cyrillic alphabet
        #$D0:
        begin
          c2 := InStr[1];
          case c2 of
            #$80..#$8F:
            begin
              new_c1 := chr(ord(c1)+1);
              new_c2  := chr(ord(c2) + $10);
            end;
            #$90..#$9F:
            begin
              new_c2 := chr(ord(c2) + $20);
            end;
            #$A0..#$AF:
            begin
              new_c1 := chr(ord(c1)+1);
              new_c2 := chr(ord(c2) - $20);
            end;
          end;
        end;
        // Archaic and non-slavic cyrillic 460-47F = D1A0-D1BF
        // These require just adding 1 to get the lowercase
        #$D1:
        begin
          if (c2 in [#$A0..#$BF]) and (ord(c2) mod 2 = 0) then
            new_c2 := chr(ord(c2) + 1);
        end;
        // Archaic and non-slavic cyrillic 480-4BF = D280-D2BF
        // These mostly require just adding 1 to get the lowercase
        #$D2:
        begin
          case c2 of
            #$80:
            begin
              new_c2 := chr(ord(c2) + 1);
            end;
            // #$81 is already lowercase
            // #$82-#$89 ???
            #$8A..#$BF:
            begin
              if ord(c2) mod 2 = 0 then
                new_c2 := chr(ord(c2) + 1);
            end;
          end;
        end;
        {
        Codepoints  04C0..04FF
        }
        #$D3:
        begin
          case c2 of
            #$80: new_c2 := #$8F;
            #$81..#$8E:
            begin
              if ord(c2) mod 2 = 1 then
                new_c2 := chr(ord(c2) + 1);
            end;
            #$90..#$BF:
            begin
              if ord(c2) mod 2 = 0 then
                new_c2 := chr(ord(c2) + 1);
            end;
          end;
        end;
        {
        Codepoints  0500..053F

        Armenian starts in 0531
        }
        #$D4:
        begin
          if ord(c2) mod 2 = 0 then
            new_c2 := chr(ord(c2) + 1);

          // Armenian
          if c2 in [#$B1..#$BF] then
          begin
            new_c1 := #$D5;
            new_c2 := chr(ord(c2) - $10);
          end;
        end;
        {
        Codepoints  0540..057F

        Armenian
        }
        #$D5:
        begin
          case c2 of
            #$80..#$8F:
            begin
              new_c2 := chr(ord(c2) + $30);
            end;
            #$90..#$96:
            begin
              new_c1 := #$D6;
              new_c2 := chr(ord(c2) - $10);
            end;
          end;
        end;
        end;
        // Common code 2-byte modifiable chars
        if (CounterDiff <> 0) then
        begin
          OutStr^ := new_c1;
          OutStr[1] := new_c2;
        end
        else
        begin
          if (new_c1 <> c1) then OutStr^ := new_c1;
          if (new_c2 <> c2) then OutStr[1] := new_c2;
        end;
        inc(InStr, 2);
        inc(OutStr, 2);
      end;
      {
      Characters with 3 bytes
      }
      #$E1:
      begin
        new_c1 := c1;
        c2 := InStr[1];
        c3 := InStr[2];
        new_c2 := c2;
        new_c3 := c3;
        {
        Georgian codepoints 10A0-10C5 => 2D00-2D25

        In UTF-8 this is:
        E1 82 A0 - E1 82 BF => E2 B4 80 - E2 B4 9F
        E1 83 80 - E1 83 85 => E2 B4 A0 - E2 B4 A5
        }
        case c2 of
        #$82:
        if (c3 in [#$A0..#$BF]) then
        begin
          new_c1 := #$E2;
          new_c2 := #$B4;
          new_c3 := chr(ord(c3) - $20);
        end;
        #$83:
        if (c3 in [#$80..#$85]) then
        begin
          new_c1 := #$E2;
          new_c2 := #$B4;
          new_c3 := chr(ord(c3) + $20);
        end;
        {
        Extra chars between 1E00..1EFF

        Blocks of chars:
          1E00..1E3F    E1 B8 80..E1 B8 BF
          1E40..1E7F    E1 B9 80..E1 B9 BF
          1E80..1EBF    E1 BA 80..E1 BA BF
          1EC0..1EFF    E1 BB 80..E1 BB BF
        }
        #$B8..#$BB:
        begin
          // Start with a default and change for some particular chars
          if ord(c3) mod 2 = 0 then
            new_c3 := chr(ord(c3) + 1);

          { Only 1E96..1E9F are different E1 BA 96..E1 BA 9F

          1E96;LATIN SMALL LETTER H WITH LINE BELOW;Ll;0;L;0068 0331;;;;N;;;;;
          1E97;LATIN SMALL LETTER T WITH DIAERESIS;Ll;0;L;0074 0308;;;;N;;;;;
          1E98;LATIN SMALL LETTER W WITH RING ABOVE;Ll;0;L;0077 030A;;;;N;;;;;
          1E99;LATIN SMALL LETTER Y WITH RING ABOVE;Ll;0;L;0079 030A;;;;N;;;;;
          1E9A;LATIN SMALL LETTER A WITH RIGHT HALF RING;Ll;0;L;<compat> 0061 02BE;;;;N;;;;;
          1E9B;LATIN SMALL LETTER LONG S WITH DOT ABOVE;Ll;0;L;017F 0307;;;;N;;;1E60;;1E60
          1E9C;LATIN SMALL LETTER LONG S WITH DIAGONAL STROKE;Ll;0;L;;;;;N;;;;;
          1E9D;LATIN SMALL LETTER LONG S WITH HIGH STROKE;Ll;0;L;;;;;N;;;;;
          1E9E;LATIN CAPITAL LETTER SHARP S;Lu;0;L;;;;;N;;;;00DF; => C3 9F
          1E9F;LATIN SMALL LETTER DELTA;Ll;0;L;;;;;N;;;;;
          }
          if (c2 = #$BA) and (c3 in [#$96..#$9F]) then new_c3 := c3;
          // LATIN CAPITAL LETTER SHARP S => to german Beta
          if (c2 = #$BA) and (c3 = #$9E) then
          begin
            inc(InStr, 3);
            OutStr^ := #$C3;
            inc(OutStr);
            OutStr^ := #$9F;
            inc(OutStr);
            inc(CounterDiff, 1);
            Continue;
          end;
        end;
        {
        Extra chars between 1F00..1FFF

        Blocks of chars:
          1E00..1E3F    E1 BC 80..E1 BC BF
          1E40..1E7F    E1 BD 80..E1 BD BF
          1E80..1EBF    E1 BE 80..E1 BE BF
          1EC0..1EFF    E1 BF 80..E1 BF BF
        }
        #$BC:
        begin
          // Start with a default and change for some particular chars
          if (ord(c3) mod $10) div 8 = 1 then
            new_c3 := chr(ord(c3) - 8);
        end;
        #$BD:
        begin
          // Start with a default and change for some particular chars
          case c3 of
          #$80..#$8F, #$A0..#$AF: if (ord(c3) mod $10) div 8 = 1 then
                        new_c3 := chr(ord(c3) - 8);
          {
          1F50;GREEK SMALL LETTER UPSILON WITH PSILI;Ll;0;L;03C5 0313;;;;N;;;;;
          1F51;GREEK SMALL LETTER UPSILON WITH DASIA;Ll;0;L;03C5 0314;;;;N;;;1F59;;1F59
          1F52;GREEK SMALL LETTER UPSILON WITH PSILI AND VARIA;Ll;0;L;1F50 0300;;;;N;;;;;
          1F53;GREEK SMALL LETTER UPSILON WITH DASIA AND VARIA;Ll;0;L;1F51 0300;;;;N;;;1F5B;;1F5B
          1F54;GREEK SMALL LETTER UPSILON WITH PSILI AND OXIA;Ll;0;L;1F50 0301;;;;N;;;;;
          1F55;GREEK SMALL LETTER UPSILON WITH DASIA AND OXIA;Ll;0;L;1F51 0301;;;;N;;;1F5D;;1F5D
          1F56;GREEK SMALL LETTER UPSILON WITH PSILI AND PERISPOMENI;Ll;0;L;1F50 0342;;;;N;;;;;
          1F57;GREEK SMALL LETTER UPSILON WITH DASIA AND PERISPOMENI;Ll;0;L;1F51 0342;;;;N;;;1F5F;;1F5F
          1F59;GREEK CAPITAL LETTER UPSILON WITH DASIA;Lu;0;L;03A5 0314;;;;N;;;;1F51;
          1F5B;GREEK CAPITAL LETTER UPSILON WITH DASIA AND VARIA;Lu;0;L;1F59 0300;;;;N;;;;1F53;
          1F5D;GREEK CAPITAL LETTER UPSILON WITH DASIA AND OXIA;Lu;0;L;1F59 0301;;;;N;;;;1F55;
          1F5F;GREEK CAPITAL LETTER UPSILON WITH DASIA AND PERISPOMENI;Lu;0;L;1F59 0342;;;;N;;;;1F57;
          }
          #$99,#$9B,#$9D,#$9F: new_c3 := chr(ord(c3) - 8);
          end;
        end;
        #$BE:
        begin
          // Start with a default and change for some particular chars
          case c3 of
          #$80..#$B9: if (ord(c3) mod $10) div 8 = 1 then
                        new_c3 := chr(ord(c3) - 8);
          {
          1FB0;GREEK SMALL LETTER ALPHA WITH VRACHY;Ll;0;L;03B1 0306;;;;N;;;1FB8;;1FB8
          1FB1;GREEK SMALL LETTER ALPHA WITH MACRON;Ll;0;L;03B1 0304;;;;N;;;1FB9;;1FB9
          1FB2;GREEK SMALL LETTER ALPHA WITH VARIA AND YPOGEGRAMMENI;Ll;0;L;1F70 0345;;;;N;;;;;
          1FB3;GREEK SMALL LETTER ALPHA WITH YPOGEGRAMMENI;Ll;0;L;03B1 0345;;;;N;;;1FBC;;1FBC
          1FB4;GREEK SMALL LETTER ALPHA WITH OXIA AND YPOGEGRAMMENI;Ll;0;L;03AC 0345;;;;N;;;;;
          1FB6;GREEK SMALL LETTER ALPHA WITH PERISPOMENI;Ll;0;L;03B1 0342;;;;N;;;;;
          1FB7;GREEK SMALL LETTER ALPHA WITH PERISPOMENI AND YPOGEGRAMMENI;Ll;0;L;1FB6 0345;;;;N;;;;;
          1FB8;GREEK CAPITAL LETTER ALPHA WITH VRACHY;Lu;0;L;0391 0306;;;;N;;;;1FB0;
          1FB9;GREEK CAPITAL LETTER ALPHA WITH MACRON;Lu;0;L;0391 0304;;;;N;;;;1FB1;
          1FBA;GREEK CAPITAL LETTER ALPHA WITH VARIA;Lu;0;L;0391 0300;;;;N;;;;1F70;
          1FBB;GREEK CAPITAL LETTER ALPHA WITH OXIA;Lu;0;L;0386;;;;N;;;;1F71;
          1FBC;GREEK CAPITAL LETTER ALPHA WITH PROSGEGRAMMENI;Lt;0;L;0391 0345;;;;N;;;;1FB3;
          1FBD;GREEK KORONIS;Sk;0;ON;<compat> 0020 0313;;;;N;;;;;
          1FBE;GREEK PROSGEGRAMMENI;Ll;0;L;03B9;;;;N;;;0399;;0399
          1FBF;GREEK PSILI;Sk;0;ON;<compat> 0020 0313;;;;N;;;;;
          }
          #$BA:
          begin
            new_c2 := #$BD;
            new_c3 := #$B0;
          end;
          #$BB:
          begin
            new_c2 := #$BD;
            new_c3 := #$B1;
          end;
          #$BC: new_c3 := #$B3;
          end;
        end;
        end;

        if (CounterDiff <> 0) then
        begin
          OutStr^ := new_c1;
          OutStr[1] := new_c2;
          OutStr[2] := new_c3;
        end
        else
        begin
          if c1 <> new_c1 then OutStr^ := new_c1;
          if c2 <> new_c2 then OutStr[1] := new_c2;
          if c3 <> new_c3 then OutStr[2] := new_c3;
        end;

        inc(InStr, 3);
        inc(OutStr, 3);
      end;
      {
      More Characters with 3 bytes, so exotic stuff between:
      $2126..$2183                    E2 84 A6..E2 86 83
      $24B6..$24CF    Result:=u+26;   E2 92 B6..E2 93 8F
      $2C00..$2C2E    Result:=u+48;   E2 B0 80..E2 B0 AE
      $2C60..$2CE2                    E2 B1 A0..E2 B3 A2
      }
      #$E2:
      begin
        new_c1 := c1;
        c2 := InStr[1];
        c3 := InStr[2];
        new_c2 := c2;
        new_c3 := c3;
        // 2126;OHM SIGN;Lu;0;L;03A9;;;;N;OHM;;;03C9; E2 84 A6 => CF 89
        if (c2 = #$84) and (c3 = #$A6) then
        begin
          inc(InStr, 3);
          OutStr^ := #$CF;
          inc(OutStr);
          OutStr^ := #$89;
          inc(OutStr);
          inc(CounterDiff, 1);
          Continue;
        end
        {
        212A;KELVIN SIGN;Lu;0;L;004B;;;;N;DEGREES KELVIN;;;006B; E2 84 AA => 6B
        }
        else if (c2 = #$84) and (c3 = #$AA) then
        begin
          inc(InStr, 3);
          if c3 = #$AA then OutStr^ := #$6B
          else OutStr^ := #$E5;
          inc(OutStr);
          inc(CounterDiff, 2);
          Continue;
        end
        {
        212B;ANGSTROM SIGN;Lu;0;L;00C5;;;;N;ANGSTROM UNIT;;;00E5; E2 84 AB => C3 A5
        }
        else if (c2 = #$84) and (c3 = #$AB) then
        begin
          inc(InStr, 3);
          OutStr^ := #$C3;
          inc(OutStr);
          OutStr^ := #$A5;
          inc(OutStr);
          inc(CounterDiff, 1);
          Continue;
        end
        {
        2160;ROMAN NUMERAL ONE;Nl;0;L;<compat> 0049;;;1;N;;;;2170; E2 85 A0 => E2 85 B0
        2161;ROMAN NUMERAL TWO;Nl;0;L;<compat> 0049 0049;;;2;N;;;;2171;
        2162;ROMAN NUMERAL THREE;Nl;0;L;<compat> 0049 0049 0049;;;3;N;;;;2172;
        2163;ROMAN NUMERAL FOUR;Nl;0;L;<compat> 0049 0056;;;4;N;;;;2173;
        2164;ROMAN NUMERAL FIVE;Nl;0;L;<compat> 0056;;;5;N;;;;2174;
        2165;ROMAN NUMERAL SIX;Nl;0;L;<compat> 0056 0049;;;6;N;;;;2175;
        2166;ROMAN NUMERAL SEVEN;Nl;0;L;<compat> 0056 0049 0049;;;7;N;;;;2176;
        2167;ROMAN NUMERAL EIGHT;Nl;0;L;<compat> 0056 0049 0049 0049;;;8;N;;;;2177;
        2168;ROMAN NUMERAL NINE;Nl;0;L;<compat> 0049 0058;;;9;N;;;;2178;
        2169;ROMAN NUMERAL TEN;Nl;0;L;<compat> 0058;;;10;N;;;;2179;
        216A;ROMAN NUMERAL ELEVEN;Nl;0;L;<compat> 0058 0049;;;11;N;;;;217A;
        216B;ROMAN NUMERAL TWELVE;Nl;0;L;<compat> 0058 0049 0049;;;12;N;;;;217B;
        216C;ROMAN NUMERAL FIFTY;Nl;0;L;<compat> 004C;;;50;N;;;;217C;
        216D;ROMAN NUMERAL ONE HUNDRED;Nl;0;L;<compat> 0043;;;100;N;;;;217D;
        216E;ROMAN NUMERAL FIVE HUNDRED;Nl;0;L;<compat> 0044;;;500;N;;;;217E;
        216F;ROMAN NUMERAL ONE THOUSAND;Nl;0;L;<compat> 004D;;;1000;N;;;;217F;
        }
        else if (c2 = #$85) and (c3 in [#$A0..#$AF]) then new_c3 := chr(ord(c3) + $10)
        {
        2183;ROMAN NUMERAL REVERSED ONE HUNDRED;Lu;0;L;;;;;N;;;;2184; E2 86 83 => E2 86 84
        }
        else if (c2 = #$86) and (c3 = #$83) then new_c3 := chr(ord(c3) + 1)
        {
        $24B6..$24CF    Result:=u+26;   E2 92 B6..E2 93 8F

        Ex: 24B6;CIRCLED LATIN CAPITAL LETTER A;So;0;L;<circle> 0041;;;;N;;;;24D0; E2 92 B6 => E2 93 90
        }
        else if (c2 = #$92) and (c3 in [#$B6..#$BF]) then
        begin
          new_c3 := #$93;
          new_c3 := chr(ord(c3) - $26);
        end
        else if (c2 = #$93) and (c3 in [#$80..#$8F]) then new_c3 := chr(ord(c3) + 26)
        {
        $2C00..$2C2E    Result:=u+48;   E2 B0 80..E2 B0 AE

        2C00;GLAGOLITIC CAPITAL LETTER AZU;Lu;0;L;;;;;N;;;;2C30; E2 B0 80 => E2 B0 B0

        2C10;GLAGOLITIC CAPITAL LETTER NASHI;Lu;0;L;;;;;N;;;;2C40; E2 B0 90 => E2 B1 80
        }
        else if (c2 = #$B0) and (c3 in [#$80..#$8F]) then new_c3 := chr(ord(c3) + $30)
        else if (c2 = #$B0) and (c3 in [#$90..#$AE]) then
        begin
          new_c2 := #$B1;
          new_c3 := chr(ord(c3) - $10);
        end
        {
        $2C60..$2CE2                    E2 B1 A0..E2 B3 A2

        2C60;LATIN CAPITAL LETTER L WITH DOUBLE BAR;Lu;0;L;;;;;N;;;;2C61; E2 B1 A0 => +1
        2C61;LATIN SMALL LETTER L WITH DOUBLE BAR;Ll;0;L;;;;;N;;;2C60;;2C60
        2C62;LATIN CAPITAL LETTER L WITH MIDDLE TILDE;Lu;0;L;;;;;N;;;;026B; => 	C9 AB
        2C63;LATIN CAPITAL LETTER P WITH STROKE;Lu;0;L;;;;;N;;;;1D7D; => E1 B5 BD
        2C64;LATIN CAPITAL LETTER R WITH TAIL;Lu;0;L;;;;;N;;;;027D; => 	C9 BD
        2C65;LATIN SMALL LETTER A WITH STROKE;Ll;0;L;;;;;N;;;023A;;023A
        2C66;LATIN SMALL LETTER T WITH DIAGONAL STROKE;Ll;0;L;;;;;N;;;023E;;023E
        2C67;LATIN CAPITAL LETTER H WITH DESCENDER;Lu;0;L;;;;;N;;;;2C68; => E2 B1 A8
        2C68;LATIN SMALL LETTER H WITH DESCENDER;Ll;0;L;;;;;N;;;2C67;;2C67
        2C69;LATIN CAPITAL LETTER K WITH DESCENDER;Lu;0;L;;;;;N;;;;2C6A; => E2 B1 AA
        2C6A;LATIN SMALL LETTER K WITH DESCENDER;Ll;0;L;;;;;N;;;2C69;;2C69
        2C6B;LATIN CAPITAL LETTER Z WITH DESCENDER;Lu;0;L;;;;;N;;;;2C6C; => E2 B1 AC
        2C6C;LATIN SMALL LETTER Z WITH DESCENDER;Ll;0;L;;;;;N;;;2C6B;;2C6B
        2C6D;LATIN CAPITAL LETTER ALPHA;Lu;0;L;;;;;N;;;;0251; => C9 91
        2C6E;LATIN CAPITAL LETTER M WITH HOOK;Lu;0;L;;;;;N;;;;0271; => C9 B1
        2C6F;LATIN CAPITAL LETTER TURNED A;Lu;0;L;;;;;N;;;;0250; => C9 90

        2C70;LATIN CAPITAL LETTER TURNED ALPHA;Lu;0;L;;;;;N;;;;0252; => C9 92
        }
        else if (c2 = #$B1) then
        begin
          case c3 of
          #$A0: new_c3 := chr(ord(c3)+1);
          #$A2,#$A4,#$AD..#$AF,#$B0:
          begin
            inc(InStr, 3);
            OutStr^ := #$C9;
            inc(OutStr);
            case c3 of
            #$A2: OutStr^ := #$AB;
            #$A4: OutStr^ := #$BD;
            #$AD: OutStr^ := #$90;
            #$AE: OutStr^ := #$B1;
            #$AF: OutStr^ := #$90;
            #$B0: OutStr^ := #$92;
            end;
            inc(OutStr);
            inc(CounterDiff, 1);
            Continue;
          end;
          #$A3:
          begin
            new_c2 := #$B5;
            new_c3 := #$BD;
          end;
          #$A7,#$A9,#$AB: new_c3 := chr(ord(c3)+1);
          {
          2C71;LATIN SMALL LETTER V WITH RIGHT HOOK;Ll;0;L;;;;;N;;;;;
          2C72;LATIN CAPITAL LETTER W WITH HOOK;Lu;0;L;;;;;N;;;;2C73;
          2C73;LATIN SMALL LETTER W WITH HOOK;Ll;0;L;;;;;N;;;2C72;;2C72
          2C74;LATIN SMALL LETTER V WITH CURL;Ll;0;L;;;;;N;;;;;
          2C75;LATIN CAPITAL LETTER HALF H;Lu;0;L;;;;;N;;;;2C76;
          2C76;LATIN SMALL LETTER HALF H;Ll;0;L;;;;;N;;;2C75;;2C75
          2C77;LATIN SMALL LETTER TAILLESS PHI;Ll;0;L;;;;;N;;;;;
          2C78;LATIN SMALL LETTER E WITH NOTCH;Ll;0;L;;;;;N;;;;;
          2C79;LATIN SMALL LETTER TURNED R WITH TAIL;Ll;0;L;;;;;N;;;;;
          2C7A;LATIN SMALL LETTER O WITH LOW RING INSIDE;Ll;0;L;;;;;N;;;;;
          2C7B;LATIN LETTER SMALL CAPITAL TURNED E;Ll;0;L;;;;;N;;;;;
          2C7C;LATIN SUBSCRIPT SMALL LETTER J;Ll;0;L;<sub> 006A;;;;N;;;;;
          2C7D;MODIFIER LETTER CAPITAL V;Lm;0;L;<super> 0056;;;;N;;;;;
          2C7E;LATIN CAPITAL LETTER S WITH SWASH TAIL;Lu;0;L;;;;;N;;;;023F; => C8 BF
          2C7F;LATIN CAPITAL LETTER Z WITH SWASH TAIL;Lu;0;L;;;;;N;;;;0240; => C9 80
          }
          #$B2,#$B5: new_c3 := chr(ord(c3)+1);
          #$BE,#$BF:
          begin
            inc(InStr, 3);
            case c3 of
            #$BE: OutStr^ := #$C8;
            #$BF: OutStr^ := #$C9;
            end;
            OutStr^ := #$C8;
            inc(OutStr);
            case c3 of
            #$BE: OutStr^ := #$BF;
            #$BF: OutStr^ := #$80;
            end;
            inc(OutStr);
            inc(CounterDiff, 1);
            Continue;
          end;
          end;
        end
        {
        2C80;COPTIC CAPITAL LETTER ALFA;Lu;0;L;;;;;N;;;;2C81; E2 B2 80 => E2 B2 81
        ...
        2CBE;COPTIC CAPITAL LETTER OLD COPTIC OOU;Lu;0;L;;;;;N;;;;2CBF; E2 B2 BE => E2 B2 BF
        2CBF;COPTIC SMALL LETTER OLD COPTIC OOU;Ll;0;L;;;;;N;;;2CBE;;2CBE
        ...
        2CC0;COPTIC CAPITAL LETTER SAMPI;Lu;0;L;;;;;N;;;;2CC1; E2 B3 80 => E2 B2 81
        2CC1;COPTIC SMALL LETTER SAMPI;Ll;0;L;;;;;N;;;2CC0;;2CC0
        ...
        2CE2;COPTIC CAPITAL LETTER OLD NUBIAN WAU;Lu;0;L;;;;;N;;;;2CE3; E2 B3 A2 => E2 B3 A3
        2CE3;COPTIC SMALL LETTER OLD NUBIAN WAU;Ll;0;L;;;;;N;;;2CE2;;2CE2 <=
        }
        else if (c2 = #$B2) then
        begin
          if ord(c3) mod 2 = 0 then new_c3 := chr(ord(c3) + 1);
        end
        else if (c2 = #$B3) and (c3 in [#$80..#$A3]) then
        begin
          if ord(c3) mod 2 = 0 then new_c3 := chr(ord(c3) + 1);
        end;

        if (CounterDiff <> 0) then
        begin
          OutStr^ := new_c1;
          OutStr[1] := new_c2;
          OutStr[2] := new_c3;
        end
        else
        begin
          if c1 <> new_c1 then OutStr^ := new_c1;
          if c2 <> new_c2 then OutStr[1] := new_c2;
          if c3 <> new_c3 then OutStr[2] := new_c3;
        end;

        inc(InStr, 3);
        inc(OutStr, 3);
      end;
      {
      FF21;FULLWIDTH LATIN CAPITAL LETTER A;Lu;0;L;<wide> 0041;;;;N;;;;FF41; EF BC A1 => EF BD 81
      ...
      FF3A;FULLWIDTH LATIN CAPITAL LETTER Z;Lu;0;L;<wide> 005A;;;;N;;;;FF5A; EF BC BA => EF BD 9A
      }
      #$EF:
      begin
        c2 := InStr[1];
        c3 := InStr[2];

        if (c2 = #$BC) and (c3 in [#$A1..#$BA]) then
        begin
          OutStr^ := c1;
          OutStr[1] := #$BD;
          OutStr[2] := chr(ord(c3) - $20);
        end;

        if (CounterDiff <> 0) then
        begin
          OutStr^ := c1;
          OutStr[1] := c2;
          OutStr[2] := c3;
        end;

        inc(InStr, 3);
        inc(OutStr, 3);
      end;
    else
      // Copy the character if the string was disaligned by previous changes
      if (CounterDiff <> 0) then OutStr^:= c1;
      inc(InStr);
      inc(OutStr);
    end; // Case InStr^
  end; // while

  // Final correction of the buffer size
  SetLength(Result,OutStr - PChar(Result));
end;

function UTF8UpperCase(const AInStr: string; ALanguage: string=''): string;
var
  i, InCounter, OutCounter: PtrInt;
  OutStr: PChar;
  CharLen: integer;
  CharProcessed: Boolean;
  NewCharLen: integer;
  NewChar, OldChar: Word;
  // Language identification
  IsTurkish: Boolean;

  procedure CorrectOutStrSize(AOldCharSize, ANewCharSize: Integer);
  begin
    if not (ANewCharSize > AOldCharSize) then Exit; // no correction needed
    if (ANewCharSize > 20) or (AOldCharSize > 20) then Exit; // sanity check
    // Fix for bug 23428
    // If the string wasn't decreased by previous char changes,
    // and our current operation will make it bigger, then for safety
    // increase the buffer
    if (ANewCharSize > AOldCharSize) and (OutCounter >= InCounter-1) then
    begin
      SetLength(Result, Length(Result)+ANewCharSize-AOldCharSize);
      OutStr := PChar(Result);
    end;
  end;

begin
  // Start with the same string, and progressively modify
  Result:=AInStr;
  UniqueString(Result);
  OutStr := PChar(Result);

  // Language identification
  IsTurkish := (ALanguage = 'tr') or (ALanguage = 'az'); // Turkish and Azeri have a special handling

  InCounter:=1; // for AInStr
  OutCounter := 0; // for Result
  while InCounter<=length(AInStr) do
  begin
    { First ASCII chars }
    if (AInStr[InCounter] <= 'z') and (AInStr[InCounter] >= 'a') then
    begin
      // Special turkish handling
      // small dotted i to capital dotted i
      if IsTurkish and (AInStr[InCounter] = 'i') then
      begin
        SetLength(Result,Length(Result)+1);// Increase the buffer
        OutStr := PChar(Result);
        OutStr[OutCounter]:=#$C4;
        OutStr[OutCounter+1]:=#$B0;
        inc(InCounter);
        inc(OutCounter,2);
      end
      else
      begin
        OutStr[OutCounter]:=chr(ord(AInStr[InCounter])-32);
        inc(InCounter);
        inc(OutCounter);
      end;
    end
    { Now everything else }
    else
    begin
      CharLen := UTF8CharacterLength(@AInStr[InCounter]);
      CharProcessed := False;
      NewCharLen := CharLen;

      if CharLen = 2 then
      begin
        OldChar := (Ord(AInStr[InCounter]) shl 8) or Ord(AInStr[InCounter+1]);
        NewChar := 0;

        // Major processing
        case OldChar of
        // Latin Characters 0000–0FFF http://en.wikibooks.org/wiki/Unicode/Character_reference/0000-0FFF
        $C39F:        NewChar := $5353; // ß => SS
        $C3A0..$C3B6,$C3B8..$C3BE: NewChar := OldChar - $20;
        $C3BF:        NewChar := $C5B8; // ÿ
        $C481..$C4B0: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // 0130 = C4 B0
        // turkish small undotted i to capital undotted i
        $C4B1:
        begin
          OutStr[OutCounter]:='I';
          NewCharLen := 1;
          CharProcessed := True;
        end;
        $C4B2..$C4B7: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // $C4B8: ĸ without upper/lower
        $C4B9..$C4BF: if OldChar mod 2 = 0 then NewChar := OldChar - 1;
        $C580: NewChar := $C4BF; // border between bytes
        $C581..$C588: if OldChar mod 2 = 0 then NewChar := OldChar - 1;
        // $C589 ŉ => ?
        $C58A..$C5B7: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // $C5B8: // Ÿ already uppercase
        $C5B9..$C5BE: if OldChar mod 2 = 0 then NewChar := OldChar - 1;
        $C5BF: // 017F
        begin
          OutStr[OutCounter]:='S';
          NewCharLen := 1;
          CharProcessed := True;
        end;
        // 0180 = C6 80 -> A convoluted part
        $C680: NewChar := $C983;
        $C682..$C685: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        $C688: NewChar := $C687;
        $C68C: NewChar := $C68B;
        // 0190 = C6 90 -> A convoluted part
        $C692: NewChar := $C691;
        $C695: NewChar := $C7B6;
        $C699: NewChar := $C698;
        $C69A: NewChar := $C8BD;
        $C69E: NewChar := $C8A0;
        // 01A0 = C6 A0 -> A convoluted part
        $C6A0..$C6A5: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        $C6A8: NewChar := $C6A7;
        $C6AD: NewChar := $C6AC;
        // 01B0 = C6 B0
        $C6B0: NewChar := $C6AF;
        $C6B3..$C6B6: if OldChar mod 2 = 0 then NewChar := OldChar - 1;
        $C6B9: NewChar := $C6B8;
        $C6BD: NewChar := $C6BC;
        $C6BF: NewChar := $C7B7;
        // 01C0 = C7 80
        $C784..$C786: NewChar := $C784;
        $C787..$C789: NewChar := $C787;
        $C78A..$C78C: NewChar := $C78A;
        $C78E: NewChar := $C78D;
        // 01D0 = C7 90
        $C790: NewChar := $C78F;
        $C791..$C79C: if OldChar mod 2 = 0 then NewChar := OldChar - 1;
        $C79D: NewChar := $C68E;
        $C79F: NewChar := $C79E;
        // 01E0 = C7 A0
        $C7A0..$C7AF: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // 01F0 = C7 B0
        $C7B2..$C7B3: NewChar := $C7B1;
        $C7B5: NewChar := $C7B4;
        $C7B8..$C7BF: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // 0200 = C8 80
        // 0210 = C8 90
        $C880..$C89F: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // 0220 = C8 A0
        // 0230 = C8 B0
        $C8A2..$C8B3: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        $C8BC: NewChar := $C8BB;
        $C8BF:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$BE;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        // 0240 = C9 80
        $C980:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$BF;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        $C982: NewChar := $C981;
        $C986..$C98F: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // 0250 = C9 90
        $C990:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$AF;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        $C991:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$AD;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        $C992:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$B0;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        $C993: NewChar := $C681;
        $C994: NewChar := $C686;
        $C996: NewChar := $C689;
        $C997: NewChar := $C68A;
        $C999: NewChar := $C68F;
        $C99B: NewChar := $C690;
        // 0260 = C9 A0
        $C9A0: NewChar := $C693;
        $C9A3: NewChar := $C694;
        $C9A5:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$EA;
          OutStr[OutCounter+1]:= #$9E;
          OutStr[OutCounter+2]:= #$8D;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        $C9A8: NewChar := $C697;
        $C9A9: NewChar := $C696;
        $C9AB:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$A2;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        $C9AF: NewChar := $C69C;
        // 0270 = C9 B0
        $C9B1:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$AE;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        $C9B2: NewChar := $C69D;
        $C9B5: NewChar := $C69F;
        $C9BD:
        begin
          CorrectOutStrSize(2, 3);
          OutStr[OutCounter]  := #$E2;
          OutStr[OutCounter+1]:= #$B1;
          OutStr[OutCounter+2]:= #$A4;
          NewCharLen := 3;
          CharProcessed := True;
        end;
        // 0280 = CA 80
        $CA80: NewChar := $C6A6;
        $CA83: NewChar := $C6A9;
        $CA88: NewChar := $C6AE;
        $CA89: NewChar := $C984;
        $CA8A: NewChar := $C6B1;
        $CA8B: NewChar := $C6B2;
        $CA8C: NewChar := $C985;
        // 0290 = CA 90
        $CA92: NewChar := $C6B7;
        {
        03A0 = CE A0

        03AC;GREEK SMALL LETTER ALPHA WITH TONOS;Ll;0;L;03B1 0301;;;;N;GREEK SMALL LETTER ALPHA TONOS;;0386;;0386
        03AD;GREEK SMALL LETTER EPSILON WITH TONOS;Ll;0;L;03B5 0301;;;;N;GREEK SMALL LETTER EPSILON TONOS;;0388;;0388
        03AE;GREEK SMALL LETTER ETA WITH TONOS;Ll;0;L;03B7 0301;;;;N;GREEK SMALL LETTER ETA TONOS;;0389;;0389
        03AF;GREEK SMALL LETTER IOTA WITH TONOS;Ll;0;L;03B9 0301;;;;N;GREEK SMALL LETTER IOTA TONOS;;038A;;038A
        }
        $CEAC: NewChar := $CE86;
        $CEAD: NewChar := $CE88;
        $CEAE: NewChar := $CE89;
        $CEAF: NewChar := $CE8A;
        {
        03B0 = CE B0

        03B0;GREEK SMALL LETTER UPSILON WITH DIALYTIKA AND TONOS;Ll;0;L;03CB 0301;;;;N;GREEK SMALL LETTER UPSILON DIAERESIS TONOS;;;;
        03B1;GREEK SMALL LETTER ALPHA;Ll;0;L;;;;;N;;;0391;;0391
        ...
        03BF;GREEK SMALL LETTER OMICRON;Ll;0;L;;;;;N;;;039F;;039F
        }
        $CEB1..$CEBF: NewChar := OldChar - $20; // Greek Characters
        {
        03C0 = CF 80

        03C0;GREEK SMALL LETTER PI;Ll;0;L;;;;;N;;;03A0;;03A0 CF 80 => CE A0
        03C1;GREEK SMALL LETTER RHO;Ll;0;L;;;;;N;;;03A1;;03A1
        03C2;GREEK SMALL LETTER FINAL SIGMA;Ll;0;L;;;;;N;;;03A3;;03A3
        03C3;GREEK SMALL LETTER SIGMA;Ll;0;L;;;;;N;;;03A3;;03A3
        03C4;GREEK SMALL LETTER TAU;Ll;0;L;;;;;N;;;03A4;;03A4
        ....
        03CB;GREEK SMALL LETTER UPSILON WITH DIALYTIKA;Ll;0;L;03C5 0308;;;;N;GREEK SMALL LETTER UPSILON DIAERESIS;;03AB;;03AB
        03CC;GREEK SMALL LETTER OMICRON WITH TONOS;Ll;0;L;03BF 0301;;;;N;GREEK SMALL LETTER OMICRON TONOS;;038C;;038C
        03CD;GREEK SMALL LETTER UPSILON WITH TONOS;Ll;0;L;03C5 0301;;;;N;GREEK SMALL LETTER UPSILON TONOS;;038E;;038E
        03CE;GREEK SMALL LETTER OMEGA WITH TONOS;Ll;0;L;03C9 0301;;;;N;GREEK SMALL LETTER OMEGA TONOS;;038F;;038F
        03CF;GREEK CAPITAL KAI SYMBOL;Lu;0;L;;;;;N;;;;03D7;
        }
        $CF80,$CF81,$CF83..$CF8B: NewChar := OldChar - $E0; // Greek Characters
        $CF82: NewChar := $CEA3;
        $CF8C: NewChar := $CE8C;
        $CF8D: NewChar := $CE8E;
        $CF8E: NewChar := $CE8F;
        {
        03D0 = CF 90

        03D0;GREEK BETA SYMBOL;Ll;0;L;<compat> 03B2;;;;N;GREEK SMALL LETTER CURLED BETA;;0392;;0392 CF 90 => CE 92
        03D1;GREEK THETA SYMBOL;Ll;0;L;<compat> 03B8;;;;N;GREEK SMALL LETTER SCRIPT THETA;;0398;;0398 => CE 98
        03D5;GREEK PHI SYMBOL;Ll;0;L;<compat> 03C6;;;;N;GREEK SMALL LETTER SCRIPT PHI;;03A6;;03A6 => CE A6
        03D6;GREEK PI SYMBOL;Ll;0;L;<compat> 03C0;;;;N;GREEK SMALL LETTER OMEGA PI;;03A0;;03A0 => CE A0
        03D7;GREEK KAI SYMBOL;Ll;0;L;;;;;N;;;03CF;;03CF => CF 8F
        03D9;GREEK SMALL LETTER ARCHAIC KOPPA;Ll;0;L;;;;;N;;;03D8;;03D8
        03DB;GREEK SMALL LETTER STIGMA;Ll;0;L;;;;;N;;;03DA;;03DA
        03DD;GREEK SMALL LETTER DIGAMMA;Ll;0;L;;;;;N;;;03DC;;03DC
        03DF;GREEK SMALL LETTER KOPPA;Ll;0;L;;;;;N;;;03DE;;03DE
        }
        $CF90: NewChar := $CE92;
        $CF91: NewChar := $CE98;
        $CF95: NewChar := $CEA6;
        $CF96: NewChar := $CEA0;
        $CF97: NewChar := $CF8F;
        $CF99..$CF9F: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        // 03E0 = CF A0
        $CFA0..$CFAF: if OldChar mod 2 = 1 then NewChar := OldChar - 1;
        {
        03F0 = CF B0

        03F0;GREEK KAPPA SYMBOL;Ll;0;L;<compat> 03BA;;;;N;GREEK SMALL LETTER SCRIPT KAPPA;;039A;;039A => CE 9A
        03F1;GREEK RHO SYMBOL;Ll;0;L;<compat> 03C1;;;;N;GREEK SMALL LETTER TAILED RHO;;03A1;;03A1 => CE A1
        03F2;GREEK LUNATE SIGMA SYMBOL;Ll;0;L;<compat> 03C2;;;;N;GREEK SMALL LETTER LUNATE SIGMA;;03F9;;03F9
        03F5;GREEK LUNATE EPSILON SYMBOL;Ll;0;L;<compat> 03B5;;;;N;;;0395;;0395 => CE 95
        03F8;GREEK SMALL LETTER SHO;Ll;0;L;;;;;N;;;03F7;;03F7
        03FB;GREEK SMALL LETTER SAN;Ll;0;L;;;;;N;;;03FA;;03FA
        }
        $CFB0: NewChar := $CE9A;
        $CFB1: NewChar := $CEA1;
        $CFB2: NewChar := $CFB9;
        $CFB5: NewChar := $CE95;
        $CFB8: NewChar := $CFB7;
        $CFBB: NewChar := $CFBA;
        // 0400 = D0 80 ... 042F everything already uppercase
        // 0430 = D0 B0
        $D0B0..$D0BF: NewChar := OldChar - $20; // Cyrillic alphabet
        // 0440 = D1 80
        $D180..$D18F: NewChar := OldChar - $E0; // Cyrillic alphabet
        // 0450 = D1 90
        $D190..$D19F: NewChar := OldChar - $110; // Cyrillic alphabet
        end;

        if NewChar <> 0 then
        begin
          OutStr[OutCounter]  := Chr(Hi(NewChar));
          OutStr[OutCounter+1]:= Chr(Lo(NewChar));
          CharProcessed := True;
        end;
      end;

      // Copy the character if the string was disaligned by previous changed
      // and no processing was done in this character
      if (InCounter <> OutCounter+1) and (not CharProcessed) then
      begin
        for i := 0 to CharLen-1 do
          OutStr[OutCounter+i]  :=AInStr[InCounter+i];
      end;

      inc(InCounter, CharLen);
      inc(OutCounter, NewCharLen);
    end;
  end; // while

  // Final correction of the buffer size
  SetLength(Result,OutCounter);
end;

{$ENDIF}

{$IFDEF FPS_FORMATDATETIME}
{******************************************************************************}
{******************************************************************************}
{                   Patch for SysUtils.FormatDateTime                          }
{******************************************************************************}
{******************************************************************************}

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
procedure DateTimeToString(out Result: string; const FormatStr: string; const DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []);
// Copied from "fpc/rtl/objpas/sysutils/datei.inc"
var
  ResultLen: integer;
  ResultBuffer: array[0..255] of char;
  ResultCurrent: pchar;
  (*                    ---- not needed here ---
{$IFDEF MSWindows}
  isEnable_E_Format : Boolean;
  isEnable_G_Format : Boolean;
  eastasiainited : boolean;

  procedure InitEastAsia;
  var     ALCID : LCID;
         PriLangID , SubLangID : Word;

  begin
    ALCID := GetThreadLocale;
    PriLangID := ALCID and $3FF;
    if (PriLangID>0) then
       SubLangID := (ALCID and $FFFF) shr 10
      else
        begin
          PriLangID := SysLocale.PriLangID;
          SubLangID := SysLocale.SubLangID;
        end;
    isEnable_E_Format := (PriLangID = LANG_JAPANESE)
                  or
                  (PriLangID = LANG_KOREAN)
                  or
                  ((PriLangID = LANG_CHINESE)
                   and
                   (SubLangID = SUBLANG_CHINESE_TRADITIONAL)
                  );
    isEnable_G_Format := (PriLangID = LANG_JAPANESE)
                  or
                  ((PriLangID = LANG_CHINESE)
                   and
                   (SubLangID = SUBLANG_CHINESE_TRADITIONAL)
                  );
    eastasiainited :=true;
  end;
{$ENDIF MSWindows}
                     *)
  procedure StoreStr(Str: PChar; Len: Integer);
  begin
    if ResultLen + Len < SizeOf(ResultBuffer) then
    begin
      StrMove(ResultCurrent, Str, Len);
      ResultCurrent := ResultCurrent + Len;
      ResultLen := ResultLen + Len;
    end;
  end;

  procedure StoreString(const Str: string);
  var Len: integer;
  begin
   Len := Length(Str);
   if ResultLen + Len < SizeOf(ResultBuffer) then
     begin
       StrMove(ResultCurrent, pchar(Str), Len);
       ResultCurrent := ResultCurrent + Len;
       ResultLen := ResultLen + Len;
     end;
  end;

  procedure StoreInt(Value, Digits: Integer);
  var
    S: string[16];
    Len: integer;
  begin
    System.Str(Value:Digits, S);
    for Len := 1 to Length(S) do
    begin
      if S[Len] = ' ' then
        S[Len] := '0'
      else
        Break;
    end;
    StoreStr(pchar(@S[1]), Length(S));
  end ;

var
  Year, Month, Day, DayOfWeek, Hour, Minute, Second, MilliSecond: word;

  procedure StoreFormat(const FormatStr: string; Nesting: Integer; TimeFlag: Boolean);
  var
    Token, lastformattoken, prevlasttoken: char;
    FormatCurrent: pchar;
    FormatEnd: pchar;
    Count: integer;
    Clock12: boolean;
    P: pchar;
    tmp: integer;
    isInterval: Boolean;

  begin
    if Nesting > 1 then  // 0 is original string, 1 is included FormatString
      Exit;

    FormatCurrent := PChar(FormatStr);
    FormatEnd := FormatCurrent + Length(FormatStr);
    Clock12 := false;
    isInterval := false;
    P := FormatCurrent;
    // look for unquoted 12-hour clock token
    while P < FormatEnd do
    begin
      Token := P^;
      case Token of
        '''', '"':
        begin
          Inc(P);
          while (P < FormatEnd) and (P^ <> Token) do
            Inc(P);
        end;
        'A', 'a':
        begin
          if (StrLIComp(P, 'A/P', 3) = 0) or
             (StrLIComp(P, 'AMPM', 4) = 0) or
             (StrLIComp(P, 'AM/PM', 5) = 0) then
          begin
            Clock12 := true;
            break;
          end;
        end;
      end;  // case
      Inc(P);
    end ;
    token := #255;
    lastformattoken := ' ';
    prevlasttoken := 'H';
    while FormatCurrent < FormatEnd do
    begin
      Token := UpCase(FormatCurrent^);
      Count := 1;
      P := FormatCurrent + 1;
      case Token of
        '''', '"':
        begin
          while (P < FormatEnd) and (p^ <> Token) do
            Inc(P);
          Inc(P);
          Count := P - FormatCurrent;
          StoreStr(FormatCurrent + 1, Count - 2);
        end ;
        'A':
        begin
          if StrLIComp(FormatCurrent, 'AMPM', 4) = 0 then
          begin
            Count := 4;
            if Hour < 12 then
              StoreString(FormatSettings.TimeAMString)
            else
              StoreString(FormatSettings.TimePMString);
          end
          else if StrLIComp(FormatCurrent, 'AM/PM', 5) = 0 then
          begin
            Count := 5;
            if Hour < 12 then StoreStr(FormatCurrent, 2)
                         else StoreStr(FormatCurrent+3, 2);
          end
          else if StrLIComp(FormatCurrent, 'A/P', 3) = 0 then
          begin
            Count := 3;
            if Hour < 12 then StoreStr(FormatCurrent, 1)
                         else StoreStr(FormatCurrent+2, 1);
          end
          else
            raise EConvertError.Create('Illegal character in format string');
        end ;
        '/': StoreStr(@FormatSettings.DateSeparator, 1);
        ':': StoreStr(@FormatSettings.TimeSeparator, 1);
	'[': if (fdoInterval in Options) then isInterval := true else StoreStr(FormatCurrent, 1);
	']': if (fdoInterval in Options) then isInterval := false else StoreStr(FormatCurrent, 1);
        ' ', 'C', 'D', 'H', 'M', 'N', 'S', 'T', 'Y', 'Z', 'F' :
        begin
          while (P < FormatEnd) and (UpCase(P^) = Token) do
            Inc(P);
          Count := P - FormatCurrent;
          case Token of
            ' ': StoreStr(FormatCurrent, Count);
            'Y': begin
              if Count > 2 then
                StoreInt(Year, 4)
              else
                StoreInt(Year mod 100, 2);
            end;
            'M': begin
	      if isInterval and ((prevlasttoken = 'H') or TimeFlag) then
	        StoreInt(Minute + (Hour + trunc(abs(DateTime))*24)*60, 0)
	      else
              if (lastformattoken = 'H') or TimeFlag then
              begin
                if Count = 1 then
                  StoreInt(Minute, 0)
                else
                  StoreInt(Minute, 2);
              end
              else
              begin
                case Count of
                  1: StoreInt(Month, 0);
                  2: StoreInt(Month, 2);
                  3: StoreString(FormatSettings.ShortMonthNames[Month]);
                else
                  StoreString(FormatSettings.LongMonthNames[Month]);
                end;
              end;
            end;
            'D': begin
              case Count of
                1: StoreInt(Day, 0);
                2: StoreInt(Day, 2);
                3: StoreString(FormatSettings.ShortDayNames[DayOfWeek]);
                4: StoreString(FormatSettings.LongDayNames[DayOfWeek]);
                5: StoreFormat(FormatSettings.ShortDateFormat, Nesting+1, False);
              else
                StoreFormat(FormatSettings.LongDateFormat, Nesting+1, False);
              end ;
            end ;
            'H':
	      if isInterval then
	        StoreInt(Hour + trunc(abs(DateTime))*24, 0)
	      else
	      if Clock12 then
              begin
                tmp := hour mod 12;
                if tmp=0 then tmp:=12;
                if Count = 1 then
                  StoreInt(tmp, 0)
                else
                  StoreInt(tmp, 2);
              end
              else begin
                if Count = 1 then
		  StoreInt(Hour, 0)
                else
                  StoreInt(Hour, 2);
              end;
            'N': if isInterval then
	           StoreInt(Minute + (Hour + trunc(abs(DateTime))*24)*60, 0)
		 else
		 if Count = 1 then
                   StoreInt(Minute, 0)
                 else
                   StoreInt(Minute, 2);
            'S': if isInterval then
	           StoreInt(Second + (Minute + (Hour + trunc(abs(DateTime))*24)*60)*60, 0)
	         else
	         if Count = 1 then
                   StoreInt(Second, 0)
                 else
                   StoreInt(Second, 2);
            'Z': if Count = 1 then
                   StoreInt(MilliSecond, 0)
                 else
		   StoreInt(MilliSecond, 3);
            'T': if Count = 1 then
		   StoreFormat(FormatSettings.ShortTimeFormat, Nesting+1, True)
                 else
	           StoreFormat(FormatSettings.LongTimeFormat, Nesting+1, True);
            'C': begin
                   StoreFormat(FormatSettings.ShortDateFormat, Nesting+1, False);
                   if (Hour<>0) or (Minute<>0) or (Second<>0) then
                     begin
                      StoreString(' ');
                      StoreFormat(FormatSettings.LongTimeFormat, Nesting+1, True);
                     end;
                 end;
            'F': begin
                   StoreFormat(FormatSettings.ShortDateFormat, Nesting+1, False);
                   StoreString(' ');
                   StoreFormat(FormatSettings.LongTimeFormat, Nesting+1, True);
                 end;
            (* ------------ not needed here...
{$IFDEF MSWindows}
            'E':
               begin
                 if not Eastasiainited then InitEastAsia;
                 if Not(isEnable_E_Format) then StoreStr(@FormatCurrent^, 1)
                  else
                   begin
                     while (P < FormatEnd) and (UpCase(P^) = Token) do
                     P := P + 1;
                     Count := P - FormatCurrent;
                     StoreString(ConvertEraYearString(Count,Year,Month,Day));
                   end;
		 prevlasttoken := lastformattoken;
                 lastformattoken:=token;
               end;
             'G':
               begin
                 if not Eastasiainited then InitEastAsia;
                 if Not(isEnable_G_Format) then StoreStr(@FormatCurrent^, 1)
                  else
                   begin
                     while (P < FormatEnd) and (UpCase(P^) = Token) do
                     P := P + 1;
                     Count := P - FormatCurrent;
                     StoreString(ConvertEraString(Count,Year,Month,Day));
                   end;
		 prevlasttoken := lastformattoken;
                 lastformattoken:=token;
               end;
{$ENDIF MSWindows}
*)
          end;
	  prevlasttoken := lastformattoken;
          lastformattoken := token;
        end;
        else
          StoreStr(@Token, 1);
      end ;
      Inc(FormatCurrent, Count);
    end;
  end;

begin          (*
{$ifdef MSWindows}
  eastasiainited:=false;
{$endif MSWindows}
*)
  DecodeDateFully(DateTime, Year, Month, Day, DayOfWeek);
  DecodeTime(DateTime, Hour, Minute, Second, MilliSecond);
  ResultLen := 0;
  ResultCurrent := @ResultBuffer[0];
  if FormatStr <> '' then
    StoreFormat(FormatStr, 0, False)
  else
    StoreFormat('C', 0, False);
  ResultBuffer[ResultLen] := #0;
  result := StrPas(@ResultBuffer[0]);
end ;

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
procedure DateTimeToString(out Result: string; const FormatStr: string;
  const DateTime: TDateTime; Options : TFormatDateTimeOptions = []);
begin
  DateTimeToString(Result, FormatStr, DateTime, DefaultFormatSettings, Options);
end;

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  Options : TFormatDateTimeOptions = []): string;
begin
  DateTimeToString(Result, FormatStr, DateTime, DefaultFormatSettings,Options);
end;

{@@
  Applies a formatting string to a date/time value and converts the number
  to a date/time string.

  This functionality is available in the SysUtils unit. But it is duplicated
  here to add a patch which is not available in stable fpc.
}
function FormatDateTime(const FormatStr: string; DateTime: TDateTime;
  const FormatSettings: TFormatSettings; Options : TFormatDateTimeOptions = []): string;
begin
  DateTimeToString(Result, FormatStr, DateTime, FormatSettings,Options);
end;
{$ENDIF}


end.

