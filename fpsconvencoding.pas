{
 *****************************************************************************
 *                                                                           *
 *  This file is part of the Lazarus Component Library (LCL)                 *
 *                                                                           *
 *  See the file COPYING.modifiedLGPL.txt, included in this distribution,    *
 *  for details about the copyright.                                         *
 *                                                                           *
 *  This program is distributed in the hope that it will be useful,          *
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of           *
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.                     *
 *                                                                           *
 *****************************************************************************
}
unit fpsconvencoding;

{$mode objfpc}{$H+}
//{$IFDEF WINDOWS}
//{$WARNING Windows/Wine/ReactOS locale conversion is not fully supported yet. Sorry.}
//{$ENDIF}

interface

//{$IFNDEF DisableIconv}
//{$IFDEF UNIX}{$IF not defined(VER2_2_0) and not defined(VER2_2_2)}{$DEFINE HasIconvEnc}{$ENDIF}{$ENDIF}
//{$ENDIF}

uses
  SysUtils, Classes, dos
  {$IFDEF HasIconvEnc},iconvenc{$ENDIF};

const
  EncodingUTF8 = 'utf8';
  EncodingAnsi = 'ansi';
  EncodingUTF8BOM = 'utf8bom'; // UTF-8 with byte order mark
  EncodingUCS2LE = 'ucs2le'; // UCS 2 byte little endian
  EncodingUCS2BE = 'ucs2le'; // UCS 2 byte big endian

type
  TConvertEncodingFunction = function(const s: string): string;
  TCharToUTF8Table = array[char] of PChar;
  TUnicodeToCharID = function(Unicode: cardinal): integer;
var
  ConvertAnsiToUTF8: TConvertEncodingFunction = nil;
  ConvertUTF8ToAnsi: TConvertEncodingFunction = nil;

function SingleByteToUTF8(const s: string; const Table: TCharToUTF8Table): string;
function UTF8BOMToUTF8(const s: string): string; // UTF8 with BOM
function ISO_8859_1ToUTF8(const s: string): string; // central europe

function UTF8ToSingleByte(const s: string;
  const UTF8CharConvFunc: TUnicodeToCharID): string;
function UTF8ToISO_8859_1(const s: string): string; // central europe

function UTF8CharacterToUnicode(p: PChar; out CharLen: integer): Cardinal;

implementation

const
  ArrayISO_8859_1ToUTF8: TCharToUTF8Table = (
    #0,                 // #0
    #1,                 // #1
    #2,                 // #2
    #3,                 // #3
    #4,                 // #4
    #5,                 // #5
    #6,                 // #6
    #7,                 // #7
    #8,                 // #8
    #9,                 // #9
    #10,                // #10
    #11,                // #11
    #12,                // #12
    #13,                // #13
    #14,                // #14
    #15,                // #15
    #16,                // #16
    #17,                // #17
    #18,                // #18
    #19,                // #19
    #20,                // #20
    #21,                // #21
    #22,                // #22
    #23,                // #23
    #24,                // #24
    #25,                // #25
    #26,                // #26
    #27,                // #27
    #28,                // #28
    #29,                // #29
    #30,                // #30
    #31,                // #31
    ' ',                // ' '
    '!',                // '!'
    '"',                // '"'
    '#',                // '#'
    '$',                // '$'
    '%',                // '%'
    '&',                // '&'
    '''',               // ''''
    '(',                // '('
    ')',                // ')'
    '*',                // '*'
    '+',                // '+'
    ',',                // ','
    '-',                // '-'
    '.',                // '.'
    '/',                // '/'
    '0',                // '0'
    '1',                // '1'
    '2',                // '2'
    '3',                // '3'
    '4',                // '4'
    '5',                // '5'
    '6',                // '6'
    '7',                // '7'
    '8',                // '8'
    '9',                // '9'
    ':',                // ':'
    ';',                // ';'
    '<',                // '<'
    '=',                // '='
    '>',                // '>'
    '?',                // '?'
    '@',                // '@'
    'A',                // 'A'
    'B',                // 'B'
    'C',                // 'C'
    'D',                // 'D'
    'E',                // 'E'
    'F',                // 'F'
    'G',                // 'G'
    'H',                // 'H'
    'I',                // 'I'
    'J',                // 'J'
    'K',                // 'K'
    'L',                // 'L'
    'M',                // 'M'
    'N',                // 'N'
    'O',                // 'O'
    'P',                // 'P'
    'Q',                // 'Q'
    'R',                // 'R'
    'S',                // 'S'
    'T',                // 'T'
    'U',                // 'U'
    'V',                // 'V'
    'W',                // 'W'
    'X',                // 'X'
    'Y',                // 'Y'
    'Z',                // 'Z'
    '[',                // '['
    '\',                // '\'
    ']',                // ']'
    '^',                // '^'
    '_',                // '_'
    '`',                // '`'
    'a',                // 'a'
    'b',                // 'b'
    'c',                // 'c'
    'd',                // 'd'
    'e',                // 'e'
    'f',                // 'f'
    'g',                // 'g'
    'h',                // 'h'
    'i',                // 'i'
    'j',                // 'j'
    'k',                // 'k'
    'l',                // 'l'
    'm',                // 'm'
    'n',                // 'n'
    'o',                // 'o'
    'p',                // 'p'
    'q',                // 'q'
    'r',                // 'r'
    's',                // 's'
    't',                // 't'
    'u',                // 'u'
    'v',                // 'v'
    'w',                // 'w'
    'x',                // 'x'
    'y',                // 'y'
    'z',                // 'z'
    '{',                // '{'
    '|',                // '|'
    '}',                // '}'
    '~',                // '~'
    #127,               // #127
    #194#128,           // #128
    #194#129,           // #129
    #194#130,           // #130
    #194#131,           // #131
    #194#132,           // #132
    #194#133,           // #133
    #194#134,           // #134
    #194#135,           // #135
    #194#136,           // #136
    #194#137,           // #137
    #194#138,           // #138
    #194#139,           // #139
    #194#140,           // #140
    #194#141,           // #141
    #194#142,           // #142
    #194#143,           // #143
    #194#144,           // #144
    #194#145,           // #145
    #194#146,           // #146
    #194#147,           // #147
    #194#148,           // #148
    #194#149,           // #149
    #194#150,           // #150
    #194#151,           // #151
    #194#152,           // #152
    #194#153,           // #153
    #194#154,           // #154
    #194#155,           // #155
    #194#156,           // #156
    #194#157,           // #157
    #194#158,           // #158
    #194#159,           // #159
    #194#160,           // #160
    #194#161,           // #161
    #194#162,           // #162
    #194#163,           // #163
    #194#164,           // #164
    #194#165,           // #165
    #194#166,           // #166
    #194#167,           // #167
    #194#168,           // #168
    #194#169,           // #169
    #194#170,           // #170
    #194#171,           // #171
    #194#172,           // #172
    #194#173,           // #173
    #194#174,           // #174
    #194#175,           // #175
    #194#176,           // #176
    #194#177,           // #177
    #194#178,           // #178
    #194#179,           // #179
    #194#180,           // #180
    #194#181,           // #181
    #194#182,           // #182
    #194#183,           // #183
    #194#184,           // #184
    #194#185,           // #185
    #194#186,           // #186
    #194#187,           // #187
    #194#188,           // #188
    #194#189,           // #189
    #194#190,           // #190
    #194#191,           // #191
    #195#128,           // #192
    #195#129,           // #193
    #195#130,           // #194
    #195#131,           // #195
    #195#132,           // #196
    #195#133,           // #197
    #195#134,           // #198
    #195#135,           // #199
    #195#136,           // #200
    #195#137,           // #201
    #195#138,           // #202
    #195#139,           // #203
    #195#140,           // #204
    #195#141,           // #205
    #195#142,           // #206
    #195#143,           // #207
    #195#144,           // #208
    #195#145,           // #209
    #195#146,           // #210
    #195#147,           // #211
    #195#148,           // #212
    #195#149,           // #213
    #195#150,           // #214
    #195#151,           // #215
    #195#152,           // #216
    #195#153,           // #217
    #195#154,           // #218
    #195#155,           // #219
    #195#156,           // #220
    #195#157,           // #221
    #195#158,           // #222
    #195#159,           // #223
    #195#160,           // #224
    #195#161,           // #225
    #195#162,           // #226
    #195#163,           // #227
    #195#164,           // #228
    #195#165,           // #229
    #195#166,           // #230
    #195#167,           // #231
    #195#168,           // #232
    #195#169,           // #233
    #195#170,           // #234
    #195#171,           // #235
    #195#172,           // #236
    #195#173,           // #237
    #195#174,           // #238
    #195#175,           // #239
    #195#176,           // #240
    #195#177,           // #241
    #195#178,           // #242
    #195#179,           // #243
    #195#180,           // #244
    #195#181,           // #245
    #195#182,           // #246
    #195#183,           // #247
    #195#184,           // #248
    #195#185,           // #249
    #195#186,           // #250
    #195#187,           // #251
    #195#188,           // #252
    #195#189,           // #253
    #195#190,           // #254
    #195#191            // #255
  );

function UTF8BOMToUTF8(const s: string): string;
begin
  Result:=copy(s,4,length(s));
end;

function ISO_8859_1ToUTF8(const s: string): string;
begin
  Result:=SingleByteToUTF8(s,ArrayISO_8859_1ToUTF8);
end;

function SingleByteToUTF8(const s: string; const Table: TCharToUTF8Table
  ): string;
var
  len: Integer;
  i: Integer;
  Src: PChar;
  Dest: PChar;
  p: PChar;
  c: Char;
begin
  if s='' then begin
    Result:=s;
    exit;
  end;
  len:=length(s);
  SetLength(Result,len*4);// UTF-8 is at most 4 bytes
  Src:=PChar(s);
  Dest:=PChar(Result);
  for i:=1 to len do begin
    c:=Src^;
    inc(Src);
    if ord(c)<128 then begin
      Dest^:=c;
      inc(Dest);
    end else begin
      p:=Table[c];
      if p<>nil then begin
        while p^<>#0 do begin
          Dest^:=p^;
          inc(p);
          inc(Dest);
        end;
      end;
    end;
  end;
  SetLength(Result,PtrUInt(Dest)-PtrUInt(Result));
end;

function UnicodeToISO_8859_1(Unicode: cardinal): integer;
begin
  case Unicode of
  0..255: Result:=Unicode;
  else Result:=-1;
  end;
end;

function UTF8ToUTF8BOM(const s: string): string;
begin
  Result:=#$EF#$BB#$BF+s;
end;

function UTF8ToISO_8859_1(const s: string): string;
begin
  Result:=UTF8ToSingleByte(s,@UnicodeToISO_8859_1);
end;


function UTF8ToSingleByte(const s: string;
  const UTF8CharConvFunc: TUnicodeToCharID): string;
var
  len: Integer;
  Src: PChar;
  Dest: PChar;
  c: Char;
  Unicode: LongWord;
  CharLen: integer;
  i: integer;
begin
  if s='' then begin
    Result:='';
    exit;
  end;
  len:=length(s);
  SetLength(Result,len);
  Src:=PChar(s);
  Dest:=PChar(Result);
  while len>0 do begin
    c:=Src^;
    if c<#128 then begin
      Dest^:=c;
      inc(Dest);
      inc(Src);
      dec(len);
    end else begin
      Unicode:=UTF8CharacterToUnicode(Src,CharLen);
      inc(Src,CharLen);
      dec(len,CharLen);
      i:=UTF8CharConvFunc(Unicode);
      if i>=0 then begin
        Dest^:=chr(i);
        inc(Dest);
      end;
    end;
  end;
  SetLength(Result,Dest-PChar(Result));
end;


procedure GetSupportedEncodings(List: TStrings);
begin
  List.Add('UTF-8');
  List.Add('UTF-8BOM');
  List.Add('Ansi');
  List.Add('CP1250');
  List.Add('CP1251');
  List.Add('CP1252');
  List.Add('CP1253');
  List.Add('CP1254');
  List.Add('CP1255');
  List.Add('CP1256');
  List.Add('CP1257');
  List.Add('CP1258');
  List.Add('CP866');
  List.Add('CP874');
  List.Add('ISO-8859-1');
  List.Add('KOI-8');
  List.Add('UCS-2LE');
  List.Add('UCS-2BE');
end;

function UTF8CharacterToUnicode(p: PChar; out CharLen: integer): Cardinal;
begin
  if p<>nil then begin
    if ord(p^)<%11000000 then begin
      // regular single byte character (#0 is a normal char, this is pascal ;)
      Result:=ord(p^);
      CharLen:=1;
    end
    else if ((ord(p^) and %11100000) = %11000000) then begin
      // could be double byte character
      if (ord(p[1]) and %11000000) = %10000000 then begin
        Result:=((ord(p^) and %00011111) shl 6)
                or (ord(p[1]) and %00111111);
        CharLen:=2;
      end else begin
        Result:=ord(p^);
        CharLen:=1;
      end;
    end
    else if ((ord(p^) and %11110000) = %11100000) then begin
      // could be triple byte character
      if ((ord(p[1]) and %11000000) = %10000000)
      and ((ord(p[2]) and %11000000) = %10000000) then begin
        Result:=((ord(p^) and %00011111) shl 12)
                or ((ord(p[1]) and %00111111) shl 6)
                or (ord(p[2]) and %00111111);
        CharLen:=3;
      end else begin
        Result:=ord(p^);
        CharLen:=1;
      end;
    end
    else if ((ord(p^) and %11111000) = %11110000) then begin
      // could be 4 byte character
      if ((ord(p[1]) and %11000000) = %10000000)
      and ((ord(p[2]) and %11000000) = %10000000)
      and ((ord(p[3]) and %11000000) = %10000000) then begin
        Result:=((ord(p^) and %00001111) shl 18)
                or ((ord(p[1]) and %00111111) shl 12)
                or ((ord(p[2]) and %00111111) shl 6)
                or (ord(p[3]) and %00111111);
        CharLen:=4;
      end else begin
        Result:=ord(p^);
        CharLen:=1;
      end;
    end
    else begin
      // invalid character
      Result:=ord(p^);
      CharLen:=1;
    end;
  end else begin
    Result:=0;
    CharLen:=0;
  end;
end;

end.
