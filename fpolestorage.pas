{
fpolestorage.pas

Writes an OLE document

AUTHORS: Felipe Monteiro de Carvalho
}
unit fpolestorage;

{$ifdef fpc}
{$mode delphi}
{$endif}

interface

{$ifdef Windows}
  {$define FPOLESTORAGE_USE_COM}
{$endif}

uses
{$ifdef FPOLESTORAGE_USE_COM}
  ActiveX, ComObj,
{$endif}
  Classes, SysUtils,
  fpsutils;

type

  { Describes an OLE Document }

  TOLEDocument = record
    Sections: array of TMemoryStream;
  end;


  { TOLEStorage }

  TOLEStorage = class
  private
{$ifdef FPOLESTORAGE_USE_COM}
    FStorage: IStorage;
    FStream: IStream;
{$endif}
    { Information filled by the write routines for the helper routines }
    FOLEDocument: TOLEDocument;
    FNumSectors: Cardinal;
    { Helper routines }
    procedure WriteOLEHeader(AStream: TStream);
    procedure WriteSectorAllocationTable(AStream: TStream);
  public
    constructor Create;
    destructor Destroy; override;
    procedure WriteOLEFile(AFileName: string; AOLEDocument: TOLEDocument);
  end;

implementation

{ TOLEStorage }

{
4.1 Compound Document Header Contents
The header is always located at the beginning of the file, and its size is exactly 512 bytes. This implies that the first
sector (with SecID 0) always starts at file offset 512.
}
procedure TOLEStorage.WriteOLEHeader(AStream: TStream);
var
  i: Integer;
begin
  {
  Contents of the compound document header structure:
  Offset Size Contents
  0 8 Compound document file identifier: D0H CFH 11H E0H A1H B1H 1AH E1H
  }
  AStream.WriteByte($D0);
  AStream.WriteByte($CF);
  AStream.WriteByte($11);
  AStream.WriteByte($E0);
  AStream.WriteByte($A1);
  AStream.WriteByte($B1);
  AStream.WriteByte($1A);
  AStream.WriteByte($E1);

  { 8 16 Unique identifier (UID) of this file (not of interest in the following, may be all 0) }
  AStream.WriteDWord(0);
  AStream.WriteDWord(0);

  { 24 2 Revision number of the file format (most used is 003EH) }
  AStream.WriteWord(WordToLE($003E));

  { 26 2 Version number of the file format (most used is 0003H) }
  AStream.WriteWord(WordToLE($0003));

  { 28 2 Byte order identifier (➜4.2): FEH FFH = Little-Endian
    FFH FEH = Big-Endian }
  AStream.WriteByte($FE);
  AStream.WriteByte($FF);

  { 30 2 Size of a sector in the compound document file (➜3.1) in power-of-two (ssz), real sector
    size is sec_size = 2ssz bytes (minimum value is 7 which means 128 bytes, most used
    value is 9 which means 512 bytes) }
  AStream.WriteWord(WordToLE($0009));

  { 32 2 Size of a short-sector in the short-stream container stream (➜6.1) in power-of-two (sssz),
    real short-sector size is short_sec_size = 2sssz bytes (maximum value is sector size
    ssz, see above, most used value is 6 which means 64 bytes) }
  AStream.WriteWord(WordToLE($0006));

  { 34 10 Not used }
  AStream.WriteDWord($0);
  AStream.WriteDWord($0);
  AStream.WriteWord($0);

  { 44 4 Total number of sectors used for the sector allocation table (➜5.2) }
  AStream.WriteDWord(DWordToLE(FNumSectors));

  { 48 4 SecID of first sector of the directory stream (➜7) }
  AStream.WriteDWord(DWordToLE($01));

  { 52 4 Not used }
  AStream.WriteDWord($0);

  { 56 4 Minimum size of a standard stream (in bytes, minimum allowed and most used size is 4096
    bytes), streams with an actual size smaller than (and not equal to) this value are stored as
    short-streams (➜6) }
  AStream.WriteDWord(DWordToLE(4096));

  { 60 4 SecID of first sector of the short-sector allocation table (➜6.2), or –2 (End Of Chain
    SecID, ➜3.1) if not extant }
  AStream.WriteDWord(DWordToLE(2));

  { 64 4 Total number of sectors used for the short-sector allocation table (➜6.2) }
  AStream.WriteDWord(DWordToLE(1));

  { 68 4 SecID of first sector of the master sector allocation table (➜5.1), or –2 (End Of Chain
    SecID, ➜3.1) if no additional sectors used }
  AStream.WriteDWord(IntegerToLE(-2));

  { 72 4 Total number of sectors used for the master sector allocation table (➜5.1) }
  AStream.WriteDWord(0);

  { 76 436 First part of the master sector allocation table (➜5.1) containing 109 SecIDs }
  AStream.WriteDWord(0);

  for i := 1 to 108 do AStream.WriteDWord($FFFFFFFF);
end;

procedure TOLEStorage.WriteSectorAllocationTable(AStream: TStream);
var
  i: Integer;
begin
  { Simple copy of an example OLE file

   00000200H  FD FF FF FF FF FF FF FF FE FF FF FF 04 00 00 00
   00000210H  05 00 00 00 06 00 00 00 07 00 00 00 08 00 00 00
   00000220H  09 00 00 00 FE FF FF FF 0B 00 00 00 FE FF FF FF

   And from now on only $FFFFFFFF covering $230 to $3FF
   for a total of $400 - $230 bytes of $FF }

  AStream.WriteDWord(DWordToLE($FFFFFFFD));
  AStream.WriteDWord($FFFFFFFF);
  AStream.WriteDWord(DWordToLE($FFFFFFFE));
  AStream.WriteDWord(DWordToLE($00000004));
  AStream.WriteDWord(DWordToLE($00000005));
  AStream.WriteDWord(DWordToLE($00000006));
  AStream.WriteDWord(DWordToLE($00000007));
  AStream.WriteDWord(DWordToLE($00000008));
  AStream.WriteDWord(DWordToLE($00000009));
  AStream.WriteDWord(DWordToLE($FFFFFFFE));
  AStream.WriteDWord(DWordToLE($0000000B));
  AStream.WriteDWord(DWordToLE($FFFFFFFE));

  for i := 1 to ($400 - $230) do AStream.WriteByte($FF);

  {
  This results in the following SecID array for the SAT:

  Array indexes 0  1  2 3 4 5 6 7 8  9 10 11 12 ...
  SecID array  –3 –1 –2 4 5 6 7 8 9 –2 11 –2 –1 ...

  As expected, sector 0 is marked with the special SAT SecID (➜3.1). Sector 1 and all sectors starting with sector 12 are
  not used (special Free SecID with value –1). }
end;

constructor TOLEStorage.Create;
begin
  inherited Create;

end;

destructor TOLEStorage.Destroy;
begin

  inherited Destroy;
end;

procedure TOLEStorage.WriteOLEFile(AFileName: string; AOLEDocument: TOLEDocument);
var
  cbWritten: Cardinal;
  AFileStream: TFileStream;
  i: Cardinal;
begin
  { Fill information for helper routines }
  FOLEDocument := AOLEDocument;
  FNumSectors := Length(AOLEDocument.Sections);

{$ifdef FPOLESTORAGE_USE_COM}
  { Initialize the Component Object Model (COM) before calling s functions }
  OleCheck(CoInitialize(nil));

  { Create a Storage Object }
  OleCheck(StgCreateDocfile(PWideChar(WideString(AFileName)),
   STGM_READWRITE or STGM_FAILIFTHERE or STGM_SHARE_EXCLUSIVE or STGM_DIRECT,
   0, FStorage));

  for i := 0 to FNumSectors do
  begin
    { Create a workbook stream in the storage.  A BIFF5 file must
      have at least a workbook stream.  This stream *must* be named 'Book' }
    OleCheck(FStorage.CreateStream('Book',
     STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_DIRECT, 0, 0, FStream));

    { Write all data }
    FStream.Write(FOLEDocument.Sections[i].Memory,
      FOLEDocument.Sections[i].Size, @cbWritten);
  end;
{$else}
  AFileStream := TFileStream.Create(AFileName, fmOpenWrite or fmCreate);
  try
    WriteOLEHeader(AFileStream);
  finally
    AFileStream.Free;
  end;
{$endif}
end;

end.

