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

uses
{$ifdef Windows}
  ActiveX, ComObj,
{$endif}
  Classes, SysUtils;

type

  { TOLEStorage }

  TOLEStorage = class
  private
{$ifdef Windows}
    FStorage: IStorage;
    FStream: IStream;
{$endif}
  public
    constructor Create;
    destructor Destroy; override;
    procedure WriteStreamToOLEFile(AFileName: string; AMemStream: TMemoryStream);
  end;

implementation

{ TOLEStorage }

constructor TOLEStorage.Create;
begin
  inherited Create;

end;

destructor TOLEStorage.Destroy;
begin

  inherited Destroy;
end;

procedure TOLEStorage.WriteStreamToOLEFile(AFileName: string; AMemStream: TMemoryStream);
var
  cbWritten: Cardinal;
begin
{$ifdef Windows}
  { Initialize the Component Object Model (COM) before calling s functions }
  OleCheck(CoInitialize(nil));

  { Create a Storage Object }
  OleCheck(StgCreateDocfile(PWideChar(WideString(AFileName)),
   STGM_READWRITE or STGM_FAILIFTHERE or STGM_SHARE_EXCLUSIVE or STGM_DIRECT,
   0, FStorage));

  { Create a workbook stream in the storage.  A BIFF5 file must
    have at least a workbook stream.  This stream *must* be named 'Book' }
  OleCheck(FStorage.CreateStream('Book',
   STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_DIRECT, 0, 0, FStream));

  { Write all data }
  FStream.Write(AMemStream.Memory, AMemStream.Size, @cbWritten);
{$endif}
end;

end.

