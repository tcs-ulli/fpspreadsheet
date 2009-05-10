{
fpolestorage.pas

Writes an OLE document using the OLE virtual layer.

Note: Compatibility with previous version (fpolestorage.pas).
}
unit fpolebasic;

interface

uses
  Classes, SysUtils,
  uvirtuallayer_ole;

type

  { Describes an OLE Document }

  TOLEDocument = record
    // Information about the document
    Stream: TMemoryStream;
  end;


  { TOLEStorage }

  TOLEStorage = class
  private
  public
    constructor Create;
    destructor Destroy; override;
    procedure WriteOLEFile(AFileName: string; AOLEDocument: TOLEDocument; const AStreamName: UTF8String='Book');
    procedure ReadOLEFile(AFileName: string; AOLEDocument: TOLEDocument; const AStreamName: UTF8String='Book');
    procedure FreeOLEDocumentData(AOLEDocument: TOLEDocument);
  end;

implementation

constructor TOLEStorage.Create;
begin
  inherited Create;
end;

destructor TOLEStorage.Destroy;
begin
  inherited Destroy;
end;

{@@
  Writes the OLE document specified in AOLEDocument
  to the file with name AFileName. The routine will fail
  if the file already exists, or if the directory where
  it should be placed doesn't exist.
}
procedure TOLEStorage.WriteOLEFile(AFileName: string;
  AOLEDocument: TOLEDocument; const AStreamName: UTF8String);
var
  RealFile: TFileStream;
  fsOLE: TVirtualLayer_OLE;
  OLEStream: TStream;
  VLAbsolutePath: UTF8String;
begin
  VLAbsolutePath:='/'+AStreamName; //Virtual layer always use absolute paths.
  RealFile:=TFileStream.Create(AFileName,fmCreate);
  fsOLE:=TVirtualLayer_OLE.Create(RealFile);
  fsOLE.Format(); //Initialize and format the OLE container.
  OLEStream:=fsOLE.CreateStream(VLAbsolutePath,fmCreate);
  AOLEDocument.Stream.Position:=0; //Ensures it is in the begining.
  OLEStream.CopyFrom(AOLEDocument.Stream,AOLEDocument.Stream.Size);
  OLEStream.Free;
  fsOLE.Free;
  RealFile.Free;
end;

{@@
  Reads an OLE file.
}
procedure TOLEStorage.ReadOLEFile(AFileName: string;
  AOLEDocument: TOLEDocument; const AStreamName: UTF8String);
var
  RealFile: TFileStream;
  fsOLE: TVirtualLayer_OLE;
  OLEStream: TStream;
  VLAbsolutePath: UTF8String;
begin
  VLAbsolutePath:='/'+AStreamName; //Virtual layer always use absolute paths.
  RealFile:=TFileStream.Create(AFileName,fmOpenRead);
  fsOLE:=TVirtualLayer_OLE.Create(RealFile);
  fsOLE.Initialize(); //Initialize the OLE container.
  OLEStream:=fsOLE.CreateStream(VLAbsolutePath,fmOpenRead);
  if not Assigned(AOLEDocument.Stream) then begin
    AOLEDocument.Stream:=TMemoryStream.Create;
  end else begin
    AOLEDocument.Stream.Clear;
  end;
  AOLEDocument.Stream.CopyFrom(OLEStream,OLEStream.Size);
  OLEStream.Free;
  fsOLE.Free;
  RealFile.Free;
end;

{@@
  Frees all internal objects storable in a TOLEDocument structure
}
procedure TOLEStorage.FreeOLEDocumentData(AOLEDocument: TOLEDocument);
begin
  if Assigned(AOLEDocument.Stream) then FreeAndNil(AOLEDocument.Stream);
end;

end.

