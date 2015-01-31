{ fpsxmlcommon.pas
  Unit shared by all xml-type reader/writer classes }

unit fpsxmlcommon;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  laz2_xmlread, laz2_DOM,
  fpspreadsheet;

type
  TsSpreadXMLReader = class(TsCustomSpreadReader)
  protected
    function GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
    function GetNodeValue(ANode: TDOMNode): String;
    procedure ReadXMLFile(out ADoc: TXMLDocument; AFileName: String);
  end;

procedure UnzipFile(AZipFileName, AZippedFile, ADestFolder: String);


implementation

uses
  fpsStreams, fpsZipper;

{ Gets value for the specified attribute. Returns empty string if attribute
  not found. }
function TsSpreadXMLReader.GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
var
  i: LongWord;
  Found: Boolean;
begin
  Result := '';
  if ANode = nil then
    exit;

  Found := false;
  i := 0;
  while not Found and (i < ANode.Attributes.Length) do begin
    if ANode.Attributes.Item[i].NodeName = AAttrName then begin
      Found := true;
      Result := ANode.Attributes.Item[i].NodeValue;
    end;
    inc(i);
  end;
end;

{ Returns the text value of a node. Normally it would be sufficient to call
  "ANode.NodeValue", but since the DOMParser needs to preserve white space
  (for the spaces in date/time formats), we have to go more into detail. }
function TsSpreadXMLReader.GetNodeValue(ANode: TDOMNode): String;
var
  child: TDOMNode;
begin
  Result := '';
  child := ANode.FirstChild;
  if Assigned(child) and (child.NodeName = '#text') then
    Result := child.NodeValue;
end;

{ We have to use our own ReadXMLFile procedure (there is one in xmlread)
  because we have to preserve spaces in element text for date/time separator.
  As a side-effect we have to skip leading spaces by ourselves. }
procedure TsSpreadXMLReader.ReadXMLFile(out ADoc: TXMLDocument; AFileName: String);
var
  parser: TDOMParser;
  src: TXMLInputSource;
  stream: TStream;
begin
  if (boBufStream in Workbook.Options) then
    stream := TBufStream.Create(AFileName, fmOpenRead + fmShareDenyWrite)
  else
    stream := TFileStream.Create(AFileName, fmOpenRead + fmShareDenyWrite);

  try
    parser := TDOMParser.Create;
    try
      parser.Options.PreserveWhiteSpace := true;    // This preserves spaces!
      src := TXMLInputSource.Create(stream);
      try
        parser.Parse(src, ADoc);
      finally
        src.Free;
      end;
    finally
      parser.Free;
    end;
  finally
    stream.Free;
  end;
end;

procedure UnzipFile(AZipFileName, AZippedFile, ADestFolder: String);
var
  list: TStringList;
  unzip: TUnzipper;
begin
  list := TStringList.Create;
  try
    list.Add(AZippedFile);
    unzip := TUnzipper.Create;
    try
      Unzip.OutputPath := ADestFolder;
      Unzip.UnzipFiles(AZipFileName, list);
    finally
      unzip.Free;
    end;
  finally
    list.Free;
  end;
end;


end.

