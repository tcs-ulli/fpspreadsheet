unit fpsxmlcommon;

{$mode objfpc}

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

implementation

uses
  fpsStreams;

{ Gets value for the specified attribute. Returns empty string if attribute
  not found. }
function TsSpreadXMLReader.GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
var
  i: integer;
  Found: Boolean;
begin
  Found := false;
  i := 0;
  Result := '';
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


end.

