unit beTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  BIFFNODE_TXO_CONTINUE1 = 1;
  BIFFNODE_TXO_CONTINUE2 = 2;

type
  { Virtual tree node data }
  TBiffNodeData = class
    Offset: Integer;
    RecordID: Integer;
    RecordName: String;
    RecordDescription: String;
    Index: Integer;
    Tag: Integer;
    destructor Destroy; override;
  end;

implementation

{ TBiffNodeData }

destructor TBiffNodeData.Destroy;
begin
  Finalize(RecordName);
  Finalize(RecordDescription);
  inherited;
end;

end.

