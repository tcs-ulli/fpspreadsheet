unit fpsclasses;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, AVL_Tree, avglvltree,
  fpstypes;

type
  { TsRowCol }
  TsRowCol = record
    Row, Col: Cardinal;
  end;
  PsRowCol = ^TsRowCol;

  { TsRowColAVLTree }
  TsRowColAVLTree = class(TAVLTree)
  private
    FOwnsData: Boolean;
  protected
    procedure DisposeData(var AData: Pointer); virtual; abstract;
    function NewData: Pointer; virtual; abstract;
  public
    constructor Create(AOwnsData: Boolean = true);
    destructor Destroy; override;
    function Add(ARow, ACol: Cardinal): PsRowCol;
    procedure Clear;
    procedure Delete(ANode: TAVLTreeNode);
    procedure DeleteRowOrCol(AIndex: Cardinal; IsRow: Boolean);
    function Find(ARow, ACol: Cardinal): PsRowCol;
    procedure InsertRowOrCol(AIndex: Cardinal; IsRow: Boolean);
    procedure Remove(ARow, ACol: Cardinal);
  end;

  { TsComments }
  TsComments = class(TsRowColAVLTree)
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    function AddComment(ARow, ACol: Cardinal; AComment: String): PsComment;
    procedure DeleteComment(ARow, ACol: Cardinal);
  end;

  { TsHyperlinks }
  TsHyperlinks = class(TsRowColAVLTree)
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    function AddHyperlink(ARow, ACol: Cardinal; ATarget: String; ATooltip: String = ''): PsHyperlink;
    procedure DeleteHyperlink(ARow, ACol: Cardinal);
  end;


implementation

uses
  fpspreadsheet;

function CompareRowCol(Item1, Item2: Pointer): Integer;
begin
  Result := LongInt(PsRowCol(Item1)^.Row) - PsRowCol(Item2)^.Row;
  if Result = 0 then
    Result := LongInt(PsRowCol(Item1)^.Col) - PsRowCol(Item2)^.Col;
end;


{******************************************************************************}
{ TsRowColAVLTree:  A specialized AVLTree working with records containing      }
{ row and column indexes.                                                      }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Constructor of the AVLTree. Installs a compare procedure for row and column
  indexes. If AOwnsData is true then the tree automatically destroys the
  data records attached to the tree nodes.
-------------------------------------------------------------------------------}
constructor TsRowColAVLTree.Create(AOwnsData: Boolean = true);
begin
  inherited Create(@CompareRowCol);
  FOwnsData := AOwnsData;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the AVLTree. Clears the tree nodes and, if the tree has been
  created with AOwnsData=true, destroys the data records
-------------------------------------------------------------------------------}
destructor TsRowColAVLTree.Destroy;
begin
  Clear;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Adds a new node to the tree identified by the specified row and column
  indexes.
-------------------------------------------------------------------------------}
function TsRowColAVLTree.Add(ARow, ACol: Cardinal): PsRowCol;
begin
  Result := Find(ARow, ACol);
  if Result = nil then
    Result := NewData;
  Result^.Row := ARow;
  Result^.Col := ACol;
  inherited Add(Result);
end;

{@@ ----------------------------------------------------------------------------
  Clears the tree, i.e, destroys the data records (if the tree has been created
  with AOwnsData = true) and removes all nodes.
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.Clear;
var
  node, nextnode: TAVLTreeNode;
begin
  node := FindLowest;
  while node <> nil do begin
    nextnode := FindSuccessor(node);
    Delete(node);
    node := nextnode;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes the specified node from the tree. If the tree has been created with
  AOwnsData = true then the data record is destroyed as well
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.Delete(ANode: TAVLTreeNode);
begin
  if FOwnsData and Assigned(ANode) then
    DisposeData(PsRowCol(ANode.Data));
  inherited Delete(ANode);
end;

{@@ ----------------------------------------------------------------------------
  This procedure adjusts row or column indexes stored in the tree nodes if a
  row or column will be deleted from the underlying worksheet.

  @param  AIndex  Index of the row (if IsRow=true) or column (if IsRow=false)
                  to be deleted
  @param  IsRow   Identifies whether AIndex refers to a row or column index
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.DeleteRowOrCol(AIndex: Cardinal; IsRow: Boolean);
var
  node, nextnode: TAVLTreeNode;
  item: PsRowCol;
begin
  node := FindLowest;
  while Assigned(node) do begin
    nextnode := FindSuccessor(node);
    item := PsRowCol(node.Data);
    if IsRow then
    begin
      // Update all RowCol records at row indexes above the deleted row
      if item^.Row > AIndex then
        dec(item^.Row)
      else
      // Remove the RowCol record if it is in the deleted row
      if item^.Row = AIndex then
        Delete(node);
    end else
    begin
      // Update all RowCol records at column indexes above the deleted column
      if item^.Col > AIndex then
        dec(item^.Col)
      else
      // Remove the RowCol record if it is in the deleted column
      if item^.Col = AIndex then
        Delete(node);
    end;
    node := nextnode;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Seeks the entire tree for a node of the specified row and column indexes and
  returns a pointer to the data record.
  Returns nil if such a node does not exist
-------------------------------------------------------------------------------}
function TsRowColAVLTree.Find(ARow, ACol: Cardinal): PsRowCol;
var
  data: TsRowCol;
  node: TAVLTreeNode;
begin
  Result := nil;
  if  (Count = 0) then
    exit;

  data.Row := ARow;
  data.Col := ACol;
  node := inherited Find(@data);
  if Assigned(node) then
    Result := PsRowCol(node.Data);
end;

{@@ ----------------------------------------------------------------------------
  This procedure adjusts row or column indexes stored in the tree nodes if a
  row or column will be inserted into the underlying worksheet.

  @param  AIndex  Index of the row (if IsRow=true) or column (if IsRow=false)
                  to be inserted
  @param  IsRow   Identifies whether AIndex refers to a row or column index
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.InsertRowOrCol(AIndex: Cardinal; IsRow: Boolean);
var
  node: TAVLTreeNode;
  item: PsRowCol;
begin
  node := FindLowest;
  while Assigned(node) do begin
    item := PsRowCol(node.Data);
    if IsRow then
    begin
      if item^.Row >= AIndex then inc(item^.Row);
    end else
    begin
      if item^.Col >= AIndex then inc(item^.Col);
    end;
    node := FindSuccessor(node);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes the node and destroys the associated data reocrd (if the tree has
  been created with AOwnsData=true) for the specified row and column indexes.
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.Remove(ARow, ACol: Cardinal);
var
  node: TAVLTreeNode;
  item: TsRowCol;
begin
  item.Row := ARow;
  item.Col := ACol;
  node := inherited Find(@item);
  Delete(node);
end;


{******************************************************************************}
{ TsComments: a AVLTree to store comment records for cells                     }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Adds a node with a new comment record to the tree. If a node already
  exists then its data will be replaced by the specified ones.
  Returns a pointer to the comment record.
-------------------------------------------------------------------------------}
function TsComments.AddComment(ARow, ACol: Cardinal;
  AComment: String): PsComment;
begin
  Result := PsComment(Add(ARow, ACol));
  Result^.Text := AComment;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for the specified row and column index along with the
  associated comment data record.
-------------------------------------------------------------------------------}
procedure TsComments.DeleteComment(ARow, ACol: Cardinal);
begin
  Remove(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Helper procedure which disposes the memory occupied by the comment data
  record attached to a tree node.
-------------------------------------------------------------------------------}
procedure TsComments.DisposeData(var AData: Pointer);
begin
  if AData <> nil then
    Dispose(PsComment(AData));
  AData := nil;
end;

{@@ ----------------------------------------------------------------------------
  Alloates memory of a comment data record.
-------------------------------------------------------------------------------}
function TsComments.NewData: Pointer;
var
  comment: PsComment;
begin
  New(comment);
  Result := comment;
end;


{******************************************************************************}
{ TsHyperlinks: a AVLTree to store hyperlink records for cells                 }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Adds a node with a new hyperlink record to the tree. If a node already
  exists then its data will be replaced by the specified ones.
  Returns a pointer to the hyperlink record.
-------------------------------------------------------------------------------}
function TsHyperlinks.AddHyperlink(ARow, ACol: Cardinal; ATarget: String;
  ATooltip: String = ''): PsHyperlink;
begin
  Result := PsHyperlink(Add(ARow, ACol));
  Result^.Target := ATarget;
  Result^.Tooltip := ATooltip;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for the specified row and column index along with the
  associated hyperlink data record.
-------------------------------------------------------------------------------}
procedure TsHyperlinks.DeleteHyperlink(ARow, ACol: Cardinal);
begin
  Remove(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Helper procedure which disposes the memory occupied by the hyperlink data
  record attached to a tree node.
-------------------------------------------------------------------------------}
procedure TsHyperlinks.DisposeData(var AData: Pointer);
begin
  if AData <> nil then
    Dispose(PsHyperlink(AData));
  AData := nil;
end;

{@@ ----------------------------------------------------------------------------
  Alloates memory of a hyperlink data record.
-------------------------------------------------------------------------------}
function TsHyperlinks.NewData: Pointer;
var
  hyperlink: PsHyperlink;
begin
  New(hyperlink);
  Result := hyperlink;
end;

end.

