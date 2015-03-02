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
    FCurrentNode: TAVLTreeNode;
  protected
    procedure DisposeData(var AData: Pointer); virtual; abstract;
    function NewData: Pointer; virtual; abstract;
  public
    constructor Create(AOwnsData: Boolean = true);
    destructor Destroy; override;
    function Add(ARow, ACol: Cardinal): PsRowCol;
    procedure Clear;
    procedure Delete(ANode: TAVLTreeNode);
    procedure DeleteRowOrCol(AIndex: Cardinal; IsRow: Boolean); virtual;
    function Find(ARow, ACol: Cardinal): PsRowCol;
    function GetFirst: PsRowCol;
    function GetNext: PsRowCol;
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

  { TsMergedCells }
  TsMergedCells = class(TsRowColAVLTree)
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    function AddRange(ARow1, ACol1, ARow2, ACol2: Cardinal): PsCellRange;
    procedure DeleteRange(ARow, ACol: Cardinal);
    procedure DeleteRowOrCol(AIndex: Cardinal; IsRow: Boolean); override;
    function FindRangeWithCell(ARow, ACol: Cardinal): PsCellRange;
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
  The combination of the methods GetFirst and GetNext allow a fast iteration
  through all nodes of the tree.
-------------------------------------------------------------------------------}
function TsRowColAVLTree.GetFirst: PsRowCol;
begin
  FCurrentNode := FindLowest;
  if FCurrentNode <> nil then
    Result := PsRowCol(FCurrentNode.Data)
  else
    Result := nil;
end;

function TsRowColAVLTree.GetNext: PsRowCol;
begin
  FCurrentNode := FindSuccessor(FCurrentNode);
  if FCurrentNode <> nil then
    Result := PsRowCol(FCurrentNode.Data)
  else
    Result := nil;
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


{******************************************************************************}
{ TsMergedCell: a AVLTree to store merged cell range records for cells         }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Adds a node with a new merge cell range record to the tree. If a node already
  exists then its data will be replaced by the specified ones.
  Returns a pointer to the cell range record.
-------------------------------------------------------------------------------}
function TsMergedCells.AddRange(ARow1, ACol1, ARow2, ACol2: Cardinal): PsCellRange;
begin
  Result := PsCellRange(Add(ARow1, ACol1));
  Result^.Row2 := ARow2;
  Result^.Col2 := ACol2;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for which the top/left corner of the cell range matches the
  specified parameters. There is only a single range fulfilling this criterion.
-------------------------------------------------------------------------------}
procedure TsMergedCells.DeleteRange(ARow, ACol: Cardinal);
begin
  Remove(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  This procedure adjusts row or column indexes stored in the tree nodes if a
  row or column will be deleted from the underlying worksheet.

  @param  AIndex  Index of the row (if IsRow=true) or column (if IsRow=false)
                  to be deleted
  @param  IsRow   Identifies whether AIndex refers to a row or column index
-------------------------------------------------------------------------------}
procedure TsMergedCells.DeleteRowOrCol(AIndex: Cardinal; IsRow: Boolean);
var
  rng, nextrng: PsCellRange;
begin
  rng := PsCellRange(GetFirst);
  while Assigned(rng) do begin
    nextrng := PsCellRange(GetNext);
    if IsRow then
    begin
      // Deleted row is above the merged range --> Shift entire range up by 1
      // NOTE:
      // The "merged" flags do not have to be changed, they move with the cells.
      if (AIndex < rng^.Row1) then begin
        dec(rng^.Row1);
        dec(rng^.Row2);
      end else
      // Single-row merged block coincides with row to be deleted
      if (AIndex = rng^.Row1) and (rng^.Row1 = rng^.Row2) then
        DeleteRange(rng^.Row1, rng^.Col1)
      else
      // Deleted row runs through the merged block --> Shift bottom row up by 1
      // NOTE: The "merged" flags disappear with the deleted cells
      if (AIndex >= rng^.Row1) and (AIndex <= rng^.Row2) then
        dec(rng^.Row2);
    end else
    begin
      // Deleted column is at the left of the merged range
      // --> Shift entire merged range to the left by 1
      // NOTE:
      // The "merged" flags do not have to be changed, they move with the cells.
      if (AIndex < rng^.Col1) then begin
        dec(rng^.Col1);
        dec(rng^.Col2);
      end else
      // Single-column block coincides with the column to be deleted
      // NOTE: The "merged" flags disappear with the deleted cells
      if (AIndex = rng^.Col1) and (rng^.Col1 = rng^.Col2) then
        DeleteRange(rng^.Row1, rng^.Col1)
      else
      // Deleted column runs through the merged block
      // --> Shift right column to the left by 1
      if (AIndex >= rng^.Col1) and (AIndex <= rng^.Col2) then
        dec(rng^.Col2);
    end;
    // Proceed with next merged range
    rng := nextrng;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Helper procedure which disposes the memory occupied by the merged cell range
  data record attached to a tree node.
-------------------------------------------------------------------------------}
procedure TsMergedCells.DisposeData(var AData: Pointer);
begin
  if AData <> nil then
    Dispose(PsCellRange(AData));
  AData := nil;
end;

{@@ ----------------------------------------------------------------------------
  Finds the cell range which contains the cell specified by its row and column
  index
-------------------------------------------------------------------------------}
function TsMergedCells.FindRangeWithCell(ARow, ACol: Cardinal): PsCellRange;
var
  node: TAVLTreeNode;
begin
  node := FindLowest;
  while Assigned(node) do
  begin
    Result := PsCellRange(node.Data);
    if (ARow >= Result^.Row1) and (ARow <= Result^.Row2) and
       (ACol >= Result^.Col1) and (ACol <= Result^.Col2) then exit;
    node := FindSuccessor(node);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Alloates memory of a merged cell range data record.
-------------------------------------------------------------------------------}
function TsMergedCells.NewData: Pointer;
var
  range: PsCellRange;
begin
  New(range);
  Result := range;
end;

end.

