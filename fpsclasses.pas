unit fpsclasses;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, AVL_Tree, avglvltree,
  fpstypes;

type
  { forward declarations }
  TsRowColAVLTree = class;

  { TsRowCol }
  TsRowCol = record
    Row, Col: LongInt;
  end;
  PsRowCol = ^TsRowCol;

  { TAVLTreeNodeStack }
  TAVLTreeNodeStack = class(TFPList)
  public
    procedure Push(ANode: TAVLTreeNode);
    function Pop: TAVLTreeNode;
  end;

  { TsRowColEnumerator }
  TsRowColEnumerator = class
  protected
    FCurrentNode: TAVLTreeNode;
    FTree: TsRowColAVLTree;
    FStartRow, FEndRow, FStartCol, FEndCol: LongInt;
    FReverse: Boolean;
    function GetCurrent: PsRowCol;
  public
    constructor Create(ATree: TsRowColAVLTree;
      AStartRow, AStartCol, AEndRow, AEndCol: LongInt; AReverse: Boolean);
    function GetEnumerator: TsRowColEnumerator; inline;
    function MoveNext: Boolean;
    property Current: PsRowCol read GetCurrent;
    property StartRow: LongInt read FStartRow;
    property EndRow: LongInt read FEndRow;
    property StartCol: LongInt read FStartCol;
    property EndCol: LongInt read FEndCol;
  end;

  { TsRowColAVLTree }
  TsRowColAVLTree = class(TAVLTree)
  private
    FOwnsData: Boolean;
    FCurrentNode: TAVLTreeNode;
    FCurrentNodeStack: TAVLTreeNodeStack;
  protected
    procedure DisposeData(var AData: Pointer); virtual; abstract;
    function NewData: Pointer; virtual; abstract;
  public
    constructor Create(AOwnsData: Boolean = true);
    destructor Destroy; override;
    function Add(ARow, ACol: LongInt): PsRowCol;
    procedure Clear;
    procedure Delete(ANode: TAVLTreeNode); overload;
    procedure Delete(ARow, ACol: LongInt); overload;
    procedure DeleteRowOrCol(AIndex: LongInt; IsRow: Boolean); virtual;
    procedure Exchange(ARow1, ACol1, ARow2, ACol2: LongInt); virtual;
    function Find(ARow, ACol: LongInt): PsRowCol; overload;
    function GetData(ANode: TAVLTreeNode): PsRowCol;
    function GetFirst: PsRowCol;
    function GetLast: PsRowCol;
    function GetNext: PsRowCol;
    function GetPrev: PsRowCol;
    procedure InsertRowOrCol(AIndex: LongInt; IsRow: Boolean);
    procedure Remove(ARow, ACol: LongInt); overload;
    procedure PushCurrent;
    procedure PopCurrent;
  end;

  { TsCells }
  TsCellEnumerator = class(TsRowColEnumerator)
  protected
    function GetCurrent: PCell;
  public
    function GetEnumerator: TsCellEnumerator; inline;
    property Current: PCell read GetCurrent;
  end;

  TsCells = class(TsRowColAVLTree)
  private
    FWorksheet: Pointer;  // Must be cast to TsWorksheet
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    constructor Create(AWorksheet: Pointer; AOwnsData: Boolean = true);
    function AddCell(ARow, ACol: LongInt): PCell;
    procedure DeleteCell(ARow, ACol: LongInt);
    function FindCell(ARow, ACol: LongInt): PCell;
    function GetFirstCell: PCell;
    function GetLastCell: PCell;
    function GetNextCell: PCell;
    function GetPrevCell: PCell;
    // enumerators
    function GetEnumerator: TsCellEnumerator;
    function GetReverseEnumerator: TsCellEnumerator;
    function GetColEnumerator(ACol: LongInt; AStartRow: Longint = 0;
      AEndRow: Longint = $7FFFFFFF): TsCellEnumerator;
    function GetRangeEnumerator(AStartRow, AStartCol,
      AEndRow, AEndCol: Longint): TsCellEnumerator;
    function GetRowEnumerator(ARow: LongInt; AStartCol:LongInt = 0;
      AEndCol: Longint = $7FFFFFFF): TsCellEnumerator;
    function GetReverseColEnumerator(ACol: LongInt; AStartRow: Longint = 0;
      AEndRow: Longint = $7FFFFFFF): TsCellEnumerator;
    function GetReverseRangeEnumerator(AStartRow, AStartCol,
      AEndRow, AEndCol: Longint): TsCellEnumerator;
    function GetReverseRowEnumerator(ARow: LongInt; AStartCol:LongInt = 0;
      AEndCol: Longint = $7FFFFFFF): TsCellEnumerator;
  end;

  { TsComments }
  TsCommentEnumerator = class(TsRowColEnumerator)
  protected
    function GetCurrent: PsComment;
  public
    function GetEnumerator: TsCommentEnumerator; inline;
    property Current: PsComment read GetCurrent;
  end;

  TsComments = class(TsRowColAVLTree)
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    function AddComment(ARow, ACol: Cardinal; AComment: String): PsComment;
    procedure DeleteComment(ARow, ACol: Cardinal);
    // enumerators
    function GetEnumerator: TsCommentEnumerator;
    function GetRangeEnumerator(AStartRow, AStartCol,
      AEndRow, AEndCol: Longint): TsCommentEnumerator;
  end;

  { TsHyperlinks }
  TsHyperlinkEnumerator = class(TsRowColEnumerator)
  protected
    function GetCurrent: PsHyperlink;
  public
    function GetEnumerator: TsHyperlinkEnumerator; inline;
    property Current: PsHyperlink read GetCurrent;
  end;

  TsHyperlinks = class(TsRowColAVLTree)
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    function AddHyperlink(ARow, ACol: Longint; ATarget: String;
      ATooltip: String = ''): PsHyperlink;
    procedure DeleteHyperlink(ARow, ACol: Longint);
    // enumerators
    function GetEnumerator: TsHyperlinkEnumerator;
    function GetRangeEnumerator(AStartRow, AStartCol,
      AEndRow, AEndCol: Longint): TsHyperlinkEnumerator;
  end;

  { TsMergedCells }
  TsMergedCells = class(TsRowColAVLTree)
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    function AddRange(ARow1, ACol1, ARow2, ACol2: Longint): PsCellRange;
    procedure DeleteRange(ARow, ACol: Longint);
    procedure DeleteRowOrCol(AIndex: Longint; IsRow: Boolean); override;
    procedure Exchange(ARow1, ACol1, ARow2, ACol2: Longint); override;
    function FindRangeWithCell(ARow, ACol: Longint): PsCellRange;
  end;

implementation

uses
  Math, fpsUtils;

function CompareRowCol(Item1, Item2: Pointer): Integer;
begin
  Result := LongInt(PsRowCol(Item1)^.Row) - PsRowCol(Item2)^.Row;
  if Result = 0 then
    Result := LongInt(PsRowCol(Item1)^.Col) - PsRowCol(Item2)^.Col;
end;


function TAVLTreeNodeStack.Pop: TAVLTreeNode;
begin
  Result := TAVLTreeNode(Items[Count-1]);
  Delete(Count-1);
end;

procedure TAVLTreeNodeStack.Push(ANode: TAVLTreeNode);
begin
  Add(ANode);
end;


{******************************************************************************}
{ TsRowColEnumerator:  A specialized enumerator for TsRowColAVLTree using the  }
{ pointers to the data records.                                                }
{******************************************************************************}

constructor TsRowColEnumerator.Create(ATree: TsRowColAVLTree;
  AStartRow, AStartCol, AEndRow, AEndCol: LongInt; AReverse: Boolean);
begin
  FTree := ATree;
  FReverse := AReverse;
  // Rearrange col/row indexes such that iteration always begins with "StartXXX"
  if AStartRow <= AEndRow then
  begin
    FStartRow := IfThen(AReverse, AEndRow, AStartRow);
    FEndRow := IfThen(AReverse, AStartRow, AEndRow);
  end else
  begin
    FStartRow := IfThen(AReverse, AStartRow, AEndRow);
    FEndRow := IfThen(AReverse, AEndRow, AStartRow);
  end;
  if AStartCol <= AEndCol then
  begin
    FStartCol := IfThen(AReverse, AEndCol, AStartCol);
    FEndCol := IfThen(AReverse, AStartCol, AEndCol);
  end else
  begin
    FStartCol := IfThen(AReverse, AStartCol, AEndCol);
    FEndCol := IfThen(AReverse, AEndCol, AStartCol);
  end;
end;

function TsRowColEnumerator.GetCurrent: PsRowCol;
begin
  if Assigned(FCurrentNode) then
    Result := PsRowCol(FCurrentNode.Data)
  else
    Result := nil;
end;

function TsRowColEnumerator.GetEnumerator: TsRowColEnumerator;
begin
  Result := self;
end;

function TsRowColEnumerator.MoveNext: Boolean;
var
  r1,c1,r2,c2: LongInt;
  item: TsRowCol;
begin
  if FCurrentNode <> nil then begin
    if FReverse then
    begin
      FCurrentNode := FTree.FindPrecessor(FCurrentNode);
      while (FCurrentNode <> nil) and
          ( (Current^.Col < FEndCol) or (Current^.Col > FStartCol) or
            (Current^.Row < FEndRow) or (Current^.Row > FStartRow) )
      do
        FCurrentNode := FTree.FindPrecessor(FCurrentNode);
    end else
    begin
      FCurrentNode := FTree.FindSuccessor(FCurrentNode);
      while (FCurrentNode <> nil) and
          ( (Current^.Col < FStartCol) or (Current^.Col > FEndCol) or
            (Current^.Row < FStartRow) or (Current^.Row > FEndRow) )
      do
        FCurrentNode := FTree.FindSuccessor(FCurrentNode);
    end;
  end else
  begin
    if FReverse and (FStartRow = $7FFFFFFF) and (FStartCol = $7FFFFFFF) then
      FCurrentNode := FTree.FindHighest
    else
    if not FReverse and (FStartRow = 0) and (FStartCol = 0) then
      FCurrentNode := FTree.FindLowest
    else
    begin
      item.Row := FStartRow;
      item.Col := FStartCol;
      FCurrentNode := FTree.Find(@item);
    end;
  end;
  Result := FCurrentNode <> nil;
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
  FCurrentNodeStack := TAVLTreeNodeStack.Create;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the AVLTree. Clears the tree nodes and, if the tree has been
  created with AOwnsData=true, destroys the data records
-------------------------------------------------------------------------------}
destructor TsRowColAVLTree.Destroy;
begin
  FCurrentNodeStack.Free;
  Clear;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Adds a new node to the tree identified by the specified row and column
  indexes.
-------------------------------------------------------------------------------}
function TsRowColAVLTree.Add(ARow, ACol: LongInt): PsRowCol;
begin
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

procedure TsRowColAVLTree.Delete(ARow, ACol: LongInt);
var
  node: TAVLTreeNode;
  cell: TCell;
begin
  cell.Row := ARow;
  cell.Col := ACol;
  node := inherited Find(@cell);
  if Assigned(node) then
    Delete(node);
end;

{@@ ----------------------------------------------------------------------------
  This procedure adjusts row or column indexes stored in the tree nodes if a
  row or column will be deleted from the underlying worksheet.

  @param  AIndex  Index of the row (if IsRow=true) or column (if IsRow=false)
                  to be deleted
  @param  IsRow   Identifies whether AIndex refers to a row or column index
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.DeleteRowOrCol(AIndex: LongInt; IsRow: Boolean);
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
  Exchanges two nodes
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.Exchange(ARow1, ACol1, ARow2, ACol2: LongInt);
var
  item1, item2: PsRowCol;
begin
  item1 := Find(ARow1, ACol1);
  item2 := Find(ARow2, ACol2);

  // There are entries for both locations: Exchange row/col indexes
  if (item1 <> nil) and (item2 <> nil) then
  begin
    Remove(item1);
    Remove(item2);
    item1^.Row := ARow2;
    item1^.Col := ACol2;
    item2^.Row := ARow1;
    item2^.Col := ACol1;
    inherited Add(item1);
    inherited Add(item2);
  end else

  // Only the 1tst item exists --> give it the row/col indexes of the 2nd item
  if (item1 <> nil) then
  begin
    Remove(item1);
    item1^.Row := ARow2;
    item1^.Col := ACol2;
    inherited Add(item1);
  end else

  // Only the 2nd item exists --> give it the row/col indexes of the 1st item
  if (item2 <> nil) then
  begin
    Remove(item2);
    item2^.Row := ARow1;
    item2^.Col := ACol1;
    inherited Add(item2);  // just adds the existing item at the new position
  end;
end;

{@@ ----------------------------------------------------------------------------
  Seeks the entire tree for a node of the specified row and column indexes and
  returns a pointer to the data record.
  Returns nil if such a node does not exist
-------------------------------------------------------------------------------}
function TsRowColAVLTree.Find(ARow, ACol: LongInt): PsRowCol;
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
  Extracts the pointer to the data record from a tree node
-------------------------------------------------------------------------------}
function TsRowColAVLTree.GetData(ANode: TAVLTreeNode): PsRowCol;
begin
  if ANode <> nil then
    Result := PsRowCol(ANode.Data)
  else
    Result := nil;
end;
                          (*
function TsRowColAVLTree.GetEnumerator: TsRowColEnumerator;
begin
  Result := TsRowColEnumerator.Create(self);
end;

function TsRowColAVLTree.GetColEnumerator(ACol: LongInt): TsRowColEnumerator;
begin
  Result := TsRowColEnumerator.Create(self, -1, ACol, -1, ACol);
end;

function TsRowColAVLTree.GetRowEnumerator(ARow: LongInt): TsRowColEnumerator;
begin
  Result := TsRowColEnumerator.Create(self, ARow, -1, ARow, -1);
end;
                            *)

{@@ ----------------------------------------------------------------------------
  The combination of the methods GetFirst and GetNext allow a fast iteration
  through all nodes of the tree.
-------------------------------------------------------------------------------}
function TsRowColAVLTree.GetFirst: PsRowCol;
begin
  FCurrentNode := FindLowest;
  Result := GetData(FCurrentNode);
end;

function TsRowColAVLTree.GetLast: PsRowCol;
begin
  FCurrentNode := FindHighest;
  Result := GetData(FCurrentNode);
end;

function TsRowColAVLTree.GetNext: PsRowCol;
begin
  if FCurrentNode <> nil then
    FCurrentNode := FindSuccessor(FCurrentNode);
  Result := GetData(FCurrentNode);
end;

function TsRowColAVLTree.GetPrev: PsRowCol;
begin
  if FCurrentNode <> nil then
    FCurrentNode := FindPrecessor(FCurrentNode);
  Result := GetData(FCurrentNode);
end;

{@@ ----------------------------------------------------------------------------
  This procedure adjusts row or column indexes stored in the tree nodes if a
  row or column will be inserted into the underlying worksheet.

  @param  AIndex  Index of the row (if IsRow=true) or column (if IsRow=false)
                  to be inserted
  @param  IsRow   Identifies whether AIndex refers to a row or column index
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.InsertRowOrCol(AIndex: LongInt; IsRow: Boolean);
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
  Removes the node, but does NOT destroy the associated data reocrd
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.Remove(ARow, ACol: LongInt);
var
  node: TAVLTreeNode;
  item: TsRowCol;
begin
  item.Row := ARow;
  item.Col := ACol;
  node := inherited Find(@item);
  Remove(node);
//  Delete(node);
end;

procedure TsRowColAVLTree.PopCurrent;
begin
  FCurrentNode := FCurrentNodeStack.Pop;
end;

procedure TsRowColAVLTree.PushCurrent;
begin
  FCurrentNodeStack.Push(FCurrentNode);
end;


{******************************************************************************}
{ TsCellEnumerator: enumerator for the TsCells AVLTree                        }
{******************************************************************************}

function TsCellEnumerator.GetEnumerator: TsCellEnumerator;
begin
  Result := self;
end;

function TsCellEnumerator.GetCurrent: PCell;
begin
  Result := PCell(inherited GetCurrent);
end;


{******************************************************************************}
{ TsCells: an AVLTree to store spreadsheet cells                               }
{******************************************************************************}

constructor TsCells.Create(AWorksheet: Pointer; AOwnsData: Boolean = true);
begin
  inherited Create(AOwnsData);
  FWorksheet := AWorksheet;
end;

{@@ ----------------------------------------------------------------------------
  Adds a node with a new TCell record to the tree.
  Returns a pointer to the cell record.
  NOTE: It must be checked first that there ia no other record at the same
        col/row. (Check omitted for better performance).
-------------------------------------------------------------------------------}
function TsCells.AddCell(ARow, ACol: LongInt): PCell;
begin
  Result := PCell(Add(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for the specified row and column index along with the
  associated cell data record.
-------------------------------------------------------------------------------}
procedure TsCells.DeleteCell(ARow, ACol: LongInt);
begin
  Delete(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Helper procedure which disposes the memory occupied by the cell data
  record attached to a tree node.
-------------------------------------------------------------------------------}
procedure TsCells.DisposeData(var AData: Pointer);
begin
  if AData <> nil then
    Dispose(PCell(AData));
  AData := nil;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a specific cell already exists
-------------------------------------------------------------------------------}
function TsCells.FindCell(ARow, ACol: Longint): PCell;
begin
  Result := PCell(Find(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Cell enumerators (use in "for ... in" syntax)
-------------------------------------------------------------------------------}
function TsCells.GetEnumerator: TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(self, 0, 0, $7FFFFFFF, $7FFFFFFF, false);
end;

function TsCells.GetColEnumerator(ACol: LongInt; AStartRow: Longint = 0;
  AEndRow: Longint = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self, AStartRow, ACol, AEndRow, ACol, false);
end;

function TsCells.GetRangeEnumerator(AStartRow, AStartCol,
  AEndRow, AEndCol: Longint): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    AStartRow, AStartCol, AEndRow, AEndCol, false);
end;

function TsCells.GetRowEnumerator(ARow: LongInt; AStartCol: Longint = 0;
  AEndCol: LongInt = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    ARow, AStartCol, ARow, AEndCol, false);
end;

function TsCells.GetReverseColEnumerator(ACol: LongInt; AStartRow: Longint = 0;
  AEndRow: Longint = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    AStartRow, ACol, AEndRow, ACol, true);
end;

function TsCells.GetReverseEnumerator: TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(self, 0, 0, $7FFFFFFF, $7FFFFFFF, true);
end;

function TsCells.GetReverseRangeEnumerator(AStartRow, AStartCol,
  AEndRow, AEndCol: Longint): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    AStartRow, AStartCol, AEndRow, AEndCol, true);
end;

function TsCells.GetReverseRowEnumerator(ARow: LongInt; AStartCol: Longint = 0;
  AEndCol: LongInt = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    ARow, AStartCol, ARow, AEndCol, true);
end;


{@@ ----------------------------------------------------------------------------
  Returns a pointer to the first cell of the tree.
  Should always be followed by GetNextCell.

  Use to iterate through all cells efficiently.
-------------------------------------------------------------------------------}
function TsCells.GetFirstCell: PCell;
begin
  Result := PCell(GetFirst);
end;

{@@ ----------------------------------------------------------------------------
  Returns a pointer to the last cell of the tree.

  Needed for efficient iteration through all nodes in reverse direction by
  calling GetPrev.
-------------------------------------------------------------------------------}
function TsCells.GetLastCell: PCell;
begin
  Result := PCell(GetLast);
end;

{@@ ----------------------------------------------------------------------------
  After beginning an iteration through all cells with GetFirstCell, the next
  available cell can be found by calling GetNextCell.

  Use to iterate througt all cells efficiently.
-------------------------------------------------------------------------------}
function TsCells.GetNextCell: PCell;
begin
  Result := PCell(GetNext);
end;

{@@ ----------------------------------------------------------------------------
  After beginning a reverse iteration through all cells with GetLastCell,
  the next available cell can be found by calling GetPrevCell.

  Use to iterate througt all cells efficiently in reverse order.
-------------------------------------------------------------------------------}
function TsCells.GetPrevCell: PCell;
begin
  Result := PCell(GetPrev);
end;

{@@ ----------------------------------------------------------------------------
  Alloates memory for a cell data record.
-------------------------------------------------------------------------------}
function TsCells.NewData: Pointer;
var
  cell: PCell;
begin
  New(cell);
  InitCell(cell^);
  cell^.Worksheet := FWorksheet;
  Result := cell;
end;


{******************************************************************************}
{ TsCommentEnumerator: enumerator for the TsComments AVLTree                   }
{******************************************************************************}

function TsCommentEnumerator.GetEnumerator: TsCommentEnumerator;
begin
  Result := self;
end;

function TsCommentEnumerator.GetCurrent: PsComment;
begin
  Result := PsComment(inherited GetCurrent);
end;


{******************************************************************************}
{ TsComments: an AVLTree to store comment records for cells                     }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Adds a node with a new comment record to the tree. If a node already
  exists then its data will be replaced by the specified ones.
  Returns a pointer to the comment record.
-------------------------------------------------------------------------------}
function TsComments.AddComment(ARow, ACol: Cardinal;
  AComment: String): PsComment;
begin
  Result := PsComment(Find(ARow, ACol));
  if Result = nil then
    Result := PsComment(Add(ARow, ACol));
  Result^.Text := AComment;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for the specified row and column index along with the
  associated comment data record.
-------------------------------------------------------------------------------}
procedure TsComments.DeleteComment(ARow, ACol: Cardinal);
begin
  Delete(ARow, ACol);
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

{@@ ----------------------------------------------------------------------------
  Comments enumerators (use in "for ... in" syntax)
-------------------------------------------------------------------------------}
function TsComments.GetEnumerator: TsCommentEnumerator;
begin
  Result := TsCommentEnumerator.Create(self, 0, 0, $7FFFFFFF, $7FFFFFFF, false);
end;

function TsComments.GetRangeEnumerator(AStartRow, AStartCol,
  AEndRow, AEndCol: Longint): TsCommentEnumerator;
begin
  Result := TsCommentEnumerator.Create(Self,
    AStartRow, AStartCol, AEndRow, AEndCol, false);
end;


{******************************************************************************}
{ TsHyperlinkEnumerator: enumerator for the TsHyperlinks AVLTree               }
{******************************************************************************}

function TsHyperlinkEnumerator.GetEnumerator: TsHyperlinkEnumerator;
begin
  Result := self;
end;

function TsHyperlinkEnumerator.GetCurrent: PsHyperlink;
begin
  Result := PsHyperlink(inherited GetCurrent);
end;

{******************************************************************************}
{ TsHyperlinks: an AVLTree to store hyperlink records for cells                 }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Adds a node with a new hyperlink record to the tree. If a node already
  exists then its data will be replaced by the specified ones.
  Returns a pointer to the hyperlink record.
-------------------------------------------------------------------------------}
function TsHyperlinks.AddHyperlink(ARow, ACol: Longint; ATarget: String;
  ATooltip: String = ''): PsHyperlink;
begin
  Result := PsHyperlink(Find(ARow, ACol));
  if Result = nil then
    Result := PsHyperlink(Add(ARow, ACol));
  Result^.Target := ATarget;
  Result^.Tooltip := ATooltip;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for the specified row and column index along with the
  associated hyperlink data record.
-------------------------------------------------------------------------------}
procedure TsHyperlinks.DeleteHyperlink(ARow, ACol: Longint);
begin
  Delete(ARow, ACol);
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
  Hyperlink enumerators (use in "for ... in" syntax)
-------------------------------------------------------------------------------}
function TsHyperlinks.GetEnumerator: TsHyperlinkEnumerator;
begin
  Result := TsHyperlinkEnumerator.Create(self, 0, 0, $7FFFFFFF, $7FFFFFFF, false);
end;

function TsHyperlinks.GetRangeEnumerator(AStartRow, AStartCol,
  AEndRow, AEndCol: Longint): TsHyperlinkEnumerator;
begin
  Result := TsHyperlinkEnumerator.Create(Self,
    AStartRow, AStartCol, AEndRow, AEndCol, false);
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
function TsMergedCells.AddRange(ARow1, ACol1, ARow2, ACol2: Longint): PsCellRange;
begin
  Result := PsCellRange(Find(ARow1, ACol1));
  if Result = nil then
    Result := PsCellRange(Add(ARow1, ACol1));
  Result^.Row2 := ARow2;
  Result^.Col2 := ACol2;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for which the top/left corner of the cell range matches the
  specified parameters. There is only a single range fulfilling this criterion.
-------------------------------------------------------------------------------}
procedure TsMergedCells.DeleteRange(ARow, ACol: Longint);
begin
  Delete(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  This procedure adjusts row or column indexes stored in the tree nodes if a
  row or column will be deleted from the underlying worksheet.

  @param  AIndex  Index of the row (if IsRow=true) or column (if IsRow=false)
                  to be deleted
  @param  IsRow   Identifies whether AIndex refers to a row or column index
-------------------------------------------------------------------------------}
procedure TsMergedCells.DeleteRowOrCol(AIndex: Longint; IsRow: Boolean);
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

procedure TsMergedCells.Exchange(ARow1, ACol1, ARow2, ACol2: Longint);
var
  rng: PsCellrange;
  dr, dc: LongInt;
begin
  rng := PsCellrange(Find(ARow1, ACol1));
  if rng <> nil then
  begin
    dr := rng^.Row2 - rng^.Row1;
    dc := rng^.Col2 - rng^.Col1;
    rng^.Row1 := ARow2;
    rng^.Col1 := ACol2;
    rng^.Row2 := ARow2 + dr;
    rng^.Col2 := ACol2 + dc;
  end;

  rng := PsCellRange(Find(ARow2, ACol2));
  if rng <> nil then
  begin
    dr := rng^.Row2 - rng^.Row1;
    dc := rng^.Col2 - rng^.Col1;
    rng^.Row1 := ARow1;
    rng^.Col1 := ACol1;
    rng^.Row2 := ARow1 + dr;
    rng^.Col2 := ACol1 + dc;
  end;

  inherited Exchange(ARow1, ACol1, ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Finds the cell range which contains the cell specified by its row and column
  index
-------------------------------------------------------------------------------}
function TsMergedCells.FindRangeWithCell(ARow, ACol: Longint): PsCellRange;
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

