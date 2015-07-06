unit fpsclasses;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, AVL_Tree, //avglvltree,
  fpstypes;

type
  { forward declarations }
  TsRowColAVLTree = class;

  { TsRowCol }
  TsRowCol = record
    Row, Col: Cardinal;
  end;
  PsRowCol = ^TsRowCol;

  { TsRowColEnumerator }
  TsRowColEnumerator = class
  private
  protected
    FCurrentNode: TAVLTreeNode;
    FTree: TsRowColAVLTree;
    FStartRow, FEndRow, FStartCol, FEndCol: LongInt;
    FDone: Boolean;
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
  protected
    procedure DisposeData(var AData: Pointer); virtual; abstract;
    function NewData: Pointer; virtual; abstract;
  public
    constructor Create(AOwnsData: Boolean = true);
    destructor Destroy; override;
    function Add(ARow, ACol: Cardinal): PsRowCol; overload;
    procedure Clear;
    procedure Delete(ANode: TAVLTreeNode); overload;
    procedure Delete(ARow, ACol: Cardinal); overload;
    procedure DeleteRowOrCol(AIndex: Cardinal; IsRow: Boolean); virtual;
    procedure Exchange(ARow1, ACol1, ARow2, ACol2: Cardinal); virtual;
    function FindByRowCol(ARow, ACol: Cardinal): PsRowCol; overload;
    function GetData(ANode: TAVLTreeNode): PsRowCol;
    function GetFirst: PsRowCol;
    function GetLast: PsRowCol;
    procedure InsertRowOrCol(AIndex: Cardinal; IsRow: Boolean);
    procedure Remove(ARow, ACol: Cardinal); overload;
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
    function AddCell(ARow, ACol: Cardinal): PCell;
    procedure DeleteCell(ARow, ACol: Cardinal);
    function FindCell(ARow, ACol: Cardinal): PCell;
    function GetFirstCell: PCell;
    function GetFirstCellOfRow(ARow: Cardinal): PCell;
    function GetLastCell: PCell;
    function GetLastCellOfRow(ARow: Cardinal): PCell;
    // enumerators
    function GetEnumerator: TsCellEnumerator;
    function GetReverseEnumerator: TsCellEnumerator;
    function GetColEnumerator(ACol: Cardinal; AStartRow: Cardinal = 0;
      AEndRow: Cardinal = $7FFFFFFF): TsCellEnumerator;
    function GetRangeEnumerator(AStartRow, AStartCol,
      AEndRow, AEndCol: Cardinal): TsCellEnumerator;
    function GetRowEnumerator(ARow: Cardinal; AStartCol:Cardinal = 0;
      AEndCol: Cardinal = $7FFFFFFF): TsCellEnumerator;
    function GetReverseColEnumerator(ACol: Cardinal; AStartRow: Cardinal = 0;
      AEndRow: Cardinal = $7FFFFFFF): TsCellEnumerator;
    function GetReverseRangeEnumerator(AStartRow, AStartCol,
      AEndRow, AEndCol: Cardinal): TsCellEnumerator;
    function GetReverseRowEnumerator(ARow: Cardinal; AStartCol:Cardinal = 0;
      AEndCol: Cardinal = $7FFFFFFF): TsCellEnumerator;
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
      AEndRow, AEndCol: Cardinal): TsCommentEnumerator;
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
    function AddHyperlink(ARow, ACol: Cardinal; ATarget: String;
      ATooltip: String = ''): PsHyperlink;
    procedure DeleteHyperlink(ARow, ACol: Cardinal);
    // enumerators
    function GetEnumerator: TsHyperlinkEnumerator;
    function GetRangeEnumerator(AStartRow, AStartCol,
      AEndRow, AEndCol: Cardinal): TsHyperlinkEnumerator;
  end;

  { TsMergedCells }
  TsCellRangeEnumerator = class(TsRowColEnumerator)
  protected
    function GetCurrent: PsCellRange;
  public
    function GetEnumerator: TsCellRangeEnumerator; inline;
    property Current: PsCellRange read GetCurrent;
  end;

  TsMergedCells = class(TsRowColAVLTree)
  protected
    procedure DisposeData(var AData: Pointer); override;
    function NewData: Pointer; override;
  public
    function AddRange(ARow1, ACol1, ARow2, ACol2: Cardinal): PsCellRange;
    procedure DeleteRange(ARow, ACol: Cardinal);
    procedure DeleteRowOrCol(AIndex: Cardinal; IsRow: Boolean); override;
    procedure Exchange(ARow1, ACol1, ARow2, ACol2: Cardinal); override;
    function FindRangeWithCell(ARow, ACol: Cardinal): PsCellRange;
    // enumerators
    function GetEnumerator: TsCellRangeEnumerator;
  end;

  { TsCellFormatList }
  TsCellFormatList = class(TFPList)
  private
    FAllowDuplicates: Boolean;
    function GetItem(AIndex: Integer): PsCellFormat;
    procedure SetItem(AIndex: Integer; const AValue: PsCellFormat);
  public
    constructor Create(AAllowDuplicates: Boolean);
    destructor Destroy; override;
    function Add(const AItem: TsCellFormat): Integer; overload;
    function Add(AItem: PsCellFormat): Integer; overload;
    procedure Clear;
    procedure Delete(AIndex: Integer);
    function FindIndexOfID(ID: Integer): Integer;
    function FindIndexOfName(AName: String): Integer;
    function IndexOf(const AItem: TsCellFormat): Integer; overload;
    property Items[AIndex: Integer]: PsCellFormat read GetItem write SetItem; default;
  end;


implementation

uses
  Math,
  fpsUtils;


{ Helper function for sorting }

function CompareRowCol(Item1, Item2: Pointer): Integer;
begin
  Result := Longint(PsRowCol(Item1)^.Row) - PsRowCol(Item2)^.Row;
  if Result = 0 then
    Result := Longint(PsRowCol(Item1)^.Col) - PsRowCol(Item2)^.Col;
end;


{******************************************************************************}
{ TsRowColEnumerator:  A specialized enumerator for TsRowColAVLTree using the  }
{ pointers to the data records.                                                }
{******************************************************************************}

constructor TsRowColEnumerator.Create(ATree: TsRowColAVLTree;
  AStartRow, AStartCol, AEndRow, AEndCol: LongInt; AReverse: Boolean);
var
  node: TAVLTreeNode;
begin
  FTree := ATree;
  FReverse := AReverse;
  if AStartRow <= AEndRow then
  begin
    FStartRow := AStartRow;
    FEndRow := AEndRow;
  end else
  begin
    FStartRow := AEndRow;
    FEndRow := AStartRow;
  end;
  if AStartCol <= AEndCol then
  begin
    FStartCol := AStartCol;
    FEndCol := AEndCol;
  end else
  begin
    FStartCol := AEndCol;
    FEndCol := AStartCol;
  end;
  if FEndRow = $7FFFFFFF then
  begin
    node := FTree.FindHighest;
    if node <> nil then
      FEndRow := PsRowCol(node.Data)^.Row;
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
  curr: PsRowCol;
  rc: TsRowCol;
begin
  Result := false;
  if FCurrentNode <> nil then begin
    if FReverse then
    begin
      FCurrentNode := FTree.FindPrecessor(FCurrentNode);
      if FCurrentNode <> nil then
      begin
        curr := PsRowCol(FCurrentNode.Data);
        if not InRange(LongInt(curr^.Col), FStartCol, FEndCol) then
        begin
          rc := curr^;
          if LongInt(rc.Col) < FStartCol then
            dec(LongInt(rc.Row));
          rc.Col := FEndCol;
          FCurrentNode := FTree.FindNearest(@rc);
          if FCurrentNode <> nil then begin
            curr := PsRowCol(FCurrentNode.Data);
            while (FCurrentNode <> nil) and
                  not (InRange(curr^.Row, FStartRow, FEndRow) and InRange(curr^.Col, FStartCol, FEndCol))
            do begin
              FCurrentNode := FTree.FindPrecessor(FCurrentNode);
              if FCurrentNode <> nil then curr := PsRowCol(FCurrentNode.Data);
            end;
          end;
        end;
      end;
    end else
    begin
      FCurrentNode := FTree.FindSuccessor(FCurrentNode);
      if FCurrentNode <> nil then
      begin
        curr := PsRowCol(FCurrentNode.Data);
        rc.Col := FStartCol;
        if LongInt(rc.Col) > FEndCol then inc(rc.Row);
        if not InRange(LongInt(curr^.Col), FStartCol, FEndCol) then
        begin
          rc := curr^;
          if LongInt(rc.Col) > FEndCol then inc(rc.Row);
          rc.Col := FStartCol;
          FCurrentNode := FTree.FindNearest(@rc);
          if FCurrentNode <> nil then
          begin
            curr := PsRowCol(FCurrentNode.Data);
            if (LongInt(curr^.Col) < FStartCol) then
              while (FCurrentNode <> nil) and not InRange(curr^.Col, FStartCol, FEndCol) do
              begin
                FCurrentNode := FTree.FindSuccessor(FCurrentNode);
                if FCurrentNode <> nil then curr := PsRowCol(FCurrentNode.Data);
              end;
            while (FCurrentNode <> nil) and
                  not (InRange(curr^.Row, FStartRow, FEndRow) and InRange(curr^.Col, FStartCol, FEndCol))
            do begin
              FCurrentNode := FTree.FindSuccessor(FCurrentNode);
              if FCurrentNode <> nil then curr := PsRowCol(FCurrentNode.Data);
            end;
          end;
        end;
      end;
    end;
    Result := (FCurrentNode <> nil) and
              InRange(curr^.Col, FStartCol, FEndCol) and
              InRange(curr^.Row, FStartRow, FEndRow);
  end else
  begin
    if FReverse then
    begin
      rc.Row := FEndRow;
      rc.Col := FEndCol;
      FCurrentNode := FTree.FindNearest(@rc);
      if FCurrentNode <> nil then
        curr := PsRowCol(FCurrentNode.Data);
      while (FCurrentNode <> nil) and
            not (InRange(curr^.Row, FStartRow, FEndRow) and InRange(curr^.Col, FStartCol, FEndCol))
      do begin
        FCurrentNode := FTree.FindPrecessor(FCurrentNode);
        if FCurrentNode <> nil then curr := PsRowCol(FCurrentNode.Data);
      end;
    end else
    begin
      rc.Row := FStartRow;
      rc.Col := FStartCol;
      FCurrentNode := FTree.FindNearest(@rc);
      if FCurrentNode <> nil then
        curr := PsRowCol(FCurrentNode.Data);
      while (FCurrentNode <> nil) and
            not (InRange(curr^.Row, FStartRow, FEndRow) and InRange(curr^.Col, FStartCol, FEndCol))
      do begin
        FCurrentNode := FTree.FindSuccessor(FCurrentNode);
        if FCurrentNode <> nil then curr := PsRowCol(FCurrentNode.Data);
      end;
    end;
    Result := (FCurrentNode <> nil);
  end;
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

procedure TsRowColAVLTree.Delete(ARow, ACol: Cardinal);
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
      // Remove and destroy the RowCol record if it is in the deleted row
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
procedure TsRowColAVLTree.Exchange(ARow1, ACol1, ARow2, ACol2: Cardinal);
var
  item1, item2: PsRowCol;
begin
  item1 := FindByRowCol(ARow1, ACol1);
  item2 := FindByRowCol(ARow2, ACol2);

  // There are entries for both locations: Exchange row/col indexes
  if (item1 <> nil) and (item2 <> nil) then
  begin
    Remove(item1);
    Remove(item2);
    item1^.Row := ARow2;
    item1^.Col := ACol2;
    item2^.Row := ARow1;
    item2^.Col := ACol1;
    inherited Add(item1);   // The items are sorted to the correct position
    inherited Add(item2);   // when they are added to the tree
  end else

  // Only the 1st item exists --> give it the row/col indexes of the 2nd item
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
function TsRowColAVLTree.FindByRowCol(ARow, ACol: Cardinal): PsRowCol;
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

{@@ ----------------------------------------------------------------------------
  The combination of the methods GetFirst and GetNext allow a fast iteration
  through all nodes of the tree.
-------------------------------------------------------------------------------}
function TsRowColAVLTree.GetFirst: PsRowCol;
begin
  Result := GetData(FindLowest);
end;

function TsRowColAVLTree.GetLast: PsRowCol;
begin
  Result := GetData(FindHighest);
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
  Removes the node, but does NOT destroy the associated data reocrd
-------------------------------------------------------------------------------}
procedure TsRowColAVLTree.Remove(ARow, ACol: Cardinal);
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
function TsCells.AddCell(ARow, ACol: Cardinal): PCell;
begin
  Result := PCell(Add(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Deletes the node for the specified row and column index along with the
  associated cell data record.
-------------------------------------------------------------------------------}
procedure TsCells.DeleteCell(ARow, ACol: Cardinal);
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
function TsCells.FindCell(ARow, ACol: Cardinal): PCell;
begin
  Result := PCell(FindByRowCol(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Cell enumerators (use in "for ... in" syntax)
-------------------------------------------------------------------------------}
function TsCells.GetEnumerator: TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(self, 0, 0, $7FFFFFFF, $7FFFFFFF, false);
end;

function TsCells.GetColEnumerator(ACol: Cardinal; AStartRow: Cardinal = 0;
  AEndRow: Cardinal = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self, AStartRow, ACol, AEndRow, ACol, false);
end;

function TsCells.GetRangeEnumerator(AStartRow, AStartCol,
  AEndRow, AEndCol: Cardinal): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    AStartRow, AStartCol, AEndRow, AEndCol, false);
end;

function TsCells.GetRowEnumerator(ARow: Cardinal; AStartCol: Cardinal = 0;
  AEndCol: Cardinal = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    ARow, AStartCol, ARow, AEndCol, false);
end;

function TsCells.GetReverseColEnumerator(ACol: Cardinal; AStartRow: Cardinal = 0;
  AEndRow: Cardinal = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    AStartRow, ACol, AEndRow, ACol, true);
end;

function TsCells.GetReverseEnumerator: TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(self, 0, 0, $7FFFFFFF, $7FFFFFFF, true);
end;

function TsCells.GetReverseRangeEnumerator(AStartRow, AStartCol,
  AEndRow, AEndCol: Cardinal): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    AStartRow, AStartCol, AEndRow, AEndCol, true);
end;

function TsCells.GetReverseRowEnumerator(ARow: Cardinal; AStartCol: Cardinal = 0;
  AEndCol: Cardinal = $7FFFFFFF): TsCellEnumerator;
begin
  Result := TsCellEnumerator.Create(Self,
    ARow, AStartCol, ARow, AEndCol, true);
end;


{@@ ----------------------------------------------------------------------------
  Returns a pointer to the first cell of the tree.
-------------------------------------------------------------------------------}
function TsCells.GetFirstCell: PCell;
begin
  Result := PCell(GetFirst);
end;

{@@ ----------------------------------------------------------------------------
  Returns a pointer to the first cell in a specified row
-------------------------------------------------------------------------------}
function TsCells.GetFirstCellOfRow(ARow: Cardinal): PCell;
begin
  Result := nil;
  // Creating the row enumerator automatically finds the first cell of the row
  for Result in GetRowEnumerator(ARow) do
    exit;
end;

{@@ ----------------------------------------------------------------------------
  Returns a pointer to the last cell of the tree.
-------------------------------------------------------------------------------}
function TsCells.GetLastCell: PCell;
begin
  Result := PCell(GetLast);
end;

{@@ ----------------------------------------------------------------------------
  Returns a pointer to the last cell of a specified row
-------------------------------------------------------------------------------}
function TsCells.GetLastCellOfRow(ARow: Cardinal): PCell;
begin
  Result := nil;
  // Creating the reverse row enumerator finds the last cell of the row
  for Result in GetReverseRowEnumerator(ARow) do
    exit;
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
  Result := PsComment(FindByRowCol(ARow, ACol));
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
  AEndRow, AEndCol: Cardinal): TsCommentEnumerator;
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
function TsHyperlinks.AddHyperlink(ARow, ACol: Cardinal; ATarget: String;
  ATooltip: String = ''): PsHyperlink;
begin
  Result := PsHyperlink(FindByRowCol(ARow, ACol));
  if Result = nil then
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
  AEndRow, AEndCol: Cardinal): TsHyperlinkEnumerator;
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
{ TsCellRangeEnumerator: enumerator for the cell range records                 }
{******************************************************************************}

function TsCellRangeEnumerator.GetEnumerator: TsCellRangeEnumerator;
begin
  Result := self;
end;

function TsCellRangeEnumerator.GetCurrent: PsCellRange;
begin
  Result := PsCellRange(inherited GetCurrent);
end;


{******************************************************************************}
{ TsMergedCells: a AVLTree to store merged cell range records for cells         }
{******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Adds a node with a new merge cell range record to the tree. If a node already
  exists then its data will be replaced by the specified ones.
  Returns a pointer to the cell range record.
-------------------------------------------------------------------------------}
function TsMergedCells.AddRange(ARow1, ACol1, ARow2, ACol2: Cardinal): PsCellRange;
begin
  Result := PsCellRange(FindByRowCol(ARow1, ACol1));
  if Result = nil then
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
  Delete(ARow, ACol);
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
  rng: PsCellRange;
  R: TsCellRange;
  node, nextnode: TAVLTreeNode;
begin
  node := FindLowest;
  while Assigned(node) do begin
    rng := PsCellRange(node.Data);
    nextnode := FindSuccessor(node);
    if IsRow then
    begin
      // Deleted row is above the merged range --> Shift entire range up by 1
      // NOTE:
      // The "merged" flags do not have to be changed, they move with the cells.
      if (AIndex < rng^.Row1) then begin
        R := rng^;        // Store range parameters
        Delete(node);     // Delete node from tree, adapt the row indexes, ...
        AddRange(R.Row1-1, R.Col1, R.Row2-1, R.Col2);         // ... and re-insert to get it sorted correctly
      end else
      // Single-row merged block coincides with row to be deleted
      if (AIndex = rng^.Row1) and (rng^.Row1 = rng^.Row2) then
        DeleteRange(rng^.Row1, rng^.Col1)
      else
      // Deleted row runs through the merged block --> Shift bottom row up by 1
      // NOTE: The "merged" flags disappear with the deleted cells
      if (AIndex >= rng^.Row1) and (AIndex <= rng^.Row2) then
        dec(rng^.Row2);  // no need to remove & re-insert because Row1 does not change
    end else
    begin
      // Deleted column is at the left of the merged range
      // --> Shift entire merged range to the left by 1
      // NOTE:
      // The "merged" flags do not have to be changed, they move with the cells.
      if (AIndex < rng^.Col1) then begin
        R := rng^;
        Delete(node);
        AddRange(R.Row1, R.Col1-1, R.Row2, R.Col2-1);
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
    node := nextnode;
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

procedure TsMergedCells.Exchange(ARow1, ACol1, ARow2, ACol2: Cardinal);
var
  rng: PsCellrange;
  dr, dc: Cardinal;
begin
  rng := PsCellrange(FindByRowCol(ARow1, ACol1));
  if rng <> nil then
  begin
    dr := rng^.Row2 - rng^.Row1;
    dc := rng^.Col2 - rng^.Col1;
    rng^.Row1 := ARow2;
    rng^.Col1 := ACol2;
    rng^.Row2 := ARow2 + dr;
    rng^.Col2 := ACol2 + dc;
  end;

  rng := PsCellRange(FindByRowCol(ARow2, ACol2));
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
  Cell range enumerator (use in "for ... in" syntax)
-------------------------------------------------------------------------------}
function TsMergedCells.GetEnumerator: TsCellRangeEnumerator;
begin
  Result := TsCellRangeEnumerator.Create(self, 0, 0, $7FFFFFFF, $7FFFFFFF, false);
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


{******************************************************************************}
{                          TsCellFormatList                                    }
{******************************************************************************}

constructor TsCellFormatList.Create(AAllowDuplicates: Boolean);
begin
  inherited Create;
  FAllowDuplicates := AAllowDuplicates;
end;

destructor TsCellFormatList.Destroy;
begin
  Clear;
  inherited;
end;

function TsCellFormatList.Add(const AItem: TsCellFormat): Integer;
var
  P: PsCellFormat;
begin
  if FAllowDuplicates then
    Result := -1
  else
    Result := IndexOf(AItem);
  if Result = -1 then begin
    New(P);
    P^.Name := AItem.Name;
    P^.ID := AItem.ID;
    P^.UsedFormattingFields := AItem.UsedFormattingFields;
    P^.FontIndex := AItem.FontIndex;
    P^.TextRotation := AItem.TextRotation;
    P^.HorAlignment := AItem.HorAlignment;
    P^.VertAlignment := AItem.VertAlignment;
    P^.Border := AItem.Border;
    P^.BorderStyles := AItem.BorderStyles;
    P^.Background := AItem.Background;
    P^.NumberFormatIndex := AItem.NumberFormatIndex;
    P^.NumberFormat := AItem.NumberFormat;
    P^.NumberFormatStr := AItem.NumberFormatStr;
    Result := inherited Add(P);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds a pointer to a FormatRecord to the list. Allows nil for the predefined
  formats which are not stored in the file.
-------------------------------------------------------------------------------}
function TsCellFormatList.Add(AItem: PsCellFormat): Integer;
begin
  if AItem = nil then
    Result := inherited Add(AItem)
  else
    Result := Add(AItem^);
end;

procedure TsCellFormatList.Clear;
var
  i: Integer;
begin
  for i:=Count-1 downto 0 do
    Delete(i);
  inherited;
end;

procedure TsCellFormatList.Delete(AIndex: Integer);
var
  P: PsCellFormat;
begin
  P := GetItem(AIndex);
  if P <> nil then
    Dispose(P);
  inherited Delete(AIndex);
end;

function TsCellFormatList.GetItem(AIndex: Integer): PsCellFormat;
begin
  Result := inherited Items[AIndex];
end;

function TsCellFormatList.FindIndexOfID(ID: Integer): Integer;
var
  P: PsCellFormat;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P <> nil) and (P^.ID = ID) then
      exit;
  end;
  Result := -1;
end;

function TsCellFormatList.FindIndexOfName(AName: String): Integer;
var
  P: PsCellFormat;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P <> nil) and (P^.Name = AName) then
      exit;
  end;
  Result := -1;
end;

function TsCellFormatList.IndexOf(const AItem: TsCellFormat): Integer;
var
  P: PsCellFormat;
  equ: Boolean;
  b: TsCellBorder;
begin
  for Result := 0 to Count-1 do
  begin
    P := GetItem(Result);
    if (P = nil) then continue;

    if (P^.UsedFormattingFields <> AItem.UsedFormattingFields) then continue;

    if (uffFont in AItem.UsedFormattingFields) then
      if (P^.FontIndex) <> (AItem.FontIndex) then continue;

    if (uffTextRotation in AItem.UsedFormattingFields) then
      if (P^.TextRotation <> AItem.TextRotation) then continue;

    if (uffHorAlign in AItem.UsedFormattingFields) then
      if (P^.HorAlignment <> AItem.HorAlignment) then continue;

    if (uffVertAlign in AItem.UsedFormattingFields) then
      if (P^.VertAlignment <> AItem.VertAlignment) then continue;

    if (uffBorder in AItem.UsedFormattingFields) then
      if (P^.Border <> AItem.Border) then continue;

    // Border styles can be set even if borders are not used --> don't check uffBorder!
    equ := true;
    for b in AItem.Border do begin
      if (P^.BorderStyles[b].LineStyle <> AItem.BorderStyles[b].LineStyle) or
         (P^.BorderStyles[b].Color <> Aitem.BorderStyles[b].Color)
      then begin
        equ := false;
        break;
      end;
    end;
    if not equ then continue;

    if (uffBackground in AItem.UsedFormattingFields) then begin
      if (P^.Background.Style <> AItem.Background.Style) then continue;
      if (P^.Background.BgColor <> AItem.Background.BgColor) then continue;
      if (P^.Background.FgColor <> AItem.Background.FgColor) then continue;
    end;

    if (uffNumberFormat in AItem.UsedFormattingFields) then begin
      if (P^.NumberFormatIndex <> AItem.NumberFormatIndex) then continue;
      if (P^.NumberFormat <> AItem.NumberFormat) then continue;
      if (P^.NumberFormatStr <> AItem.NumberFormatStr) then continue;
    end;

    // If we arrive here then the format records match.
    exit;
  end;

  // We get here if no record matches
  Result := -1;
end;

procedure TsCellFormatList.SetItem(AIndex: Integer; const AValue: PsCellFormat);
begin
  inherited Items[AIndex] := AValue;
end;

end.

