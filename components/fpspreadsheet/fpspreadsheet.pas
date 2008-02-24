{
fpspreadsheet.pas

Writes an spreadsheet document

AUTHORS: Felipe Monteiro de Carvalho
}
unit fpspreadsheet;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils;

type
  TsSpreadsheetFormat = (sfExcel2, sfExcel3, sfExcel4, sfExcel5, sfExcel8,
   sfOOXML, sfOpenDocument, sfCSV);

const
  { Default extensions }
  STR_EXCEL_EXTENSION = '.xls';

const
  { TokenID values }

  { Binary Operator Tokens }
  INT_EXCEL_TOKEN_TADD    = $03;
  INT_EXCEL_TOKEN_TSUB    = $04;
  INT_EXCEL_TOKEN_TMUL    = $05;
  INT_EXCEL_TOKEN_TDIV    = $06;
  INT_EXCEL_TOKEN_TPOWER  = $07;

  { Constant Operand Tokens }
  INT_EXCEL_TOKEN_TNUM    = $1F;

  { Operand Tokens }
  INT_EXCEL_TOKEN_TREFR   = $24;
  INT_EXCEL_TOKEN_TREFV   = $44;
  INT_EXCEL_TOKEN_TREFA   = $64;

type

  {@@ A Token of a RPN Token array for formulas }

  TRPNToken = record
    TokenID: Byte;
    Col: Byte;
    Row: Word;
    DoubleValue: double;
  end;

  {@@ RPN Token array for formulas }

  TRPNFormula = array of TRPNToken;

  {@@ Describes the type of content of a cell on a TsWorksheet }

  TCellContentType = (cctFormula, cctNumber, cctString, cctWideString);
  
  {@@ Cell structure for TsWorksheet }

  TCell = record
    Row, Col: Cardinal;
    ContentType: TCellContentType;
    FormulaValue: TRPNFormula;
    NumberValue: double;
    StringValue: string;
    WideStringValue: widestring;
  end;

  PCell = ^TCell;

type

  TsCustomSpreadWriter = class;

  {@@ TsWorksheet }

  TsWorksheet = class
  private
    procedure RemoveCallback(data, arg: pointer);
  public
    FCells: TFPList;
    Name: string;
    { Base methods }
    constructor Create;
    destructor Destroy; override;
    { Data manipulation methods }
    function  FindCell(ARow, ACol: Cardinal): PCell;
    function  GetCell(ARow, ACol: Cardinal): PCell;
    procedure RemoveAllCells;
    procedure WriteAnsiText(ARow, ACol: Cardinal; AText: ansistring);
    procedure WriteNumber(ARow, ACol: Cardinal; ANumber: double);
    procedure WriteRPNFormula(ARow, ACol: Cardinal; AFormula: TRPNFormula);
  end;

  {@@ TsWorkbook }

  TsWorkbook = class
  private
    FWorksheets: TFPList;
    procedure RemoveCallback(data, arg: pointer);
  public
    { Base methods }
    constructor Create;
    destructor Destroy; override;
    function  CreateSpreadWriter(AFormat: TsSpreadsheetFormat): TsCustomSpreadWriter;
    procedure WriteToFile(AFileName: string; AFormat: TsSpreadsheetFormat);
    procedure WriteToStream(AStream: TStream; AFormat: TsSpreadsheetFormat);
    { Worksheet list handling methods }
    function  AddWorksheet(AName: string): TsWorksheet;
    function  GetFirstWorksheet: TsWorksheet;
    function  GetWorksheetByIndex(AIndex: Cardinal): TsWorksheet;
    function  GetWorksheetCount: Cardinal;
    procedure RemoveAllWorksheets;
  end;

  {@@ TsSpreadReader class reference type }

  TsSpreadReaderClass = class of TsCustomSpreadReader;
  
  {@@ TsCustomSpreadReader }

  TsCustomSpreadReader = class
  protected
    FWorkbook: TsWorkbook;
    FCurrentWorksheet: TsWorksheet;
  public
    { General writing methods }
    procedure ReadFromFile(AFileName: string; AData: TsWorkbook); virtual;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); virtual; abstract;
    { Record reading methods }
    procedure ReadFormula(AStream: TStream); virtual; abstract;
    procedure ReadLabel(AStream: TStream); virtual; abstract;
    procedure ReadNumber(AStream: TStream); virtual; abstract;
  end;

  {@@ TsSpreadWriter class reference type }

  TsSpreadWriterClass = class of TsCustomSpreadWriter;

  {@@ TsCustomSpreadWriter }

  TsCustomSpreadWriter = class
  public
    { General writing methods }
    procedure WriteCellCallback(data, arg: pointer);
    procedure WriteCellsToStream(AStream: TStream; ACells: TFPList);
    procedure WriteToFile(AFileName: string; AData: TsWorkbook); virtual;
    procedure WriteToStream(AStream: TStream; AData: TsWorkbook); virtual; abstract;
    { Record writing methods }
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Word; const AFormula: TRPNFormula); virtual; abstract;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Word; const AValue: string); virtual; abstract;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double); virtual; abstract;
  end;

  {@@ List of registered formats }

  TsSpreadFormatData = record
    ReaderClass: TsSpreadReaderClass;
    WriterClass: TsSpreadWriterClass;
    Format: TsSpreadsheetFormat;
  end;
  
var
  GsSpreadFormats: array of TsSpreadFormatData;

procedure RegisterSpreadFormat(
  AReaderClass: TsSpreadReaderClass;
  AWriterClass: TsSpreadWriterClass;
  AFormat: TsSpreadsheetFormat);

implementation

{@@
  Registers a new reader/writer pair for a format
}
procedure RegisterSpreadFormat(
  AReaderClass: TsSpreadReaderClass;
  AWriterClass: TsSpreadWriterClass;
  AFormat: TsSpreadsheetFormat);
var
  len: Integer;
begin
  len := Length(GsSpreadFormats);
  SetLength(GsSpreadFormats, len + 1);
  
  GsSpreadFormats[len].ReaderClass := AReaderClass;
  GsSpreadFormats[len].WriterClass := AWriterClass;
  GsSpreadFormats[len].Format := AFormat;
end;

{ TsWorksheet }

{@@
  Helper method for clearing the records in a spreadsheet.
}
procedure TsWorksheet.RemoveCallback(data, arg: pointer);
begin
  FreeMem(data);
end;

{@@
  Constructor.
}
constructor TsWorksheet.Create;
begin
  inherited Create;

  FCells := TFPList.Create;
end;

{@@
  Destructor.
}
destructor TsWorksheet.Destroy;
begin
  RemoveAllCells;

  FCells.Free;

  inherited Destroy;
end;

{@@
  Tryes to locate a Cell in the list of already
  written Cells

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell

  @return Nil if no existing cell was found,
          otherwise a pointer to the desired Cell

  @see    TCell
}
function TsWorksheet.FindCell(ARow, ACol: Cardinal): PCell;
var
  i: Integer;
  ACell: PCell;
begin
  i := 0;
  Result := nil;
  
  while (i < FCells.Count) do
  begin
    ACell := PCell(FCells.Items[i]);
    
    if (ACell^.Row = ARow) and (ACell^.Col = ACol) then
    begin
      Result := ACell;
      Exit;
    end;
    
    Inc(i);
  end;
end;

{@@
  Obtains an allocated cell at the desired location.

  If the Cell already exists, a pointer to it will
  be returned.

  If not, then new memory for the cell will be allocated,
  a pointer to it will be returned and it will be added
  to the list of Cells.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell

  @return A pointer to the Cell on the desired location.

  @see    TCell
}
function TsWorksheet.GetCell(ARow, ACol: Cardinal): PCell;
begin
  Result := FindCell(ARow, ACol);
  
  if (Result = nil) then
  begin
    Result := GetMem(SizeOf(TCell));
    FillChar(Result^, SizeOf(TCell), #0);

    Result^.Row := ARow;
    Result^.Col := ACol;

    FCells.Add(Result);
  end;
end;

{@@
  Clears the list of Cells and releases their memory.
}
procedure TsWorksheet.RemoveAllCells;
begin
  FCells.ForEachCall(RemoveCallback, nil);
end;

{@@
  Writes ansi text to a determined cell.
  
  The text must be encoded on the system default encoding.
  
  On formats the support unicode the text will be converted to the unicode
  encoding that the format supports.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AText     The text to be written encoded with the system encoding
}
procedure TsWorksheet.WriteAnsiText(ARow, ACol: Cardinal; AText: ansistring);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.ContentType := cctString;
  ACell^.StringValue := AText;
end;

{@@
  Writes a floating-point number to a determined cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  ANumber   The number to be written
}
procedure TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: double);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.ContentType := cctNumber;
  ACell^.NumberValue := ANumber;
end;

{@@
  Writes a formula to a determined cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AFormula  The formula in RPN array format
}
procedure TsWorksheet.WriteRPNFormula(ARow, ACol: Cardinal; AFormula: TRPNFormula);
var
  ACell: PCell;
begin
  ACell := GetCell(ARow, ACol);

  ACell^.ContentType := cctFormula;
  ACell^.FormulaValue := AFormula;
end;

{ TsWorkbook }

{@@
  Helper method for clearing the spreadsheet list.
}
procedure TsWorkbook.RemoveCallback(data, arg: pointer);
begin
  TsWorksheet(data).Free;
end;

{@@
  Constructor.
}
constructor TsWorkbook.Create;
begin
  inherited Create;
  
  FWorksheets := TFPList.Create;
end;

{@@
  Destructor.
}
destructor TsWorkbook.Destroy;
begin
  RemoveAllWorksheets;

  FWorksheets.Free;

  inherited Destroy;
end;

{@@
  Convenience method which creates the correct
  writer object for a given spreadsheet format.
}
function TsWorkbook.CreateSpreadWriter(AFormat: TsSpreadsheetFormat): TsCustomSpreadWriter;
var
  i: Integer;
begin
  Result := nil;

  for i := 0 to Length(GsSpreadFormats) - 1 do
    if GsSpreadFormats[i].Format = AFormat then
    begin
      Result := GsSpreadFormats[i].WriterClass.Create;
    
      Break;
    end;
    
  if Result = nil then raise Exception.Create('Unsuported spreadsheet format.');
end;

{@@
  Writes the document to a file.

  If the file doesn't exist, it will be created.
}
procedure TsWorkbook.WriteToFile(AFileName: string; AFormat: TsSpreadsheetFormat);
var
  AWriter: TsCustomSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);

  try
    AWriter.WriteToFile(AFileName, Self);
  finally
    AWriter.Free;
  end;
end;

{@@
  Writes the document to a stream
}
procedure TsWorkbook.WriteToStream(AStream: TStream; AFormat: TsSpreadsheetFormat);
var
  AWriter: TsCustomSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);

  try
    AWriter.WriteToStream(AStream, Self);
  finally
    AWriter.Free;
  end;
end;

{@@
  Adds a new worksheet to the workbook

  It is added to the end of the list of worksheets

  @param  AName     The name of the new worksheet

  @return The instace of the newly created worksheet

  @see    TsWorkbook
}
function TsWorkbook.AddWorksheet(AName: string): TsWorksheet;
begin
  Result := TsWorksheet.Create;

  Result.Name := AName;

  FWorksheets.Add(Pointer(Result));
end;

{@@
  Quick helper routine which returns the first worksheet

  @return A TsWorksheet instance if at least one is present.
          nil otherwise.

  @see    TsWorkbook.GetWorksheetByIndex
  @see    TsWorksheet
}
function TsWorkbook.GetFirstWorksheet: TsWorksheet;
begin
  Result := TsWorksheet(FWorksheets.First);
end;

{@@
  Gets the worksheet with a given index

  The index is zero-based, so the first worksheet
  added has index 0, the second 1, etc.

  @param  AIndex    The index of the worksheet (0-based)

  @return A TsWorksheet instance if one is present at that index.
          nil otherwise.

  @see    TsWorkbook.GetFirstWorksheet
  @see    TsWorksheet
}
function TsWorkbook.GetWorksheetByIndex(AIndex: Cardinal): TsWorksheet;
begin
  Result := TsWorksheet(FWorksheets.Items[AIndex]);
end;

{@@
  The number of worksheets on the workbook

  @see    TsWorksheet
}
function TsWorkbook.GetWorksheetCount: Cardinal;
begin
  Result := FWorksheets.Count;
end;

{@@
  Clears the list of Worksheets and releases their memory.
}
procedure TsWorkbook.RemoveAllWorksheets;
begin
  FWorksheets.ForEachCall(RemoveCallback, nil);
end;

{ TsCustomSpreadReader }

procedure TsCustomSpreadReader.ReadFromFile(AFileName: string; AData: TsWorkbook);
begin

end;

{ TsCustomSpreadWriter }

{@@
  Helper function for the spreadsheet writers.

  @see    TsCustomSpreadWriter.WriteCellsToStream
}
procedure TsCustomSpreadWriter.WriteCellCallback(data, arg: pointer);
var
  ACell: PCell;
  AStream: TStream;
begin
  ACell := PCell(data);
  AStream := TStream(arg);

  case ACell.ContentType of
    cctFormula: WriteFormula(AStream, ACell^.Row, ACell^.Col, ACell^.FormulaValue);
    cctNumber:  WriteNumber(AStream, ACell^.Row, ACell^.Col, ACell^.NumberValue);
    cctString:  WriteLabel(AStream, ACell^.Row, ACell^.Col, ACell^.StringValue);
  end;
end;

{@@
  Helper function for the spreadsheet writers.

  Iterates all cells on a list, calling the appropriate write method for them.

  @param  AStream The output stream.
  @param  ACells  List of cells to be writeen
}
procedure TsCustomSpreadWriter.WriteCellsToStream(AStream: TStream; ACells: TFPList);
begin
  ACells.ForEachCall(WriteCellCallback, Pointer(AStream));
end;

{@@
  Default file writting method.

  Opens the file and calls WriteToStream

  @param  AFileName The output file name.
                   If the file already exists it will be replaced.
  @param  AData     The Workbook to be saved.

  @see    TsWorkbook
}
procedure TsCustomSpreadWriter.WriteToFile(AFileName: string; AData: TsWorkbook);
var
  OutputFile: TFileStream;
begin
  OutputFile := TFileStream.Create(AFileName, fmCreate or fmOpenWrite);
  try
    WriteToStream(OutputFile, AData);
  finally
    OutputFile.Free;
  end;
end;

finalization

  SetLength(GsSpreadFormats, 0);

end.

