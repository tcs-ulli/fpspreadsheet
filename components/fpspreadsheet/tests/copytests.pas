unit copytests;

{$mode objfpc}{$H+}

interface
{ Tests for copying cells
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8, fpsopendocument, {and a project requirement for lclbase for utf8 handling}
  testsutility;

var
  SourceCells: Array[0..6] of TCell;

procedure InitCopyData;

type
  { TSpreadCopyTests }
  TSpreadCopyTests = class(TTestCase)
  private

  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;

    procedure Test_Copy(ATestKind: Integer);

  published
    procedure Test_CopyValuesToEmptyCells;
//    procedure Test_Copy_Format;
//    procedure Test_Copy_Formula;
  end;

implementation

uses
  TypInfo, Math, fpsutils;

const
  CopyTestSheet = 'Copy';

function InitNumber(ANumber: Double): TCell;
begin
  InitCell(Result);
  Result.ContentType := cctNumber;
  Result.Numbervalue := ANumber;
end;

function InitString(AString: String): TCell;
begin
  InitCell(Result);
  Result.ContentType := cctUTF8String;
  Result.UTF8StringValue := AString;
end;

function InitFormula(AFormula: String; ANumberResult: Double): TCell;
begin
  InitCell(Result);
  Result.FormulaValue := AFormula;
  Result.NumberValue := ANumberResult;
  Result.ContentType := cctNumber;
end;

procedure InitCopyData;
begin
  SourceCells[0] := InitNumber(1.0);   // will be in A1
  SourceCells[1] := InitNumber(2.0);
  SourceCells[2] := InitNumber(3.0);
  SourceCells[3] := InitString('Lazarus');
  SourceCells[4] := InitFormula('A1+1', 2.0);
  InitCell(SourceCells[5]);  // empty but existing
end;


{ TSpreadCopyTests }

procedure TSpreadCopyTests.SetUp;
begin
  inherited SetUp;
  InitCopyData;
end;

procedure TSpreadCopyTests.TearDown;
begin
  inherited TearDown;
end;

{ This test prepares a worksheet and copies Values (ATestKind = 1), Formats
  (AWhat = 2), or Formulas (AWhat = 3). The worksheet is saved, reloaded
  and compared to expectated data }
procedure TSpreadCopyTests.Test_Copy(ATestKind: Integer);
const
  AFormat = sfExcel8;
var
  TempFile: string;
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  cell: PCell;

begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boAutoCalc];

    MyWorkSheet:= MyWorkBook.AddWorksheet(CopyTestSheet);

    // Create two identical columns A and B
    for row := 0 to High(SourceCells) do
      for col := 0 to 1 do
      begin
        case SourceCells[row].ContentType of
          cctNumber:
            cell := MyWorksheet.WriteNumber(row, col, SourceCells[row].NumberValue);
          cctUTF8String:
            cell := Myworksheet.WriteUTF8Text(row, col, SourceCells[row].UTF8StringValue);
          cctEmpty:
            cell := MyWorksheet.WriteBlank(row, col);
        end;
        if SourceCells[row].FormulaValue <> '' then
          Myworksheet.WriteFormula(row, col, SourceCells[row].FormulaValue);
      end;

    MyWorksheet.CalcFormulas;

    case ATestKind of
      1: // copy the source cell values to the empty column C
         for row := 0 to High(SourceCells) do
           Myworksheet.CopyValue(MyWorksheet.FindCell(row, 0), row, 2);
    end;

    // Write to file
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boAutoCalc, boReadFormulas];
    // Read spreadsheet file...
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    MyWorksheet := MyWorkbook.GetFirstWorksheet;

    case ATestKind of
      1: // Copied values in first colum to empty third column
         // The formula cell should contain the result of A1+1 (only value copied)
         begin
           col := 2;
           // Number cells
           for row := 0 to High(SourceCells) do
           begin
             cell := MyWorksheet.FindCell(row, col);
             if (SourceCells[row].ContentType in [cctNumber, cctUTF8String, cctEmpty]) then
               CheckEquals(
                 GetEnumName(TypeInfo(TCellContentType), Integer(SourceCells[row].ContentType)),
                 GetEnumName(TypeInfo(TCellContentType), Integer(cell^.ContentType)),
                 'Content type mismatch, cell '+CellNotation(MyWorksheet, row, col));

             case SourceCells[row].ContentType of
               cctNumber:
                 CheckEquals(
                   SourceCells[row].NumberValue,
                   cell^.NumberValue,
                   'Number value mismatch, cell ' + CellNotation(MyWorksheet, row, col));
               cctUTF8String:
                 CheckEquals(
                   SourceCells[row].UTF8StringValue,
                   cell^.UTF8StringValue,
                   'String value mismatch, cell ' + CellNotation(MyWorksheet, row, col));
             end;

             if HasFormula(@SourceCells[row]) then
               CheckEquals(
                 SourceCells[0].NumberValue + 1,
                 cell^.NumberValue,
                 'Result of copied formula mismatch, cell ' + CellNotation(MyWorksheet, row, col));

           end;
         end;
    end;

  finally
    MyWorkbook.Free;
  end;

  DeleteFile(TempFile);
end;

{ Copy given cell values to empty cells }
procedure TSpreadCopyTests.Test_CopyValuesToEmptyCells;
begin
  Test_Copy(1);
end;


initialization
  RegisterTest(TSpreadCopyTests);
  InitCopyData;

end.

