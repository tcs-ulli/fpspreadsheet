unit copytests;

{$mode objfpc}{$H+}

interface
{ Tests for copying cells
  NOTE: The code in these tests is very fragile because the test results are
  hard-coded. Any modification in "InitCopyData" must be carefully verified!
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8, fpsopendocument, {and a project requirement for lclbase for utf8 handling}
  testsutility;

type
  PCellRecord = ^TCellRecord;
  TCellRecord = record
    ContentType: TCellContentType;
    Numbervalue: Double;
    UTF8StringValue: String;
    FormulaValue: String;
    UsedFormattingFields: TsUsedFormattingFields;
    SharedFormulaBase: PCellRecord;
    BackgroundColor: TsColor;
  end;

var
  SourceCells: Array[0..8] of TCellRecord;

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
    procedure Test_CopyValuesToOccupiedCells;

    procedure Test_CopyFormatsToEmptyCells;
    procedure Test_CopyFormatsToOccupiedCells;

    procedure Test_CopyFormulasToEmptyCells;
    procedure Test_CopyFormulasToOccupiedCells;
  end;

implementation

uses
  TypInfo, fpsutils;

const
  CopyTestSheet = 'Copy';

procedure MyInitCellRecord(out ACell: TCellRecord);
begin
  ACell.Contenttype := cctEmpty;
  ACell.NumberValue := 0.0;
  ACell.UTF8StringValue := '';
  ACell.FormulaValue := '';
  ACell.UsedformattingFields := [];
  ACell.BackgroundColor := scTransparent;
  ACell.SharedFormulaBase := nil;
end;

function MyHasFormula(ACell: PCellRecord): Boolean;
begin
  Result := Assigned(ACell) and (
    (ACell^.SharedFormulaBase <> nil) or (Length(ACell^.FormulaValue) > 0)
  );
end;

function InitNumber(ANumber: Double; ABkColor: TsColor): TCellRecord;
begin
  MyInitCellRecord(Result);
  Result.ContentType := cctNumber;
  Result.Numbervalue := ANumber;
  if (ABkColor <> scTransparent) then begin
    Result.UsedFormattingFields := Result.UsedFormattingFields + [uffBackground];
    Result.BackgroundColor := ABkColor;
  end;
end;

function InitString(AString: String; ABkColor: TsColor): TCellRecord;
begin
  MyInitCellRecord(Result);
  Result.ContentType := cctUTF8String;
  Result.UTF8StringValue := AString;
  if (ABkColor <> scTransparent) then begin
    Result.UsedFormattingFields := Result.usedFormattingFields + [uffBackground];
    Result.BackgroundColor := ABkColor;
  end;
end;

function InitFormula(AFormula: String; ANumberResult: Double;
  ABkColor: TsColor): TCellRecord;
begin
  MyInitCellRecord(Result);
  Result.FormulaValue := AFormula;
  Result.NumberValue := ANumberResult;
  Result.ContentType := cctNumber;
  if (ABkColor <> scTransparent) then begin
    Result.UsedFormattingFields := Result.usedFormattingFields + [uffBackground];
    Result.BackgroundColor := ABkColor;
  end;
end;

{ IMPORTANT: Carefully check the Test_Copy method if anything is changed here.
  The expected test results are hard-coded in this method! }
procedure InitCopyData;
begin
  SourceCells[0] := InitNumber(1.0, scTransparent);   // will be in A1
  SourceCells[1] := InitNumber(2.0, scTransparent);
  SourceCells[2] := InitNumber(3.0, scYellow);
  SourceCells[3] := InitString('Lazarus', scRed);
  SourceCells[4] := InitFormula('A1+1', 2.0, scTransparent);
  SourceCells[5] := InitFormula('$A1+1', 2.0, scTransparent);
  SourceCells[6] := InitFormula('A$1+1', 2.0, scTransparent);
  SourceCells[7] := InitFormula('$A$1+1', 2.0, scGray);
  MyInitCellRecord(SourceCells[8]);  // empty but existing
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

{ This test prepares a worksheet and copies Values (ATestKind = 1 or 2), Formats
  (AWhat = 3 or 4), or Formulas (AWhat = 5 or 6). The odd ATestKind number
  copy the data to the empty column C, the even value copy them to the
  occupied column B which contains the source data (in column A) shifted down
  by 1 cell. "The worksheet is saved, reloaded and compared to expectated data }
procedure TSpreadCopyTests.Test_Copy(ATestKind: Integer);
const
  AFormat = sfExcel8;
var
  TempFile: string;
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  i, row, col: Integer;
  cell: PCell;
  expectedFormula: String;

begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
//    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving]; //boAutoCalc];

    MyWorkSheet:= MyWorkBook.AddWorksheet(CopyTestSheet);

    // Prepare the worksheet in which cells are copied:
    // Store the SourceCells to column A and B; in B shifted down by 1 cell
    {      A                B                   C
      1   1.0
      2   2.0              1.0
      3   3.0 (yellow)     2.0
      4   Lazarus (red)    3.0
      5   A1+1             Lazarus
      6   $A1+1            A1+1
      7   A$1+1            $A1+1
      8   $A$1+1 (gray)    A$1+1
      9   (empty)          $A$1+1 (gray)
     10                   (empty)
    }
    for col := 0 to 1 do
      for row := 0 to High(SourceCells) do
      begin
        // Adding the col to the row index shifts the data in the second column
        // down. Offsetting the second column is done to avoid that the "copy"
        // action operates on cells having a different content afterwards.
        case SourceCells[row].ContentType of
          cctNumber:
            cell := MyWorksheet.WriteNumber(row+col, col, SourceCells[row].NumberValue);
          cctUTF8String:
            cell := Myworksheet.WriteUTF8Text(row+col, col, SourceCells[row].UTF8StringValue);
          cctEmpty:
            cell := MyWorksheet.WriteBlank(row+col, col);
        end;
        if SourceCells[row].FormulaValue <> '' then
          Myworksheet.WriteFormula(row+col, col, SourceCells[row].FormulaValue);
        if SourceCells[row].BackgroundColor <> scTransparent then
          MyWorksheet.WriteBackgroundColor(cell, SourceCells[row].BackgroundColor);
      end;

    MyWorksheet.CalcFormulas;

    // Now perform the "copy" operations
    for row := 0 to High(SourceCells) do
    begin
      cell := Myworksheet.FindCell(row, 0);
      case ATestKind of
        1: MyWorksheet.CopyValue(cell, row, 2);
        2: MyWorksheet.CopyValue(cell, row, 1);
        3: MyWorksheet.CopyFormat(cell, row, 2);
        4: MyWorksheet.CopyFormat(cell, row, 1);
        5: MyWorksheet.CopyFormula(cell, row, 2);
        6: MyWorksheet.CopyFormula(cell, row, 1);
      end;
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

    if odd(ATestKind) then col := 2 else col := 1;

    for i:=0 to Length(SourceCells) do  // the "-1" is dropped to catch the down-shifted column!
    begin
      row := i;
      cell := MyWorksheet.FindCell(row, col);

      // (1) -- Compare values ---

      case ATestKind of
        1, 2:  // Copied values
          if (cell <> nil) and (row <= High(SourceCells)) then
          begin
            // Check formula results
            if MyHasFormula(@SourceCells[row]) then
              CheckEquals(
                SourceCells[0].NumberValue + 1,   // +1 because that's what the formula does...
                cell^.NumberValue,
                'Result of copied formula mismatch, cell ' + CellNotation(MyWorksheet, row, col)
              )
            else
            if (SourceCells[row].ContentType in [cctNumber, cctUTF8String, cctEmpty]) then
              CheckEquals(
                GetEnumName(TypeInfo(TCellContentType), Integer(SourceCells[row].ContentType)),
                GetEnumName(TypeInfo(TCellContentType), Integer(cell^.ContentType)),
                'Content type mismatch, cell '+CellNotation(MyWorksheet, row, col)
              );
            case SourceCells[row].ContentType of
              cctNumber:
                CheckEquals(
                  SourceCells[row].NumberValue,
                  cell^.NumberValue,
                  'Number value mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                );
              cctUTF8String:
                CheckEquals(
                  SourceCells[row].UTF8StringValue,
                  cell^.UTF8StringValue,
                  'String value mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                );
            end;
          end;

        3: // Copied formats to empty column -> there must not be any content
          if (cell <> nil) and (cell^.ContentType <> cctEmpty) then
            CheckEquals(
              true,     // true = "the cell has no content"
              (cell = nil) or (cell^.ContentType = cctEmpty),
              'No content mismatch, cell ' + CellNotation(MyWorksheet, row,col)
            );

        4: // Copied formats to occupied column --> data must be equal to source
           // cells, but offset by 1 cell
          if (row = 0) then
            CheckEquals(
              true,    // true = "the cell has no content"
              (cell = nil) or (cell^.ContentType = cctEmpty),
              'No content mismatch, cell ' + CellNotation(MyWorksheet, row, col)
            )
          else begin
            if (SourceCells[i+col-2].ContentType in [cctNumber, cctUTF8String, cctEmpty]) then
              CheckEquals(
                GetEnumName(TypeInfo(TCellContentType), Integer(SourceCells[i+col-2].ContentType)),
                GetEnumName(TypeInfo(TCellContentType), Integer(cell^.ContentType)),
                'Content type mismatch, cell '+CellNotation(MyWorksheet, row, col)
              );
            case SourceCells[i+col-2].ContentType of
              cctNumber:
                CheckEquals(
                  SourceCells[i+col-2].NumberValue,
                  cell^.NumberValue,
                  'Number value mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                );
              cctUTF8String:
                CheckEquals(
                  SourceCells[i+col-2].UTF8StringValue,
                  cell^.UTF8StringValue,
                  'String value mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                );
            end;
          end;
      end;

      // (2) -- Compare formatting ---

      case ATestKind of
        1, 5:
          CheckEquals(
           true,
           (cell = nil) or (MyWorksheet.ReadUsedFormatting(cell) = []),
           'Default formatting mismatch, cell ' + CellNotation(MyWorksheet, row, col)
          );
        2, 6:
          if (row = 0) then
            CheckEquals(
              true,
              (cell = nil) or (MyWorksheet.ReadUsedFormatting(cell) = []),
              'Default formatting mismatch, cell ' + CellNotation(MyWorksheet, row, col)
            )
          else
          begin
            CheckEquals(
              true,
              SourceCells[i+(col-2)].UsedFormattingFields = MyWorksheet.ReadUsedFormatting(cell),
              'Used formatting fields mismatch, cell ' + CellNotation(myWorksheet, row, col)
            );
            if (uffBackground in SourceCells[i+(col-2)].UsedFormattingFields) then
              CheckEquals(
                SourceCells[i+(col-2)].BackgroundColor,
                MyWorksheet.ReadBackgroundColor(cell),
                'Background color mismatch, cell ' + CellNotation(Myworksheet, row, col)
              );
          end;
        3, 4:
          if (i = Length(SourceCells)) and (ATestKind = 4) then
            CheckEquals(
              true,
              (cell = nil) or (MyWorksheet.ReadUsedFormatting(cell) = []),
              'Default formatting mismatch, cell ' + CellNotation(MyWorksheet, row, col)
            )
          else
          if (cell <> nil) then
          begin
            CheckEquals(
              true,
              SourceCells[i].UsedFormattingFields = MyWorksheet.ReadUsedFormatting(cell),
              'Used formatting fields mismatch, cell ' + CellNotation(MyWorksheet, row, col)
            );
            if (uffBackground in SourceCells[i].UsedFormattingFields) then
              CheckEquals(
                SourceCells[i].BackgroundColor,
                MyWorksheet.ReadBackgroundColor(cell),
                'Background color mismatch, cell ' + CellNotation(Myworksheet, row, col)
              );
          end;
      end;

      // (3) --- Check formula ---

      case ATestKind of
        1, 2, 3:
          CheckEquals(
            false,
            HasFormula(cell),
            'No formula mismatch, cell ' + CellNotation(MyWorksheet, row, col)
          );
        4:
          if (row = 0) then
            CheckEquals(
              false,
              (cell <> nil) and HasFormula(cell),
              'No formula mismatch, cell ' + CellNotation(Myworksheet, row, col)
            )
          else
            CheckEquals(
              SourceCells[i+col-2].FormulaValue,
              cell^.Formulavalue,
              'Formula mismatch, cell ' + CellNotation(MyWorksheet, row, col)
            );
        5:
          if cell <> nil then
          begin
            case SourceCells[i].FormulaValue of
              'A1+1' : expectedFormula := 'C1+1';
              'A$1+1': expectedFormula := 'C$1+1';
              else     expectedFormula := SourceCells[i].FormulaValue;
            end;
            CheckEquals(
              expectedFormula,
              cell^.FormulaValue,
              'Formula mismatch, cell ' + Cellnotation(Myworksheet, row, col)
            );
          end;
        6:
          begin
            if (row = 0) or (i >= Length(SourceCells)) then
              expectedFormula := ''
            else
            begin
              case SourceCells[i].FormulaValue of
                'A1+1' : expectedFormula := 'B1+1';
                'A$1+1': expectedFormula := 'B$1+1';
                '$A1+1': expectedFormula := '$A1+1';
                else     expectedFormula := SourceCells[i].FormulaValue;
              end;
              CheckEquals(
                expectedFormula,
                cell^.FormulaValue,
                'Formula mismatch, cell ' + Cellnotation(Myworksheet, row, col)
              );
            end;
          end;
      end;
    end; // For

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

{ Copy given cell values to occupied cells }
procedure TSpreadCopyTests.Test_CopyValuesToOccupiedCells;
begin
  Test_Copy(2);
end;

{ Copy given cell formats to empty cells }
procedure TSpreadCopyTests.Test_CopyFormatsToEmptyCells;
begin
  Test_Copy(3);
end;

{ Copy given cell formats to occupied cells }
procedure TSpreadCopyTests.Test_CopyFormatsToOccupiedCells;
begin
  Test_Copy(4);
end;

{ Copy given cell formulas to empty cells }
procedure TSpreadCopyTests.Test_CopyFormulasToEmptyCells;
begin
  Test_Copy(5);
end;

{ Copy given cell formulas to occupied cells }
procedure TSpreadCopyTests.Test_CopyFormulasToOccupiedCells;
begin
  Test_Copy(6);
end;


initialization
  RegisterTest(TSpreadCopyTests);
  InitCopyData;

end.

