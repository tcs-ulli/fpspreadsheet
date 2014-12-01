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
  SourceCells: Array[0..9] of TCell;

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
  end;

implementation

uses
  TypInfo, fpsutils;

const
  CopyTestSheet = 'Copy';

function InitNumber(ANumber: Double; ABkColor: TsColor): TCell;
begin
  InitCell(Result);
  Result.ContentType := cctNumber;
  Result.Numbervalue := ANumber;
  if (ABkColor <> scNotDefined) and (ABkColor <> scTransparent) then
  begin
    Result.UsedFormattingFields := Result.UsedFormattingFields + [uffBackgroundColor];
    Result.BackgroundColor := ABkColor;
  end;
end;

function InitString(AString: String; ABkColor: TsColor): TCell;
begin
  InitCell(Result);
  Result.ContentType := cctUTF8String;
  Result.UTF8StringValue := AString;
  if (ABkColor <> scNotDefined) and (ABkColor <> scTransparent) then
  begin
    Result.UsedFormattingFields := Result.UsedFormattingFields + [uffBackgroundColor];
    Result.BackgroundColor := ABkColor;
  end;
end;

function InitFormula(AFormula: String; ANumberResult: Double; ABkColor: TsColor): TCell;
begin
  InitCell(Result);
  Result.FormulaValue := AFormula;
  Result.NumberValue := ANumberResult;
  Result.ContentType := cctNumber;
  if (ABkColor <> scNotDefined) and (ABkColor <> scTransparent) then
  begin
    Result.UsedFormattingFields := Result.UsedFormattingFields + [uffBackgroundColor];
    Result.BackgroundColor := ABkColor;
  end;
end;

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
  InitCell(SourceCells[8]);  // empty but existing
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

    // Prepare the worksheet in which cells are copied:
    // Store the SourceCells to column A and B; in B shifted down by 1 cell
    {      A              B
      1   1.0
      2   2.0            1.0
      3   3.0 (yellow)   2.0
      4   Lazarus (red)  3.0
      5   A1+1           Lazarus
      6   $A1+1          A1+1
      7   A$1+1          $A1+1
      8   $A$1+1 (gray)  A$1+1
      9   (empty)        $A$1+1 (gray)
     10                 (empty)
    }
    for col := 0 to 1 do
      for row := 0 to High(SourceCells) do
      begin
        // Why is there a row index of "row + col" below? The first column has the
        // data starting at the top, in cell A1. In the second column each row
        // index is incremented by 1, i.e. the data are shifted down by 1 cell.
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
        if (uffBackgroundColor in SourceCells[row].UsedFormattingFields) then
          MyWorksheet.WriteBackgroundColor(cell, SourceCells[row].BackgroundColor);
      end;

    MyWorksheet.CalcFormulas;

    // Now perform the "copy" operations
    case ATestKind of
      1, 2:
        // copy the source cell VALUES to the empty column C (ATestKind = 1)
        // or occupied column B (ATestKind = 2)
        begin
          if ATestKind = 1 then col := 2 else col := 1;
          for row := 0 to High(SourceCells) do
            Myworksheet.CopyValue(MyWorksheet.FindCell(row, 0), row, col);
        end;
      3, 4:
        // copy the source cell FORMATS to the empty column C (ATestKind = 1)
        // or occupied column B (ATestKind = 2)
        begin
          if ATestKind = 1 then col := 2 else col := 1;
          for row := 0 to High(SourceCells) do
            MyWorksheet.CopyFormat(MyWorksheet.FindCell(row, 0), row, col);
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

    case ATestKind of
      1, 2:
      // Copied VALUES in first colum to empty third column (ATestKind = 1) or
      // occuopied second column (ATestKind = 2)
      // The formula cell should contain the result of A1+1 (only value copied)
        begin
          if ATestKind = 1 then col := 2 else col := 1;
          for row := 0 to Length(SourceCells) do
          begin
            cell := MyWorksheet.FindCell(row, col);

            if row < Length(SourceCells) then
            begin
              // Check content type
              if (SourceCells[row].ContentType in [cctNumber, cctUTF8String, cctEmpty]) then
                CheckEquals(
                  GetEnumName(TypeInfo(TCellContentType), Integer(SourceCells[row].ContentType)),
                  GetEnumName(TypeInfo(TCellContentType), Integer(cell^.ContentType)),
                  'Content type mismatch, cell '+CellNotation(MyWorksheet, row, col)
                );

              // Check values
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

              // Check formula results
              if HasFormula(@SourceCells[row]) then
                CheckEquals(
                  SourceCells[0].NumberValue + 1,
                  cell^.NumberValue,
                  'Result of copied formula mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                );
            end;

            // Check format: it should not be changed when copying only values
            case ATestKind of
              1:  // Copy to empty column --> no formatting
                  CheckEquals(
                    true,     // true = "the cell has default formatting"
                    (cell = nil) or (cell^.UsedFormattingFields = []),
                    'Default format mismatch, cell ' + CellNotation(MyWorksheet, row,col)
                  );
              2:  // Copy to occupied column --> format like source, but shifted down 1 cvell
                  if row = 0 then  // this cell should not be formatted
                    CheckEquals(
                      true,
                      cell^.UsedFormattingFields = [],
                      'Formatting fields mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                    )
                  else
                  begin
                    CheckEquals(
                      true,
                      SourceCells[row-1].UsedFormattingFields = cell^.UsedFormattingFields,
                      'Formatting fields mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                    );
                    if (uffBackgroundColor in cell^.UsedFormattingFields) then
                      CheckEquals(
                        SourceCells[row-1].BackgroundColor,
                        cell^.BackgroundColor,
                        'Background color mismatch, cell '+ CellNotation(MyWorksheet, row, col)
                      );
                  end;
              end;
          end;
        end;

      { ------------------------------------------------ }

      3: // FORMATs copied from first column to empty third column
        begin
          col := 2;
          for row :=0 to Length(SourceCells)-1 do
          begin
            cell := MyWorksheet.FindCell(row, col);

            // There should not be any content because the column was empty and
            // we had copied only formats
            CheckEquals(
              true,     // true = "the cell has no content"
              (cell = nil) or (cell^.ContentType = cctEmpty),
              'No content mismatch, cell ' + CellNotation(MyWorksheet, row,col)
            );

            // Check the format: it should be identical to that in column A
            if cell <> nil then
            begin
              CheckEquals(
                true,
                SourceCells[row].UsedFormattingFields = cell^.UsedFormattingFields,
                'Formatting fields mismatch, cell ' + CellNotation(MyWorksheet, row, col)
              );
              if (uffBackgroundColor in cell^.UsedFormattingFields) then
                CheckEquals(
                  SourceCells[row].BackgroundColor,
                  cell^.BackgroundColor,
                  'Background color mismatch, cell '+ CellNotation(MyWorksheet, row, col)
                );
            end;
          end;
        end;

      { ---------------------------- }

      4: // FORMATs copied from 1st to second column.
        begin
          col := 1;

          // Check values: they should be unchanged, i.e. identical to column A,
          // but there is a vertical offset by 1 cell
          cell := MyWorksheet.FindCell(0, col);
          CheckEquals(
            true,     // true = "the cell has no content"
            (cell = nil) or (cell^.ContentType = cctEmpty),
            'No content mismatch, cell ' + CellNotation(MyWorksheet, row,col)
          );
          for row := 1 to Length(SourceCells) do
          begin
            cell := MyWorksheet.FindCell(row, col);
            // Check content type
            if (SourceCells[row-1].ContentType in [cctNumber, cctUTF8String, cctEmpty]) then
              CheckEquals(
                GetEnumName(TypeInfo(TCellContentType), Integer(SourceCells[row-1].ContentType)),
                GetEnumName(TypeInfo(TCellContentType), Integer(cell^.ContentType)),
                'Content type mismatch, cell '+CellNotation(MyWorksheet, row, col)
              );
            // Check values
            case SourceCells[row-1].ContentType of
              cctNumber:
                CheckEquals(
                  SourceCells[row-1].NumberValue,
                  cell^.NumberValue,
                  'Number value mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                );
              cctUTF8String:
                CheckEquals(
                  SourceCells[row-1].UTF8StringValue,
                  cell^.UTF8StringValue,
                  'String value mismatch, cell ' + CellNotation(MyWorksheet, row, col)
                );
            end;
            // Check formula results
            if HasFormula(@SourceCells[row-1]) then
              CheckEquals(
                SourceCells[0].NumberValue + 1,
                cell^.NumberValue,
                'Result of copied formula mismatch, cell ' + CellNotation(MyWorksheet, row, col)
              );
          end;

          // Now check formatting - it should be equal to first column
          for row := 0 to Length(SourceCells)-1 do
          begin
            cell := MyWorksheet.FindCell(row, col);
            CheckEquals(
              true,
              SourceCells[row].UsedFormattingFields = cell^.UsedFormattingFields,
              'Formatting fields mismatch, cell ' + CellNotation(MyWorksheet, row, col)
            );

            if (uffBackgroundColor in cell^.UsedFormattingFields) then
              CheckEquals(
                SourceCells[row].BackgroundColor,
                cell^.BackgroundColor,
                'Background color mismatch, cell '+ CellNotation(MyWorksheet, row, col)
              );
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


initialization
  RegisterTest(TSpreadCopyTests);
  InitCopyData;

end.

