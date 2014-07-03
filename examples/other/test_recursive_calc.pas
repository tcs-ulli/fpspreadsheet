{ This demo is a test for recursive calculation of cells. The cell formulas
  are constructed such that the first cell depends on the second, and the second
  cell depends on the third one. Only the third cell contains a number.
  Therefore calculation has to be done recursively until the independent third
  cell is found. }

program test_recursive_calc;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Classes, laz_fpspreadsheet
  { you can add units after this },
  math, fpspreadsheet, fpsfunc;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;

begin
  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Calc_test');
    worksheet.Options := worksheet.Options + [soCalcBeforeSaving];
    worksheet.WriteColWidth(0, 20);

    // A1
    worksheet.WriteUTF8Text(0, 0, '=B2+1');
    // B1
    worksheet.WriteRPNFormula(0, 1, CreateRPNFormula(
      RPNCellValue('B2',
      RPNNumber(1,
      RPNFunc(fekAdd, nil)))));

    // A2
    worksheet.WriteUTF8Text(1, 0, '=B3+1');
    // B2
    worksheet.WriteRPNFormula(1, 1, CreateRPNFormula(
      RPNCellValue('B3',
      RPNNumber(1,
      RPNFunc(fekAdd, nil)))));

    // A3
    worksheet.WriteUTF8Text(2, 0, '(not dependent)');
    // B3
    worksheet.WriteNumber(2, 1, 1);

    workbook.WriteToFile('test_calc.xls', sfExcel8, true);
  finally
    workbook.Free;
  end;
end.

