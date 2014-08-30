{ This demo is a test for recursive calculation of cells. The cell formulas
  are constructed such that the first cell depends on the second, and the second
  cell depends on the third one. Only the third cell contains a number.
  Therefore calculation has to be done recursively until the independent third
  cell is found. }

program demo_recursive_calc;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  {$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}
  {$ENDIF}
  Classes,
  math, fpspreadsheet, fpsfunc, xlsbiff8;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
const
  OutputFile='test_recursive.xls';

begin
  writeln('Starting program.');
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boCalcBeforeSaving];

    worksheet := workbook.AddWorksheet('Calc_test');
    worksheet.WriteColWidth(0, 20);

    // A1
    worksheet.WriteUTF8Text(0, 0, '=B2+1');
    // B1
    worksheet.WriteFormula(0, 1, 'B2+1');
    {
    worksheet.WriteRPNFormula(0, 1, CreateRPNFormula(
      RPNCellValue('B2',
      RPNInteger(1,
      RPNFunc(fekAdd, nil)))));
     }
    // A2
    worksheet.WriteUTF8Text(1, 0, '=B3+1');
    // B2
    worksheet.WriteFormula(1, 1, 'B3+1');
    {
    worksheet.WriteRPNFormula(1, 1, CreateRPNFormula(
      RPNCellValue('B3',
      RPNInteger(1,
      RPNFunc(fekAdd, nil)))));
     }
    // A3
    worksheet.WriteUTF8Text(2, 0, '(not dependent)');
    // B3
    worksheet.WriteNumber(2, 1, 1);

    workbook.WriteToFile(OutputFile, sfExcel8, true);
    writeln('Finished.');
    writeln;
    writeln('Please open "'+OutputFile+'" in "fpsgrid".');
    writeLn('It should show calculation results in cells B1 and B2.');
  finally
    workbook.Free;
  end;
end.

