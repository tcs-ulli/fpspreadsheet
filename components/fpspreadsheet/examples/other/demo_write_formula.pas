{
test_write_formula.pas

Demonstrates how to write a formula using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program demo_write_formula;

{$mode delphi}{$H+}

uses
  Classes, SysUtils,
  fpspreadsheet, xlsbiff5, xlsbiff8, fpsopendocument;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  MyCell: PCell;

procedure WriteFirstWorksheet();
var
  MyFormula: TsFormula;
  MyRPNFormula: TsRPNFormula;
begin
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet1');

  // Write some cells
  MyWorksheet.WriteUTF8Text(0, 1, 'Text Formula');// B1
  MyWorksheet.WriteUTF8Text(0, 2, 'RPN');// C1

  MyWorksheet.WriteNumber(0, 4, -3.14);  // E1
  MyWorksheet.WriteNumber(1, 4, 100);    // E2
  MyWorksheet.WriteNumber(2, 4, 200);    // E3
  Myworksheet.WriteNumber(3, 4, 300);    // E4
  MyWorksheet.WriteNumber(4, 4, 250);    // E5

  // =Sum(E2:e5)
  MyWorksheet.WriteUTF8Text(1, 0, '=Sum(E2:e5)'); // A2
  //
  MyFormula.FormulaStr := '=Sum(E2:e5)';
  MyFormula.DoubleValue := 0.0;
  MyWorksheet.WriteFormula(1, 1, MyFormula);    // B2
  //
  MyWorksheet.WriteRPNFormula(1, 2, CreateRPNFormula(  // C2
    RPNCellRange('E2:E5',
    RPNFunc(fekSum, 1, nil))));

  // Write the formula =ABS(E1)
  MyWorksheet.WriteUTF8Text(2, 0, '=ABS(E1)'); // A3
  //
  MyWorksheet.WriteRPNFormula(2, 2, CreateRPNFormula(  // C3
    RPNCellValue('E1',
    RPNFunc(fekAbs, nil))));

  // Write the formula =4+5
  MyWorksheet.WriteUTF8Text(3, 0, '=4+5'); // A4
  //
  MyWorksheet.WriteRPNFormula(3, 2, CreateRPNFormula(  //C4
    RPNNumber(4.0,
    RPNNumber(5.0,
    RPNFunc(fekAdd, nil)))));
end;

procedure WriteSecondWorksheet();
begin
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet2');

  // Write some cells

  // Line 1

  MyWorksheet.WriteUTF8Text(1, 1, 'Relatório');
  MyCell := MyWorksheet.GetCell(1, 1);
  MyWorksheet.WriteBorders(MyCell, [cbNorth, cbWest, cbSouth]);
  Myworksheet.WriteBackgroundColor(MyCell, scGray20pct);
end;

const
  TestFile='test_formula.xls';

{$R *.res}

begin
  writeln('Starting program.');
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;

  WriteFirstWorksheet();

  WriteSecondWorksheet();

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + TestFile, sfExcel8, True);
//  MyWorkbook.WriteToFile(MyDir + 'test_formula.odt', sfOpenDocument, False);
  MyWorkbook.Free;
  writeln('Finished. Please open "'+Testfile+'" in your spreadsheet program.');
end.

