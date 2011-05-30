{
test_write_formula.pas

Demonstrates how to write an formula using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program test_write_formula;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff8, fpsopendocument,
  laz_fpspreadsheet, fpsconvencoding;

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
  MyWorksheet.WriteUTF8Text(1, 0, 'Text Formulas');// A2

  MyWorksheet.WriteUTF8Text(1, 1, '=Sum(D2:d5) Text Formula');    // B2

  MyFormula.FormulaStr := '=Sum(D2:d5)';
  MyFormula.DoubleValue := 0.0;
  MyWorksheet.WriteFormula(1, 2, MyFormula);    // C2

  MyWorksheet.WriteUTF8Text(1, 1, '=Sum(D2:d5) RPN');    // B3

  MyFormula.FormulaStr := '=Sum(D2:d5)';
  MyFormula.DoubleValue := 0.0;
  MyWorksheet.WriteFormula(1, 2, MyFormula);    // C3

  SetLength(MyRPNFormula, 2);
  MyRPNFormula[0].ElementKind := fekOpSUM;
  MyRPNFormula[1].ElementKind := fekCellRange;
  MyRPNFormula[1].Row := 1;
  MyRPNFormula[1].Row := 4;
  MyRPNFormula[1].Col := 3;
  MyRPNFormula[1].Col := 3;
  MyWorksheet.WriteRPNFormula(1, 2, MyRPNFormula);    // C2
end;

procedure WriteSecondWorksheet();
begin
{  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet2');

  // Write some cells

  // Line 1

  MyWorksheet.WriteUTF8Text(1, 1, 'Relatório');
  MyCell := MyWorksheet.GetCell(1, 1);
  MyCell^.Border := [cbNorth, cbWest, cbSouth];
  MyCell^.BackgroundColor := scGrey20pct;
  MyCell^.UsedFormattingFields := [uffBorder, uffBackgroundColor, uffBold];}
end;

begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;

  WriteFirstWorksheet();

  WriteSecondWorksheet();

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test_formula.xls', sfExcel8, False);
//  MyWorkbook.WriteToFile(MyDir + 'test_formula.odt', sfOpenDocument, False);
  MyWorkbook.Free;
end.

