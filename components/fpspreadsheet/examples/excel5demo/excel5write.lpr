{
excel5write.dpr

Demonstrates how to write an Excel 5.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel5write;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff5,
  laz_fpspreadsheet;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyFormula: TsFormula;
  MyDir: string;
  i: Integer;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  // Write some cells
  MyWorksheet.WriteNumber(0, 0, 1.0);// A1
  MyWorksheet.WriteNumber(0, 1, 2.0);// B1
  MyWorksheet.WriteNumber(0, 2, 3.0);// C1
  MyWorksheet.WriteNumber(0, 3, 4.0);// D1
  MyWorksheet.WriteUTF8Text(4, 2, 'Total:');// C5
  MyWorksheet.WriteNumber(4, 3, 10);        // D5

{ Uncommend this to test large XLS files
  for i := 2 to 20 do
  begin
    MyWorksheet.WriteAnsiText(i, 0, ParamStr(0));
    MyWorksheet.WriteAnsiText(i, 1, ParamStr(0));
    MyWorksheet.WriteAnsiText(i, 2, ParamStr(0));
    MyWorksheet.WriteAnsiText(i, 3, ParamStr(0));
  end;
}

  // Write the formula E1 = A1 + B1
//  MyFormula.FormulaStr := '';
//  MyWorksheet.WriteFormula(0, 4, MyFormula);

  // Creates a new worksheet
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet 2');

{  // Write some string cells
  MyWorksheet.WriteUTF8Text(0, 0, 'First');
  MyWorksheet.WriteUTF8Text(0, 1, 'Second');
  MyWorksheet.WriteUTF8Text(0, 2, 'Third');
  MyWorksheet.WriteUTF8Text(0, 3, 'Fourth');}

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.xls', sfExcel5);
  MyWorkbook.Free;
end.

