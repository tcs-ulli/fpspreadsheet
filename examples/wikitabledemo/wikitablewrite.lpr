{
wikitablewrite.lpr

Demonstrates how to write a wikitable file using the fpspreadsheet library
Note: the output written by wikitablewrite cannot yet be read by the
wikitableread demo.
}
program wikitablewrite;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, wikitable,
  laz_fpspreadsheet;

const
  Str_First = 'First';
  Str_Second = 'Second';
  Str_Third = 'Third';
  Str_Fourth = 'Fourth';
  Str_Worksheet1 = 'Meu Relatório';
  Str_Worksheet2 = 'My Worksheet 2';
  Str_Total = 'Total:';
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyRPNFormula: TsRPNFormula;
  MyDir: string;
  number: Double;
  lCell: PCell;
  lCol: TCol;
  i: Integer;
  r: Integer = 10;
  s: String;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;

  MyWorksheet := MyWorkbook.AddWorksheet(Str_Worksheet1);

  // Write some cells
  MyWorksheet.WriteUTF8Text(0, 0, 'This is a text:');
  MyWorksheet.WriteUTF8Text(0, 1, 'Hello world!');
  MyWorksheet.WriteUTF8Text(1, 0, 'This is a number:');
  MyWorksheet.WriteNumber(1, 1, 3.141592);
  MyWorksheet.WriteUTF8Text(2, 0, 'This is a date:');
  Myworksheet.WriteDateTime(2, 1, date());

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.wikitable_wikimedia', sfWikitable_wikimedia);
  MyWorkbook.Free;
end.

