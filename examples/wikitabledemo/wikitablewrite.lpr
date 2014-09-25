{
wikitablewrite.lpr

Demonstrates how to write a wikitable file using the fpspreadsheet library
Note: the output written by wikitablewrite cannot yet be read by the
wikitableread demo.
}
program wikitablewrite;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, wikitable;

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

  MyWorksheet.WriteUTF8Text(1, 0, 'This is bold text:');
  Myworksheet.WriteUTF8Text(1, 1, 'Hello world!');
  Myworksheet.WriteFontStyle(1, 1, [fssBold]);

  MyWorksheet.WriteUTF8Text(2, 0, 'This is a number:');
  MyWorksheet.WriteNumber(2, 1, 3.141592);
  MyWorksheet.WriteBackgroundColor(2, 1, scMagenta);
  Myworksheet.WriteHorAlignment(2, 1, haRight);

  MyWorksheet.WriteUTF8Text(3, 0, 'This is a date:');
  Myworksheet.WriteDateTime(3, 1, date());

  MyWorksheet.WriteUTF8Text(4, 0, 'This is a long text:');
  MyWorksheet.WriteUTF8Text(4, 1, 'A very, very, very, very long text, indeed');

  MyWorksheet.WriteUTF8Text(5, 0, 'This is long text with line break:');
  Myworksheet.WriteVertAlignment(5, 0, vaTop);

  MyWorksheet.WriteUTF8Text(5, 1, 'A very, very, very, very long text,<br /> indeed');

  MyWorksheet.WriteUTF8Text(6, 0, 'Merged rows');
  Myworksheet.MergeCells(6, 0, 7, 0);
  MyWorksheet.WriteUTF8Text(6, 1, 'A');
  MyWorksheet.WriteUTF8Text(7, 1, 'B');

  MyWorksheet.WriteUTF8Text(8, 0, 'Merged columns');
  MyWorksheet.WriteHorAlignment(8, 0, haCenter);
  MyWorksheet.MergeCells(8, 0, 8, 1);

  MyWorksheet.WriteUTF8Text(10, 0, 'Right borders:');
  MyWorksheet.WriteBorders(10, 0, [cbEast]);

  MyWorksheet.WriteUTF8Text(10, 1, 'medium / blue');
  MyWorksheet.WriteBorders(10, 1, [cbEast]);
  MyWorksheet.WriteBorderLineStyle(10, 1, cbEast, lsMedium);
  MyWorksheet.WriteBorderColor(10, 1, cbEast, scBlue);

  MyWorksheet.WriteUTF8Text(11, 0, 'Top borders:');
  MyWorksheet.WriteBorders(11, 0, [cbNorth]);
  MyWorksheet.WriteBorderLineStyle(11, 0, cbNorth, lsDashed);

  MyWorksheet.WriteUTF8Text(11, 1, '(dotted)');
  MyWorksheet.WriteBorders(11, 1, [cbNorth]);
  MyWorksheet.WriteBorderLineStyle(11, 1, cbNorth, lsDotted);

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.wikitable_wikimedia', sfWikitable_wikimedia);
  MyWorkbook.Free;
end.

