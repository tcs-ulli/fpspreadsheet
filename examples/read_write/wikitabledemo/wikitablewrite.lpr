{
wikitablewrite.lpr

Demonstrates how to write a wikitable file using the fpspreadsheet library
Note: the output written by wikitablewrite cannot yet be read by the
wikitableread demo.
}
program wikitablewrite;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, wikitable;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  row: Integer;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('Sheet');

  // Use first row and column as headers
  Myworksheet.Options := MyWorksheet.Options + [soHasFrozenPanes];
  Myworksheet.TopPaneHeight := 1;
  Myworksheet.LeftPaneWidth := 1;

  // Write colwidth
  Myworksheet.WriteColWidth(1, 25);  // 25 characters

  // Write some cells
  row := 0;

  MyWorksheet.WriteBlank(row, 0);
  MyWorksheet.WriteUTF8Text(row, 1, 'Description');
  MyWorksheet.WriteUTF8Text(row, 2, 'Example');
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'This is a text:');
  MyWorksheet.WriteUTF8Text(row, 2, 'Hello world!');
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'This is bold text:');
  Myworksheet.WriteUTF8Text(row, 2, 'Hello world!');
  Myworksheet.WriteFontStyle(row, 2, [fssBold]);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'This is a number:');
  MyWorksheet.WriteNumber(row, 2, 3.141592);
  Myworksheet.WriteHorAlignment(row, 2, haRight);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'This is a date:');
  Myworksheet.WriteDateTime(row, 2, date());
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'This is a long text:');
  MyWorksheet.WriteUTF8Text(row, 2, 'A very, very, very, very long text, indeed');
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Centered text:');
  MyWorksheet.WriteUTF8Text(row, 2, 'I am in the center.');
  MyWorksheet.WriteHorAlignment(row, 2, haCenter);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'This is a long text with line break:');
  Myworksheet.WriteVertAlignment(row, 1, vaTop);
  MyWorksheet.WriteUTF8Text(row, 2, 'A very, very, very, very long text,<br /> indeed');
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Merged rows');
  Myworksheet.MergeCells(row, 1, row+1, 1);
  MyWorksheet.WriteUTF8Text(row, 2, 'A');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 2, 'B');
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Merged columns');
  MyWorksheet.WriteHorAlignment(row, 1, haCenter);
  MyWorksheet.MergeCells(row, 1, row, 2);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Right border:');
  MyWorksheet.WriteUTF8Text(row, 2, 'medium / red');
  MyWorksheet.WriteBorders(row, 2, [cbEast]);
  MyWorksheet.WriteBorderStyle(row, 2, cbEast, lsMedium, scRed);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Top border:');
  MyWorksheet.WriteUTF8Text(row, 2, 'top / dashed');
  MyWorksheet.WriteBorders(row, 2, [cbNorth]);
  MyWorksheet.WriteBorderLineStyle(row, 2, cbNorth, lsDashed);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Left border:');
  MyWorksheet.WriteUTF8Text(row, 2, 'left / dotted');
  MyWorksheet.WriteBorders(row, 2, [cbWest]);
  MyWorksheet.WriteBorderLineStyle(row, 2, cbWest, lsDotted);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Bottom border:');
  MyWorksheet.WriteUTF8Text(row, 2, 'bottom / double');
  MyWorksheet.WriteBorders(row, 2, [cbSouth]);
  MyWorksheet.WriteBorderLineStyle(row, 2, cbSouth, lsDouble);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'This row is high');
  MyWorksheet.WriteRowHeight(row, 5);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Colors:');
  MyWorksheet.WriteUTF8Text(row, 2, 'yellow on blue');
  MyWorksheet.WriteFontColor(row, 2, scYellow);
  MyWorksheet.WriteBackgroundColor(row, 2, scBlue);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'RGB background color:');
  MyWorksheet.WriteUTF8Text(row, 2, 'color #FF77C3');  // HTML colors are big-endian
  MyWorksheet.WriteBackgroundColor(row, 2, $C377FF);   // fps colors are little-endian
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Bold text:');
  MyWorksheet.WriteUTF8Text(row, 2, 'Bold text');
  MyWorksheet.WriteFontStyle(row, 2, [fssBold]);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Italic text:');
  MyWorksheet.WriteUTF8Text(row, 2, 'Italic text');
  MyWorksheet.WriteFontStyle(row, 2, [fssItalic]);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Underlined text:');
  MyWorksheet.WriteUTF8Text(row, 2, 'Underlined text');
  MyWorksheet.WriteFontStyle(row, 2, [fssUnderline]);
  inc(row);

  MyWorksheet.WriteNumber(row, 0, row);
  MyWorksheet.WriteUTF8Text(row, 1, 'Strike-through text:');
  MyWorksheet.WriteUTF8Text(row, 2, 'Strike-through text');
  MyWorksheet.WriteFontStyle(row, 2, [fssStrikeout]);
  inc(row);

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.wikitable_wikimedia', sfWikitable_wikimedia);
  MyWorkbook.Free;
end.

