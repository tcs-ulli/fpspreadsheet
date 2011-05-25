{
excel8write.dpr

Demonstrates how to write an Excel 8+ file using the fpspreadsheet library

Adds formatting to the file

AUTHORS: Felipe Monteiro de Carvalho
}
program excel8write_format;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff8,
  laz_fpspreadsheet, fpsconvencoding;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  MyCell: PCell;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet1');

  // Write some cells
  MyWorksheet.WriteUTF8Text(1, 0, 'Border');// A2
  MyWorksheet.WriteUTF8Text(1, 1, '[]');    // B2
  MyWorksheet.WriteUTF8Text(1, 2, '[North]');// C2
  MyWorksheet.WriteUTF8Text(1, 3, '[West]');// D2
  MyWorksheet.WriteUTF8Text(1, 4, '[East]');// E2
  MyWorksheet.WriteUTF8Text(1, 5, '[South]');// F2

  // Format them

  MyCell := MyWorksheet.GetCell(1, 1);
  MyCell^.Border := [];
  MyCell^.UsedFormattingFields := [uffBorder];

  MyCell := MyWorksheet.GetCell(1, 2);
  MyCell^.Border := [cbNorth];
  MyCell^.UsedFormattingFields := [uffBorder];

  MyCell := MyWorksheet.GetCell(1, 3);
  MyCell^.Border := [cbWest];
  MyCell^.UsedFormattingFields := [uffBorder];

  MyCell := MyWorksheet.GetCell(1, 4);
  MyCell^.Border := [cbEast];
  MyCell^.UsedFormattingFields := [uffBorder];

  MyCell := MyWorksheet.GetCell(1, 5);
  MyCell^.Border := [cbSouth];
  MyCell^.UsedFormattingFields := [uffBorder];

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test3.xls', sfExcel8, False);
  MyWorkbook.Free;
end.

