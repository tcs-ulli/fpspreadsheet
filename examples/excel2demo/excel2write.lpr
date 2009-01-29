{
excel2write.dpr

Demonstrates how to write an Excel 2.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel2write;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff2, laz_fpspreadsheet;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
begin
  // Open the output file
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  // Write some number cells
  MyWorksheet.WriteNumber(1, 1, 1.0);
  MyWorksheet.WriteNumber(1, 2, 2.0);
  MyWorksheet.WriteNumber(1, 3, 3.0);
  MyWorksheet.WriteNumber(1, 4, 4.0);

  // Write some string cells
  MyWorksheet.WriteUTF8Text(2, 1, 'First');
  MyWorksheet.WriteUTF8Text(2, 2, 'Second');
  MyWorksheet.WriteUTF8Text(2, 3, 'Third');
  MyWorksheet.WriteUTF8Text(2, 4, 'Fourth');

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test' + STR_EXCEL_EXTENSION, sfExcel2);
  MyWorkbook.Free;
end.

