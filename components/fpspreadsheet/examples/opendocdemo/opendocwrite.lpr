{
opendocwrite.dpr

Demonstrates how to write an OpenDocument file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program opendocwrite;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, fpsallformats,
  laz_fpspreadsheet;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  // Write some number cells
  MyWorksheet.WriteNumber(0, 0, 1.0);
  MyWorksheet.WriteNumber(0, 1, 2.0);
  MyWorksheet.WriteNumber(0, 2, 3.0);
  MyWorksheet.WriteNumber(0, 3, 4.0);

  // Write some string cells
  MyWorksheet.WriteUTF8Text(4, 2, 'Total:');
  MyWorksheet.WriteNumber(4, 3, 10.0);

  // Creates a new worksheet
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet 2');

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.ods',
    sfOpenDocument);
  MyWorkbook.Free;
end.

